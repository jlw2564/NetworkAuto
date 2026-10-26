"""FastAPI Web 层：任务提交、实时进度（SSE）、结果预览、Excel 下载、历史记录。

启动：
    python -m net_port_tool.webapp
    # 或
    uvicorn net_port_tool.webapp:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import asyncio
import json
import logging
import queue
import shutil
import threading
import time
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from fastapi import FastAPI, HTTPException, Request, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import config
from .report import ALL_SECTIONS
from .service import RunOptions, run as run_service

LOG = logging.getLogger('net_port_tool.web')

STATIC_DIR = Path(__file__).resolve().parent / 'static'
RUNS_DIR = config.RUNS_DIR
UPLOAD_DIR = RUNS_DIR / 'uploads'

SECTION_LABELS = {
    'name': '命名验证',
    'links': '交换机连线',
    'terminal': '终端定位/接入表',
    'usage': '端口使用表',
    'suspicious': '可疑端口清单',
}
PREVIEW_ROWS = 200

app = FastAPI(title='交换机端口资源管理工具', version='2.0')
app.mount('/static', StaticFiles(directory=str(STATIC_DIR)), name='static')


# --------------------------------------------------------------------------
# 任务模型
# --------------------------------------------------------------------------
class RunRequest(BaseModel):
    switch_file: Optional[str] = None
    terminal_file: Optional[str] = None
    sections: Optional[List[str]] = None
    concurrency: int = 6
    discover: bool = True


class Job:
    def __init__(self, options: RunOptions, raw: dict):
        self.id = uuid.uuid4().hex[:12]
        self.created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.options = options
        self.raw = raw
        self.status = 'running'
        self.stage = '排队中'
        self.current = 0
        self.total = 0
        self.detail = ''
        self.logs: List[dict] = []
        self.result: Optional[dict] = None
        self.error = ''
        self.output_file = ''
        self.run_dir = RUNS_DIR / self.id
        self._lock = threading.Lock()
        self._subscribers: List['queue.Queue'] = []

    # ---- 事件 ----
    def emit(self, kind: str, payload: dict) -> None:
        event = {'kind': kind, 'ts': time.time(), **payload}
        with self._lock:
            if kind == 'log':
                self.logs.append(event)
                if len(self.logs) > 4000:
                    self.logs = self.logs[-2000:]
            for q in list(self._subscribers):
                try:
                    q.put_nowait(event)
                except Exception:  # noqa: BLE001
                    pass

    def subscribe(self) -> 'queue.Queue':
        q: 'queue.Queue' = queue.Queue()
        with self._lock:
            self._subscribers.append(q)
        return q

    def unsubscribe(self, q: 'queue.Queue') -> None:
        with self._lock:
            if q in self._subscribers:
                self._subscribers.remove(q)

    def snapshot(self) -> dict:
        return {
            'job_id': self.id, 'created_at': self.created_at, 'status': self.status,
            'stage': self.stage, 'current': self.current, 'total': self.total,
            'detail': self.detail, 'error': self.error,
            'output_file': self.output_file, 'result': self.result,
            'log_count': len(self.logs),
        }

    def save_meta(self) -> None:
        try:
            self.run_dir.mkdir(parents=True, exist_ok=True)
            meta = self.snapshot()
            meta['options'] = {
                'switch_file': str(self.options.switch_file or ''),
                'terminal_file': str(self.options.terminal_file or ''),
                'concurrency': self.options.concurrency,
                'discover': self.options.discover,
                'sections': list(self.options.sections),
            }
            (self.run_dir / 'meta.json').write_text(
                json.dumps(meta, ensure_ascii=False, indent=2), encoding='utf-8')
        except Exception as exc:  # noqa: BLE001
            LOG.warning('写入任务元数据失败: %s', exc)


JOBS: Dict[str, Job] = {}
JOBS_LOCK = threading.Lock()
MAX_JOBS = 100


def _register(job: Job) -> None:
    with JOBS_LOCK:
        JOBS[job.id] = job
        if len(JOBS) > MAX_JOBS:
            old = sorted(JOBS.values(), key=lambda j: j.created_at)[:len(JOBS) - MAX_JOBS]
            for j in old:
                if j.status != 'running':
                    JOBS.pop(j.id, None)


def _resolve_input(value: Optional[str], default: Path, label: str) -> Optional[Path]:
    """把前端传来的路径解析成存在的文件；空值表示使用默认。"""
    if not value:
        return default if default.exists() else None
    p = Path(value)
    if not p.is_absolute():
        p = config.BASE_DIR / value
    if not p.exists():
        raise HTTPException(status_code=400, detail=f'{label} 不存在: {p}')
    return p


# --------------------------------------------------------------------------
# 路由：页面与静态资源
# --------------------------------------------------------------------------
@app.get('/', response_class=HTMLResponse)
def index() -> HTMLResponse:
    html = (STATIC_DIR / 'index.html').read_text(encoding='utf-8')
    return HTMLResponse(html)


@app.get('/api/config')
def get_config() -> dict:
    return {
        'base_dir': str(config.BASE_DIR),
        'default_switch_file': str(config.SWITCH_TABLE_FILE),
        'default_switch_exists': config.SWITCH_TABLE_FILE.exists(),
        'default_terminal_file': str(config.TERMINAL_IP_FILE),
        'default_terminal_exists': config.TERMINAL_IP_FILE.exists(),
        'default_output_file': str(config.OUTPUT_FILE),
        'sections': [{'key': k, 'label': SECTION_LABELS[k]} for k in ALL_SECTIONS],
        'runs_dir': str(RUNS_DIR),
    }


@app.post('/api/upload')
async def upload(file: UploadFile = File(...)) -> dict:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = Path(file.filename or 'upload.xlsx').name
    target = UPLOAD_DIR / f'{uuid.uuid4().hex[:8]}_{safe_name}'
    with target.open('wb') as fh:
        shutil.copyfileobj(file.file, fh)
    return {'path': str(target), 'name': safe_name}


# --------------------------------------------------------------------------
# 路由：任务
# --------------------------------------------------------------------------
@app.post('/api/run')
def start_run(req: RunRequest) -> dict:
    with JOBS_LOCK:
        if any(j.status == 'running' for j in JOBS.values()):
            raise HTTPException(status_code=409, detail='已有任务在运行，请等待完成后再提交')

    switch_file = _resolve_input(req.switch_file, config.SWITCH_TABLE_FILE, '交换机清单')
    if switch_file is None:
        raise HTTPException(status_code=400, detail=f'默认交换机清单不存在: {config.SWITCH_TABLE_FILE}')

    sections = req.sections or list(ALL_SECTIONS)
    sections = [s for s in sections if s in ALL_SECTIONS] or list(ALL_SECTIONS)
    terminal_file = None
    if 'terminal' in sections:
        terminal_file = _resolve_input(req.terminal_file, config.TERMINAL_IP_FILE, '终端IP表')
        if terminal_file is None:
            sections = [s for s in sections if s != 'terminal']

    options = RunOptions(
        switch_file=switch_file,
        terminal_file=terminal_file,
        output_file=None,
        concurrency=max(1, min(int(req.concurrency or 6), 32)),
        discover=bool(req.discover),
        sections=tuple(sections),
    )
    job = Job(options, req.model_dump() if hasattr(req, 'model_dump') else req.dict())
    _register(job)
    threading.Thread(target=_execute_job, args=(job,), daemon=True).start()
    return job.snapshot()


def _execute_job(job: Job) -> None:
    job.run_dir.mkdir(parents=True, exist_ok=True)
    output_file = job.run_dir / '交换机端口资源管理表.xlsx'
    job.options.output_file = output_file
    job.output_file = str(output_file)

    def on_log(level: str, message: str) -> None:
        job.emit('log', {'level': level, 'message': message})

    def on_progress(stage: str, current: int, total: int, detail: str) -> None:
        job.stage = stage
        job.current = current
        job.total = total
        job.detail = detail
        job.emit('progress', {'stage': stage, 'current': current, 'total': total, 'detail': detail})

    job.emit('log', {'level': 'INFO', 'message': f'任务 {job.id} 开始'})
    try:
        result = run_service(job.options, on_log=on_log, on_progress=on_progress)
        job.status = 'done' if result.ok else 'failed'
        job.error = result.error
        job.result = {
            'stats': result.stats, 'duration': result.duration, 'core_name': result.core_name,
            'devices': result.devices, 'unknown_switches': result.unknown_switches,
            'output_file': result.output_file,
        }
        job.emit('done', job.result)
    except Exception as exc:  # noqa: BLE001
        job.status = 'failed'
        job.error = str(exc)
        job.emit('log', {'level': 'ERROR', 'message': '任务异常:\n' + traceback.format_exc()})
        job.emit('done', {'ok': False, 'error': str(exc)})
    finally:
        job.save_meta()


@app.get('/api/jobs')
def list_jobs() -> dict:
    with JOBS_LOCK:
        jobs = sorted(JOBS.values(), key=lambda j: j.created_at, reverse=True)
    return {'jobs': [j.snapshot() for j in jobs]}


@app.get('/api/jobs/{job_id}')
def get_job(job_id: str) -> dict:
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail='任务不存在')
    snap = job.snapshot()
    snap['logs'] = job.logs[-500:]
    return snap


@app.get('/api/jobs/{job_id}/events')
async def job_events(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail='任务不存在')
    q = job.subscribe()

    async def event_stream():
        # 先补发历史日志，保证刷新页面也能看到已产生的内容
        for ev in list(job.logs[-500:]):
            yield f'data: {json.dumps(ev, ensure_ascii=False)}\n\n'
        yield f'data: {json.dumps({"kind": "progress", "stage": job.stage, "current": job.current, "total": job.total, "detail": job.detail}, ensure_ascii=False)}\n\n'
        try:
            while True:
                try:
                    ev = q.get_nowait()
                except queue.Empty:
                    if job.status != 'running':
                        yield f'data: {json.dumps({"kind": "done", **(job.result or {}), "status": job.status, "error": job.error}, ensure_ascii=False)}\n\n'
                        break
                    await asyncio.sleep(0.2)
                    continue
                yield f'data: {json.dumps(ev, ensure_ascii=False)}\n\n'
                if ev.get('kind') == 'done':
                    break
        finally:
            job.unsubscribe(q)

    return StreamingResponse(event_stream(), media_type='text/event-stream',
                             headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.get('/api/jobs/{job_id}/download')
def download(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail='任务不存在')
    path = Path(job.output_file) if job.output_file else None
    if not path or not path.exists():
        raise HTTPException(status_code=404, detail='结果文件不存在')
    return FileResponse(str(path), filename=f'交换机端口资源管理表_{job_id}.xlsx',
                        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.get('/api/jobs/{job_id}/sheets')
def list_sheets(job_id: str) -> dict:
    path = _job_output(job_id)
    xls = pd.ExcelFile(path)
    return {'sheets': xls.sheet_names}


@app.get('/api/jobs/{job_id}/sheets/{sheet_name}')
def preview_sheet(job_id: str, sheet_name: str, limit: int = PREVIEW_ROWS) -> dict:
    path = _job_output(job_id)
    xls = pd.ExcelFile(path)
    if sheet_name not in xls.sheet_names:
        raise HTTPException(status_code=404, detail=f'Sheet 不存在: {sheet_name}')
    df = pd.read_excel(xls, sheet_name=sheet_name).fillna('')
    total = int(len(df))
    head = df.head(max(1, min(limit, 1000)))
    return {
        'sheet': sheet_name,
        'total': total,
        'columns': [str(c) for c in head.columns.tolist()],
        'rows': [[_cell(v) for v in row] for row in head.itertuples(index=False, name=None)],
    }


def _job_output(job_id: str) -> Path:
    job = JOBS.get(job_id)
    path = Path(job.output_file) if job and job.output_file else (RUNS_DIR / job_id / '交换机端口资源管理表.xlsx')
    if not path.exists():
        raise HTTPException(status_code=404, detail='结果文件不存在')
    return path


def _cell(value):
    if value is None:
        return ''
    if isinstance(value, float) and value != value:  # NaN
        return ''
    return value if isinstance(value, (int, float, str, bool)) else str(value)


# --------------------------------------------------------------------------
# 历史记录（跨进程重启：从 RUNS_DIR 读取 meta.json）
# --------------------------------------------------------------------------
@app.get('/api/history')
def history() -> dict:
    items = []
    if RUNS_DIR.exists():
        for d in sorted(RUNS_DIR.iterdir(), reverse=True):
            meta_file = d / 'meta.json'
            if not d.is_dir() or not meta_file.exists():
                continue
            try:
                meta = json.loads(meta_file.read_text(encoding='utf-8'))
            except Exception:  # noqa: BLE001
                continue
            meta['has_output'] = (d / '交换机端口资源管理表.xlsx').exists()
            meta['job_id'] = d.name
            items.append(meta)
    return {'history': items[:50]}


@app.get('/api/history/{job_id}/download')
def history_download(job_id: str) -> FileResponse:
    path = RUNS_DIR / job_id / '交换机端口资源管理表.xlsx'
    if not path.exists():
        raise HTTPException(status_code=404, detail='历史结果文件不存在')
    return FileResponse(str(path), filename=f'交换机端口资源管理表_{job_id}.xlsx',
                        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.get('/api/history/{job_id}/sheets/{sheet_name}')
def history_preview(job_id: str, sheet_name: str, limit: int = PREVIEW_ROWS) -> dict:
    path = RUNS_DIR / job_id / '交换机端口资源管理表.xlsx'
    if not path.exists():
        raise HTTPException(status_code=404, detail='历史结果文件不存在')
    xls = pd.ExcelFile(path)
    if sheet_name not in xls.sheet_names:
        raise HTTPException(status_code=404, detail=f'Sheet 不存在: {sheet_name}')
    df = pd.read_excel(xls, sheet_name=sheet_name).fillna('')
    head = df.head(max(1, min(limit, 1000)))
    return {
        'sheet': sheet_name, 'total': int(len(df)),
        'columns': [str(c) for c in head.columns.tolist()],
        'rows': [[_cell(v) for v in row] for row in head.itertuples(index=False, name=None)],
    }


@app.get('/api/history/{job_id}/sheets')
def history_sheets(job_id: str) -> dict:
    path = RUNS_DIR / job_id / '交换机端口资源管理表.xlsx'
    if not path.exists():
        raise HTTPException(status_code=404, detail='历史结果文件不存在')
    return {'sheets': pd.ExcelFile(path).sheet_names}


@app.exception_handler(Exception)
async def unhandled(_request: Request, exc: Exception):
    LOG.error('未处理异常: %s', exc)
    return JSONResponse(status_code=500, content={'detail': str(exc)})


def main() -> None:
    import uvicorn
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
    uvicorn.run(app, host='0.0.0.0', port=8000, log_level='info')


if __name__ == '__main__':
    main()
