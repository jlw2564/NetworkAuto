"""任务编排层：CLI 与 Web 共用同一套流程，通过回调暴露进度与日志。"""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence

from . import analyzer, collector, config, report
from .report import ALL_SECTIONS

LOG = logging.getLogger('net_port_tool')


# --------------------------------------------------------------------------
# 回调适配
# --------------------------------------------------------------------------
class _CallbackLogHandler(logging.Handler):
    """把日志转发给外部回调（Web 端用于实时滚动日志）。"""

    def __init__(self, callback: Callable[[str, str], None]):
        super().__init__()
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.callback(record.levelname, self.format(record))
        except Exception:  # noqa: BLE001 - 回调异常不能影响主流程
            pass


def _configure_logging(on_log: Optional[Callable[[str, str], None]] = None) -> logging.Logger:
    logger = logging.getLogger('net_port_tool')
    logger.setLevel(logging.INFO)
    logger.propagate = False
    for h in list(logger.handlers):
        logger.removeHandler(h)
    console = logging.StreamHandler()
    console.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', '%H:%M:%S'))
    logger.addHandler(console)
    if on_log:
        handler = _CallbackLogHandler(on_log)
        handler.setFormatter(logging.Formatter('%(message)s'))
        logger.addHandler(handler)
    return logger


@dataclass
class RunOptions:
    """一次运行的参数。"""
    switch_file: Optional[Path] = None
    terminal_file: Optional[Path] = None
    output_file: Optional[Path] = None
    concurrency: int = 6
    discover: bool = True
    sections: Sequence[str] = field(default_factory=lambda: tuple(ALL_SECTIONS))


@dataclass
class RunResult:
    ok: bool = False
    output_file: str = ''
    stats: Dict[str, int] = field(default_factory=dict)
    duration: float = 0.0
    error: str = ''
    core_name: str = ''
    devices: List[dict] = field(default_factory=list)
    unknown_switches: List[str] = field(default_factory=list)


def run(options: Optional[RunOptions] = None,
        on_log: Optional[Callable[[str, str], None]] = None,
        on_progress: Optional[Callable[[str, int, int, str], None]] = None) -> RunResult:
    """执行完整流程。任何异常都会被捕获并写进 RunResult.error。"""
    opts = options or RunOptions()
    log = _configure_logging(on_log)
    result = RunResult()
    started = time.time()

    def progress(stage: str, current: int, total: int, message: str = '') -> None:
        if on_progress:
            try:
                on_progress(stage, current, total, message)
            except Exception:  # noqa: BLE001
                pass

    try:
        # 1) 读取清单
        progress('load', 0, 1, '读取设备清单')
        log.info('读取交换机清单: %s', opts.switch_file or config.SWITCH_TABLE_FILE)
        devices = config.load_switches(opts.switch_file)
        if not devices:
            raise RuntimeError('交换机清单为空，请检查 switch_table.xlsx')
        terminals = []
        if 'terminal' in set(opts.sections):
            terminals = config.load_terminals(opts.terminal_file)
        log.info('交换机 %d 台，终端 %d 台', len(devices), len(terminals))
        core = next((d for d in devices if d.is_core), None)
        result.core_name = core.name if core else ''
        if not core:
            log.warning('未配置核心交换机（是否核心交换机=是），部分分析将以弱模式运行')
        progress('load', 1, 1, f'交换机 {len(devices)} 台 / 终端 {len(terminals)} 台')

        # 2) 并发采集
        log.info('开始并发采集（并发度=%d）...', opts.concurrency)
        done_count = {'n': 0}

        def _on_collect_done(dev, payload, err):
            done_count['n'] += 1
            if err is not None:
                progress('collect', done_count['n'], len(devices), f'{dev.name} 采集异常')
            else:
                state = '成功' if getattr(payload, 'ok', False) else f'失败({getattr(payload, "error", "")[:40]})'
                progress('collect', done_count['n'], len(devices), f'{dev.name} {state}')

        snapshots = collector.collect_all(
            devices, max_workers=opts.concurrency, logger=log, on_done=_on_collect_done)

        # 3) 发现未知交换机（可选）
        discovered_devices = []
        if opts.discover and core is not None:
            core_snap = next((s for s in snapshots if s.name == core.name), None)
            creds = []
            seen = set()
            for d in devices:
                if d.username and d.password and (d.username, d.password) not in seen:
                    seen.add((d.username, d.password))
                    creds.append((d.username, d.password))
            for c in config.fallback_credentials():
                if c not in seen:
                    seen.add(c)
                    creds.append(c)
            progress('discover', 0, 1, '从核心发现未知交换机')
            known_hosts = {d.host for d in devices}
            known_names = {d.name for d in devices}
            discovered_devices = collector.discover_unknown(
                core_snap, known_hosts, known_names, creds, logger=log)
            if discovered_devices:
                log.info('新发现 %d 台交换机，追加采集', len(discovered_devices))
                extra = collector.collect_all(discovered_devices, max_workers=opts.concurrency, logger=log)
                snapshots.extend(extra)
            progress('discover', 1, 1, f'新发现 {len(discovered_devices)} 台')

        # 4) 分析
        progress('analyze', 0, 1, '分析连线/终端/端口使用')
        analysis = analyzer.analyze(
            snapshots, terminals, core_name=result.core_name,
            discovered_names={d.name for d in discovered_devices}, logger=log)
        progress('analyze', 1, 1, '分析完成')

        # 5) 输出
        out_path = Path(opts.output_file) if opts.output_file else config.OUTPUT_FILE
        progress('report', 0, 1, '生成 Excel 报告')
        result.stats = report.write_report(
            out_path, analysis, core_name=result.core_name,
            sections=opts.sections, logger=log)
        result.output_file = str(out_path)
        progress('report', 1, 1, f'报告已生成: {out_path.name}')

        result.devices = [{
            'name': s.name, 'host': s.host, 'is_core': s.is_core,
            'ok': s.ok, 'error': s.error, 'sysname': s.sysname,
            'ports': len(s.brief_rows), 'macs': len(s.mac_index),
        } for s in snapshots]
        result.unknown_switches = analysis.unknown_switches
        result.ok = True
        log.info('全部完成：成功采集 %d 台，报告=%s',
                 sum(1 for s in snapshots if s.ok), out_path.name)
    except Exception as exc:  # noqa: BLE001 - 顶层兜底，交给调用方展示
        result.ok = False
        result.error = str(exc)
        log.exception('运行失败: %s', exc)
    finally:
        result.duration = round(time.time() - started, 2)
    return result
