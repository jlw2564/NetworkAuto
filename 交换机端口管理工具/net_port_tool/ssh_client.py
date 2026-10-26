"""SSH 会话封装。

核心优化点：
- 一台设备只登录一次（原来每台设备被反复登录 4 次以上，且每台终端都会新建连接）；
- `screen-length disable` / `terminal logging level 6` 只在登录后下发一次，不再每条命令前重复；
- 提供 `gather_concurrently` 支持按交换机并发采集。
"""
from __future__ import annotations

import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from netmiko import ConnectHandler

from .config import SwitchDevice

LOG = logging.getLogger('net_port_tool.ssh')

# 登录后需要执行一次的会话级准备命令
SESSION_SETUP_CMDS: Sequence[Tuple[str, int]] = (
    ('screen-length disable', 1),
    ('terminal logging level 6', 1),
)


class SessionError(RuntimeError):
    """登录或会话准备失败。"""


class SwitchSession:
    """一次 SSH 会话。推荐用 with 语句，确保连接释放。"""

    def __init__(self, device: SwitchDevice, logger: Optional[logging.Logger] = None):
        self.device = device
        self.log = logger or LOG
        self._conn = None

    # ---------- 上下文管理 ----------
    def __enter__(self) -> 'SwitchSession':
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        self.close()
        return False

    # ---------- 连接 ----------
    def connect(self) -> None:
        try:
            self._conn = ConnectHandler(**self.device.conn_params())
        except Exception as exc:  # noqa: BLE001 - 需要把原始异常信息带给上层
            raise SessionError(f'登录失败: {str(exc).strip()[:120]}') from exc
        for cmd, delay in SESSION_SETUP_CMDS:
            self._safe(cmd, delay=delay, read_timeout=15)
        self.log.debug('已登录 %s (%s)', self.device.name, self.device.host)

    def close(self) -> None:
        if self._conn is not None:
            try:
                self._conn.disconnect()
            except Exception:  # noqa: BLE001 - 关闭失败不影响结果
                pass
            self._conn = None

    # ---------- 命令 ----------
    def _safe(self, cmd: str, delay: int = 2, read_timeout: int = 30) -> str:
        if self._conn is None:
            return ''
        try:
            return self._conn.send_command(
                cmd, delay_factor=delay, expect_string=r'>', read_timeout=read_timeout,
            ) or ''
        except Exception as exc:  # noqa: BLE001 - 单条命令失败不应中断整机采集
            self.log.warning('[%s] 命令失败: %s (%s)', self.device.name, cmd, str(exc)[:80])
            return ''

    def run(self, cmd: str, delay: int = 2, read_timeout: int = 30) -> str:
        return self._safe(cmd, delay=delay, read_timeout=read_timeout)

    def run_many(self, cmds: Iterable[Tuple[str, int, int]]) -> Dict[str, str]:
        """批量下发命令，返回 {命令: 输出}。cmds 元素为 (命令, delay, read_timeout)。"""
        result: Dict[str, str] = {}
        for cmd, delay, read_timeout in cmds:
            result[cmd] = self._safe(cmd, delay=delay, read_timeout=read_timeout)
        return result

    def is_alive(self) -> bool:
        return self._conn is not None


def probe_login(host: str, username: str, password: str,
                device_type: str = 'hp_comware', logger: Optional[logging.Logger] = None
                ) -> Tuple[bool, str, str]:
    """尝试登录并读取 sysname。返回 (是否成功, sysname, 错误信息)。"""
    dev = SwitchDevice(name=host, host=host, username=username, password=password,
                       device_type=device_type)
    try:
        with SwitchSession(dev, logger=logger) as sess:
            out = sess.run('display current-configuration | include sysname',
                           delay=2, read_timeout=20)
    except SessionError as exc:
        return False, '', str(exc)
    m = None
    import re as _re
    for pattern in (r'^\s*sysname\s+(\S+)', r'sysname\s+(\S+)'):
        m = _re.search(pattern, out, _re.MULTILINE)
        if m:
            break
    if not m:
        return False, '', '读取设备名失败'
    return True, m.group(1).strip(), ''


def gather_concurrently(items: List, worker: Callable, max_workers: int = 6,
                        on_done: Optional[Callable] = None):
    """并发执行 worker(item)，结果按输入顺序返回 [(item, result, exc)]。

    单台失败不影响其他设备；worker 内部异常会被捕获并原样返回。
    """
    results: List[Optional[Tuple]] = [None] * len(items)
    if not items:
        return []
    workers = max(1, min(int(max_workers), len(items)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_map = {pool.submit(worker, item): idx for idx, item in enumerate(items)}
        for fut in as_completed(future_map):
            idx = future_map[fut]
            item = items[idx]
            try:
                payload, err = fut.result(), None
            except Exception as exc:  # noqa: BLE001 - 汇总失败原因给上层
                payload, err = None, exc
            results[idx] = (item, payload, err)
            if on_done:
                try:
                    on_done(item, payload, err)
                except Exception:  # noqa: BLE001 - 进度回调异常不影响采集
                    pass
    return results


def parallel_run(device: SwitchDevice, cmds: Sequence[Tuple[str, int, int]],
                 logger: Optional[logging.Logger] = None) -> Tuple[bool, Dict[str, str], str]:
    """登录一次并把命令全部取回。返回 (是否成功, {命令: 输出}, 错误信息)。"""
    try:
        with SwitchSession(device, logger=logger) as sess:
            return True, sess.run_many(cmds), ''
    except SessionError as exc:
        return False, {}, str(exc)


def sleep_between(retry_index: int) -> None:
    """重试间隔（递增），用于发现未知交换机时避免密集打设备。"""
    time.sleep(0.5 + retry_index * 0.5)
