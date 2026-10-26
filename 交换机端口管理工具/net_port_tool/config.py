"""配置与设备清单加载。

与原脚本的差异：
- 路径默认基于项目根目录解析，不再依赖“当前工作目录”，避免同名文件取错；
- 允许运行时覆盖输入/输出文件路径（CLI 参数、Web 上传都会用到）；
- 凭据只从 Excel 读取，代码内不再硬编码明文密码。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

# 项目根目录：net_port_tool/ 的上一级
BASE_DIR = Path(__file__).resolve().parent.parent

# 默认文件路径
SWITCH_TABLE_FILE = BASE_DIR / 'switch_table.xlsx'
TERMINAL_IP_FILE = BASE_DIR / '终端IP表.xlsx'
OUTPUT_FILE = BASE_DIR / '交换机端口资源管理表.xlsx'

# 运行产物目录（Web 任务结果、日志）
RUNS_DIR = BASE_DIR / 'net_port_tool_runs'

# Excel 列名兼容表（现场表格列名不统一，按候选顺序匹配第一个命中的）
SWITCH_COLUMNS = {
    'name': ['设备名称', '交换机名称', '名称', 'hostname', 'sysname'],
    'host': ['IP地址', 'IP', '管理IP', 'ip_address'],
    'username': ['用户名', '账号', 'user', 'username'],
    'password': ['密码', '口令', 'password', 'pwd'],
    'is_core': ['是否核心交换机', '是否核心', '核心交换机', 'is_core'],
}
TERMINAL_COLUMNS = {
    'name': ['设备名称', '名称', '终端名称'],
    'host': ['IP地址', 'IP', 'ip_address'],
}

_TRUE_WORDS = ('是', 'yes', 'true', '1', 'y', '核心')

# 可选兜底凭据：仅当显式设置环境变量 NET_PORT_FALLBACK_CREDS 时启用
# 格式： "user1:pwd1,user2:pwd2"
FALLBACK_CREDS_ENV = 'NET_PORT_FALLBACK_CREDS'


@dataclass
class SwitchDevice:
    """一台交换机。"""
    name: str
    host: str
    username: str
    password: str
    is_core: bool = False
    device_type: str = 'hp_comware'
    raw: Dict[str, str] = field(default_factory=dict)

    def conn_params(self) -> Dict[str, object]:
        """构造 netmiko ConnectHandler 参数（不含自定义键）。"""
        return {
            'device_type': self.device_type,
            'host': self.host,
            'username': self.username,
            'password': self.password,
            'timeout': 30,
            'conn_timeout': 15,
            'banner_timeout': 20,
            'auth_timeout': 20,
            'fast_cli': False,
            'global_delay_factor': 1.5,
        }


@dataclass
class TerminalDevice:
    """一台终端（来自终端 IP 表）。"""
    name: str
    ip: str


def _pick_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    """按候选顺序（忽略大小写与首尾空白）找到第一个存在的列名。"""
    normalized = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in normalized:
            return normalized[key]
    return None


def load_switches(path: Optional[os.PathLike] = None) -> List[SwitchDevice]:
    """读取交换机清单。缺列或空行会被跳过，不抛异常中断整体流程。"""
    src = Path(path) if path else SWITCH_TABLE_FILE
    df = pd.read_excel(src)
    col = {k: _pick_column(list(df.columns), v) for k, v in SWITCH_COLUMNS.items()}
    if not col['name'] or not col['host']:
        raise ValueError(
            f'{src.name} 缺少必需列：设备名称 / IP地址。当前列名={list(df.columns)}'
        )

    devices: List[SwitchDevice] = []
    seen_hosts = set()
    for _, row in df.iterrows():
        name = str(row.get(col['name'], '')).strip()
        host = str(row.get(col['host'], '')).strip()
        if not name or not host or host.lower() in ('nan', 'none'):
            continue
        if host in seen_hosts:
            continue
        seen_hosts.add(host)
        username = str(row.get(col['username'], '')).strip() if col['username'] else ''
        password = str(row.get(col['password'], '')).strip() if col['password'] else ''
        if password.lower() == 'nan':
            password = ''
        is_core = False
        if col['is_core']:
            is_core = str(row.get(col['is_core'], '')).strip().lower() in _TRUE_WORDS
        devices.append(SwitchDevice(
            name=name, host=host, username=username, password=password, is_core=is_core,
        ))
    return devices


def load_terminals(path: Optional[os.PathLike] = None) -> List[TerminalDevice]:
    """读取终端 IP 表。"""
    src = Path(path) if path else TERMINAL_IP_FILE
    df = pd.read_excel(src)
    col = {k: _pick_column(list(df.columns), v) for k, v in TERMINAL_COLUMNS.items()}
    if not col['host']:
        raise ValueError(f'{src.name} 缺少必需列：IP地址。当前列名={list(df.columns)}')

    terminals: List[TerminalDevice] = []
    seen = set()
    for _, row in df.iterrows():
        ip = str(row.get(col['host'], '')).strip()
        if not ip or ip.lower() in ('nan', 'none'):
            continue
        if ip in seen:
            continue
        seen.add(ip)
        name = str(row.get(col['name'], '')).strip() if col['name'] else ''
        if name.lower() == 'nan':
            name = ''
        terminals.append(TerminalDevice(name=name or '未知设备', ip=ip))
    return terminals


def fallback_credentials() -> List[tuple]:
    """兜底凭据（用于发现未知交换机）。默认空，需用环境变量显式开启。"""
    raw = os.environ.get(FALLBACK_CREDS_ENV, '').strip()
    if not raw:
        return []
    creds = []
    for item in raw.split(','):
        if ':' in item:
            u, p = item.split(':', 1)
            u, p = u.strip(), p.strip()
            if u and p:
                creds.append((u, p))
    return creds


def ensure_runs_dir() -> Path:
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    return RUNS_DIR
