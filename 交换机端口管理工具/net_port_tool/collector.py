"""并发采集层：每台交换机只登录一次，一次性取回分析所需全部数据。"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

from . import parsers
from .config import SwitchDevice
from .ssh_client import SessionError, SwitchSession, gather_concurrently, probe_login

LOG = logging.getLogger('net_port_tool.collector')

# (命令, delay_factor, read_timeout)
BASE_CMDS: Tuple[Tuple[str, int, int], ...] = (
    ('display current-configuration | include sysname', 2, 20),
    ('display lldp neighbor-information list', 5, 30),
    ('display link-aggregation verbose', 5, 30),
    ('display interface brief', 5, 40),
    ('display mac-address all', 6, 30),
)
CORE_CMDS: Tuple[Tuple[str, int, int], ...] = (
    ('display arp all', 12, 40),
)
MAC_FALLBACK_CMD = 'display mac-address'
ARP_FALLBACK_CMD = 'display arp'


@dataclass
class SwitchSnapshot:
    """一台交换机的采集结果（纯数据，便于缓存与并发）。"""
    name: str
    host: str
    is_core: bool = False
    ok: bool = False
    error: str = ''
    sysname: str = ''
    lldp_entries: List[dict] = field(default_factory=list)
    lldp_any: Dict[str, dict] = field(default_factory=dict)
    lldp_managed: set = field(default_factory=set)
    member_to_agg: Dict[str, str] = field(default_factory=dict)
    agg_to_members: Dict[str, List[str]] = field(default_factory=dict)
    brief_rows: List[dict] = field(default_factory=list)
    brief_by_if: Dict[str, dict] = field(default_factory=dict)
    port_to_macs: Dict[str, List[str]] = field(default_factory=dict)
    port_to_mac_vlan: Dict[str, Dict[str, str]] = field(default_factory=dict)
    mac_index: Dict[str, dict] = field(default_factory=dict)
    arp_mac_to_ip: Dict[str, str] = field(default_factory=dict)
    arp_ip_to_mac: Dict[str, str] = field(default_factory=dict)
    raw: Dict[str, str] = field(default_factory=dict)

    def uplink_ports(self) -> set:
        """本机作为上联的口（LLDP 邻居是已纳管交换机）。"""
        return self.lldp_managed


def collect_switch(device: SwitchDevice, logger: Optional[logging.Logger] = None) -> SwitchSnapshot:
    """采集单台交换机（一次登录）。失败时返回 ok=False 的快照，不抛异常。"""
    log = logger or LOG
    snap = SwitchSnapshot(name=device.name, host=device.host, is_core=device.is_core)
    cmds = list(BASE_CMDS) + (list(CORE_CMDS) if device.is_core else [])
    try:
        with SwitchSession(device, logger=log) as sess:
            raw = sess.run_many(cmds)
            # MAC 表兜底
            mac_out = raw.get('display mac-address all', '')
            if not parsers.RE_MAC.search((mac_out or '').lower()):
                raw[MAC_FALLBACK_CMD] = sess.run(MAC_FALLBACK_CMD, delay=5, read_timeout=25)
            # ARP 兜底（仅核心）
            if device.is_core:
                arp_out = raw.get('display arp all', '')
                if not parsers.RE_IP.search(arp_out or ''):
                    raw[ARP_FALLBACK_CMD] = sess.run(ARP_FALLBACK_CMD, delay=10, read_timeout=30)
    except SessionError as exc:
        snap.error = str(exc)
        log.warning('采集失败 %s (%s): %s', device.name, device.host, exc)
        return snap

    _fill_snapshot(snap, device, raw)
    snap.raw = raw
    snap.ok = True
    return snap


def _fill_snapshot(snap: SwitchSnapshot, device: SwitchDevice, raw: Dict[str, str]) -> None:
    """把原始输出解析进快照。"""
    sys_out = raw.get('display current-configuration | include sysname', '')
    m = re.search(r'^\s*sysname\s+(\S+)', sys_out, re.MULTILINE) or re.search(r'sysname\s+(\S+)', sys_out)
    snap.sysname = m.group(1).strip() if m else ''

    snap.lldp_entries = parsers.parse_lldp_neighbor_list(
        raw.get('display lldp neighbor-information list', ''))

    mac_outputs = [raw.get('display mac-address all', ''), raw.get(MAC_FALLBACK_CMD, '')]
    snap.port_to_macs, snap.port_to_mac_vlan, snap.mac_index = parsers.parse_mac_table(mac_outputs)

    snap.brief_rows = parsers.parse_interface_brief(raw.get('display interface brief', ''))
    snap.brief_by_if = {r['interface']: r for r in snap.brief_rows}

    snap.member_to_agg, snap.agg_to_members = parsers.parse_link_aggregation_verbose(
        raw.get('display link-aggregation verbose', ''))

    if device.is_core:
        arp_outputs = [raw.get('display arp all', ''), raw.get(ARP_FALLBACK_CMD, '')]
        maps = [parsers.parse_arp_to_maps(x) for x in arp_outputs]
        snap.arp_mac_to_ip, snap.arp_ip_to_mac = parsers.merge_mac_ip_maps(maps)


def collect_all(devices: List[SwitchDevice], max_workers: int = 6,
                logger: Optional[logging.Logger] = None,
                on_done: Optional[Callable] = None) -> List[SwitchSnapshot]:
    """并发采集多台交换机，返回与输入顺序一致、且数量一致的快照列表。"""
    log = logger or LOG
    if not devices:
        return []

    def _worker(dev: SwitchDevice) -> SwitchSnapshot:
        return collect_switch(dev, logger=log)

    ordered = gather_concurrently(devices, _worker, max_workers=max_workers, on_done=on_done)
    return [payload for _item, payload, _err in ordered]


def discover_unknown(core_snap: Optional[SwitchSnapshot], known_hosts: set, known_names: set,
                     credentials: List[Tuple[str, str]], logger: Optional[logging.Logger] = None,
                     on_found: Optional[Callable] = None) -> List[SwitchDevice]:
    """借助核心的 ARP + LLDP 发现清单外交换机。

    候选 = LLDP 邻居的 chassis MAC 在核心 ARP 表中出现过的 IP（且不在已知清单）。
    用凭据依次试登录，成功且 sysname 未纳管则视为新设备。
    """
    log = logger or LOG
    if core_snap is None or not core_snap.arp_mac_to_ip or not credentials:
        return []

    mac_to_ip = core_snap.arp_mac_to_ip
    candidates: List[str] = []
    for entry in core_snap.lldp_entries:
        mac = str(entry.get('chassis_id', '')).strip().lower()
        if not mac:
            continue
        ip = mac_to_ip.get(mac)
        if ip and ip not in known_hosts and ip not in candidates:
            candidates.append(ip)
    if not candidates:
        return []

    log.info('发现 %d 个潜在未知 IP，将用 %d 组凭据尝试登录',
             len(candidates), len(credentials))
    found: List[SwitchDevice] = []
    for ip in candidates:
        for username, password in credentials:
            ok, sysname, err = probe_login(ip, username, password, logger=log)
            if ok and sysname:
                if sysname in known_names:
                    log.info('  [跳过] %s (%s) 已在清单中', sysname, ip)
                    break
                log.info('  [发现] %s (%s) 使用账号 %s', sysname, ip, username)
                dev = SwitchDevice(name=sysname, host=ip, username=username, password=password)
                found.append(dev)
                known_names.add(sysname)
                known_hosts.add(ip)
                if on_found:
                    try:
                        on_found(dev)
                    except Exception:  # noqa: BLE001 - 回调异常不影响发现流程
                        pass
                break
            if err:
                log.debug('  登录 %s 失败(%s): %s', ip, username, err)
    return found
