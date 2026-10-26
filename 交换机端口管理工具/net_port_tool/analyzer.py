"""分析层：完全基于采集快照做连线 / 终端定位 / 端口使用 / 可疑端口判定。

与原实现的差别：
- 不再需要任何额外 SSH 连接，终端定位靠 `mac_index` 直接反查（原来每台终端都要重连一次）；
- `link_by_port` / `devices_by_port` / 终端 IP 字典全部预建，去掉原来的 O(n²)~O(n³) 扫描。
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd

from .collector import SwitchSnapshot
from .config import TerminalDevice
from .parsers import FUZZY_TERMINAL_KEYWORDS, is_lag_or_virtual, is_physical_port, port_sort_key

LOG = logging.getLogger('net_port_tool.analyzer')

# 可疑端口判定参数
MIN_MAC_SUSPICIOUS = 1
SUSPICIOUS_ONLY_UP = True

TERMINAL_TABLE_COLS = ['设备类型', '设备名称', 'IP地址', 'MAC地址', '所在交换机', '所在端口', 'VLAN', '备注']
ANOMALY_EXTRA_COLS = ['异常类型', '原因', '处理建议']
SUSPICIOUS_COLS = [
    '交换机', '端口', 'VLAN', '链路状态', '端口描述', '学到MAC数', '首个MAC',
    '涉及VLAN(MAC维度)', 'LLDP邻居', '聚合组', '判定依据', '建议操作',
]
PORT_USAGE_COLS = ['交换机', '端口', 'VLAN', '端口使用情况', '终端设备名称', '终端IP', '聚合组']
NAME_VERIFY_COLS = ['期望名称', '实际名称', 'IP地址', '状态']
LINK_COLS = ['本地交换机', '本地端口', '邻居交换机', '邻居端口', '连接类型']

_EMPTY = ('', '未知', 'nan', 'None', '未找到')


@dataclass
class AnalysisResult:
    name_verification: List[dict] = field(default_factory=list)
    links: List[dict] = field(default_factory=list)
    port_data: List[dict] = field(default_factory=list)
    port_usage: List[dict] = field(default_factory=list)
    suspicious: List[dict] = field(default_factory=list)
    unknown_switches: List[str] = field(default_factory=list)


def _clean(value) -> str:
    return str(value if value is not None else '').strip()


def _is_empty(value) -> bool:
    v = _clean(value)
    return v == '' or v in _EMPTY


def analyze(snapshots: Sequence[SwitchSnapshot], terminals: Sequence[TerminalDevice],
            core_name: str = '', discovered_names: Optional[set] = None,
            logger: Optional[logging.Logger] = None) -> AnalysisResult:
    """主分析入口。"""
    log = logger or LOG
    result = AnalysisResult()
    discovered_names = discovered_names or set()
    snap_by_name: Dict[str, SwitchSnapshot] = {s.name: s for s in snapshots}
    managed_names = set(snap_by_name.keys())
    core_snap = snap_by_name.get(core_name)

    # 补齐每台设备的 LLDP 派生信息
    for snap in snapshots:
        snap.lldp_any, snap.lldp_managed = _lldp_maps(snap, managed_names)

    result.name_verification = _verify_names(snapshots, discovered_names)

    result.links, link_by_port = _build_links(snapshots, managed_names, core_name)
    result.unknown_switches = sorted(
        {ln['邻居交换机'] for ln in result.links if ln['邻居交换机'] not in managed_names}
    )

    port_data: List[dict] = []
    seen_rows = set()

    def _add_row(row: dict) -> None:
        key = (row.get('所在交换机', ''), row.get('所在端口', ''),
               str(row.get('MAC地址', '')).lower(), row.get('IP地址', ''))
        if key in seen_rows:
            return
        seen_rows.add(key)
        port_data.append(row)

    _locate_terminals(terminals, snap_by_name, core_snap, managed_names, _add_row, log)
    occupied = _build_occupied_ports(port_data)

    known_uplinks = _build_known_uplinks(result.links)
    _scan_extra_up_ports(snapshots, core_snap, core_name, occupied, known_uplinks,
                         link_by_port, _add_row, log)

    result.port_data = port_data

    devices_by_port: Dict[Tuple[str, str], List[dict]] = {}
    for row in port_data:
        sw = _clean(row.get('所在交换机'))
        port = _clean(row.get('所在端口'))
        if not sw or not port or _is_empty(sw) or _is_empty(port):
            continue
        devices_by_port.setdefault((sw, port), []).append(row)

    usage, suspicious = _build_port_usage_and_suspicious(
        snapshots, core_name, link_by_port, devices_by_port, seen_rows, _add_row, log)
    result.port_usage = usage
    result.suspicious = suspicious
    result.port_data = port_data
    return result


# --------------------------------------------------------------------------
# LLDP 派生
# --------------------------------------------------------------------------
def _lldp_maps(snap: SwitchSnapshot, managed_names: set) -> Tuple[Dict[str, dict], set]:
    any_neighbor: Dict[str, dict] = {}
    managed = set()
    for e in snap.lldp_entries or []:
        port = _clean(e.get('local_port'))
        if not port:
            continue
        key = port.upper()
        name = _clean(e.get('neighbor_name'))
        any_neighbor[key] = {'neighbor_name': name, 'neighbor_port': _clean(e.get('neighbor_port'))}
        if name and name in managed_names:
            managed.add(key)
    return any_neighbor, managed


# --------------------------------------------------------------------------
# 命名验证
# --------------------------------------------------------------------------
def _verify_names(snapshots: Sequence[SwitchSnapshot], discovered_names: set) -> List[dict]:
    rows = []
    for snap in snapshots:
        expected = snap.name
        actual = snap.sysname or expected
        if snap.name in discovered_names:
            status = '自动发现'
        elif not snap.ok:
            err = snap.error or '未知错误'
            status = err if err.startswith('登录失败') else f'登录失败: {err[:80]}'
        elif not snap.sysname:
            status = '读取设备名失败'
        elif snap.sysname != expected:
            status = f'名称不一致(实际={snap.sysname})'
        else:
            status = '匹配'
        rows.append({
            '期望名称': expected, '实际名称': actual,
            'IP地址': snap.host, '状态': status,
        })
    return rows


# --------------------------------------------------------------------------
# 连线
# --------------------------------------------------------------------------
def _build_links(snapshots: Sequence[SwitchSnapshot], managed_names: set,
                 core_name: str) -> Tuple[List[dict], Dict[Tuple[str, str], dict]]:
    links: List[dict] = []
    link_by_port: Dict[Tuple[str, str], dict] = {}
    seen_pairs = set()
    # 核心优先，保证核心的连线先出现
    ordered = sorted(snapshots, key=lambda s: (not s.is_core, s.name))
    for snap in ordered:
        for entry in snap.lldp_entries or []:
            neighbor = _clean(entry.get('neighbor_name'))
            if not neighbor:
                continue
            local_port = _clean(entry.get('local_port'))
            neighbor_port = _clean(entry.get('neighbor_port'))
            pair = frozenset({snap.name, neighbor})
            if len(pair) == 2 and pair in seen_pairs:
                continue
            if len(pair) == 2:
                seen_pairs.add(pair)
            row = {
                '本地交换机': snap.name,
                '本地端口': local_port.upper(),
                '邻居交换机': neighbor,
                '邻居端口': neighbor_port if neighbor_port else '待查',
                '连接类型': '交换机级联',
            }
            links.append(row)
            link_by_port[(snap.name, local_port.upper())] = {
                'neighbor_switch': neighbor, 'neighbor_port': neighbor_port,
            }
            link_by_port[(neighbor, neighbor_port.upper())] = {
                'neighbor_switch': snap.name, 'neighbor_port': local_port,
            }
    return links, link_by_port


def _build_known_uplinks(links: Sequence[dict]) -> Dict[str, set]:
    known: Dict[str, set] = {}
    for link in links:
        a_sw, a_port = _clean(link.get('本地交换机')), _clean(link.get('本地端口'))
        b_sw, b_port = _clean(link.get('邻居交换机')), _clean(link.get('邻居端口'))
        if a_sw and a_port:
            known.setdefault(a_sw, set()).add(a_port.upper())
        if b_sw and b_port:
            known.setdefault(b_sw, set()).add(b_port.upper())
    return known


def _build_occupied_ports(port_data: Sequence[dict]) -> Dict[str, set]:
    occupied: Dict[str, set] = {}
    for row in port_data:
        sw, port = _clean(row.get('所在交换机')), _clean(row.get('所在端口'))
        if sw and port and not _is_empty(sw) and not _is_empty(port):
            occupied.setdefault(sw, set()).add(port)
    return occupied


# --------------------------------------------------------------------------
# 终端定位（零 SSH：靠核心 ARP + 各交换机 mac_index 反查）
# --------------------------------------------------------------------------
def _resolve_vlan(snap: Optional[SwitchSnapshot], mac: str, port: str,
                  preferred: str = '') -> str:
    if preferred and preferred not in _EMPTY:
        return preferred
    if snap is not None:
        mac_vlan = (snap.port_to_mac_vlan.get(port) or {}).get(mac, '')
        if mac_vlan and mac_vlan not in _EMPTY:
            return mac_vlan
        brief_vlan = _clean((snap.brief_by_if.get(port) or {}).get('vlan', ''))
        if brief_vlan:
            return brief_vlan
    return '未知'


def _locate_terminals(terminals, snap_by_name, core_snap, managed_names, add_row, log) -> None:
    ip_to_mac = dict(core_snap.arp_ip_to_mac) if core_snap else {}
    for dev in terminals:
        ip = _clean(dev.ip)
        if not ip or ip == 'nan':
            continue
        mac = ip_to_mac.get(ip, '')
        row = {
            '设备类型': '终端设备', '设备名称': dev.name, 'IP地址': ip,
            'MAC地址': mac if mac else '未发现',
            '所在交换机': '未找到', '所在端口': '未找到', 'VLAN': '未找到',
            '备注': 'ARP存在' if mac else '无ARP记录',
        }
        if not mac:
            row['备注'] = '无ARP记录, 未发现MAC'
            add_row(row)
            continue

        core_port = ''
        if core_snap is not None:
            core_port = _clean((core_snap.mac_index.get(mac) or {}).get('port', ''))
        if not core_port:
            add_row(row)
            continue

        core_neighbor = ''
        if core_snap is not None:
            core_neighbor = _clean((core_snap.lldp_any.get(core_port.upper()) or {})
                                   .get('neighbor_name', ''))

        # 优先在“核心该端口的 LLDP 邻居”那台接入交换机上找
        access_snap = snap_by_name.get(core_neighbor) if core_neighbor in managed_names else None
        if access_snap is not None:
            info = access_snap.mac_index.get(mac)
            if info:
                port = _clean(info.get('port'))
                uplinks = access_snap.lldp_managed
                if port and port.upper() not in uplinks:
                    mac_count = len(access_snap.port_to_macs.get(port, []) or [])
                    row['所在交换机'] = access_snap.name
                    row['所在端口'] = port
                    row['VLAN'] = _resolve_vlan(access_snap, mac, port, info.get('vlan', ''))
                    remark = row['备注'] if row['备注'] else ''
                    remark = (remark + ', ' if remark else '') + '已定位'
                    if mac_count > 3:
                        remark += f', 端口下共{mac_count}个MAC(可能级联)'
                    row['备注'] = remark
                    add_row(row)
                    continue

        # 回退：核心直连
        row['所在交换机'] = core_snap.name if core_snap else ''
        row['所在端口'] = core_port
        row['VLAN'] = _resolve_vlan(core_snap, mac, core_port)
        row['备注'] = (row['备注'] + ', ' if row['备注'] else '') + '核心直连'
        add_row(row)


# --------------------------------------------------------------------------
# 额外 UP 端口扫描（清单里没有的终端 / 无 MAC 的空口）
# --------------------------------------------------------------------------
def _up_physical_ports(snap: SwitchSnapshot, known_uplinks: set) -> List[str]:
    ports = []
    for r in snap.brief_rows:
        port = _clean(r.get('interface'))
        if not port or not is_physical_port(port):
            continue
        port_up = port.upper()
        if port_up in snap.lldp_managed:
            continue
        if port_up in known_uplinks or port in known_uplinks:
            continue
        if _clean(r.get('link')).upper() == 'UP':
            ports.append(port)
    return sorted(ports, key=port_sort_key)


def _scan_extra_up_ports(snapshots, core_snap, core_name, occupied, known_uplinks,
                         link_by_port, add_row, log) -> None:
    mac_to_ip = dict(core_snap.arp_mac_to_ip) if core_snap else {}

    def _emit(snap: SwitchSnapshot, port: str, tag_no_mac: str, tag_has: str,
              skip_with_lldp: bool) -> None:
        if skip_with_lldp and _port_has_lldp(snap, port):
            return
        if port in occupied.get(snap.name, set()):
            return
        macs = list(dict.fromkeys(snap.port_to_macs.get(port, []) or []))
        port_vlan = _clean((snap.brief_by_if.get(port) or {}).get('vlan', '')) or '未知'
        mac_vlan_map = snap.port_to_mac_vlan.get(port, {})
        if not macs:
            add_row({
                '设备类型': '未知设备', '设备名称': '未知设备', 'IP地址': '未知', 'MAC地址': '未知',
                '所在交换机': snap.name, '所在端口': port, 'VLAN': port_vlan, '备注': tag_no_mac,
            })
            return
        for m in macs:
            extra_ip = mac_to_ip.get(m, '未知')
            add_row({
                '设备类型': '终端设备',
                '设备名称': '发现设备' if extra_ip == '未知' else '发现设备',
                'IP地址': extra_ip, 'MAC地址': m,
                '所在交换机': snap.name, '所在端口': port,
                'VLAN': _clean(mac_vlan_map.get(m, '')) or port_vlan,
                '备注': tag_has,
            })

    for snap in snapshots:
        if snap.is_core or snap.name == core_name:
            continue
        for port in _up_physical_ports(snap, known_uplinks.get(snap.name, set())):
            _emit(snap, port, '额外Up端口 (无MAC)', '额外Up端口 (有数据)', skip_with_lldp=False)

    if core_snap is not None:
        for port in _up_physical_ports(core_snap, set()):
            _emit(core_snap, port, '核心额外直连端口 (无MAC)',
                  '核心额外直连端口 (有数据)', skip_with_lldp=True)


def _port_has_lldp(snap: SwitchSnapshot, port: str) -> bool:
    info = snap.lldp_any.get(port.upper())
    return bool(info and _clean(info.get('neighbor_name')))


# --------------------------------------------------------------------------
# 端口使用表 + 可疑端口
# --------------------------------------------------------------------------
def _build_port_usage_and_suspicious(snapshots, core_name, link_by_port, devices_by_port,
                                     seen_rows, add_row, log):
    usage_rows: List[dict] = []
    suspicious_rows: List[dict] = []
    managed_names = {s.name for s in snapshots}

    for snap in snapshots:
        sw_name = snap.name
        for r in sorted(snap.brief_rows, key=lambda x: port_sort_key(x.get('interface', ''))):
            port = _clean(r.get('interface'))
            if not port or not is_physical_port(port):
                continue
            port_up = port.upper()
            link_status = _clean(r.get('link'))
            desc = _clean(r.get('description'))
            vlan_from_brief = _clean(r.get('vlan'))
            macs = list(dict.fromkeys(snap.port_to_macs.get(port, []) or []))
            mac_vlan_map = snap.port_to_mac_vlan.get(port, {})
            mac_vlans = [v for v in mac_vlan_map.values() if v and v != '未知']
            mac_vlan_str = ';'.join(sorted({_clean(v) for v in mac_vlans if _clean(v)}))

            neighbor_info = link_by_port.get((sw_name, port_up)) or link_by_port.get((sw_name, port))
            if not neighbor_info:
                lldp_n = snap.lldp_any.get(port_up)
                if lldp_n:
                    n_name = _clean(lldp_n.get('neighbor_name'))
                    if n_name and n_name not in managed_names:
                        neighbor_info = {'neighbor_switch': n_name,
                                         'neighbor_port': _clean(lldp_n.get('neighbor_port'))}

            devices = devices_by_port.get((sw_name, port), [])
            terminal_names, terminal_ips, terminal_vlans = _collect_terminal_fields(devices)

            vlan = vlan_from_brief or (terminal_vlans[0] if terminal_vlans else mac_vlan_str or '未知')

            agg_group = _clean(snap.member_to_agg.get(port_up, ''))

            is_cascade = bool(neighbor_info) or (port_up in snap.lldp_managed)
            if not is_cascade and agg_group:
                agg_up = agg_group.upper()
                if (agg_up in snap.lldp_managed or agg_up in snap.lldp_any
                        or link_by_port.get((sw_name, agg_up)) or link_by_port.get((sw_name, agg_group))):
                    is_cascade = True

            suspicious, reasons = _judge_suspicious(
                link_status, is_cascade, macs, desc, port_up, snap)

            if suspicious:
                first_mac = macs[0] if macs else ''
                first_mac_vlan = _clean(mac_vlan_map.get(first_mac, '')) if first_mac else ''
                suspicious_vlan = first_mac_vlan if first_mac_vlan and first_mac_vlan != '未知' else vlan
                lldp_info = snap.lldp_any.get(port_up)
                suspicious_rows.append({
                    '交换机': sw_name, '端口': port, 'VLAN': suspicious_vlan,
                    '链路状态': link_status, '端口描述': desc, '学到MAC数': len(macs),
                    '首个MAC': first_mac, '涉及VLAN(MAC维度)': mac_vlan_str,
                    'LLDP邻居': _clean(lldp_info.get('neighbor_name')) if lldp_info else '',
                    '聚合组': agg_group, '判定依据': '；'.join(reasons),
                    '建议操作': '现场确认端口下接了什么设备；若无合法终端则关端口/查环路/查傻瓜交换机',
                })
                for idx_m, m in enumerate(macs):
                    v_m = _clean(mac_vlan_map.get(m, '')) or suspicious_vlan
                    dev_name = '可疑终端(待确认)' if len(macs) == 1 else f'可疑终端{idx_m + 1}(待确认)'
                    add_row({
                        '设备类型': '终端设备', '设备名称': dev_name,
                        'IP地址': '未知(核心ARP无，请现场确认)', 'MAC地址': m,
                        '所在交换机': sw_name, '所在端口': port, 'VLAN': v_m,
                        '备注': f'可疑端口兜底识别；{";".join(reasons)}',
                    })

            usage = _decide_usage(is_cascade, terminal_ips, suspicious, macs, agg_group, r, link_status)
            terminal_ip_str = ''
            terminal_name_str = ''
            if usage == '终端':
                if not terminal_ips:
                    terminal_ips = ['(可疑, 核心ARP无记录)']
                    terminal_names = ['(可疑终端, 需确认)']
                terminal_ip_str = ';'.join(dict.fromkeys(terminal_ips))
                terminal_name_str = ';'.join(dict.fromkeys(terminal_names))

            usage_rows.append({
                '交换机': sw_name, '端口': port, 'VLAN': vlan, '端口使用情况': usage,
                '终端设备名称': terminal_name_str, '终端IP': terminal_ip_str, '聚合组': agg_group,
            })
    return usage_rows, suspicious_rows


def _collect_terminal_fields(devices: Sequence[dict]):
    names, ips, vlans = [], [], []
    for d in devices:
        d_type = _clean(d.get('设备类型'))
        n = _clean(d.get('设备名称'))
        ip = _clean(d.get('IP地址'))
        v = _clean(d.get('VLAN'))
        if (d_type == '终端设备' or not _is_empty(ip)
                or not _is_empty(d.get('MAC地址'))):
            if n and n not in ('未知设备', '发现设备', '未知', 'nan', ''):
                names.append(n)
            if ip and ip not in ('未知', 'nan', ''):
                ips.append(ip)
            if v and v not in ('未知', '未找到', 'nan', ''):
                vlans.append(v)
    return names, ips, vlans


def _judge_suspicious(link_status, is_cascade, macs, desc, port_up, snap):
    suspicious = False
    reasons: List[str] = []
    if SUSPICIOUS_ONLY_UP and str(link_status).upper() != 'UP':
        return False, reasons
    if is_cascade:
        return False, reasons
    mac_cnt = len(macs)
    if mac_cnt >= MIN_MAC_SUSPICIOUS:
        suspicious = True
        reasons.append(f'物理口UP且学到{mac_cnt}个MAC，非级联口')
    desc_lower = desc.lower()
    if any(kw in desc_lower for kw in FUZZY_TERMINAL_KEYWORDS):
        if not suspicious:
            suspicious = True
            reasons.append('端口描述匹配终端类关键词')
        else:
            reasons.append('描述含终端类关键词')
    if suspicious and port_up not in snap.lldp_any:
        reasons.append('无LLDP邻居信息')
    return suspicious, reasons


def _decide_usage(is_cascade, terminal_ips, suspicious, macs, agg_group, brief_row, link_status):
    if is_cascade:
        return '级联'
    if terminal_ips:
        return '终端'
    if suspicious and macs:
        return '终端'
    if agg_group:
        return '聚合'
    port_type = _clean(brief_row.get('type')).upper()
    if port_type == 'T':
        return 'Trunk'
    if port_type == 'H':
        return 'Hybrid'
    return '空闲' if str(link_status).upper() == 'UP' else (link_status if link_status else '未知')


# --------------------------------------------------------------------------
# 终端接入表过滤 / 异常分析
# --------------------------------------------------------------------------
def filter_access_terminals(df: Optional[pd.DataFrame], core_name: str = ''):
    """保持原有过滤规则：剔除聚合/虚拟口、级联透传、核心直连等异常行。"""
    if df is None or df.empty:
        empty_df = pd.DataFrame(columns=TERMINAL_TABLE_COLS)
        empty_anom = pd.DataFrame(columns=TERMINAL_TABLE_COLS + ANOMALY_EXTRA_COLS)
        summary = pd.DataFrame([
            {'项目': '原始接入终端总数', '数量': 0},
            {'项目': '过滤后保留终端数', '数量': 0},
            {'项目': '异常/重复终端总数', '数量': 0},
            {'项目': '级联聚合口导致重复', '数量': 0},
            {'项目': '同一物理口挂多终端(傻瓜交换机/多MAC)', '数量': 0},
            {'项目': '其他异常', '数量': 0},
        ])
        return empty_df, empty_anom, summary

    data = df.copy()
    for c in TERMINAL_TABLE_COLS:
        if c in data.columns:
            data[c] = data[c].astype(str).str.strip()

    keep_mask = pd.Series([True] * len(data), index=data.index)
    anomaly_rows: List[dict] = []

    # (1) 同一 IP+MAC 出现在接入口与聚合/虚拟口 → 丢弃聚合那条
    for (ip, mac), g in data.groupby(['IP地址', 'MAC地址']):
        if len(g) <= 1:
            continue
        ports = g['所在端口'].tolist()
        flags = [is_lag_or_virtual(p) for p in ports]
        if not any(flags):
            continue
        for idx, f in zip(g.index.tolist(), flags):
            if not f:
                continue
            keep_mask.loc[idx] = False
            row = data.loc[idx].to_dict()
            row['异常类型'] = 'IP/MAC重复（聚合/级联口穿透）'
            row['原因'] = (f'同一IP/MAC({ip} / {mac})同时出现在接入口和聚合/虚拟口 {ports} 中，'
                           f'聚合/级联口那条不是终端真正直连接口，属于“透传数据”')
            row['处理建议'] = '丢弃（保留物理接入口那条记录）'
            anomaly_rows.append(row)

    # (2) 虚拟/聚合口行直接丢弃
    data_sw_port = data.copy()
    data_sw_port['_port_virtual'] = data_sw_port['所在端口'].map(is_lag_or_virtual)
    for idx in data_sw_port[data_sw_port['_port_virtual']].index.tolist():
        if not keep_mask.loc[idx]:
            continue
        keep_mask.loc[idx] = False
        r = data.loc[idx].to_dict()
        if any(a['IP地址'] == r['IP地址'] and a['MAC地址'] == r['MAC地址'] for a in anomaly_rows):
            continue
        r['异常类型'] = '端口为聚合/虚拟口（非终端直连物理口）'
        r['原因'] = (f"终端所在端口 {r['所在端口']} 是聚合口(BAGG/AGG等)或虚拟口，属于级联/透传，"
                     f'不应出现在“终端接入表(接入口-终端一一对应)”中')
        r['处理建议'] = ('丢弃（如需保留，请确认该端口是否真的直连终端；若下面接了傻瓜交换机/多级级联，'
                         f'请在端口使用表中作为级联口标注）')
        anomaly_rows.append(r)

    # (3) 同一物理口挂多终端
    port_groups = data_sw_port[~data_sw_port['_port_virtual']].groupby(['所在交换机', '所在端口'])
    for (sw, pt), g in port_groups:
        if len(g) <= 1:
            continue
        for idx in g.index.tolist():
            r = data.loc[idx].to_dict()
            r['异常类型'] = '同一物理口挂多终端（傻瓜交换机/多MAC设备）'
            r['原因'] = (f'交换机 {sw} 端口 {pt} 下检测到 {len(g)} 个终端(IP/MAC组合)，'
                         f'物理口一般只接一个终端，多终端通常是下面接了傻瓜交换机/摄像机/多网口服务器等')
            r['处理建议'] = '保留在终端接入表；若数量明显异常(>10)，需人工确认是不是级联口漏标'
            anomaly_rows.append(r)

    # (4) 其他异常
    for idx, r in data.iterrows():
        extra = []
        if core_name and str(r.get('所在交换机', '')) == str(core_name):
            extra.append('终端挂在核心交换机（通常应只保留接入交换机直连）')
        remark = str(r.get('备注', ''))
        if '核心直连' in remark:
            extra.append('备注含"核心直连"')
        if '可能级联' in remark:
            extra.append('备注含"可能级联"（疑似级联口透传数据）')
        for c in ['IP地址', 'MAC地址', '所在交换机', '所在端口']:
            if str(r.get(c, '')) in ['', 'nan', '未知', '未找到', 'None']:
                extra.append(f'字段 {c} 为空/异常')
        if not extra:
            continue
        row = r.to_dict()
        row['异常类型'] = '其他异常'
        row['原因'] = '；'.join(extra)
        row['处理建议'] = '人工核对'
        anomaly_rows.append(row)
        if '核心直连' in remark or '可能级联' in remark:
            keep_mask.loc[idx] = False

    kept_df = data[keep_mask].copy()
    anom_df = pd.DataFrame(anomaly_rows)
    if not anom_df.empty:
        cols = [c for c in list(data.columns) + ANOMALY_EXTRA_COLS if c in anom_df.columns]
        anom_df = anom_df[cols]

    cat_cascade = cat_multi = cat_other = 0
    if not anom_df.empty and '异常类型' in anom_df.columns:
        t = anom_df['异常类型'].astype(str)
        cat_cascade = int(t.str.contains('聚合|级联口|聚合/虚拟口', na=False).sum())
        cat_multi = int(t.str.contains('同一物理口挂多终端', na=False).sum())
        cat_other = int(t.str.contains('其他异常', na=False).sum())

    summary = pd.DataFrame([
        {'项目': '原始接入终端总数', '数量': int(len(data))},
        {'项目': '过滤后保留终端数', '数量': int(len(kept_df))},
        {'项目': '异常/重复终端总数', '数量': int(len(anom_df))},
        {'项目': '级联聚合口导致重复', '数量': cat_cascade},
        {'项目': '同一物理口挂多终端(傻瓜交换机/多MAC)', '数量': cat_multi},
        {'项目': '其他异常（核心直连/字段空等）', '数量': cat_other},
    ])
    return kept_df, anom_df, summary


def build_raw_terminal_rows(port_data: Sequence[dict], core_name: str) -> List[dict]:
    """挑出可进入“终端接入表”的原始行（排除核心、虚拟口）。"""
    rows = []
    for row in port_data:
        sw = _clean(row.get('所在交换机'))
        port = _clean(row.get('所在端口'))
        if not sw or not port or _is_empty(sw) or _is_empty(port):
            continue
        if core_name and sw == core_name:
            continue
        if is_lag_or_virtual(port):
            continue
        d_type = _clean(row.get('设备类型'))
        ip = _clean(row.get('IP地址'))
        mac = _clean(row.get('MAC地址'))
        eligible = d_type == '终端设备' or (
            ip not in ('', '未知', 'nan') or mac not in ('', '未知', 'nan', '未找到', '查询失败'))
        if eligible:
            rows.append(dict(row))
    return rows
