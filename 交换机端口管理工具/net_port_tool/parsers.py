"""Comware 输出纯解析函数。

全部为无副作用的模块级函数，便于单独测试与复用（Web 层不关心解析细节）。
正则统一预编译，避免在密集循环里重复编译。
"""
from __future__ import annotations

import re
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

# ---------- 预编译正则 ----------
RE_MAC = re.compile(r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})')
RE_IP = re.compile(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})')
RE_SWITCH_PORT = re.compile(
    r'^(?:'
    r'(?:GE|XGE|Eth|Ten-G|MGE|FGE|HGE|TGE)\d+/\d+/\d+'
    r'|(?:BAGG|RAGG)\d+'
    r'|Bridge-Aggregation\d+'
    r'|Route-Aggregation\d+'
    r'|Port-channel\d+'
    r')$',
    re.IGNORECASE,
)
RE_PHYSICAL_PORT = re.compile(r'^[A-Za-z]+\d+/\d+/\d+$')
RE_LAG_OR_VIRTUAL = re.compile(
    r'BAGG|RAGG|AGG|BRIDGE-AGGREGATION|VLAN|VLANIF|NULL|INLOOP|LOOP|M-GE|MGE|AUX|CON',
    re.IGNORECASE,
)
RE_PORT_NUMBERS = re.compile(r'\d+')
RE_NON_DIGIT = re.compile(r'\d+')

# 终端类关键词（端口描述命中即视为可疑/终端候选）
FUZZY_TERMINAL_KEYWORDS = (
    'cctv', 'cam', 'ipc', 'nvr', 'server', 'ap', 'wifi',
    '终端', '用户', 'pc', '打印机', '门禁', '报警',
)


def is_switch_port_interface(if_name: str) -> bool:
    """判断是否为交换机端口名（物理口或聚合口）。"""
    if not if_name:
        return False
    return bool(RE_SWITCH_PORT.match(str(if_name).strip()))


def is_physical_port(if_name: str) -> bool:
    """是否为物理口形态（x/y/z）。"""
    return bool(RE_PHYSICAL_PORT.match(str(if_name or '').strip()))


def is_lag_or_virtual(port: str) -> bool:
    """是否为聚合口 / 虚拟口（终端接入表应剔除）。"""
    if not port:
        return False
    s = str(port).upper().strip()
    if not s:
        return False
    if RE_PHYSICAL_PORT.match(s):
        return False
    return bool(RE_LAG_OR_VIRTUAL.search(s))


def port_sort_key(port: str):
    """端口自然排序键：GE1/0/10 排在 GE1/0/2 之后。"""
    s = str(port)
    return (RE_NON_DIGIT.sub('', s), [int(x) for x in RE_PORT_NUMBERS.findall(s)])


def parse_lldp_neighbor_list(output: str) -> List[Dict[str, str]]:
    """解析 display lldp neighbor-information list。"""
    entries: List[Dict[str, str]] = []
    header_found = False
    for line in (output or '').splitlines():
        if 'Local Interface' in line and 'Chassis ID' in line and 'System Name' in line:
            header_found = True
            continue
        if not header_found:
            continue
        stripped = line.strip()
        if not stripped or stripped.startswith('-'):
            continue
        parts = line.split()
        if len(parts) < 4:
            continue
        local_port = parts[0]
        chassis_id = parts[1]
        neighbor_port = parts[2]
        sys_name = ' '.join(parts[3:]).strip()
        if sys_name == '-' or not sys_name:
            sys_name = ''
        mac_match = RE_MAC.search(str(chassis_id).lower())
        entries.append({
            'local_port': local_port,
            'chassis_id': mac_match.group(1) if mac_match else '',
            'neighbor_port': neighbor_port if neighbor_port != '-' else '',
            'neighbor_name': sys_name,
        })
    return entries


def lldp_maps(entries: List[Dict[str, str]], managed_names: set) -> Tuple[Dict[str, dict], set]:
    """把 LLDP 条目整理成 {端口: {neighbor_name, neighbor_port}} 与“已纳管上联口”集合。"""
    any_neighbor: Dict[str, dict] = {}
    managed_uplinks = set()
    for e in entries or []:
        local_port = str(e.get('local_port', '')).strip()
        if not local_port:
            continue
        key = local_port.upper()
        neighbor_name = str(e.get('neighbor_name', '')).strip()
        neighbor_port = str(e.get('neighbor_port', '')).strip()
        any_neighbor[key] = {'neighbor_name': neighbor_name, 'neighbor_port': neighbor_port}
        if neighbor_name and neighbor_name in managed_names:
            managed_uplinks.add(key)
    return any_neighbor, managed_uplinks


def _skip_arp_line(stripped: str) -> bool:
    low = stripped.lower()
    if not stripped:
        return True
    if stripped.startswith('>') or 'display arp' in low:
        return True
    if 'ip address' in low or 'mac address' in low or 'type:' in low:
        return True
    if stripped.startswith('----') or stripped.startswith('===='):
        return True
    return False


def parse_arp_to_maps(output: str) -> Tuple[Dict[str, str], Dict[str, str]]:
    """解析 display arp 输出，返回 (mac->ip, ip->mac)。"""
    mac_to_ip: Dict[str, str] = {}
    ip_to_mac: Dict[str, str] = {}
    if not output:
        return mac_to_ip, ip_to_mac
    for line in output.splitlines():
        stripped = line.strip()
        if _skip_arp_line(stripped):
            continue
        low = stripped.lower()
        ip_m = RE_IP.search(low)
        mac_m = RE_MAC.search(low)
        if ip_m and mac_m:
            ip, mac = ip_m.group(1), mac_m.group(1)
            mac_to_ip.setdefault(mac, ip)
            ip_to_mac.setdefault(ip, mac)
    return mac_to_ip, ip_to_mac


def merge_mac_ip_maps(maps_list) -> Tuple[Dict[str, str], Dict[str, str]]:
    """合并多份 (mac->ip, ip->mac)，先到先得。"""
    m2i: Dict[str, str] = {}
    i2m: Dict[str, str] = {}
    for mm, im in maps_list:
        for k, v in (mm or {}).items():
            if k and v:
                m2i.setdefault(k, v)
        for k, v in (im or {}).items():
            if k and v:
                i2m.setdefault(k, v)
    return m2i, i2m


def parse_vlan_from_arp(output: str, ip: Optional[str] = None) -> str:
    """从 display arp 输出中取某 IP 所在 VLAN。"""
    for line in (output or '').splitlines():
        stripped = line.strip()
        if _skip_arp_line(stripped):
            continue
        if ip and ip not in stripped:
            continue
        parts = stripped.split()
        if len(parts) >= 3 and RE_IP.search(parts[0]):
            return parts[2]
    return '未知'


def extract_port_from_mac_lines(output: str, mac: Optional[str] = None) -> Optional[str]:
    """从 mac-address 输出里挑一个端口，优先物理口。"""
    if not output:
        return None
    best = None
    best_is_physical = False
    for line in output.splitlines():
        line_lower = line.lower()
        if mac and mac not in line_lower:
            continue
        for token in line.split():
            if not is_switch_port_interface(token):
                continue
            is_phy = is_physical_port(token)
            if best is None or (is_phy and not best_is_physical):
                best, best_is_physical = token, is_phy
    return best


def parse_interface_brief(output: str) -> List[dict]:
    """解析 display interface brief（兼容 route/bridge 双段、定宽表头）。"""
    lines = (output or '').splitlines()
    header_indices = [
        i for i, line in enumerate(lines)
        if 'Interface' in line and 'Link' in line
    ]
    if not header_indices:
        return _parse_brief_loose(lines)

    candidates = [
        ('interface', ['interface']),
        ('link', ['link']),
        ('protocol', ['protocol']),
        ('speed', ['speed']),
        ('duplex', ['duplex']),
        ('type', ['type']),
        ('primary_ip', ['primary ip', 'primary-ip', 'ip address']),
        ('vlan', ['pvid', 'vid', 'vlan']),
        ('description', ['description', 'desc']),
    ]

    def build_spans(header_line: str):
        header_lower = str(header_line).lower()
        positions = []
        for key, names in candidates:
            for n in names:
                p = header_lower.find(n)
                if p != -1:
                    positions.append((p, key))
                    break
        positions = sorted(set(positions))
        if not positions:
            return []
        spans = []
        for idx, (pos, key) in enumerate(positions):
            end = positions[idx + 1][0] if idx + 1 < len(positions) else None
            spans.append((key, pos, end))
        return spans

    by_interface: Dict[str, dict] = {}
    for idx, header_idx in enumerate(header_indices):
        spans = build_spans(lines[header_idx])
        if not spans:
            continue
        next_idx = header_indices[idx + 1] if idx + 1 < len(header_indices) else len(lines)
        for raw in lines[header_idx + 1:next_idx]:
            row_line = raw.rstrip('\n')
            stripped = row_line.strip()
            if not stripped or stripped.startswith('>'):
                continue
            if 'Brief information on interfaces' in stripped:
                continue
            first = stripped.split(':', 1)[0].strip().lower()
            if first in ('link', 'protocol', 'speed', 'duplex', 'type'):
                continue
            if re.match(r'^-+\s*$', stripped):
                continue
            parsed = {}
            for key, start, end in spans:
                cell = row_line[start:end].strip() if end is not None else row_line[start:].strip()
                parsed[key] = cell
            if_name = parsed.get('interface', '')
            if not is_switch_port_interface(if_name):
                continue
            row = {
                'interface': if_name,
                'link': parsed.get('link', ''),
                'protocol': parsed.get('protocol', ''),
                'speed': parsed.get('speed', ''),
                'duplex': parsed.get('duplex', ''),
                'type': parsed.get('type', ''),
                'primary_ip': parsed.get('primary_ip', ''),
                'vlan': parsed.get('vlan', ''),
                'description': parsed.get('description', ''),
            }
            existed = by_interface.get(if_name)
            by_interface[if_name] = _merge_brief_rows(existed, row) if existed else row
    return list(by_interface.values())


def _merge_brief_rows(old: dict, new: dict) -> dict:
    merged = dict(old)
    for k, v in new.items():
        if k == 'interface':
            continue
        if str(v).strip() and not str(merged.get(k, '')).strip():
            merged[k] = v
        if k == 'link' and str(v).strip():
            merged[k] = v
    return merged


def _parse_brief_loose(lines: List[str]) -> List[dict]:
    """无标准表头时的兜底解析。"""
    rows = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith('>'):
            continue
        parts = stripped.split()
        if len(parts) < 2:
            continue
        if_name = parts[0]
        if not is_switch_port_interface(if_name):
            continue
        rows.append({
            'interface': if_name,
            'link': parts[1],
            'protocol': parts[2] if len(parts) > 2 else '',
            'speed': '', 'duplex': '', 'type': '', 'primary_ip': '',
            'vlan': '', 'description': '',
        })
    return rows


def parse_mac_table(outputs: List[str]) -> Tuple[Dict[str, List[str]], Dict[str, Dict[str, str]], Dict[str, dict]]:
    """解析 display mac-address [all]，返回：
        port_to_macs   {端口: [mac, ...]}
        port_to_mac_vlan  {端口: {mac: vlan}}
        mac_index      {mac: {'port': 端口, 'vlan': vlan}}  ← 用于零成本反查终端所在端口
    """
    port_to_macs: Dict[str, List[str]] = defaultdict(list)
    port_to_mac_vlan: Dict[str, Dict[str, str]] = defaultdict(dict)
    mac_index: Dict[str, dict] = {}
    seen = set()
    for output in outputs or []:
        if not output:
            continue
        for line in output.splitlines():
            mac_m = RE_MAC.search(line.lower())
            if not mac_m:
                continue
            parts = line.split()
            vlan = parts[1] if len(parts) > 1 else '未知'
            port = None
            for token in parts:
                if is_switch_port_interface(token):
                    port = token
                    break
            if not port:
                continue
            mac = mac_m.group(1)
            key = (port.upper(), mac)
            if key in seen:
                continue
            seen.add(key)
            port_to_macs[port].append(mac)
            vlan_clean = str(vlan).strip()
            if vlan_clean and vlan_clean not in ('未知', ''):
                port_to_mac_vlan[port][mac] = vlan_clean
            # mac_index 只记物理口优先
            prev = mac_index.get(mac)
            if prev is None or (is_physical_port(port) and not is_physical_port(prev['port'])):
                mac_index[mac] = {'port': port, 'vlan': vlan_clean if vlan_clean not in ('未知', '') else ''}
    return dict(port_to_macs), {k: dict(v) for k, v in port_to_mac_vlan.items()}, mac_index


def parse_link_aggregation_verbose(output: str) -> Tuple[Dict[str, str], Dict[str, List[str]]]:
    """解析 display link-aggregation verbose，返回 (成员口->聚合组, 聚合组->成员口列表)。"""
    member_to_agg: Dict[str, str] = {}
    agg_to_members: Dict[str, List[str]] = defaultdict(list)
    current_agg = None
    for line in (output or '').splitlines():
        m_agg = re.search(r'Aggregate Interface:\s*(\S+)', line, re.IGNORECASE)
        if m_agg:
            current_agg = m_agg.group(1).strip()
            continue
        if not current_agg:
            continue
        m_port = re.match(r'^\s*(\S+)\s+([SUI])\s+\d+\s+\d+', line.strip())
        if not m_port:
            continue
        port = re.sub(r'\(.*?\)$', '', m_port.group(1).strip()).strip()
        if not port:
            continue
        member_to_agg[port.upper()] = current_agg
        agg_to_members[current_agg].append(port)
    return member_to_agg, dict(agg_to_members)
