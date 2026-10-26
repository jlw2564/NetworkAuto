# -*- coding: utf-8 -*-
import pandas as pd
import re
import os
import sys
import time
import json
import traceback
from pathlib import Path
from datetime import datetime
from collections import defaultdict, OrderedDict
from netmiko import ConnectHandler
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
SWITCH_TABLE_FILE = BASE_DIR / 'switch_table.xlsx'
OUTPUT_DIR = BASE_DIR / '采集结果'
CONFIG_DIR = OUTPUT_DIR / '交换机配置文件'
LEDGER_FILE = OUTPUT_DIR / '交换机设备台账.xlsx'
HEALTH_FILE = OUTPUT_DIR / '交换机健康状态总表.xlsx'
TOPO_FILE = OUTPUT_DIR / '网络拓扑图.html'
RUN_LOG = OUTPUT_DIR / '采集运行日志.log'
CACHE_FILE = OUTPUT_DIR / '_采集数据缓存.json'

WARN_FILL = PatternFill(start_color='FFFFF2CC', end_color='FFFFF2CC', fill_type='solid')
ERR_FILL = PatternFill(start_color='FFFBCCDC', end_color='FFFBCCDC', fill_type='solid')
OK_FILL = PatternFill(start_color='FFC6EFCE', end_color='FFC6EFCE', fill_type='solid')
HEAD_FILL = PatternFill(start_color='FF4472C4', end_color='FF4472C4', fill_type='solid')
THIN = Side(style='thin', color='FFBFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CPU_WARN, CPU_ERR = 60.0, 80.0
MEM_WARN, MEM_ERR = 60.0, 85.0
TEMP_WARN, TEMP_ERR = 55.0, 70.0
BANDWIDTH_WARN, BANDWIDTH_ERR = 60.0, 85.0
RX_PWR_MIN, RX_PWR_MAX = -18.0, 3.0
TX_PWR_MIN, TX_PWR_MAX = -9.0, 3.0


def log(msg):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{ts}] {msg}'
    print(line)
    try:
        with open(RUN_LOG, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass


def norm_col(x):
    return re.sub(r'\s+', '', str(x)).strip().lower()


def pick(row, *candidates, default=''):
    for k in candidates:
        for col, v in row.items():
            if norm_col(col) == norm_col(k) and str(v) not in ('', 'nan', 'None', 'NaN'):
                return str(v).strip()
    return str(default).strip()


def read_switch_table():
    log(f'读取交换机清单: {SWITCH_TABLE_FILE}')
    df = pd.read_excel(str(SWITCH_TABLE_FILE), dtype=str).fillna('')
    devices = []
    check_col_candidates = ['是否需要检查', '需检查', '需要检查', '是否采集', '需要采集', '是否运行', 'check', 'need_check', 'run']
    check_col_actual = None
    for col in df.columns:
        nc = norm_col(col)
        for cand in check_col_candidates:
            if nc == norm_col(cand):
                check_col_actual = col
                break
        if check_col_actual:
            break
    for idx, (_, row) in enumerate(df.iterrows(), 1):
        row = row.to_dict()
        name = pick(row, '设备名称', '名称', 'hostname', 'switch_name')
        host = pick(row, 'IP地址', '管理IP', '管理地址', 'IP', 'Host', '主机')
        user = pick(row, '用户名', '账号', 'SSH用户名', 'username', 'user')
        pwd = pick(row, '密码', 'password', 'pwd', 'pass', '登录密码')
        is_core = pick(row, '是否核心交换机', '是否核心', '核心', 'role', '角色')
        need_check_raw = ''
        if check_col_actual and check_col_actual in row:
            need_check_raw = str(row[check_col_actual] or '').strip()
        if not name or not host or not user or not pwd:
            continue
        nc_raw = norm_col(need_check_raw)
        if nc_raw in ('否', 'no', 'false', '0', '跳过', '不需要', '不需要检查', 'skip', '无需检查'):
            need_check = False
        elif nc_raw in ('是', 'yes', 'true', '1', '需要', '需要检查', '要', '强制', '必须'):
            need_check = True
        else:
            # 空值 / 未知字符串 → 默认需要检查（兼容老表）
            need_check = True
        devices.append({
            'name': name, 'host': host, 'username': user, 'password': pwd,
            'is_core': norm_col(is_core) in ['是', 'yes', 'true', '1', 'core', '核心'],
            'need_check': need_check,
            'need_check_raw': need_check_raw,
            '_row_index': idx
        })
    log(f'共读取 {len(devices)} 台交换机；是否需要检查列={check_col_actual or "(无，默认全部需要检查)"}')
    _s = sum(1 for d in devices if d['need_check'])
    _sk = len(devices) - _s
    if _sk > 0:
        log(f'  本次将采集 {_s} 台；按列标记跳过 {_sk} 台（如需强制重跑，把对应行“是否需要检查”改为“是”即可）')
    return devices


def judge_need_check(d):
    reasons = []
    if str(d.get('登录状态', '')).strip() not in ('成功', 'OK', 'ok', 'True', 'true', '1'):
        reasons.append(f'登录失败：{d.get("登录状态","失败")}')
    vi = d.get('version_info') or {}
    # 【manuinfo 回显缺失兜底】：如果 _manuinfo_missing_sn=True（display device manuinfo SSH 回显本身没收到完整 SN 段）
    #   且 型号/硬件/软件 3 项都非未知非空，就跳过 SN=未知 的“解析失败”，不算 need_check。
    missing_sn_ok = (bool(vi.get('_manuinfo_missing_sn'))
                     and str(vi.get('型号', '')).strip() not in ('', '未知')
                     and str(vi.get('硬件版本', '')).strip() not in ('', '未知')
                     and str(vi.get('软件版本', '')).strip() not in ('', '未知'))
    for k, friendly in [('型号', '型号'), ('序列号SN', 'SN'), ('硬件版本', '硬件'), ('软件版本', '软件')]:
        v = str(vi.get(k, '')).strip()
        if k == '序列号SN' and missing_sn_ok:
            continue
        if not v or v == '未知':
            reasons.append(f'{friendly}解析失败({v})')
    try:
        cpu = float(d.get('cpu_pct', 0.0) or 0.0)
        if cpu >= CPU_ERR:
            reasons.append(f'CPU过高{cpu:.1f}% ≥ {CPU_ERR:.0f}%')
        elif cpu >= CPU_WARN:
            reasons.append(f'CPU告警{cpu:.1f}% ≥ {CPU_WARN:.0f}%')
    except Exception:
        pass
    try:
        mem = float(d.get('mem_pct', 0.0) or 0.0)
        if mem >= MEM_ERR:
            reasons.append(f'内存过高{mem:.1f}% ≥ {MEM_ERR:.0f}%')
        elif mem >= MEM_WARN:
            reasons.append(f'内存告警{mem:.1f}% ≥ {MEM_WARN:.0f}%')
    except Exception:
        pass
    thr_w = d.get('温度告警阈值') if isinstance(d.get('温度告警阈值'), dict) else {}
    thr_e = d.get('温度错误阈值') if isinstance(d.get('温度错误阈值'), dict) else {}
    for t in (d.get('temps') or []):
        tv = None
        name = ''
        if isinstance(t, dict):
            name = str(t.get('name') or t.get('sensor') or t.get('label') or '')
            try:
                tv = float(t.get('value') or 0.0)
            except Exception:
                tv = None
        elif isinstance(t, (tuple, list)) and len(t) >= 2:
            name = str(t[0] or '')
            try:
                tv = float(t[1] or 0.0)
            except Exception:
                tv = None
        if tv is None:
            continue
        wv = ev = None
        for cand in [name, name.replace(' ', '_'), name.lower(), name.upper()]:
            if not cand:
                continue
            if wv is None and cand in thr_w:
                try:
                    wv = float(thr_w[cand])
                except Exception:
                    wv = None
            if ev is None and cand in thr_e:
                try:
                    ev = float(thr_e[cand])
                except Exception:
                    ev = None
        if wv is None:
            wv = TEMP_WARN
        if ev is None:
            ev = TEMP_ERR
        try:
            if tv >= ev:
                reasons.append(f"温度告警({name or '传感器'}={tv:.1f}℃ ≥ Alarm {ev:.0f}℃)")
            elif tv >= wv:
                reasons.append(f"温度偏高({name or '传感器'}={tv:.1f}℃ ≥ Warning {wv:.0f}℃)")
        except Exception:
            if tv >= TEMP_ERR:
                reasons.append(f"温度过高{name or '传感器'}={tv:.1f}℃")
            elif tv >= TEMP_WARN:
                reasons.append(f"温度偏高{name or '传感器'}={tv:.1f}℃")
    for p in (d.get('ports') or []):
        if not isinstance(p, dict):
            continue
        try:
            in_u = float(p.get('bandwidth_in_util_pct', 0.0) or 0.0)
            out_u = float(p.get('bandwidth_out_util_pct', 0.0) or 0.0)
        except Exception:
            in_u = out_u = 0.0
        name = str(p.get('port') or p.get('interface') or '')
        if in_u >= BANDWIDTH_ERR or out_u >= BANDWIDTH_ERR:
            reasons.append(f'端口{name}带宽过载 入{in_u:.1f}%/出{out_u:.1f}% ≥ {BANDWIDTH_ERR:.0f}%')
        elif in_u >= BANDWIDTH_WARN or out_u >= BANDWIDTH_WARN:
            reasons.append(f'端口{name}带宽偏高 入{in_u:.1f}%/出{out_u:.1f}% ≥ {BANDWIDTH_WARN:.0f}%')
        rx = p.get('rx_power_dbm')
        tx = p.get('tx_power_dbm')
        try:
            if rx is not None and str(rx).strip() != '' and rx != 'N/A':
                rxv = float(rx)
                if rxv < RX_PWR_MIN or rxv > RX_PWR_MAX:
                    reasons.append(f'端口{name}收光异常 RX={rxv:.2f}dBm(正常[{RX_PWR_MIN:.0f},{RX_PWR_MAX:.0f}])')
        except Exception:
            pass
        try:
            if tx is not None and str(tx).strip() != '' and tx != 'N/A':
                txv = float(tx)
                if txv < TX_PWR_MIN or txv > TX_PWR_MAX:
                    reasons.append(f'端口{name}发光异常 TX={txv:.2f}dBm(正常[{TX_PWR_MIN:.0f},{TX_PWR_MAX:.0f}])')
        except Exception:
            pass
        if p.get('transceiver_status') and str(p['transceiver_status']).strip() not in ('正常', '', 'OK', 'ok'):
            reasons.append(f'端口{name}光模块{p["transceiver_status"]}')
        if p.get('counter_overflow') or str(p.get('errors', '')).strip() == 'Overflow':
            reasons.append(f'端口{name}计数Overflow')
        elif p.get('errors') and str(p['errors']).strip() not in ('', '0', '正常', '无'):
            reasons.append(f'端口{name}计数错误{p["errors"]}')
    for e in (d.get('errors') or []):
        es = str(e).strip()
        if es and es not in ('正常', '', '无'):
            reasons.append(es[:120])
    need = len(reasons) > 0
    return need, reasons


def write_switch_table_check_flag(all_data):
    try:
        from openpyxl import load_workbook
    except Exception as e:
        log(f'[标记列回写] 跳过（openpyxl 不可用）：{e}')
        return
    if not SWITCH_TABLE_FILE.exists():
        log(f'[标记列回写] 跳过：{SWITCH_TABLE_FILE} 不存在')
        return
    backup = SWITCH_TABLE_FILE.with_name(SWITCH_TABLE_FILE.stem + '_bak_' +
                                         time.strftime('%Y%m%d_%H%M%S') + SWITCH_TABLE_FILE.suffix)
    try:
        import shutil
        shutil.copy2(str(SWITCH_TABLE_FILE), str(backup))
    except Exception as e:
        log(f'[标记列回写] 备份失败（继续写原文件）：{e}')
        backup = None
    try:
        wb = load_workbook(str(SWITCH_TABLE_FILE))
    except Exception as e:
        log(f'[标记列回写] 打开 xlsx 失败：{e}')
        return
    ws = wb.active
    max_col = ws.max_column or 1
    headers = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=1, column=c).value
        if v is None:
            continue
        headers[norm_col(str(v))] = c
    ip_cols = [headers[norm_col(x)] for x in ['IP地址', '管理IP', '管理地址', 'IP', 'Host', '主机'] if norm_col(x) in headers]
    name_cols = [headers[norm_col(x)] for x in ['设备名称', '名称', 'hostname', 'switch_name'] if norm_col(x) in headers]
    check_col = None
    for cand in ['是否需要检查', '需检查', '需要检查', '是否采集', '需要采集', '是否运行', 'check', 'need_check', 'run']:
        if norm_col(cand) in headers:
            check_col = headers[norm_col(cand)]
            break
    if check_col is None:
        check_col = max_col + 1
        ws.cell(row=1, column=check_col, value='是否需要检查')
    data_by_host = {}
    data_by_name = {}
    for d in all_data:
        h = str(d.get('host', '')).strip()
        n = str(d.get('name', '')).strip()
        need, reasons = judge_need_check(d)
        val = '是' if need else '否'
        entry = {'val': val, 'need': need, 'reasons': reasons, 'name': n, 'host': h}
        if h:
            data_by_host[h] = entry
        if n:
            data_by_name[n] = entry
    written = 0
    matched_rows = 0
    for r in range(2, (ws.max_row or 1) + 1):
        host_vals = [str(ws.cell(row=r, column=c).value or '').strip() for c in ip_cols] if ip_cols else []
        name_vals = [str(ws.cell(row=r, column=c).value or '').strip() for c in name_cols] if name_cols else []
        entry = None
        for hv in host_vals:
            if hv and hv in data_by_host:
                entry = data_by_host[hv]
                break
        if not entry:
            for nv in name_vals:
                if nv and nv in data_by_name:
                    entry = data_by_name[nv]
                    break
        if not entry:
            continue
        matched_rows += 1
        existing = str(ws.cell(row=r, column=check_col).value or '').strip()
        ws.cell(row=r, column=check_col, value=entry['val'])
        if existing != entry['val']:
            written += 1
            note_line = '；'.join(entry['reasons'][:3])
            if len(entry['reasons']) > 3:
                note_line += f'；另有{len(entry["reasons"])-3}项异常…'
            if entry['need'] and note_line:
                try:
                    comment_cell = ws.cell(row=r, column=check_col)
                    from openpyxl.comments import Comment
                    old_comment = comment_cell.comment.text if comment_cell.comment else ''
                    new_text = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {note_line}"
                    if old_comment:
                        new_text = old_comment.rstrip() + "\n" + new_text
                    comment_cell.comment = Comment(new_text[:1000], 'switch_monitor')
                except Exception:
                    pass
            log(f'  回写：{entry["name"]}({entry["host"]}) → 是否需要检查 = {entry["val"]}'
                + (f'（原因：{note_line}）' if entry['need'] and note_line else ''))
    try:
        wb.save(str(SWITCH_TABLE_FILE))
    except Exception as e:
        log(f'[标记列回写] 保存 xlsx 失败（可能文件正被 Excel 打开，请关闭后下次运行自动回写）：{e}')
        if backup:
            try:
                import shutil
                shutil.copy2(str(backup), str(SWITCH_TABLE_FILE))
                log(f'[标记列回写] 已从备份恢复原文件')
            except Exception:
                pass
        return
    log(f'[标记列回写] 完成：匹配 {matched_rows} 行，变更 {written} 行；备份文件：{backup.name if backup else "(未备份)"})')
    if backup:
        log(f'  备份路径：{backup}（确认无误后可手动删除）')


def dump_all_data_cache(all_data):
    try:
        import json
        safe = []
        for d in all_data:
            if not isinstance(d, dict):
                continue
            item = {}
            for k, v in d.items():
                if k in ('config_txt', '_raw_version', '_raw_manuinfo'):
                    # 超大字段不存缓存，避免 JSON 上百 MB；配置 txt 文件在磁盘上已保留
                    continue
                try:
                    json.dumps({k: v}, ensure_ascii=False)
                    item[k] = v
                except Exception:
                    try:
                        item[k] = str(v)
                    except Exception:
                        continue
            safe.append(item)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        with open(str(CACHE_FILE), 'w', encoding='utf-8') as f:
            json.dump(safe, f, ensure_ascii=False, indent=2, default=str)
        log(f'[缓存] 已写入采集数据缓存（{len(safe)} 台）：{CACHE_FILE}（下次单台/过滤模式会自动合并，不会再覆盖旧数据）')
    except Exception as e:
        log(f'[缓存] 写入缓存失败（忽略，不影响本次输出）：{e}')


def load_all_data_cache():
    out = {}
    if not CACHE_FILE.exists():
        return out
    try:
        import json
        with open(str(CACHE_FILE), 'r', encoding='utf-8') as f:
            arr = json.load(f)
        for d in arr:
            if not isinstance(d, dict):
                continue
            h = str(d.get('host', '') or '').strip()
            if h:
                out[h] = d
        log(f'[缓存] 已加载采集数据缓存（{len(out)} 台），将作为旧数据合并基础。')
        return out
    except Exception as e:
        log(f'[缓存] 读取缓存失败（将忽略缓存，尝试从 Excel 还原）：{e}')
        return {}


def _map_col(columns):
    m = {}
    for col in columns:
        nc = norm_col(str(col))
        m[nc] = col
    return m


def load_health_from_excel():
    out = {}
    if not (LEDGER_FILE.exists() and HEALTH_FILE.exists()):
        return out
    try:
        import pandas as _pd
        log('[合并兜底] 缓存不存在/加载失败，将从已有 Excel（台账+健康总表 4 Sheet）尽力还原旧数据。')
        # 1. 台账 Sheet1
        df1 = _pd.read_excel(LEDGER_FILE, sheet_name=0, dtype=str).fillna('')
        col_map1 = _map_col(df1.columns)
        for _, row in df1.iterrows():
            h = str(row.get(col_map1.get('管理ip', '管理IP'), '')).strip() or str(
                row.get(col_map1.get('ip地址', 'IP地址'), '')).strip()
            if not h:
                continue
            is_core_val = str(row.get(col_map1.get('是否核心交换机', '是否核心交换机'), '')).strip()
            out[h] = {
                'name': str(row.get(col_map1.get('设备名称', '设备名称'), '')).strip(),
                'host': h,
                'is_core': ('核心' in is_core_val) or ('是' in is_core_val) or ('core' in is_core_val.lower()),
                '登录状态': str(row.get(col_map1.get('登录状态', '登录状态'), '失败')).strip() or '失败',
                '采集时间': str(row.get(col_map1.get('采集时间', '采集时间'), '')).strip(),
                'version_info': {
                    '型号': str(row.get(col_map1.get('型号', '型号'), '')).strip() or '未知',
                    '序列号SN': str(row.get(col_map1.get('序列号sn', '序列号SN'), '')).strip() or '未知',
                    '硬件版本': str(row.get(col_map1.get('硬件版本', '硬件版本'), '')).strip() or '未知',
                    '软件版本': str(row.get(col_map1.get('软件版本', '软件版本'), '')).strip() or '未知',
                },
                '备注': str(row.get(col_map1.get('备注', '备注'), '')).strip(),
                'cpu_pct': 0.0, 'mem_pct': 0.0, 'temps': [], 'ports': [], 'lldp': [], 'errors': [],
                'counters_method': '', 'counters_sample_sec': 30, 'config_txt': ''
            }
        # 2. 健康 Sheet1 设备健康概览 → 回填 CPU/MEM/温度阈值
        try:
            df2 = _pd.read_excel(HEALTH_FILE, sheet_name=0, dtype=str).fillna('')
        except Exception:
            df2 = None
        if df2 is not None and not df2.empty:
            col_map2 = _map_col(df2.columns)
            for _, row in df2.iterrows():
                h = str(row.get(col_map2.get('管理ip', '管理IP'), '')).strip() or str(
                    row.get(col_map2.get('ip地址', 'IP地址'), '')).strip()
                if not h or h not in out:
                    continue
                try:
                    out[h]['cpu_pct'] = float(row.get(col_map2.get('cpu使用率(%)', 'CPU使用率(%)'), 0.0) or 0.0)
                except Exception:
                    pass
                try:
                    out[h]['mem_pct'] = float(row.get(col_map2.get('内存使用率(%)', '内存使用率(%)'), 0.0) or 0.0)
                except Exception:
                    pass
                # 温度列: 格式 Slot1_hotspot1_MAC(℃)，我们尽力还原成 list[tuple(name, value)]
                temps = []
                thr_w = {}
                thr_e = {}
                for raw_col in df2.columns:
                    col = str(raw_col)
                    if '℃' in col and '温度' in col:
                        sensor = col.split('温度', 1)[0].strip() or raw_col
                        val = str(row.get(raw_col, '')).strip()
                        try:
                            fv = float(val)
                            temps.append((sensor, fv))
                        except Exception:
                            pass
                out[h]['temps'] = temps
                out[h].setdefault('温度告警阈值', thr_w)
                out[h].setdefault('温度错误阈值', thr_e)
        # 3. 健康 Sheet2 端口带宽与光功率 → 按 host 分组还原 ports 列表
        try:
            df3 = _pd.read_excel(HEALTH_FILE, sheet_name=1, dtype=str).fillna('')
        except Exception:
            df3 = None
        port_by_host = {}
        if df3 is not None and not df3.empty:
            col_map3 = _map_col(df3.columns)
            for _, row in df3.iterrows():
                h = str(row.get(col_map3.get('管理ip', '管理IP'), '')).strip() or str(
                    row.get(col_map3.get('ip地址', 'IP地址'), '')).strip()
                if not h:
                    continue
                p = {}
                col_pairs = [
                    ('port', ('端口', '端口名', '端口名称', 'interface', 'port')),
                    ('link', ('链路状态', '链路', 'link')),
                    ('speed', ('速率', '协商速率', 'speed')),
                    ('duplex', ('双工', 'duplex')),
                    ('pvid', ('pvid', '所属vlan', 'vlan')),
                    ('description', ('描述', 'description', '端口描述')),
                    ('bandwidth_in_util_pct', ('入带宽利用率(%)', '入利用率(%)', '入带宽', '入方向利用率(%)')),
                    ('bandwidth_out_util_pct', ('出带宽利用率(%)', '出利用率(%)', '出带宽', '出方向利用率(%)')),
                    ('peak_in_mbps', ('入峰值(mbps)', '入峰值')),
                    ('peak_out_mbps', ('出峰值(mbps)', '出峰值')),
                    ('negotiated_speed_mbps', ('真实协商速率(mbps)', '真实速率', '协商速率')),
                    ('rx_power_dbm', ('收光(dbm)', '收光功率(dbm)', 'rxpower(dbm)', 'rx')),
                    ('tx_power_dbm', ('发光(dbm)', '发光功率(dbm)', 'txpower(dbm)', 'tx')),
                    ('transceiver_status', ('光模块状态', '光模块诊断', 'transceiver')),
                    ('errors', ('错误数', 'err', '计数错误')),
                    ('counter_overflow', ('计数溢出', 'overflow')),
                    ('counters_method', ('采集方法', '带宽方法')),
                ]
                for key, cands in col_pairs:
                    val = ''
                    for c in cands:
                        if norm_col(c) in col_map3:
                            val = str(row.get(col_map3[norm_col(c)], '')).strip()
                            if val != '':
                                break
                    if val != '':
                        # 尽力转成 float/int/bool
                        if 'util' in key or 'mbps' in key or 'power' in key:
                            try:
                                val = float(val)
                            except Exception:
                                pass
                        elif key == 'counter_overflow':
                            val = (val.lower() in ('是', 'true', '1', 'yes', '有', '溢出'))
                        p[key] = val
                if p:
                    port_by_host.setdefault(h, []).append(p)
        for h, plist in port_by_host.items():
            if h in out:
                out[h]['ports'] = plist
        # 4. 健康 Sheet3 异常汇总 → 按 host 分组还原 errors 列表
        try:
            df4 = _pd.read_excel(HEALTH_FILE, sheet_name=2, dtype=str).fillna('')
        except Exception:
            df4 = None
        if df4 is not None and not df4.empty:
            col_map4 = _map_col(df4.columns)
            host_col = None
            for c in ('管理ip', 'ip地址', '管理地址', '主机', 'host'):
                if norm_col(c) in col_map4:
                    host_col = col_map4[norm_col(c)]
                    break
            desc_col = None
            for c in ('异常描述', '异常信息', '描述', '详情', '错误信息', 'error'):
                if norm_col(c) in col_map4:
                    desc_col = col_map4[norm_col(c)]
                    break
            if host_col and desc_col:
                for _, row in df4.iterrows():
                    h = str(row.get(host_col, '')).strip()
                    if not h or h not in out:
                        continue
                    msg = str(row.get(desc_col, '')).strip()
                    if msg:
                        out[h].setdefault('errors', []).append(msg)
        log(f'[合并兜底] Excel 还原完成，共 {len(out)} 台设备基础数据，其中 {len(port_by_host)} 台已还原端口明细。')
        return out
    except Exception as e:
        log(f'[合并兜底] Excel 还原失败（将忽略旧数据，全量重写）：{e}')
        return {}


def build_conn(d):
    return {
        'device_type': 'hp_comware',
        'host': d['host'], 'username': d['username'], 'password': d['password'],
        'timeout': 30, 'global_delay_factor': 1.5,
        'conn_timeout': 15, 'banner_timeout': 20, 'auth_timeout': 20,
        'fast_cli': False
    }


def safe_send(ssh, cmd, delay=1, timeout=60, read_timeout=None, max_attempts=1):
    last = ''
    for i in range(max(1, max_attempts)):
        try:
            r = ssh.send_command(cmd, delay_factor=delay, expect_string=r'>',
                                 read_timeout=read_timeout or timeout) or ''
            if len(r) > len(last):
                last = r
            if any(tag in r for tag in ['DEVICE_SERIAL_NUMBER', 'DEVICE_NAME', 'BOARD TYPE', 'H3C Comware Software']):
                return r
        except Exception as e:
            last = f'ERROR: {e}'
    return last


def log_cmd(sw, cmd, lines=0):
    s = f'[采集] {sw} | {cmd}'
    if lines:
        s += f' ({lines} lines)'
    log(s)


def diagnose_template(ver_out, manuinfo_out, version_info):
    """返回 (template_name, template_desc, retry_manuinfo_times)
    retry_manuinfo_times ∈ [0,1,2,3]，0 表示不重试，大于0表示 collect_one 里再跑多少次 dis device manuinfo 补采
    """
    vi = version_info or {}
    ver_str = str(ver_out or '')
    manu_str = str(manuinfo_out or '')
    v_model = str(vi.get('型号', '')).strip()
    v_hw = str(vi.get('硬件版本', '')).strip()
    v_sw = str(vi.get('软件版本', '')).strip()
    v_sn = str(vi.get('序列号SN', '')).strip()

    tri_ok = (v_model and v_model != '未知') and (v_hw and v_hw != '未知') and (v_sw and v_sw != '未知')
    sn_ok = v_sn and v_sn != '未知'
    manu_has_sn = 'DEVICE_SERIAL_NUMBER' in manu_str
    multi_slot = (bool(re.search(r'Chassis\s+[\w\s\-]+self', manu_str, re.I))
                  or len(re.findall(r'Slot\s+\d+(?:\s+CPU\s+\d+)?\s*[:：]?', manu_str, re.I)) >= 2)
    # ver_truncated：只有 version 里 真·首字节被截（isplay/erminal/el 6）才算 C 模板
    # 正常 CLI 回显第一行就是 "<SW-XYZ>display version" 不算截流（那是 netmiko 正常把命令和prompt echo 进输出里）
    ver_truncated = any(re.search(pat, ver_str, re.MULTILINE) for pat in [
        r'^\s*isplay\s+version\s*$',
        r'^\s*erminal\s+logging\s+level\s+6\s*$',
        r'^\s*el\s+6\s*$',
        r'^\s*isplay\s+\S',
        r'^\s*erminal\s+\S',
    ])

    if multi_slot:
        return ('D-多Slot框式', 'manuinfo 分段结构(Chassis self/Slot N)，按板卡/Slot聚合', 0)
    if sn_ok or manu_has_sn:
        return ('A-完整', 'manuinfo+version 完整返回 4 字段', 0)
    if ver_truncated and tri_ok:
        return ('C-版本截流', 'version段首字节被截/混terminal logging level 6，已通过预处理成功解析3字段，若manuinfo也缺失将一并重试补采SN', max(3, 2 if sn_ok else 3))
    if tri_ok and not sn_ok:
        if bool(vi.get('_manuinfo_missing_sn')):
            return ('B-manuinfo截流(重试)', '型号/HW/SW 3字段正确，manuinfo回显未收到 SN 段，将自动补采最多3次 dis device manuinfo', 3)
        return ('B-manuinfo截流', '型号/HW/SW 3字段正确，SN因回显缺失未知，已跳过SN解析失败', 0)
    if not tri_ok:
        return ('Z-未知', '版本/manuinfo信号不足以解析全部字段，请把原始输出贴回补正则', 1)
    return ('Z-未知', '未命中以上模板，请贴原始输出', 0)


PORT_SHORT_MAP = [
    ('GigabitEthernet', 'GE'),
    ('Ten-GigabitEthernet', 'XGE'),
    ('FortyGigE', 'FGE'),
    ('HundredGigE', 'HGE'),
    ('Twenty-FiveGigE', 'TGE'),
    ('Bridge-Aggregation', 'BAGG'),
    ('Route-Aggregation', 'RAGG'),
]


def port_to_short(name):
    name = str(name).strip()
    for long, short in PORT_SHORT_MAP:
        if name.lower().startswith(long.lower()):
            return short + name[len(long):]
    return name


def port_to_long(name):
    name = str(name).strip()
    for long, short in PORT_SHORT_MAP:
        if name.lower().startswith(short.lower() + '1' if False else short.lower()):
            m = re.match(r'^(' + short + r')(\d.*)$', name, re.I)
            if m:
                return long + m.group(2)
    return name


def _split_columns_by_header(header_line, wanted_keys):
    """按表头列名定位每列的起止字符位置，返回 {key:(start,end)}
    """
    mapping = {}
    for k in wanted_keys:
        idx = header_line.lower().find(k.lower())
        if idx >= 0:
            mapping[k] = [idx, None]
    positions = sorted(mapping.items(), key=lambda kv: kv[1][0])
    for i, (k, rng) in enumerate(positions):
        if i + 1 < len(positions):
            rng[1] = positions[i + 1][1][0]
        else:
            rng[1] = len(header_line)
    return {k: tuple(rng) for k, rng in positions}


def _cut(line, rng):
    s, e = rng
    return line[s:e].strip()


def parse_device_manuinfo(out):
    info = {}
    slots = []
    cur_block = None
    cur_fields = {}

    def flush_block():
        nonlocal cur_block, cur_fields
        if not cur_block or not cur_fields:
            cur_block = None
            cur_fields = {}
            return
        if cur_block.lower().startswith('chassis') and 'self' in cur_block.lower():
            for k, v in cur_fields.items():
                info.setdefault(k, v)
            info['chassis_fields'] = dict(cur_fields)
        elif cur_block.lower().startswith('slot'):
            d = dict(cur_fields)
            d['block'] = cur_block
            slots.append(d)
            for k, v in cur_fields.items():
                info.setdefault(k, v)
        else:
            pass
        cur_block = None
        cur_fields = {}

    CN_ALIASES = {
        '设备型号': 'DEVICE_NAME',
        '型号名称': 'DEVICE_NAME',
        '设备序列号': 'DEVICE_SERIAL_NUMBER',
        '序列号': 'DEVICE_SERIAL_NUMBER',
        '设备MAC地址': 'MAC_ADDRESS',
        'MAC地址': 'MAC_ADDRESS',
        '硬件版本': 'HARDWARE_VERSION',
        '软件版本': 'SOFTWARE_VERSION',
        '生产日期': 'MANUFACTURING_DATE',
        '制造日期': 'MANUFACTURING_DATE',
        '供应商名称': 'VENDOR_NAME',
    }

    for raw in out.splitlines():
        line = raw.strip()
        if not line:
            continue
        block_hit = False
        m = re.match(r'^(Chassis\s+[\w\s\-]+|Slot\s+\d+(?:\s+CPU\s+\d+)?|Fan\s+\d+|Power\s+\d+)\s*[:：]?$', line, re.I)
        if m:
            flush_block()
            cur_block = m.group(1).strip()
            cur_fields = {}
            block_hit = True
        if block_hit:
            continue
        if line.lower().startswith('the operation is not supported') or '% unrecognized' in line.lower():
            continue
        mkv = re.match(r'([A-Za-z0-9_\u4e00-\u9fa5]+)\s*[:：]\s*(.*)', line)
        if not mkv:
            continue
        raw_k, raw_v = mkv.group(1).strip(), mkv.group(2).strip()
        k = CN_ALIASES.get(raw_k, raw_k)
        if cur_block is None:
            info.setdefault(k, raw_v)
        else:
            cur_fields[k] = raw_v
    flush_block()
    if slots:
        info['slots'] = slots
    has_chassis = bool(info.get('chassis_fields'))
    if not has_chassis and slots:
        first = slots[0]
        for gk, lk in [('DEVICE_NAME', 'DEVICE_NAME'),
                       ('DEVICE_SERIAL_NUMBER', 'DEVICE_SERIAL_NUMBER'),
                       ('MAC_ADDRESS', 'MAC_ADDRESS'),
                       ('MANUFACTURING_DATE', 'MANUFACTURING_DATE'),
                       ('VENDOR_NAME', 'VENDOR_NAME'),
                       ('PRODUCT_ID', 'PRODUCT_ID'),
                       ('HARDWARE_VERSION', 'HARDWARE_VERSION'),
                       ('SOFTWARE_VERSION', 'SOFTWARE_VERSION')]:
            if not info.get(gk) and first.get(lk):
                info[gk] = first[lk]
        info.setdefault('chassis_fields', {})
        for k, v in first.items():
            info['chassis_fields'].setdefault(k, v)
    if not info.get('DEVICE_SERIAL_NUMBER') and info.get('chassis_fields', {}).get('DEVICE_SERIAL_NUMBER'):
        info['DEVICE_SERIAL_NUMBER'] = info['chassis_fields']['DEVICE_SERIAL_NUMBER']
    if not info.get('DEVICE_NAME') and info.get('chassis_fields', {}).get('DEVICE_NAME'):
        info['DEVICE_NAME'] = info['chassis_fields']['DEVICE_NAME']
    # 【兜底】manuinfo 命令无返回/被截断（回显混着 version 段）时，从 version 文本抓 BOARD TYPE/PCB/Release 填型号/硬件
    if not info.get('DEVICE_NAME'):
        m = re.search(r'BOARD\s+TYPE:\s*([\w\-\.\+\/]+)', out, re.I)
        if m:
            info['DEVICE_NAME'] = m.group(1).strip()
            info.setdefault('chassis_fields', {})['DEVICE_NAME'] = info['DEVICE_NAME']
    if not info.get('HARDWARE_VERSION'):
        m = re.search(r'PCB\s+(?:\d+\s+)?Version:\s*([\w\.\-]+)', out, re.I)
        if m:
            info['HARDWARE_VERSION'] = m.group(1).strip()
            info.setdefault('chassis_fields', {})['HARDWARE_VERSION'] = info['HARDWARE_VERSION']
    if not info.get('SOFTWARE_VERSION'):
        m = re.search(r'H3C\s+Comware\s+Software,\s+Version\s+[\w\.\,]+\s*Release\s+([\w\.]+)', out, re.I)
        if m:
            info['SOFTWARE_VERSION'] = m.group(1).strip()
            info.setdefault('chassis_fields', {})['SOFTWARE_VERSION'] = info['SOFTWARE_VERSION']
    if not info.get('DEVICE_SERIAL_NUMBER'):
        # 明确标记缺失，避免 parse_version 误以为 manuinfo 里有 SN 而混用
        info['_manuinfo_missing_sn'] = True
        info.setdefault('备注', '')
        if 'manuinfo无输出' not in info['备注']:
            info['备注'] = (info['备注'] + '；' if info['备注'] else '') + 'display device manuinfo 未返回完整内容，SN 需补采'
    return info


def parse_version(out, manuinfo=None):
    model, sn, hw_ver, sw_ver = '未知', '未知', '未知', '未知'
    # 【预处理】：去掉 netmiko 残留回显（terminal logging level 6 / 前一行开头残字），以及 <hostname>display version prompt 前缀垃圾
    cleaned_lines = []
    prompt_prefix_pat = re.compile(r'^<[\w\-\.]+>\s*')
    for raw in out.splitlines():
        line = raw.strip()
        if not line:
            continue
        # 去 prompt 前缀 <XXX>
        if prompt_prefix_pat.match(line):
            line = prompt_prefix_pat.sub('', line).strip()
        # 去 terminal logging level 6 残留
        if line.lower().startswith('terminal logging level'):
            continue
        if not line:
            continue
        cleaned_lines.append(line)
    cleaned_out = '\n'.join(cleaned_lines)
    has_real_signature = False
    for pat in [r'H3C\s+[\w\-\+\.\/]+\s+uptime is',
                r'BOARD\s+TYPE:\s*[\w\-\.\+\/]+',
                r'Release\s+Version:\s*H3C\s+',
                r'H3C\s+Comware\s+Software,\s+Version',
                r'Uptime for\s+[\w\-\.\/\+]+',
                r'Hardware\s+Version\s+is\s+',
                r'Boot\s+image\s+version:']:
        if re.search(pat, cleaned_out, re.I):
            has_real_signature = True
            break
    # 型号正则（允许行首任意残字/空格，比如 "isplay version\nH3C ... uptime is" 或 "erminal logging level 6" 去掉后命中）
    m = re.search(r'H3C\s+([\w\-\+\.\/]+)\s+uptime\s+is', cleaned_out, re.I | re.M)
    if m and m.group(1):
        candidate = m.group(1).strip()
        if candidate.lower() not in ('comware', 'switch', 'router', '', 'h3c'):
            model = candidate
    if model == '未知' and has_real_signature:
        m = re.search(r'BOARD\s+TYPE:\s*([\w\-\.\+\/]+)', cleaned_out, re.I)
        if m:
            model = m.group(1).strip()
    if model == '未知':
        for pat in [r'Uptime for\s+([\w\-\.\/\+]+)\s+is',
                    r'Switch\s+([\w\-\.\/\+]+)\s+is\s+up',
                    r'Hardware\s+Version\s+is\s+([\w\-\.]+)']:
            mm = re.search(pat, cleaned_out, re.I)
            if mm and mm.group(1):
                model = mm.group(1).strip()
                break
    if model == '未知':
        m = re.search(r'板卡型号\s*[:：]\s*([\w\-\.\+\/]+)', cleaned_out, re.I)
        if m:
            model = m.group(1).strip()
    # Release Version: H3C S5048PV6-EI-PWR-3507P28（贪婪取型号，避免 non-greedy 把型号截断）
    m = re.search(
        r'Release\s+Version:\s*H3C\s+([\w\-\+\/\.]+?)-(\d{2,6}[A-Z]?\d{0,4}[A-Z0-9]*)\s*$', cleaned_out, re.I | re.M)
    if not m:
        m = re.search(r'Release\s+Version:\s*H3C\s+([\w\-\+\/\.]+)-(\d{2,6}[A-Z]?\d{0,4}[A-Z0-9]*)', cleaned_out, re.I)
    if m:
        if model == '未知':
            model = m.group(1).strip()
        sw_ver = m.group(2).strip()
    if sw_ver == '未知':
        m = re.search(r'H3C\s+Comware\s+Software,\s+Version\s+[\w\.\,]+\s*Release\s+([\w\.\-]+)', cleaned_out, re.I)
        if m:
            sw_ver = m.group(1).strip()
    if sw_ver == '未知':
        m = re.search(r'Release\s+Version:\s*([\w\-\.\+]+)', cleaned_out, re.I)
        if m:
            sw_ver = m.group(1).strip()
    if sw_ver == '未知':
        m = re.search(r'Boot image version:\s*[\w\.\,]+\s*Release\s+([\w\.]+)', cleaned_out, re.I)
        if m:
            sw_ver = m.group(1).strip()
    if sw_ver == '未知':
        m = re.search(r'System image version:\s*[\w\.\,]+\s*Release\s+([\w\.]+)', cleaned_out, re.I)
        if m:
            sw_ver = m.group(1).strip()
    m = re.search(r'软件版本\s*[:：]\s*([\w\-\.\+Release\s]+)', cleaned_out, re.I)
    if m and sw_ver == '未知':
        sw_ver = m.group(1).strip()
    m = re.search(r'PCB\s+(?:\d+\s+)?Version:\s*([\w\.\-]+)', cleaned_out, re.I)
    if m:
        hw_ver = m.group(1).strip()
    m = re.search(r'硬件版本\s*[:：]\s*([\w\.\-]+)', cleaned_out, re.I)
    if m and hw_ver == '未知':
        hw_ver = m.group(1).strip()
    for pat in [r'DEVICE_SERIAL_NUMBER\s*[:：]\s*([A-Za-z0-9\-]+)',
                r'System\s+Serial\s+Number\s*[:：]\s*([A-Za-z0-9\-]+)',
                r'Slot\s+\d+\s+Serial\s+Number\s*[:：]\s*([A-Za-z0-9\-]+)',
                r'设备序列号\s*[:：]\s*([A-Za-z0-9\-]+)',
                r'序列号\s*[:：]\s*([A-Za-z0-9\-]+)']:
        mm = re.search(pat, cleaned_out, re.I)
        if mm:
            sn = mm.group(1).strip()
            break
    if sn == '未知' and manuinfo and not manuinfo.get('_manuinfo_missing_sn') and manuinfo.get('DEVICE_SERIAL_NUMBER'):
        sn = str(manuinfo['DEVICE_SERIAL_NUMBER']).strip()
    if not manuinfo:
        return {'型号': model, '序列号SN': sn, '硬件版本': hw_ver, '软件版本': sw_ver}
    mfg = manuinfo or {}
    chassis_model = None
    if mfg.get('chassis_fields'):
        chassis_model = mfg['chassis_fields'].get('DEVICE_NAME')
    if chassis_model:
        model = chassis_model
    elif mfg.get('DEVICE_NAME') and (model == '未知' or re.match(r'^[A-Z]{2,3}-\d{3,}', model)):
        # hostname 兜底形如 EBG-XXX 就不要覆盖 manuinfo 型号
        if not re.match(r'^[A-Z]{2,3}-[\w\-]+$', model):
            model = mfg['DEVICE_NAME']
    if sn == '未知' and not mfg.get('_manuinfo_missing_sn') and mfg.get('chassis_fields') and mfg['chassis_fields'].get('DEVICE_SERIAL_NUMBER'):
        sn = str(mfg['chassis_fields']['DEVICE_SERIAL_NUMBER']).strip()
    if hw_ver == '未知' and mfg.get('chassis_fields') and mfg['chassis_fields'].get('HARDWARE_VERSION'):
        hw_ver = str(mfg['chassis_fields']['HARDWARE_VERSION']).strip()
    if sw_ver == '未知' and mfg.get('chassis_fields') and mfg['chassis_fields'].get('SOFTWARE_VERSION'):
        sw_ver = str(mfg['chassis_fields']['SOFTWARE_VERSION']).strip()
    out2 = {'型号': model,
            '序列号SN': sn,
            '硬件版本': hw_ver, '软件版本': sw_ver}
    if mfg.get('_manuinfo_missing_sn'):
        out2['_manuinfo_missing_sn'] = True
    if mfg.get('slots'):
        out2['板卡清单'] = [{'槽位': s.get('block', ''), '板卡型号': s.get('DEVICE_NAME', ''), '板卡SN': s.get('DEVICE_SERIAL_NUMBER', ''),
                         '板卡MAC': s.get('MAC_ADDRESS', ''), '生产日期': s.get('MANUFACTURING_DATE', '')} for s in mfg['slots']]
    if mfg.get('chassis_fields'):
        out2['机箱MAC'] = mfg['chassis_fields'].get('MAC_ADDRESS', '')
        out2['生产日期'] = mfg['chassis_fields'].get('MANUFACTURING_DATE', '')
    if mfg.get('备注'):
        if not out2.get('备注'):
            out2['备注'] = mfg['备注']
    return out2


def parse_cpu(out):
    best = 0.0
    seen = False
    m = re.search(r'Statistics\s+interval\s+in\s+seconds?:\s*\d+[\s\S]*?5\s+seconds?:\s*(\d+(?:\.\d+)?)%', out, re.I)
    if m:
        return float(m.group(1))
    m = re.search(r'CPU\s+utilization\s+for\s+five\s+seconds:\s*(\d+(?:\.\d+)?)%', out, re.I)
    if m:
        return float(m.group(1))
    m = re.search(r'Slot\s+\d+\s+CPU\s+\d+\s+CPU\s+usage:[\s\S]*?(\d+(?:\.\d+)?)%', out, re.I)
    if m:
        return float(m.group(1))
    m = re.search(r'CPU\s+Usage.*?(\d+(?:\.\d+)?)\s*%', out, re.I)
    if m:
        return float(m.group(1))
    return best


def parse_memory(out):
    slot_rows = []
    lines = out.splitlines()
    cur_slot = None
    header_cols = None
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(r'Slot\s+(\d+):', line)
        if m:
            cur_slot = int(m.group(1))
        if re.match(r'^\s*Total\s+Used\s+Free\s+', line):
            header_cols = line
            if i + 1 < len(lines) and lines[i+1].strip().startswith('Mem:'):
                vals = lines[i+1].strip().split()
                if len(vals) >= 8:
                    try:
                        total = float(vals[1])
                        used = float(vals[2])
                        free = float(vals[3])
                        free_ratio = float(vals[-1].rstrip('%'))
                        used_ratio = round(100.0 - free_ratio, 2)
                        slot_rows.append((cur_slot, used_ratio, used, total))
                    except Exception:
                        pass
        i += 1
    if slot_rows:
        used_ratio = max(r[1] for r in slot_rows)
        used = sum(r[2] for r in slot_rows)
        total = sum(r[3] for r in slot_rows)
        return used_ratio, used, total
    m = re.search(r'Memory\s+Using\s+(\d+(?:\.\d+)?)\s*%', out, re.I)
    if m:
        return float(m.group(1)), 0.0, 0.0
    used = 0.0
    total = 0.0
    for line in lines:
        line = line.strip()
        if not line:
            continue
        m_num = list(map(float, re.findall(r'(\d+(?:\.\d+)?)', line)))
        if len(m_num) >= 3 and any(k in line.lower() for k in ['mem:', 'memory', '内存']):
            total, used = m_num[0], m_num[1]
            break
    pct = round((used / total * 100.0), 2) if total > 0 else 0.0
    return pct, used, total


def parse_temperature(out):
    """支持 V7 标准环境表:
    Slot  Sensor      Name   Temperature  Lower  Warning  Alarm  Shutdown
    或者 (Sensor列合并)
    Slot  Sensor    Temperature  Lower  Warning  Alarm  Shutdown
    """
    res = []
    lines = out.splitlines()
    header_idx = None
    cols_map = None
    for i, line in enumerate(lines):
        if 'Temperature' in line and 'Warning' in line and 'Alarm' in line:
            header_idx = i
            cols_map = _split_columns_by_header(line,
                                                ['Slot', 'Sensor', 'Name', 'Temperature', 'Lower', 'Warning', 'Alarm', 'Shutdown'])
            break
    if header_idx is None or cols_map is None:
        for line in lines:
            s = line.strip()
            if not s:
                continue
            m = re.findall(r'(\d+(?:\.\d+)?)', s)
            if any(k in s.lower() for k in ['sensor', '温度', 'temp', 'cpu']) and not any(k in s.lower() for k in ['warning', 'alarm', 'shutdown', 'lower']):
                for n in m:
                    try:
                        v = float(n)
                        if 10 <= v <= 130:
                            name = s.split()[0] if s.split() else 'Sensor'
                            res.append((name, v))
                            break
                    except ValueError:
                        pass
        return res
    has_name = 'Name' in cols_map
    for line in lines[header_idx + 1:]:
        if len(line) < 4:
            continue
        if line.strip().startswith('---') or line.strip().startswith('<') or not line.strip():
            continue
        parts = re.findall(r'\d+\.?\d*|NA|N/A', line)
        if not parts:
            continue
        try:
            temp_cell = _cut(line, cols_map['Temperature'])
            m = re.search(r'\-?\d+(?:\.\d+)?', temp_cell)
            if not m:
                continue
            t = float(m.group(0))
            if t < -20 or t > 200:
                continue
            slot = _cut(line, cols_map.get('Slot', (0, 8))) if 'Slot' in cols_map else ''
            sensor = _cut(line, cols_map.get('Sensor', (0, 8))) if 'Sensor' in cols_map else ''
            name = _cut(line, cols_map.get('Name', (0, 8))) if has_name else ''
            slot_s = re.sub(r'\s+', '', str(slot or '')).strip()
            sensor_clean = re.sub(r'\s+', ' ', str(sensor or '')).strip()
            name_clean = re.sub(r'\s+', ' ', str(name or '')).strip()
            labels_set = set()
            # 1. 单 Slot 只保留 sensor/name（不带 Slot 前缀，便于匹配阈值 dict）
            if sensor_clean and sensor_clean.lower() != 'na':
                labels_set.add(sensor_clean)
            if name_clean and name_clean.lower() != 'na' and name_clean not in labels_set:
                labels_set.add(name_clean)
            # 2. 多 Slot 设备：Slot-Sensor / Slot_Name / SlotSensor 三种组合
            if slot_s:
                base_parts = [slot_s, sensor_clean or name_clean]
                if len([x for x in base_parts if x]) >= 2:
                    labels_set.add('-'.join([x for x in base_parts if x]))
                    labels_set.add('_'.join([x for x in base_parts if x]))
                    labels_set.add(''.join([x for x in base_parts if x]))
            labels = list(labels_set)
            if not labels:
                labels = [(line.split()[0] if line.split() else 'Sensor')]
            # 第一个 label 用最直观的 sensor/name，不要带 Slot 前缀，用户更友好（健康表备注用它）
            primary_label = ''
            for cand in [sensor_clean, name_clean]:
                if cand and cand.lower() != 'na':
                    primary_label = cand
                    break
            if not primary_label and labels:
                primary_label = labels[0]
            if not primary_label:
                primary_label = (line.split()[0] if line.split() else 'Sensor')
            # 剔除明显是阈值的行 (Warning/Alarm/Shutdown列对应值不等于Temperature)
            try:
                warn_cell = _cut(line, cols_map['Warning'])
                wm = re.search(r'\-?\d+(?:\.\d+)?', warn_cell)
                if wm and abs(float(wm.group(0)) - t) < 0.01:
                    continue
                alarm_cell = _cut(line, cols_map['Alarm'])
                am = re.search(r'\-?\d+(?:\.\d+)?', alarm_cell)
                if am and abs(float(am.group(0)) - t) < 0.01:
                    continue
                low_cell = _cut(line, cols_map['Lower'])
                lm = re.search(r'\-?\d+(?:\.\d+)?', low_cell)
                if lm and abs(float(lm.group(0)) - t) < 0.01:
                    continue
            except Exception:
                pass
            res.append((primary_label, t))
        except Exception:
            continue
    return res


def parse_interface_brief(out):
    rows = []
    lines = out.splitlines()
    section = None
    i = 0
    while i < len(lines):
        line = lines[i]
        if 'Brief information on interfaces in route mode' in line:
            section = 'route'
            i += 1
            continue
        if 'Brief information on interfaces in bridge mode' in line:
            section = 'bridge'
            i += 1
            continue
        if section == 'bridge' and ('Interface' in line and 'Link' in line and 'Speed' in line and 'Duplex' in line):
            cols = _split_columns_by_header(line,
                                            ['Interface', 'Link', 'Speed', 'Duplex', 'Type', 'PVID', 'Description'])
            j = i + 1
            while j < len(lines):
                row_line = lines[j]
                if not row_line.strip():
                    j += 1
                    continue
                if row_line.strip().startswith('<'):
                    break
                if 'Brief information on interfaces' in row_line or row_line.strip().startswith('Speed:') or row_line.strip().startswith('Link:') or row_line.strip().startswith('Type:') or row_line.strip().startswith('Duplex:'):
                    j += 1
                    continue
                if len(row_line) < 4:
                    j += 1
                    continue
                try:
                    port = _cut(row_line, cols['Interface'])
                    link = _cut(row_line, cols['Link'])
                    speed = _cut(row_line, cols['Speed'])
                    duplex = _cut(row_line, cols['Duplex'])
                    ptype = _cut(row_line, cols['Type'])
                    pvid = _cut(row_line, cols.get('PVID', (0, 8))) if 'PVID' in cols else ''
                    desc = _cut(row_line, cols.get('Description', (0, 8))) if 'Description' in cols else ''
                    if not re.match(r'^[A-Za-z\-]+\d+(?:/\d+)*\d*$', port):
                        j += 1
                        continue
                    rows.append({'端口': port, '链路状态': link, '速率配置': speed, '双工模式': duplex,
                                 '接口类型': ptype, 'PVID': pvid, '描述': desc})
                except Exception:
                    pass
                j += 1
            i = j
            continue
        i += 1
    if not rows:
        header_idx = None
        for idx, line in enumerate(lines):
            if 'Link' in line and 'Speed' in line and 'Type' in line:
                header_idx = idx
                break
        if header_idx is not None:
            for line in lines[header_idx + 1:]:
                s = line.strip()
                if not s:
                    continue
                if s.startswith('<') or s.startswith('['):
                    break
                parts = s.split()
                if len(parts) < 5:
                    continue
                port = parts[0]
                if not re.match(r'^[A-Za-z\-]+\d+(?:/\d+)*\d*$', port):
                    continue
                if len(parts) >= 7 and (parts[3] in ('A', 'T', 'H', 'F(a)', 'H(a)', 'A(a)')):
                    link, speed, duplex, ptype, pvid = parts[1], parts[2], parts[3], parts[4], parts[5]
                    desc = ' '.join(parts[6:]) if len(parts) > 6 else ''
                    rows.append({'端口': port, '链路状态': link, '速率配置': speed,
                                '双工模式': duplex, '接口类型': ptype, 'PVID': pvid, '描述': desc})
                else:
                    link, speed, ptype = parts[-3], parts[-2], parts[-1]
                    rows.append({'端口': port, '链路状态': link, '速率配置': speed,
                                '双工模式': '', '接口类型': ptype, 'PVID': '', '描述': ''})
    return rows


def parse_counters_simple(out):
    """解析 display counters inbound interface / outbound interface
    返回 {port: (total_pkts, err_pkts_or_None, overflow_flag)}
    """
    res = {}
    lines = out.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if 'Interface' in line and 'Total' in line and 'Err' in line:
            header_idx = idx
            break
    if header_idx is None:
        return res
    for line in lines[header_idx + 1:]:
        s = line.strip()
        if not s:
            break
        if s.startswith('<') or s.startswith('Overflow:') or s.startswith('--:'):
            break
        parts = s.split()
        if len(parts) < 2:
            continue
        port = parts[0]
        if not re.match(r'^[A-Za-z\-]+\d+(?:/\d+)*\d*$', port):
            continue
        overflow = False
        err = None
        # Err列: 最后一列可能是数字或 Overflow
        last = parts[-1]
        if last.lower() == 'overflow':
            overflow = True
        else:
            try:
                err = int(last)
            except Exception:
                pass
        try:
            total = int(re.sub(r'Overflow', '', parts[1]))
        except Exception:
            total = 0
        res[port] = (total, err if err is not None else 0, overflow)
    return res


def estimate_utilization_from_counters(t1, t2, gap_seconds, speed_mbps, mtu_guess=1500):
    """用两次 counters 的总报文数差估算带宽利用率（%）
    假设平均包长 = mtu_guess * 1.1 （含报头）
    """
    if not t1 or not t2:
        return 0.0
    try:
        diff = abs(int(t2) - int(t1))
    except Exception:
        return 0.0
    if gap_seconds <= 0 or speed_mbps <= 0 or diff <= 0:
        return 0.0
    avg_packet_bytes = mtu_guess + 24  # 以太网前导+IFG
    bytes_per_sec = diff * avg_packet_bytes / gap_seconds
    bits_per_sec = bytes_per_sec * 8.0
    cap_bps = speed_mbps * 1_000_000.0
    return round(min(999.0, bits_per_sec * 100.0 / cap_bps), 2)


def parse_counters_rate(out):
    map_in, map_out = {}, {}
    header_idx = None
    for idx, line in enumerate(out.splitlines()):
        if 'Interface' in line and ('In' in line or 'Out' in line or 'input' in line or 'output' in line):
            header_idx = idx
            break
    if header_idx is None:
        return map_in, map_out
    cols = _split_columns_by_header(out.splitlines()[header_idx], [
                                    'Interface', 'In', 'Out', 'InUti', 'OutUti', 'input', 'output'])
    for line in out.splitlines()[header_idx + 1:]:
        s = line.strip()
        if not s:
            continue
        if s.startswith('<'):
            break
        if 'Interface' in s:
            continue
        parts = s.split()
        if len(parts) < 2:
            continue
        port = parts[0]
        if not re.match(r'^[A-Za-z\-]+\d+(?:/\d+)*\d*$', port):
            continue
        nums = []
        for p in parts[1:]:
            m = re.match(r'(\d+(?:\.\d+)?)', p)
            if m:
                nums.append(float(m.group(1)))
        if len(nums) >= 2:
            map_in[port] = nums[-2]
            map_out[port] = nums[-1]
    return map_in, map_out


def parse_transceiver(out):
    items = defaultdict(dict)
    LONG_PORT_RE = re.compile(
        r'^(GigabitEthernet|Ten-GigabitEthernet|FortyGigE|HundredGigE|Twenty-FiveGigE|XGE|GE|FGE|HGE|TGE)\s*\d+(?:/\d+)+(?:[:/]\d+)?',
        re.I
    )
    lines = out.splitlines()
    i = 0
    cur_port = None
    skip_mode = None  # None / 'alarm' / 'warning' → 跳过阈值段的数值行
    NUM_RE = re.compile(r'\-?\d+(?:\.\d+)?')
    while i < len(lines):
        s = lines[i].strip()
        low_s = s.lower()
        if s:
            tok0 = s.split()[0]
            if LONG_PORT_RE.match(tok0) and ('transceiver' in low_s or LONG_PORT_RE.fullmatch(tok0)):
                cur_port = port_to_short(tok0)
                skip_mode = None
            elif LONG_PORT_RE.match(tok0):
                cur_port = port_to_short(tok0)
                skip_mode = None
        if cur_port and s:
            if 'does not support this function' in low_s:
                items[cur_port]['光模块状态'] = '不支持数字诊断'
            if 'is absent' in low_s:
                items[cur_port]['光模块状态'] = '未插入'
            if low_s.startswith('alarm thresholds'):
                skip_mode = 'alarm'
            elif low_s.startswith('warning thresholds'):
                skip_mode = 'warning'
            header_hit = (('Temp.' in s or 'Temperature' in s or '温度' in s)
                          and ('Voltage' in s or '电压' in s)
                          and ('Bias' in s or 'Current' in s or '电流' in s)
                          and ('RX power' in s or '收光功率' in s)
                          and ('TX power' in s or '发光功率' in s))
            if header_hit:
                # 当前段的表头下找 Current 数值行（排除 Alarm/Warning 段）
                j = i + 1
                vals_row = None
                while j < len(lines) and j <= i + 4:
                    candidate = lines[j]
                    cs = candidate.strip()
                    if not cs:
                        j += 1
                        continue
                    c_low = cs.lower()
                    if c_low.startswith(('alarm', 'warning')):
                        break
                    if c_low.startswith(('current', 'note', 'the', '---', 'high', 'low')):
                        j += 1
                        continue
                    if LONG_PORT_RE.match(cs.split()[0] if cs.split() else ''):
                        break
                    nums = NUM_RE.findall(cs)
                    if len(nums) >= 5:
                        vals_row = cs
                        break
                    j += 1
                if vals_row and skip_mode is None:
                    all_nums = NUM_RE.findall(vals_row)
                    if len(all_nums) >= 5:
                        mapping_key = ['温度(光模块)', '电压(V)', '电流(mA)', '收光(dBm)', '发光(dBm)']
                        for idx, key in enumerate(mapping_key):
                            try:
                                items[cur_port][key] = float(all_nums[idx])
                            except (ValueError, IndexError):
                                pass
        i += 1
    return items


def _looks_like_mac(s):
    s = str(s).strip().lower()
    if re.match(r'^[0-9a-f]{4}\-[0-9a-f]{4}\-[0-9a-f]{4}$', s):
        return True
    if re.match(r'^[0-9a-f]{2}(?:\:[0-9a-f]{2}){5}$', s):
        return True
    return False


def parse_lldp_list(out):
    entries = []
    lines = out.splitlines()
    header_idx = None
    cols = None
    header_line = None
    for i, line in enumerate(lines):
        if 'Local Interface' in line and ('System Name' in line or 'Chassis ID' in line):
            header_idx = i
            header_line = line.rstrip('\n').rstrip('\r')
            cols = _split_columns_by_header(header_line,
                                            ['Local Interface', 'Chassis ID', 'Port ID', 'System Name'])
            break
    if header_idx is None:
        for line in lines:
            s = line.strip()
            if not s:
                continue
            if not re.match(r'^(?:GigabitEthernet|Ten-GigabitEthernet|GE|XGE|FGE|HGE|FortyGigE|HundredGigE|Bridge-Aggregation)\s*\d', s, re.I):
                continue
            parts = s.split()
            if len(parts) < 4:
                continue
            loc = parts[0]
            neigh_port = parts[-2]
            neigh_name = parts[-1]
            loc_s = port_to_short(loc) if not _looks_like_mac(loc) else loc
            nport_s = port_to_short(neigh_port) if not _looks_like_mac(neigh_port) else neigh_port
            if neigh_name == '-':
                neigh_name = ''
            entries.append({'本地端口': loc_s, '邻居系统名': neigh_name, '邻居端口': nport_s})
        return entries
    # 用 header 定宽 + 最后一列取到行尾兜底，避免 System Name 被截断
    loc_range = cols.get('Local Interface')
    cid_range = cols.get('Chassis ID')
    pid_range = cols.get('Port ID')
    sysname_start = cols.get('System Name', (None, None))[0] if cols.get('System Name') else None
    for line in lines[header_idx + 1:]:
        if len(line) < 6:
            continue
        st = line.strip()
        if st.startswith('<') or st.startswith('Chassis ID') or st.startswith('Default --') or st.startswith('*') or st.startswith('#'):
            continue
        if not st:
            continue
        try:
            loc_raw = _cut(line, loc_range) if loc_range else ''
            neigh_port_raw = _cut(line, pid_range) if pid_range else ''
            if sysname_start is not None:
                neigh_name = line[sysname_start:].rstrip()
            else:
                neigh_name = ''
            if not neigh_name:
                # fallback: 用空格切，最后一段当 sysname，倒数第二段当 port id
                parts = st.split()
                if len(parts) >= 4:
                    neigh_name = parts[-1]
                    if neigh_port_raw in ('', None) or _looks_like_mac(neigh_port_raw) is False and re.search(r'\d', neigh_port_raw) is None:
                        neigh_port_raw = parts[-2]
            loc = port_to_short(loc_raw) if not _looks_like_mac(loc_raw) else loc_raw
            neigh_port = port_to_short(neigh_port_raw) if not _looks_like_mac(neigh_port_raw) else neigh_port_raw
            if neigh_name == '-':
                neigh_name = ''
            if not re.match(r'^[A-Za-z\-]+\d', loc):
                continue
            entries.append({'本地端口': loc, '邻居系统名': str(neigh_name).strip(), '邻居端口': neigh_port})
        except Exception:
            continue
    return entries


def parse_interface_util(out):
    """解析 display interface 文本，返回 {短端口名: {入带宽%,出带宽%,协商速率Mbps,峰值入Mbps,峰值出Mbps,真实速率Mbps,物理状态,协议状态}}
    """
    res = {}
    lines = out.splitlines()
    cur_port = None
    i = 0
    LONG_PORT_PREFIX = tuple(h for h, _ in PORT_SHORT_MAP) + tuple(s for _, s in PORT_SHORT_MAP)
    while i < len(lines):
        s = lines[i].rstrip()
        head = s.strip()
        if head and not head.startswith((' ', '\t')):
            tok0 = head.split()[0]
            m = re.match(
                r'^(GigabitEthernet|Ten-GigabitEthernet|FortyGigE|HundredGigE|Twenty-FiveGigE|XGE|GE|FGE|HGE|TGE|BAGG|RAGG|Bridge-Aggregation|Route-Aggregation|LoopBack|NULL|Vlan|MGE|Reg)\s*\d+(?:/\d+)*(?:\:\d+)?',
                tok0, re.I
            )
            if m or any(tok0.lower().startswith(p.lower()) for p in LONG_PORT_PREFIX):
                cur_port = port_to_short(tok0)
                res.setdefault(cur_port, {})
        if cur_port:
            low = s.lower()
            if 'current state:' in low:
                res[cur_port]['物理状态'] = 'UP' if 'current state: up' in low else 'DOWN'
            if 'line protocol state:' in low:
                res[cur_port]['协议状态'] = 'UP' if 'line protocol state: up' in low else 'DOWN'
            m_bw = re.search(r'bandwidth:\s*(\d+(?:\.\d+)?)\s*kbps', low)
            if m_bw:
                res[cur_port]['协商速率Mbps'] = round(float(m_bw.group(1)) / 1000.0, 3)
            m_spd = re.search(r'(\d+(?:\.\d+)?)\s*(mbps|gbps|kbps)\-speed\s*mode', low)
            if m_spd:
                val = float(m_spd.group(1))
                unit = m_spd.group(2).lower()
                if unit == 'gbps':
                    res[cur_port]['真实速率Mbps'] = val * 1000.0
                elif unit == 'kbps':
                    res[cur_port]['真实速率Mbps'] = val / 1000.0
                else:
                    res[cur_port]['真实速率Mbps'] = val
            pkin = re.search(r'peak\s+input\s+rate:\s*(\d+(?:\.\d+)?)\s*bytes/sec', low)
            if pkin:
                res[cur_port]['峰值入Mbps'] = round(float(pkin.group(1)) * 8.0 / 1_000_000.0, 2)
            pkout = re.search(r'peak\s+output\s+rate:\s*(\d+(?:\.\d+)?)\s*bytes/sec', low)
            if pkout:
                res[cur_port]['峰值出Mbps'] = round(float(pkout.group(1)) * 8.0 / 1_000_000.0, 2)
            if 'last 300 seconds input' in low:
                m_in = re.search(r'last\s+300\s+seconds\s+input\b[\s\S]*?(\d+(?:\.\d+)?)\s*%', low)
                if m_in:
                    try:
                        res[cur_port]['入带宽%'] = float(m_in.group(1))
                    except ValueError:
                        pass
            if 'last 300 seconds output' in low:
                m_out = re.search(r'last\s+300\s+seconds\s+output\b[\s\S]*?(\d+(?:\.\d+)?)\s*%', low)
                if m_out:
                    try:
                        res[cur_port]['出带宽%'] = float(m_out.group(1))
                    except ValueError:
                        pass
        i += 1
    return res


def classify_numeric(v, warn, err, higher_bad=True, ok_max=None, ok_min=None):
    try:
        v = float(v)
    except Exception:
        return '未知', v
    if ok_min is not None and ok_max is not None:
        if not (ok_min <= v <= ok_max):
            if v < ok_min or v > ok_max:
                return ('错误', v) if (v < ok_min - 3 or v > ok_max + 3) else ('告警', v)
        return '正常', v
    if higher_bad:
        if v >= err:
            return '错误', v
        if v >= warn:
            return '告警', v
    else:
        if v <= err:
            return '错误', v
        if v <= warn:
            return '告警', v
    return '正常', v


def speed_to_mbps(s):
    s = str(s).strip().lower()
    if not s or 'auto' in s:
        return 0.0
    m = re.match(r'(\d+(?:\.\d+)?)\s*([gmk])?', s)
    if not m:
        return 0.0
    v = float(m.group(1))
    u = (m.group(2) or '').lower()
    if u == 'g':
        return v * 1000.0
    if u == 'm' or u == '':
        return v
    if u == 'k':
        return v / 1000.0
    return v


def collect_one(d, counter_sample_seconds=30):
    sw = d['name']
    res = {
        'name': sw, 'host': d['host'], 'is_core': d['is_core'],
        '采集时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '登录状态': '失败', 'version_info': {}, 'config_txt': '',
        'cpu_pct': 0.0, 'mem_pct': 0.0, 'mem_used': 0.0, 'mem_total': 0.0,
        'temps': [], 'ports': [], 'lldp': [], 'errors': [],
        'counters_method': '', 'counters_sample_sec': counter_sample_seconds,
        '_raw_version': '', '_raw_manuinfo': ''
    }
    try:
        with ConnectHandler(**build_conn(d)) as ssh:
            safe_send(ssh, 'screen-length disable', delay=1, timeout=20)
            safe_send(ssh, 'terminal logging level 6', delay=1, timeout=15)
            ver_out = safe_send(ssh, 'display version', delay=2, timeout=60)
            log_cmd(sw, 'display version', len(ver_out.splitlines()))
            res['_raw_version'] = ver_out
            manuinfo_out = safe_send(ssh, 'display device manuinfo', delay=2, timeout=40)
            log_cmd(sw, 'display device manuinfo', len([l for l in manuinfo_out.splitlines() if l.strip()]))
            res['_raw_manuinfo'] = manuinfo_out
            # 【兜底】两段回显截断/复用（manuinfo 没返回时），把 version 文本拼进 manuinfo 再解析，保证 4 字段不空
            combined_manuinfo = (manuinfo_out or '').strip()
            if (not combined_manuinfo
                or combined_manuinfo.startswith('display version')
                or (ver_out.strip() and combined_manuinfo in ver_out.strip())
                or combined_manuinfo.lower().startswith('isplay version')
                    or '<' in combined_manuinfo[:80] and combined_manuinfo.strip().startswith('<') and 'display version' in combined_manuinfo[:400]):
                combined_manuinfo = (ver_out.strip() + '\n' + combined_manuinfo).strip()
            combined_version = (ver_out or '').strip()
            if manuinfo_out and manuinfo_out.strip() and ('DEVICE_NAME' in manuinfo_out or 'DEVICE_SERIAL_NUMBER' in manuinfo_out or 'BOARD TYPE' in manuinfo_out):
                combined_version = (combined_version + '\n' + manuinfo_out.strip()).strip()
            manuinfo = parse_device_manuinfo(combined_manuinfo)
            vi_preliminary = parse_version(combined_version, manuinfo)
            t_name, t_desc, retry_times = diagnose_template(ver_out, manuinfo_out, vi_preliminary)
            # 【模板触发自动补采】manuinfo截流/版本截流/未知模板 最多 retry_times 次重跑 dis device manuinfo
            if retry_times > 0 and 'DEVICE_SERIAL_NUMBER' not in (manuinfo_out or ''):
                retry_cmd_pool = [
                    ('display device manuinfo', 2, 50),
                    ('display device manuinfo verbose', 2, 60),
                    ('dis device manuinfo', 2, 50),
                    ('display device manuinfo', 3, 90),
                ]
                retry_ok = False
                for attempt in range(min(retry_times, len(retry_cmd_pool))):
                    c_c, c_d, c_t = retry_cmd_pool[attempt]
                    extra = ''
                    if attempt >= 1:
                        extra = f'（第{attempt+1}次尝试: read_timeout={c_t}s delay_factor={c_d}）'
                    log(f'[模板{t_name}补采{sw}] {c_c}{extra}')
                    time.sleep(1.0 + attempt * 0.5)
                    try_retry = safe_send(ssh, c_c, delay=c_d, timeout=c_t, read_timeout=c_t, max_attempts=1)
                    if (try_retry
                        and 'DEVICE_SERIAL_NUMBER' in try_retry
                            and len(try_retry) > len(manuinfo_out or '')):
                        manuinfo_out = try_retry
                        res['_raw_manuinfo'] = manuinfo_out
                        combined_manuinfo = (ver_out.strip() + '\n' + manuinfo_out.strip()).strip()
                        combined_version = (combined_version.strip() + '\n' + manuinfo_out.strip()).strip()
                        manuinfo = parse_device_manuinfo(combined_manuinfo)
                        vi_preliminary = parse_version(combined_version, manuinfo)
                        log(f'[模板{t_name}补采{sw}] ✅ 成功收到 DEVICE_SERIAL_NUMBER 段（{attempt+1}次命中）')
                        retry_ok = True
                        break
                if not retry_ok:
                    log(f'[模板{t_name}补采{sw}] ⚠️ 重试 {retry_times} 次仍未收到 manuinfo SN 段（不影响 型号/HW/SW 解析，将跳过SN解析失败）')
                    if bool(vi_preliminary.get('_manuinfo_missing_sn')) is False:
                        vi_preliminary['_manuinfo_missing_sn'] = True
                        v_note = str(vi_preliminary.get('备注') or '').strip()
                        extra_note = 'manuinfo 补采重试失败，SN 需补采'
                        vi_preliminary['备注'] = (v_note + '；' + extra_note if v_note else extra_note)
            # 最终模板判定（可能补采后升级为A-完整）
            t_name, t_desc, _ = diagnose_template(ver_out, manuinfo_out, vi_preliminary)
            vi_preliminary['诊断模板名'] = t_name
            vi_preliminary['诊断模板说明'] = t_desc
            log(f'[{sw}] 最终命中模板: {t_name}｜{t_desc}')
            res['version_info'] = vi_preliminary
            cfg_out = safe_send(ssh, 'display current-configuration', delay=3, timeout=300, read_timeout=300)
            log_cmd(sw, 'display current-configuration', len(cfg_out.splitlines()))
            res['config_txt'] = cfg_out
            cpu_out = safe_send(ssh, 'display cpu-usage', delay=2, timeout=40)
            res['cpu_pct'] = parse_cpu(cpu_out)
            log_cmd(sw, f"display cpu-usage -> {res['cpu_pct']}%")
            mem_out = safe_send(ssh, 'display memory', delay=2, timeout=40)
            res['mem_pct'], res['mem_used'], res['mem_total'] = parse_memory(mem_out)
            log_cmd(sw, f"display memory -> {res['mem_pct']}%")
            tmp = ''
            for c in ['display environment', 'display temperature', 'display fan', 'display power']:
                o = safe_send(ssh, c, delay=2, timeout=45)
                if any(k in o.lower() for k in ['temp', '温度', 'sensor', 'degree', '°c', ' c ']):
                    tmp = o
                    break
            res['temps'] = parse_temperature(tmp) if tmp else []
            if tmp:
                w_t, e_t = parse_environment_thresholds(tmp, model_name=res['version_info'].get('型号', ''))
                res['温度告警阈值'] = w_t
                res['温度错误阈值'] = e_t
            else:
                res['温度告警阈值'] = None
                res['温度错误阈值'] = None
            log_cmd(sw, '温度采集', len(res['temps']))
            brief_out = safe_send(ssh, 'display interface brief', delay=3, timeout=60)
            res['ports'] = parse_interface_brief(brief_out)
            log_cmd(sw, 'display interface brief', len(res['ports']))
            method_used = ''
            rate_out = safe_send(ssh, 'display counters rate inbound outbound interface', delay=3, timeout=90)
            if 'ERROR' in rate_out or 'invalid' in rate_out.lower() or 'Unrecognized' in rate_out or 'Incomplete' in rate_out:
                rate_out = safe_send(ssh, 'display counters rate', delay=3, timeout=90)
            if rate_out and not any(k in rate_out for k in ['Unrecognized', 'Incomplete', 'Invalid', '% Unrecognized', '% Incomplete']):
                map_in, map_out = parse_counters_rate(rate_out)
                method_used = 'rate'
            else:
                map_in, map_out = {}, {}
            if not map_in and not map_out:
                in1 = safe_send(ssh, 'display counters inbound interface', delay=2, timeout=60)
                out1 = safe_send(ssh, 'display counters outbound interface', delay=2, timeout=60)
                t1 = time.time()
                time.sleep(max(1, counter_sample_seconds))
                in2 = safe_send(ssh, 'display counters inbound interface', delay=2, timeout=60)
                out2 = safe_send(ssh, 'display counters outbound interface', delay=2, timeout=60)
                t2 = time.time()
                gap = max(1, t2 - t1)
                m_in1 = parse_counters_simple(in1)
                m_out1 = parse_counters_simple(out1)
                m_in2 = parse_counters_simple(in2)
                m_out2 = parse_counters_simple(out2)
                method_used = f'两次采样{int(round(gap,0))}s'
                res['counters_method'] = f'counters_in_out_sample_{int(round(gap,0))}s'
                for p in res['ports']:
                    port = p['端口']
                    speed_mbps = speed_to_mbps(p.get('速率配置', ''))
                    if speed_mbps <= 0:
                        try:
                            if p.get('接口类型') in ('T', 'H'):
                                speed_mbps = 1000.0
                        except Exception:
                            pass
                    t_in1 = m_in1.get(port, (0, 0, False))[0]
                    t_in2 = m_in2.get(port, (0, 0, False))[0]
                    t_out1 = m_out1.get(port, (0, 0, False))[0]
                    t_out2 = m_out2.get(port, (0, 0, False))[0]
                    in_util = estimate_utilization_from_counters(t_in1, t_in2, gap, speed_mbps)
                    out_util = estimate_utilization_from_counters(t_out1, t_out2, gap, speed_mbps)
                    map_in[port] = in_util
                    map_out[port] = out_util
                    p['入方向计数溢出'] = '是' if (m_in1.get(port, (0, 0, False))[2]
                                           or m_in2.get(port, (0, 0, False))[2]) else '否'
                    p['出方向计数溢出'] = '是' if (m_out1.get(port, (0, 0, False))[2]
                                           or m_out2.get(port, (0, 0, False))[2]) else '否'
                    p['入方向错误包'] = m_in2.get(port, (0, 0, False))[1] or 0
                    p['出方向错误包'] = m_out2.get(port, (0, 0, False))[1] or 0
            else:
                res['counters_method'] = 'counters_rate'
                for p in res['ports']:
                    p['入方向计数溢出'] = ''
                    p['出方向计数溢出'] = ''
                    p['入方向错误包'] = ''
                    p['出方向错误包'] = ''
            for p in res['ports']:
                p.setdefault('入带宽利用率%', map_in.get(p['端口'], 0.0))
                p.setdefault('出带宽利用率%', map_out.get(p['端口'], 0.0))
                p['带宽方法_原始'] = res['counters_method']
            log_cmd(sw, f'带宽采集:{method_used or "未命中"}', len(map_in))
            # 追加 display interface 官方300s口径（若能抓到 util%，则优先级最高；同时写入真实协商速率）
            try:
                if_out = safe_send(ssh, 'display interface', delay=4, timeout=300, read_timeout=300)
                util_map = parse_interface_util(if_out) if if_out else {}
                for p in res['ports']:
                    u = util_map.get(p['端口'], {})
                    p['display_interface_info'] = u
                    if u.get('入带宽%') is not None:
                        p['入带宽利用率%'] = float(u['入带宽%'])
                        p['带宽方法_入'] = 'display_interface_300s'
                    if u.get('出带宽%') is not None:
                        p['出带宽利用率%'] = float(u['出带宽%'])
                        p['带宽方法_出'] = 'display_interface_300s'
                    if u.get('真实速率Mbps'):
                        p['真实协商速率Mbps'] = float(u['真实速率Mbps'])
                    elif u.get('协商速率Mbps'):
                        p['真实协商速率Mbps'] = float(u['协商速率Mbps'])
                    if u.get('峰值入Mbps') is not None:
                        p['峰值入Mbps'] = float(u['峰值入Mbps'])
                    if u.get('峰值出Mbps') is not None:
                        p['峰值出Mbps'] = float(u['峰值出Mbps'])
                log_cmd(sw, 'display interface 解析', len(util_map))
            except Exception as ee:
                log(f'[WARN] {sw} display interface 失败: {ee}')
            trans = ''
            for c in ['display transceiver diagnosis interface', 'display transceiver interface',
                      'display transceiver diagnosis', 'display transceiver']:
                o = safe_send(ssh, c, delay=3, timeout=120)
                if any(k in o.lower() for k in ['power', 'rx', 'tx', '收光', '发光', 'bias', 'voltage', 'current']):
                    trans = o
                    break
            diag_map = parse_transceiver(trans) if trans else {}
            for p in res['ports']:
                p['光模块诊断'] = diag_map.get(p['端口'], {})
            log_cmd(sw, '光模块诊断采集', len(diag_map))
            lldp_out = safe_send(ssh, 'display lldp neighbor-information list', delay=3, timeout=60)
            res['lldp'] = parse_lldp_list(lldp_out)
            log_cmd(sw, 'display lldp neighbor-information list', len(res['lldp']))
            res['登录状态'] = '成功'
    except Exception as e:
        tb = traceback.format_exc()
        res['errors'].append(str(e) + '\n' + tb)
        log(f'[ERROR] {sw}: {e}')
    return res


def write_config_txt(data, collected_hosts=None):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    cnt = 0
    for d in data:
        if collected_hosts is not None and d.get('host') not in collected_hosts:
            # 本次没有重新采集的设备，保留旧配置 txt 文件，不重写不删除
            continue
        name = re.sub(r'[\\/:*?"<>|]+', '_', d['name'])
        p = CONFIG_DIR / f"{name}_{d['host']}.txt"
        try:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(f"设备名称: {d['name']}\n管理IP: {d['host']}\n采集时间: {d['采集时间']}\n")
                f.write(f"登录状态: {d['登录状态']}\n型号: {d['version_info'].get('型号', '未知')}\n")
                f.write(f"SN: {d['version_info'].get('序列号SN', '未知')}\n{''.join(['-'] * 80)}\n\n")
                f.write(d['config_txt'] or '')
                if d.get('errors'):
                    f.write('\n\n==== 采集异常 =====\n' + '\n'.join(d['errors']))
            cnt += 1
        except Exception as e:
            log(f'[WARN] 写配置文件失败 {name}: {e}')
    log(f'已输出 {cnt} 台设备的配置文件到: {CONFIG_DIR}'
        + (f'（仅本次新采集的 {len(collected_hosts)} 台；其他设备配置文件已保留，未重写）' if collected_hosts is not None else ''))


def _hs(wb, sheet, cols):
    ws = wb[sheet]
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i, value=c)
        cell.font = Font(bold=True, color='FFFFFFFF')
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = max(16, min(56, int(len(str(cols[i - 1])) * 2.4) + 6))
    ws.freeze_panes = 'A2'


def write_ledger(data):
    wb = Workbook()
    ws = wb.active
    ws.title = '设备台账'
    cols = ['序号', '设备名称', '管理IP', '是否核心', '登录状态', '型号', '序列号SN', '硬件版本', '软件版本', '诊断模板名', '采集时间', '备注']
    _hs(wb, '设备台账', cols)
    for idx, d in enumerate(data, 1):
        v = d['version_info']
        model = str(v.get('型号', ''))
        sn = str(v.get('序列号SN', ''))
        hw = str(v.get('硬件版本', ''))
        sw = str(v.get('软件版本', ''))
        tmpl = str(v.get('诊断模板名', ''))
        note_parts = []
        if d['errors']:
            note_parts.append(';'.join(d['errors']))
        # 如果诊断模板不是 A-完整/D-多Slot框式，把模板说明写进备注
        if tmpl and tmpl not in ('A-完整', 'D-多Slot框式'):
            tmpl_note = f"【诊断模板】{tmpl}：{v.get('诊断模板说明') or ''}"
            # B/C模板下，SN=未知但型号3字段全对就不贴冗长原始输出了
            if tmpl.startswith('B-') or tmpl == 'C-版本截流':
                missing_core = [k for k, val in [('型号', model), ('硬件', hw), ('软件', sw)] if not val or val == '未知']
                if not missing_core:
                    note_parts.append(tmpl_note)
                    note = '\n'.join(note_parts)[:4000]
                    row = [idx, d['name'], d['host'], '是' if d['is_core'] else '否', d['登录状态'],
                           model, sn, hw, sw, tmpl, d['采集时间'], note]
                    for c, val in enumerate(row, 1):
                        cell = ws.cell(row=idx + 1, column=c, value=val)
                        cell.border = BORDER
                        cell.alignment = Alignment(vertical='center', wrap_text=True)
                        if d['登录状态'] != '成功':
                            cell.fill = ERR_FILL
                        elif tmpl in ('A-完整', 'D-多Slot框式'):
                            cell.fill = OK_FILL
                        elif tmpl.startswith('B-') or tmpl == 'C-版本截流':
                            cell.fill = WARN_FILL
                        else:
                            cell.fill = ERR_FILL
                    continue
                else:
                    note_parts.append(tmpl_note)
            elif tmpl.startswith('Z-'):
                note_parts.append(tmpl_note)
        if d['登录状态'] == '成功' and (model == '未知' or sn == '未知' or hw == '未知' or sw == '未知'):
            diag = []
            missing = [k for k, val in [('型号', model), ('SN', sn), ('硬件', hw), ('软件', sw)] if val == '未知']
            diag.append(f"[诊断] 以下字段解析失败: {','.join(missing)}；请把本段原始输出贴回补正则")
            diag.append('===== display version 原始输出（前1200字）=====')
            raw_v = (d.get('_raw_version') or '')[:1200]
            if raw_v:
                diag.append(raw_v)
            diag.append('===== display device manuinfo 原始输出（前800字）=====')
            raw_m = (d.get('_raw_manuinfo') or '')[:800]
            if raw_m:
                diag.append(raw_m)
            note_parts.append('\n'.join(diag))
        note = '\n'.join(note_parts)[:4000]
        row = [idx, d['name'], d['host'], '是' if d['is_core'] else '否', d['登录状态'],
               model, sn, hw, sw, tmpl,
               d['采集时间'], note]
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=idx + 1, column=c, value=val)
            cell.border = BORDER
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            if d['登录状态'] != '成功':
                cell.fill = ERR_FILL
            elif tmpl in ('A-完整', 'D-多Slot框式'):
                cell.fill = OK_FILL
            elif tmpl.startswith('B-') or tmpl == 'C-版本截流':
                cell.fill = WARN_FILL
            else:
                cell.fill = ERR_FILL
    wb.save(str(LEDGER_FILE))
    log(f'设备台账已输出: {LEDGER_FILE}')


def _f(c, tag):
    if tag == '错误':
        c.fill = ERR_FILL
    elif tag == '告警':
        c.fill = WARN_FILL
    else:
        c.fill = OK_FILL


def parse_environment_thresholds(out, model_name=''):
    """从 display environment 表中取每个 Sensor 专属 Warning/Alarm
    返回 (warn_thresholds: dict[sensor_name→float(Warning*0.92)],
            err_thresholds: dict[sensor_name→float(Alarm*0.95)])
    同时保留向后兼容：如果调用方解包成两个变量也能用；judge_need_check / _health_header_rows 已按 dict[sensor] 取值
    """
    warn_dict, err_dict = {}, {}
    lines = out.splitlines()
    header_idx, cols_map = None, None
    for i, line in enumerate(lines):
        if 'Temperature' in line and 'Warning' in line and 'Alarm' in line:
            header_idx = i
            cols_map = _split_columns_by_header(
                line, ['Slot', 'Sensor', 'Name', 'Temperature', 'Lower', 'Warning', 'Alarm', 'Shutdown'])
            break
    if header_idx is None or not cols_map:
        return None, None
    for line in lines[header_idx + 1:]:
        if len(line) < 4:
            continue
        if line.strip().startswith('<'):
            continue
        try:
            sensor_name = ''
            if 'Sensor' in cols_map:
                sc = _cut(line, cols_map['Sensor'])
                if sc:
                    sensor_name = re.sub(r'\s+', ' ', sc).strip()
            if 'Name' in cols_map and cols_map['Name']:
                nc = _cut(line, cols_map['Name'])
                if nc and nc.strip() and (not sensor_name or sensor_name.isdigit()):
                    sensor_name = re.sub(r'\s+', ' ', nc).strip()
            slot_s = ''
            if 'Slot' in cols_map:
                slot_c = _cut(line, cols_map['Slot'])
                if slot_c:
                    slot_s = slot_c.strip()
            # 合并 Slot+Sensor：若多 Slot 相同 Sensor 名，按 Slot-Sensor 做 key，单 Slot 直接 Sensor name 也行
            base_names = [sensor_name]
            if slot_s and sensor_name:
                base_names.insert(0, f'{slot_s}-{sensor_name}')
            w_val = a_val = None
            w_cell = _cut(line, cols_map['Warning'])
            m = re.search(r'\-?\d+(?:\.\d+)?', w_cell)
            if m and m.group(0) not in ('0', 'NA', 'N/A'):
                try:
                    w_val = round(float(m.group(0)) * 0.92, 2)
                except ValueError:
                    w_val = None
            a_cell = _cut(line, cols_map['Alarm'])
            m = re.search(r'\-?\d+(?:\.\d+)?', a_cell)
            if m and m.group(0) not in ('0', 'NA', 'N/A'):
                try:
                    a_val = round(float(m.group(0)) * 0.95, 2)
                except ValueError:
                    a_val = None
            if not base_names:
                continue
            for bn in base_names:
                if not bn:
                    continue
                if w_val is not None:
                    warn_dict.setdefault(bn, w_val)
                if a_val is not None:
                    err_dict.setdefault(bn, a_val)
                # 别名（judge_need_check 里 4 种变形匹配，这里多 key 一次存入兜底）
                variants = [bn.replace(' ', '_'), bn.lower(), bn.upper(), bn.replace(' ', '')]
                for v in variants:
                    if v == bn:
                        continue
                    if w_val is not None:
                        warn_dict.setdefault(v, w_val)
                    if a_val is not None:
                        err_dict.setdefault(v, a_val)
        except Exception:
            continue
    return (warn_dict or None), (err_dict or None)


def _health_header_rows(data):
    rows = []
    for d in data:
        tag_cpu, _ = classify_numeric(d['cpu_pct'], CPU_WARN, CPU_ERR)
        tag_mem, _ = classify_numeric(d['mem_pct'], MEM_WARN, MEM_ERR)
        dev_warn, dev_err = d.get('温度告警阈值'), d.get('温度错误阈值')
        thr_w = dev_warn if isinstance(dev_warn, dict) else {}
        thr_e = dev_err if isinstance(dev_err, dict) else {}
        # 兼容旧结构：如果 dev_warn 是 number(老版本 minWarning*0.92 标量)，转成 dict 全部 sensors 都用这个数
        if not thr_w and isinstance(dev_warn, (int, float)):
            fallback_w = float(dev_warn)
            thr_w = {f'__fallback__': fallback_w}
        if not thr_e and isinstance(dev_err, (int, float)):
            fallback_e = float(dev_err)
            thr_e = {f'__fallback__': fallback_e}
        tags_t = []
        temp_max_val = 0.0
        temps_src = d['temps'] or []
        if not temps_src and d['登录状态'] == '成功':
            temps_src = [('采集无', -1.0)]
        elif not temps_src:
            temps_src = [('未知', -1.0)]
        for t in temps_src:
            t_val = None
            t_name = ''
            if isinstance(t, dict):
                t_name = str(t.get('name') or t.get('sensor') or t.get('label') or '')
                try:
                    t_val = float(t.get('value') or 0.0)
                except Exception:
                    t_val = None
            elif isinstance(t, (tuple, list)) and len(t) >= 2:
                t_name = str(t[0] or '')
                try:
                    t_val = float(t[1] or 0.0)
                except Exception:
                    t_val = None
            if t_val is None:
                continue
            if t_val > temp_max_val:
                temp_max_val = t_val
            wv = ev = None
            for cand in [t_name, t_name.replace(' ', '_'), t_name.lower(), t_name.upper(), t_name.replace(' ', ''), '__fallback__']:
                if not cand:
                    continue
                if wv is None and cand in thr_w:
                    try:
                        wv = float(thr_w[cand])
                    except Exception:
                        wv = None
                if ev is None and cand in thr_e:
                    try:
                        ev = float(thr_e[cand])
                    except Exception:
                        ev = None
            if wv is None:
                wv = TEMP_WARN
            if ev is None:
                ev = TEMP_ERR
            try:
                tt = classify_numeric(t_val, wv, ev)[0]
            except Exception:
                tt = '正常'
            tags_t.append(tt)
        if not tags_t:
            tags_t = ['正常']

        def worst(x): return '错误' if '错误' in x else ('告警' if '告警' in x else '正常')
        worst_temp = worst(tags_t)
        if thr_w:
            try:
                if '__fallback__' in thr_w:
                    disp_tw = round(float(thr_w['__fallback__']), 1)
                else:
                    reals = [float(v) for v in thr_w.values() if isinstance(v, (int, float))]
                    disp_tw = round(min(reals), 1) if reals else TEMP_WARN
            except Exception:
                disp_tw = TEMP_WARN
        else:
            disp_tw = TEMP_WARN
        if thr_e:
            try:
                if '__fallback__' in thr_e:
                    disp_te = round(float(thr_e['__fallback__']), 1)
                else:
                    reals = [float(v) for v in thr_e.values() if isinstance(v, (int, float))]
                    disp_te = round(min(reals), 1) if reals else TEMP_ERR
            except Exception:
                disp_te = TEMP_ERR
        else:
            disp_te = TEMP_ERR
        rows.append({
            '设备名称': d['name'], '管理IP': d['host'], '是否核心': '是' if d['is_core'] else '否',
            '登录状态': d['登录状态'], 'CPU利用率%': round(d['cpu_pct'], 2), 'CPU状态': tag_cpu,
            '内存利用率%': round(d['mem_pct'], 2), '内存状态': tag_mem,
            '温度(最高传感器)': round(temp_max_val, 1) if (temp_max_val > 0 or (temps_src and temps_src[0][1] not in (-1.0,))) else (temps_src[0][0] if temps_src else '未知'),
            '温度状态': worst_temp,
            '温度告警阈值': f'{round(disp_tw,1)}C' if disp_tw else TEMP_WARN,
            '温度错误阈值': f'{round(disp_te,1)}C' if disp_te else TEMP_ERR,
            'UP端口数': sum(1 for p in d['ports'] if (isinstance(p, dict) and str(p.get('链路状态', '')).upper() in ('UP', 'U'))),
            '总物理端口数': len(d['ports']),
            '带宽采集方法': d.get('counters_method', ''),
            '采集时间': d['采集时间'],
            '备注': ('温度传感器:' + ';'.join([f"{t[0]}={t[1]}C" for t in d['temps'][:6]])) if d['temps'] else ''
        })
    return rows


def _port_rows(d):
    out = []
    dev_warn, dev_err = d.get('温度告警阈值'), d.get('温度错误阈值')
    for p in d['ports']:
        up = p['链路状态'].upper() in ('UP', 'U')
        in_u = float(p.get('入带宽利用率%', 0.0))
        out_u = float(p.get('出带宽利用率%', 0.0))
        bw_max = max(in_u, out_u)
        bw_tag, _ = classify_numeric(bw_max, BANDWIDTH_WARN, BANDWIDTH_ERR) if up else ('未启用', bw_max)
        diag = p.get('光模块诊断') or {}
        rx = diag.get('收光(dBm)')
        tx = diag.get('发光(dBm)')
        rx_tag = '未采集'
        if rx is not None:
            if not (RX_PWR_MIN <= rx <= RX_PWR_MAX):
                rx_tag = '告警' if (RX_PWR_MIN - 5 <= rx <= RX_PWR_MAX + 2) else '错误'
            else:
                rx_tag = '正常'
        tx_tag = '未采集'
        if tx is not None:
            tx_tag = '正常' if (TX_PWR_MIN <= tx <= TX_PWR_MAX) else (
                '告警' if (TX_PWR_MIN - 3 <= tx <= TX_PWR_MAX + 1) else '错误')
        anomalous = []
        if up and bw_tag in ('告警', '错误'):
            anomalous.append(f'带宽{bw_tag}')
        if rx_tag in ('告警', '错误'):
            anomalous.append(f'收光{rx_tag}')
        if tx_tag in ('告警', '错误'):
            anomalous.append(f'发光{tx_tag}')
        if p.get('入方向计数溢出') == '是':
            anomalous.append('入计数器溢出')
        if p.get('出方向计数溢出') == '是':
            anomalous.append('出计数器溢出')
        try:
            if int(p.get('入方向错误包') or 0) > 0:
                anomalous.append('入方向错误包')
            if int(p.get('出方向错误包') or 0) > 0:
                anomalous.append('出方向错误包')
        except Exception:
            pass
        # 光模块温度也判断
        mod_temp = diag.get('温度(光模块)')
        mod_temp_tag = ''
        if mod_temp is not None:
            try:
                mod_t = float(mod_temp)
                if dev_err and mod_t >= dev_err:
                    mod_temp_tag = '错误'
                elif dev_warn and mod_t >= dev_warn:
                    mod_temp_tag = '告警'
                elif mod_t >= TEMP_ERR:
                    mod_temp_tag = '错误'
                elif mod_t >= TEMP_WARN:
                    mod_temp_tag = '告警'
                if mod_temp_tag in ('告警', '错误'):
                    anomalous.append(f'光模块温度{mod_temp_tag}')
            except Exception:
                pass
        out.append({
            '设备名称': d['name'], '管理IP': d['host'], '端口': p['端口'], '链路状态': p['链路状态'],
            '速率配置': p.get('速率配置', ''), '双工': p.get('双工模式', ''),
            '接口类型': p.get('接口类型', ''), 'PVID': p.get('PVID', ''),
            '描述': p.get('描述', ''),
            '入带宽利用率%': round(in_u, 2),
            '出带宽利用率%': round(out_u, 2),
            '带宽状态': bw_tag,
            '收光(dBm)': rx if rx is not None else '-',
            '发光(dBm)': tx if tx is not None else '-',
            '收光状态': rx_tag, '发光状态': tx_tag,
            '光模块温度': mod_temp if mod_temp is not None else ('-'),
            '光模块电压(V)': diag.get('电压(V)') if diag.get('电压(V)') is not None else '-',
            '光模块电流(mA)': diag.get('电流(mA)') if diag.get('电流(mA)') is not None else '-',
            '入方向错误包': p.get('入方向错误包', ''),
            '出方向错误包': p.get('出方向错误包', ''),
            '入方向计数溢出': p.get('入方向计数溢出', ''),
            '出方向计数溢出': p.get('出方向计数溢出', ''),
            '异常标记': '; '.join(anomalous) if anomalous else ''
        })
    return out


def _anomaly_rows(data, port_rows):
    devs = []
    for d in data:
        reasons = []
        if d['登录状态'] != '成功':
            reasons.append(f"登录失败:{d['登录状态']}")
        tag = classify_numeric(d['cpu_pct'], CPU_WARN, CPU_ERR)[0]
        if tag in ('告警', '错误'):
            reasons.append(f'CPU{tag}')
        tag = classify_numeric(d['mem_pct'], MEM_WARN, MEM_ERR)[0]
        if tag in ('告警', '错误'):
            reasons.append(f'内存{tag}')
        dev_warn, dev_err = d.get('温度告警阈值'), d.get('温度错误阈值')
        tw = dev_warn if dev_warn else TEMP_WARN
        te = dev_err if dev_err else TEMP_ERR
        for t in d['temps']:
            tag = classify_numeric(t[1], tw, te)[0]
            if tag in ('告警', '错误'):
                reasons.append(f'温度传感器{t[0]}={t[1]}C-{tag}')
        if reasons:
            devs.append({'异常级别': '错误' if any('错误' in r for r in reasons) else '告警',
                         '对象类型': '设备', '对象标识': f"{d['name']}({d['host']})",
                         '异常内容': '; '.join(reasons),
                         '建议操作': '立即登录控制台检查/升级补丁/清理风扇/检查上联链路',
                         '采集时间': d['采集时间']})
    for p in port_rows:
        if p['异常标记']:
            level = '错误' if '错误' in p['异常标记'] else '告警'
            devs.append({'异常级别': level, '对象类型': '端口',
                         '对象标识': f"{p['设备名称']}:{p['端口']}",
                         '异常内容': p['异常标记'],
                         '建议操作': '检查链路/光模块衰减/交换机转发队列/下挂傻瓜交换机环路',
                         '采集时间': ''})
    return devs


def write_health(data):
    wb = Workbook()
    wb.remove(wb.active)
    wb.create_sheet('设备健康概览')
    wb.create_sheet('端口带宽与光功率')
    wb.create_sheet('异常汇总')
    overview_cols = ['序号', '设备名称', '管理IP', '是否核心', '登录状态', 'CPU利用率%', 'CPU状态', '内存利用率%',
                     '内存状态', '温度(最高传感器)', '温度状态', '温度告警阈值', '温度错误阈值',
                     'UP端口数', '总物理端口数', '带宽采集方法', '采集时间', '备注']
    _hs(wb, '设备健康概览', overview_cols)
    hrows = _health_header_rows(data)
    for idx, r in enumerate(hrows, 1):
        for c, k in enumerate(overview_cols[1:], 2):
            cell = wb['设备健康概览'].cell(row=idx + 1, column=c, value=r.get(k, ''))
            cell.border = BORDER
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        wb['设备健康概览'].cell(row=idx + 1, column=1, value=idx).border = BORDER
        for col_key, cell_col in [('CPU状态', 7), ('内存状态', 9), ('温度状态', 11), ('登录状态', 5)]:
            c = wb['设备健康概览'].cell(row=idx + 1, column=cell_col)
            if col_key == '登录状态':
                if str(c.value) != '成功':
                    c.fill = ERR_FILL
                    continue
            else:
                _f(c, str(c.value))
    port_cols = ['设备名称', '管理IP', '端口', '链路状态', '速率配置', '双工', '接口类型', 'PVID', '描述',
                 '入带宽利用率%', '出带宽利用率%',
                 '带宽状态', '收光(dBm)', '发光(dBm)', '收光状态', '发光状态', '光模块温度', '光模块电压(V)',
                 '光模块电流(mA)', '入方向错误包', '出方向错误包', '入方向计数溢出', '出方向计数溢出',
                 '异常标记']
    _hs(wb, '端口带宽与光功率', port_cols)
    all_ports = []
    for d in data:
        all_ports.extend(_port_rows(d))
    for idx, r in enumerate(all_ports, 1):
        for c, k in enumerate(port_cols, 1):
            cell = wb['端口带宽与光功率'].cell(row=idx + 1, column=c, value=r.get(k, ''))
            cell.border = BORDER
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        for col_key, col_idx in [('带宽状态', 12), ('收光状态', 15), ('发光状态', 16)]:
            c = wb['端口带宽与光功率'].cell(row=idx + 1, column=col_idx)
            _f(c, str(c.value))
        if r['异常标记']:
            for col_idx in [12, 15, 16, 24]:
                c = wb['端口带宽与光功率'].cell(row=idx + 1, column=col_idx)
                if c.fill == PatternFill() or c.fill.start_color.rgb in ('00000000', None):
                    c.fill = WARN_FILL
    anom_cols = ['异常级别', '对象类型', '对象标识', '异常内容', '建议操作', '采集时间']
    _hs(wb, '异常汇总', anom_cols)
    for idx, r in enumerate(_anomaly_rows(data, all_ports), 1):
        for c, k in enumerate(anom_cols, 1):
            cell = wb['异常汇总'].cell(row=idx + 1, column=c, value=r.get(k, ''))
            cell.border = BORDER
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            if r.get('异常级别') == '错误':
                cell.fill = ERR_FILL
            elif r.get('异常级别') == '告警':
                cell.fill = WARN_FILL
    wb.save(str(HEALTH_FILE))
    log(f'健康状态总表已输出: {HEALTH_FILE}')


def build_topo(data):
    names = {d['name'] for d in data}
    ips = {d['host']: d['name'] for d in data}
    nodes = []
    device_health = {}
    for d in data:
        hrows = _health_header_rows([d])[0]
        worst = '正常'
        for tag in [hrows['CPU状态'], hrows['内存状态'], hrows['温度状态']]:
            if tag == '错误':
                worst = '错误'
                break
            if tag == '告警':
                worst = '告警'
        if d['登录状态'] != '成功':
            worst = '错误'
        device_health[d['name']] = worst
        color = {'正常': '#2E7D32', '告警': '#F57C00', '错误': '#C62828'}[worst]
        shape = 'box' if d['is_core'] else 'ellipse'
        nodes.append({
            'id': d['name'], 'label': f"{d['name']}\n{d['host']}",
            'shape': shape, 'color': {'background': color, 'border': '#111111'},
            'title': f"型号: {d['version_info'].get('型号','')}<br>SN: {d['version_info'].get('序列号SN','')}<br>"
                     f"CPU:{hrows['CPU利用率%']}% 内存:{hrows['内存利用率%']}% 温度:{hrows['温度(最高传感器)']}<br>"
                     f"登录状态:{d['登录状态']} 健康级别:{worst}",
            'size': 26 if d['is_core'] else 20, 'font': {'color': '#FFFFFF', 'face': 'Microsoft YaHei'}
        })
    edges = []
    seen = set()
    for d in data:
        for l in d['lldp']:
            a = d['name']
            neigh = l['邻居系统名']
            b = neigh if neigh in names else (ips.get(neigh, neigh))
            if a == b:
                continue
            a_port = l['本地端口']
            b_port = l['邻居端口']
            key = tuple(sorted([(a, a_port), (b, b_port)]))
            if key in seen:
                continue
            seen.add(key)
            anomalous = False
            for p in d['ports']:
                if p['端口'] == a_port and p.get('异常标记'):
                    anomalous = True
                    break
            color = '#C62828' if anomalous else '#666666'
            edges.append({
                'from': a, 'to': b,
                'label': f"{a_port}<->\n{b_port}",
                'title': f"{a}({a_port}) <-> {b}({b_port})",
                'color': {'color': color},
                'font': {'size': 11, 'face': 'Microsoft YaHei'}
            })
    return nodes, edges, device_health


def write_topo_html(nodes, edges, data_list):
    import shutil as _shutil
    topo_dir = TOPO_FILE.parent
    local_static_dir = topo_dir / '_static'
    local_vis_src = BASE_DIR / '_static' / 'vis-network.min.js'
    local_vis_dst = local_static_dir / 'vis-network.min.js'
    if local_vis_src.exists():
        local_static_dir.mkdir(parents=True, exist_ok=True)
        try:
            _shutil.copy(str(local_vis_src), str(local_vis_dst))
        except Exception:
            pass
    have_local = local_vis_dst.exists()
    stat = defaultdict(int)
    for d in data_list:
        w = '正常'
        if d['登录状态'] != '成功':
            w = '错误'
        else:
            hrows = _health_header_rows([d])[0]
            if '错误' in [hrows['CPU状态'], hrows['内存状态'], hrows['温度状态']]:
                w = '错误'
            elif '告警' in [hrows['CPU状态'], hrows['内存状态'], hrows['温度状态']]:
                w = '告警'
        stat[w] += 1
    summary_html = f"<div class='summary'><b>统计概览</b>：总设备 {len(data_list)}；正常 {stat['正常']}，告警 {stat['告警']}，错误 {stat['错误']}，未采集到LLDP链路的设备以孤立节点显示。{'【离线模式】' if have_local else ''}</div>"
    vis_load_scripts = []
    if have_local:
        vis_load_scripts.append('<script type="text/javascript" src="_static/vis-network.min.js"></script>')
    else:
        vis_load_scripts.append(
            '<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/vis-network@9.1.6/standalone/umd/vis-network.min.js" onerror="var s=document.createElement(\'script\');s.src=\'https://unpkg.com/vis-network@9.1.6/standalone/umd/vis-network.min.js\';s.onerror=function(){var s2=document.createElement(\'script\');s2.src=\'https://cdnjs.cloudflare.com/ajax/libs/vis-network/9.1.6/vis-network.min.js\';document.body.appendChild(s2);};document.body.appendChild(s);"></script>')
    load_err_block = """
<div id="visErr" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:99999;align-items:center;justify-content:center;">
  <div style="background:#fff;padding:24px 28px;max-width:560px;border-radius:6px;box-shadow:0 6px 24px rgba(0,0,0,.25);">
    <h3 style="margin:0 0 10px;color:#C62828;">拓扑图资源加载失败</h3>
    <p style="margin:0 0 12px;">原因：当前机器无法访问 jsdelivr / unpkg / cdnjs 这三个公共 CDN，且没有本地离线 vis-network 文件。</p>
    <p style="margin:0 0 6px;">解决方法（任选其一）：</p>
    <ol style="margin:0 0 10px 20px;line-height:1.8;">
      <li>回到运行脚本的机器，确认 交换机健康采集工具/_static/vis-network.min.js 文件存在（700KB左右），然后重新跑脚本生成新的拓扑图 HTML（会在同级生成 _static/ 目录，一起拷走）。</li>
      <li>或者：把 HTML 和 _static 目录放在同一级目录，双击 HTML 即可离线打开。</li>
      <li>临时应急：HTML 白屏不影响数据，所有链路/带宽/光功率都在 交换机健康状态总表.xlsx 的「端口带宽与光功率」和「异常汇总」Sheet 里。</li>
    </ol>
    <button onclick="document.getElementById('visErr').style.display='none'">关闭</button>
  </div>
</div>
<script>
  setTimeout(function(){
    if (!window.vis) {
      var e = document.getElementById('visErr');
      e && (e.style.display = 'flex');
    }
  }, 4000);
</script>
"""
    html = f"""<!DOCTYPE html>
<html lang='zh-CN'><head>
<meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=edge'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>网络拓扑图 - {datetime.now().strftime('%Y-%m-%d')}</title>
<style>
body{{font-family:'Microsoft YaHei',Arial;margin:0;background:#F7F9FC;color:#222;}}
#toolbar{{padding:10px 16px;background:#fff;border-bottom:1px solid #e6e6e6;display:flex;gap:12px;flex-wrap:wrap;align-items:center}}
.legend span{{display:inline-block;width:14px;height:14px;margin-right:4px;vertical-align:middle;border:1px solid #333;border-radius:2px}}
.core-bg{{background:#2E7D32}}.warn-bg{{background:#F57C00}}.err-bg{{background:#C62828}}.edge-err{{border:2px solid #C62828;display:inline-block;width:22px;height:0;vertical-align:middle}}
.summary{{padding:8px 16px;background:#FFF8E1;border-bottom:1px solid #FFE58F;font-size:13px}}
#net{{width:100%;height:calc(100vh - 100px);background:#fff}}
button{{padding:6px 12px;border:1px solid #d0d7de;border-radius:4px;background:#fff;cursor:pointer}}
button:hover{{background:#f6f8fa}}
</style>
{chr(10).join(vis_load_scripts)}
</head>
<body>
<div id="toolbar">
  <div><b>华三交换机网络拓扑</b>（可拖拽/缩放，悬停节点看详情，悬停边看链路端口）</div>
  <div class="legend">
    图例: <span class="core-bg"></span>健康/正常设备 &nbsp;
    <span class="warn-bg"></span>告警设备 &nbsp;
    <span class="err-bg"></span>错误/异常设备 &nbsp;
    <span class="edge-err"></span>异常链路（端口带宽/光功率异常）
  </div>
  <div style="margin-left:auto">
    <button onclick="network.fit()">全图适配</button>
    <button onclick="network.setOptions({{physics:{{enabled:!network.physics.options.enabled}}}})">物理模拟开关</button>
    <button onclick="savePNG()">导出PNG截图</button>
  </div>
</div>
{summary_html}
<div id="net"></div>
{load_err_block}
<script>
if (!window.vis) {{ document.getElementById('visErr').style.display = 'flex'; }}
const nodes = new vis.DataSet({json.dumps(nodes, ensure_ascii=False)});
const edges = new vis.DataSet({json.dumps(edges, ensure_ascii=False)});
const container = document.getElementById('net');
const data = {{nodes, edges}};
const options = {{
  interaction: {{hover: true, tooltipDelay: 100, zoomView: true, dragView: true}},
  physics: {{barnesHut: {{gravitationalConstant: -8000, springLength: 220, springConstant: 0.04}},
            stabilization: {{iterations: 200}}, solver: 'barnesHut'}},
  edges: {{smooth: {{type: 'cubicBezier', roundness: 0.55}}, arrows: {{to: {{enabled: false}}}}}},
  layout: {{improvedLayout: true}}
}};
const network = new vis.Network(container, data, options);
function savePNG(){{
  try {{
    network.once('afterDrawing', function(){{
      const c = document.querySelector('#net canvas');
      if (!c) {{ alert('当前渲染器未生成Canvas，请直接浏览器Ctrl+P打印或截图保存'); return; }}
      const url = c.toDataURL('image/png');
      const a = document.createElement('a'); a.href = url; a.download='拓扑截图_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png';
      document.body.appendChild(a); a.click(); a.remove();
    }});
    network.redraw();
  }} catch(e){{ alert(e); }}
}}
</script>
</body></html>"""
    with open(TOPO_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    log(f'拓扑图已输出: {TOPO_FILE}')
    log(f'  离线模式: {"已启用（HTML同级目录有 _static/vis-network.min.js，可以双击直接打开，断网也能用）" if have_local else "未启用（HTML必须能访问 CDN；如果打不开，请确认 交换机健康采集工具/_static/ 目录下有 vis-network.min.js，再重跑脚本）"}')


def main():
    import argparse as _argparse
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if RUN_LOG.exists():
        try:
            RUN_LOG.unlink()
        except Exception:
            pass
    log('===== 华三交换机配置采集 + 健康检查 + 拓扑图 启动 =====')

    parser = _argparse.ArgumentParser(description='华三交换机配置+健康+拓扑采集工具', formatter_class=_argparse.RawTextHelpFormatter,
                                      epilog="""示例：
  1) 默认跑 switch_table.xlsx 全部设备：
     python h3c_switch_health_topology.py
  2) 只跑 172.17.115.80 这台（按 IP）：
     python h3c_switch_health_topology.py --ip 172.17.115.80
  3) 只跑名字叫 EBG-3FJ01-SA-SWITCH-03 这台（按名称，精确匹配/包含都行）：
     python h3c_switch_health_topology.py --name EBG-3FJ01-SA-SWITCH-03
  4) 只跑 switch_table.xlsx 的第 21 行（序号=21，或 Excel 行号）：
     python h3c_switch_health_topology.py --row 21
  5) 组合使用：--ip 172.17.115.80,172.17.115.92 同时跑多台（逗号分隔）
        注意：使用 --ip/--name/--row 任一过滤时，其他没被选中的设备在台账/健康总表里的已有内容会保留，
              只更新目标设备；配置文件只重写目标设备的；拓扑图按所有已知设备重建。""")
    parser.add_argument('--ip', default='', help='只采集指定 IP（多台用英文逗号分隔）')
    parser.add_argument('--name', default='', help='只采集指定设备名称（支持子串匹配，多台用英文逗号分隔）')
    parser.add_argument('--row', default='', help='只采集 switch_table.xlsx 指定行号（1 为表头后第 1 行，多台用英文逗号分隔）')
    args = parser.parse_args()

    devices_all = read_switch_table()
    if not devices_all:
        log('[错误] 没有读取到任何交换机清单。请确保 switch_table.xlsx 至少包含列：设备名称 / IP地址 / 用户名 / 密码')
        sys.exit(1)

    filter_mode = bool(args.ip or args.name or args.row)
    skip_need_false = []
    run_default = []
    need_merge = False
    devices = []
    if not filter_mode:
        skip_need_false = [d for d in devices_all if not d.get('need_check', True)]
        run_default = [d for d in devices_all if d.get('need_check', True)]
        if skip_need_false:
            log(f'[是否需要检查列] 跳过 {len(skip_need_false)} 台（列值=否/跳过/0/不需要）：')
            for d in skip_need_false[:20]:
                log(f'  跳过: {d["name"]} ({d["host"]})')
            if len(skip_need_false) > 20:
                log(f'  …（另有 {len(skip_need_false)-20} 台省略）')
            log(f'  如需强制采集，请把对应行“是否需要检查”改为“是”，或使用 --ip/--name/--row 参数。')
        devices = list(run_default)
        need_merge = bool(skip_need_false) or (len(devices) < len(devices_all))
        if need_merge and not CACHE_FILE.exists() and not (LEDGER_FILE.exists() and HEALTH_FILE.exists()):
            need_merge = False
    else:
        keep = [False] * len(devices_all)
        if args.ip:
            ips = {x.strip() for x in args.ip.split(',') if x.strip()}
            for i, d in enumerate(devices_all):
                if d['host'] in ips:
                    keep[i] = True
        if args.name:
            names = [x.strip() for x in args.name.split(',') if x.strip()]
            for i, d in enumerate(devices_all):
                if keep[i]:
                    continue
                for pat in names:
                    if not pat:
                        continue
                    if pat == d['name'] or pat in d['name']:
                        keep[i] = True
                        break
        if args.row:
            try:
                rows = {int(x.strip()) for x in args.row.split(',') if x.strip()}
            except Exception:
                rows = set()
            for i, d in enumerate(devices_all):
                if keep[i]:
                    continue
                if (d.get('_row_index') or (i + 1)) in rows:
                    keep[i] = True
        devices = [d for d, flag in zip(devices_all, keep) if flag]
        if not devices:
            log(f'[错误] 按过滤条件（ip={args.ip!r} name={args.name!r} row={args.row!r}）没有匹配到任何设备，请检查 switch_table.xlsx 内容。')
            sys.exit(1)
        log(f'[手动指定过滤] 将只采集 {len(devices)} 台设备（忽略“是否需要检查”列），其他 {len(devices_all)-len(devices)} 台的台账/健康/端口/拓扑/异常 将保留旧值（若存在）。')
        for d in devices:
            log(f'  目标设备: {d["name"]} ({d["host"]})')
        need_merge = True

    # 【合并旧数据】（过滤模式或 need_check 跳过模式都启用）：先读无损 JSON 缓存，再读 Excel 兜底 4 Sheet
    existing_by_host = {}
    if need_merge:
        existing_by_host = load_all_data_cache()
        cache_hosts = set(existing_by_host.keys())
        excel_existing = load_health_from_excel()
        for h, ed in excel_existing.items():
            if h not in existing_by_host:
                existing_by_host[h] = ed
            else:
                for k in ('name', '登录状态', '采集时间', 'version_info', '备注', 'cpu_pct', 'mem_pct', 'temps', 'ports', 'lldp', 'errors', 'counters_method', 'counters_sample_sec'):
                    if not existing_by_host[h].get(k) and ed.get(k):
                        existing_by_host[h][k] = ed[k]
        log(f'[合并旧数据] 最终得到 {len(existing_by_host)} 台历史设备：缓存{len(cache_hosts)}台 + Excel还原{len(excel_existing)}台。')

    # 最终安全检查：devices 必须是 list 且有内容（否则直接退出，避免 NameError 或空跑）
    if not isinstance(devices, list) or len(devices) == 0:
        log('[错误] 本次没有待采集的设备（可能"是否需要检查"列全为否且没有手动指定过滤参数）。如需强制采集请指定 --ip/--name/--row。')
        sys.exit(1)

    all_data = []
    collected_hosts = set()
    for idx, d in enumerate(devices, 1):
        log(f'==== [{idx}/{len(devices)}] 开始采集: {d["name"]} ({d["host"]}) ====')
        r = collect_one(d)
        all_data.append(r)
        collected_hosts.add(r['host'])
        time.sleep(0.6)

    # 合并：把本次采集到的结果替换进 existing_by_host，然后整体输出
    if need_merge:
        for r in all_data:
            existing_by_host[r['host']] = r
        merged = []
        order_hosts = list(dict.fromkeys([d['host'] for d in devices_all] + list(existing_by_host.keys())))
        for h in order_hosts:
            if h not in existing_by_host:
                continue
            item = existing_by_host[h]
            if not isinstance(item, dict):
                try:
                    item = dict(item)
                except Exception:
                    continue
            for k, v in {'name': item.get('name', ''), 'host': item.get('host', ''), 'is_core': item.get('is_core', False),
                         '采集时间': item.get('采集时间', ''), '登录状态': item.get('登录状态', '失败'), 'version_info': item.get('version_info', {}) if isinstance(item.get('version_info', {}), dict) else {},
                         'config_txt': item.get('config_txt', ''), 'cpu_pct': float(item.get('cpu_pct', 0.0) or 0.0),
                         'mem_pct': float(item.get('mem_pct', 0.0) or 0.0), 'mem_used': float(item.get('mem_used', 0.0) or 0.0),
                         'mem_total': float(item.get('mem_total', 0.0) or 0.0),
                         'temps': item.get('temps', []) if isinstance(item.get('temps', []), list) else [],
                         'ports': item.get('ports', []) if isinstance(item.get('ports', []), list) else [],
                         'lldp': item.get('lldp', []) if isinstance(item.get('lldp', []), list) else [],
                         'errors': item.get('errors', []) if isinstance(item.get('errors', []), list) else [],
                         'counters_method': item.get('counters_method', '') or '',
                         'counters_sample_sec': int(item.get('counters_sample_sec', 30) or 30),
                         '温度告警阈值': item.get('温度告警阈值') if isinstance(item.get('温度告警阈值'), dict) else {},
                         '温度错误阈值': item.get('温度错误阈值') if isinstance(item.get('温度错误阈值'), dict) else {},
                         '备注': item.get('备注', '')}.items():
                item.setdefault(k, v)
            merged.append(item)
        all_data = merged

    log('采集完成，开始写出结果...')
    try:
        write_config_txt(all_data, collected_hosts if need_merge else None)
    except Exception as e:
        log(f'[WARN] 写配置文件出错：{e}')
    try:
        write_ledger(all_data)
    except Exception as e:
        log(f'[WARN] 写台账出错：{e}')
    try:
        write_health(all_data)
    except Exception as e:
        log(f'[WARN] 写健康总表出错：{e}')
    try:
        nodes, edges, _ = build_topo(all_data)
        write_topo_html(nodes, edges, all_data)
    except Exception as e:
        log(f'[WARN] 生成拓扑图出错：{e}')
    try:
        dump_all_data_cache(all_data)
    except Exception as e:
        log(f'[WARN] 写入缓存失败（忽略）：{e}')
    try:
        log('开始回写 switch_table.xlsx 的"是否需要检查"列...')
        write_switch_table_check_flag(all_data)
    except Exception as e:
        log(f'[WARN] 回写"是否需要检查"列出错（通常是 Excel 被打开，关闭即可）：{e}')
    total_ok = sum(1 for d in all_data if isinstance(d, dict) and (d.get('登录状态', '') == '成功'))
    log(f'===== 全部完成：成功采集 {total_ok}/{len(all_data)} 台 =====')
    log(f'台账: {LEDGER_FILE}')
    log(f'健康: {HEALTH_FILE}')
    log(f'配置: {CONFIG_DIR}')
    log(f'拓扑图: {TOPO_FILE}（直接用浏览器打开）')
    if filter_mode:
        log(f'[手动指定过滤] 实际本次新采集 {len(collected_hosts)} 台，其他 {max(0,len(all_data)-len(collected_hosts))} 台保留了旧数据（端口/拓扑/异常/健康全部未被覆盖）。')
    elif len(devices_all) - len(devices) > 0:
        log(f'[是否需要检查列] 实际本次新采集 {len(collected_hosts)} 台，自动跳过 {max(0,len(devices_all)-len(devices))} 台（列值=否/跳过/0/不需要），已保留旧数据。')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        log(f'[FATAL] {e}\n{traceback.format_exc()}')
        sys.exit(2)
