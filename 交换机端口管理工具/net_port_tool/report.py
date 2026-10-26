"""报告输出层：Sheet 名称与列名保持与原脚本一致，便于平滑替换。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, Optional, Sequence

import pandas as pd

from .analyzer import (
    ANOMALY_EXTRA_COLS,
    LINK_COLS,
    NAME_VERIFY_COLS,
    PORT_USAGE_COLS,
    SUSPICIOUS_COLS,
    TERMINAL_TABLE_COLS,
    AnalysisResult,
    build_raw_terminal_rows,
    filter_access_terminals,
)
from .parsers import port_sort_key

LOG = logging.getLogger('net_port_tool.report')

PRINT_COLS = ['交换机', '端口', 'VLAN', '端口使用情况', '终端设备名称', '终端IP']


def _sort_by_switch_port(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or '端口' not in df.columns:
        return df
    try:
        return df.sort_values(
            by=['交换机', '端口'],
            key=lambda s: s.map(port_sort_key) if s.name == '端口' else s.astype(str),
        )
    except Exception:  # noqa: BLE001 - 排序失败不影响导出
        return df


ALL_SECTIONS = ('name', 'links', 'terminal', 'usage', 'suspicious')


def write_report(output_path, result: AnalysisResult, core_name: str = '',
                 sections: Optional[Sequence[str]] = None,
                 logger: Optional[logging.Logger] = None) -> Dict[str, int]:
    """写出 Excel 报告，返回各 Sheet 行数统计。

    sections 用于按“功能模块”勾选导出，取值见 ALL_SECTIONS；None 表示全选。
    """
    log = logger or LOG
    out = Path(output_path)
    enabled = set(sections) if sections is not None else set(ALL_SECTIONS)

    port_usage_df = pd.DataFrame(result.port_usage)
    if port_usage_df.empty:
        port_usage_df = pd.DataFrame(columns=PORT_USAGE_COLS)

    port_usage_print_df = port_usage_df[
        [c for c in PRINT_COLS if c in port_usage_df.columns]
    ].copy() if not port_usage_df.empty else pd.DataFrame(columns=PRINT_COLS)

    port_usage_df = _sort_by_switch_port(port_usage_df)
    port_usage_print_df = _sort_by_switch_port(port_usage_print_df)

    raw_rows = build_raw_terminal_rows(result.port_data, core_name)
    raw_df = pd.DataFrame(raw_rows)
    kept_df, anomaly_df, summary_df = filter_access_terminals(raw_df, core_name)

    access_usage_df, core_usage_df = _split_by_core(port_usage_df, core_name, PORT_USAGE_COLS)
    access_usage_print_df, core_usage_print_df = _split_by_core(
        port_usage_print_df, core_name, PRINT_COLS)

    suspicious_df = pd.DataFrame(result.suspicious)
    if suspicious_df.empty:
        suspicious_df = pd.DataFrame(columns=SUSPICIOUS_COLS)
    else:
        for c in SUSPICIOUS_COLS:
            if c not in suspicious_df.columns:
                suspicious_df[c] = ''
        suspicious_df = suspicious_df[SUSPICIOUS_COLS]
        suspicious_df = _sort_by_switch_port(suspicious_df)

    stats: Dict[str, int] = {}
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        wrote_any = False
        if 'name' in enabled:
            pd.DataFrame(result.name_verification or [], columns=NAME_VERIFY_COLS).to_excel(
                writer, sheet_name='命名验证', index=False)
            wrote_any = True
            stats['命名验证'] = len(result.name_verification or [])
        if 'links' in enabled:
            pd.DataFrame(result.links or [], columns=LINK_COLS).to_excel(
                writer, sheet_name='交换机连线', index=False)
            wrote_any = True
            stats['交换机连线'] = len(result.links or [])
        if 'terminal' in enabled:
            summary_df.to_excel(writer, sheet_name='终端接入-分析统计', index=False)
            (kept_df if not kept_df.empty else pd.DataFrame(columns=TERMINAL_TABLE_COLS)).to_excel(
                writer, sheet_name='终端接入表', index=False)
            if not anomaly_df.empty:
                anomaly_df.to_excel(writer, sheet_name='终端接入-异常明细', index=False)
            else:
                pd.DataFrame(columns=TERMINAL_TABLE_COLS + ANOMALY_EXTRA_COLS).to_excel(
                    writer, sheet_name='终端接入-异常明细', index=False)
            wrote_any = True
            stats['终端接入表'] = int(len(kept_df))
            stats['终端接入-异常明细'] = int(len(anomaly_df))
        if 'suspicious' in enabled:
            suspicious_df.to_excel(writer, sheet_name='可疑端口清单', index=False)
            wrote_any = True
            stats['可疑端口清单'] = int(len(suspicious_df))
        if 'usage' in enabled:
            (access_usage_df if not access_usage_df.empty else pd.DataFrame(columns=PORT_USAGE_COLS)).to_excel(
                writer, sheet_name='端口使用表', index=False)
            if not access_usage_print_df.empty:
                access_usage_print_df.to_excel(writer, sheet_name='端口使用表-打印', index=False)
            if not core_usage_df.empty:
                core_usage_df.to_excel(writer, sheet_name='核心端口使用表', index=False)
            if not core_usage_print_df.empty:
                core_usage_print_df.to_excel(writer, sheet_name='核心端口使用表-打印', index=False)
            wrote_any = True
            stats['端口使用表'] = int(len(access_usage_df))
            stats['核心端口使用表'] = int(len(core_usage_df))
        if not wrote_any:
            pd.DataFrame([{'提示': '未勾选任何输出模块'}]).to_excel(
                writer, sheet_name='说明', index=False)

    log.info('报告已输出: %s', out)
    return stats


def _split_by_core(df: pd.DataFrame, core_name: str, cols: Sequence[str]):
    if df is None or df.empty or '交换机' not in df.columns or not core_name:
        return df if df is not None else pd.DataFrame(columns=list(cols)), pd.DataFrame(columns=list(cols))
    core_df = df[df['交换机'] == core_name].copy()
    access_df = df[df['交换机'] != core_name].copy()
    return access_df, core_df
