import os
import sys
import shutil
import datetime
import pandas as pd
from openpyxl import load_workbook

if __package__ in (None, ''):
    sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from net_port_tool import config as np_config
from net_port_tool.analyzer import filter_access_terminals

OUTPUT_FILE = str(np_config.OUTPUT_FILE)


def guess_core_name_from_workbook(path):
    """优先从命名验证/交换机连线sheet判断核心交换机名。"""
    try:
        df_nv = pd.read_excel(path, sheet_name='命名验证')
        if not df_nv.empty:
            for col in ['期望名称', '实际名称']:
                if col in df_nv.columns:
                    vals = [str(x).strip() for x in df_nv[col].tolist() if str(x).strip() not in ('nan', '')]
                    for v in vals:
                        up = v.upper()
                        if any(k in up for k in ['CORE', 'ROOM', '核心', 'IPSAN', 'SERVER']):
                            return v
            # 退而求其次：如果连线里出现最多次的“本地交换机”，大概率是核心
            try:
                df_lk = pd.read_excel(path, sheet_name='交换机连线')
                if not df_lk.empty and '本地交换机' in df_lk.columns:
                    top = df_lk['本地交换机'].astype(str).value_counts().idxmax()
                    if str(top).strip() not in ('', 'nan'):
                        return str(top).strip()
            except Exception:
                pass
            return vals[0] if vals else ''
    except Exception:
        return ''


def ensure_file_not_locked(path):
    """尝试改名再改回来判断是否被 Excel 占用；尝试关闭当前 Excel 中同名工作簿。"""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client as win32
            excel = win32.GetActiveObject('Excel.Application')
            for wb in list(excel.Workbooks):
                full = os.path.abspath(path)
                try:
                    wb_full = os.path.abspath(wb.FullName)
                except Exception:
                    continue
                if os.path.normcase(wb_full) == os.path.normcase(full):
                    print(f'[WARN] 检测到 Excel 已打开: {full}，尝试自动关闭...')
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass


def main():
    xlsx_path = OUTPUT_FILE
    if not os.path.isfile(xlsx_path):
        print(f'[FAIL] 找不到输出文件: {xlsx_path}')
        return 1
    ensure_file_not_locked(xlsx_path)

    # 备份
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = xlsx_path.replace('.xlsx', '') + f'.reanalyze_{ts}.bak.xlsx'
    shutil.copy(xlsx_path, backup)
    print(f'[OK] 已备份原文件: {backup}')

    print(f'读取: {xlsx_path}')
    df_orig = pd.read_excel(xlsx_path, sheet_name='终端接入表')
    print(f'当前“终端接入表”行数: {len(df_orig)}')

    core_name = guess_core_name_from_workbook(xlsx_path)
    if not core_name:
        print('[WARN] 无法自动判断核心交换机名，核心异常分类会减弱。')
    else:
        print(f'[OK] 识别核心交换机名: {core_name}')

    kept, anom, summary = filter_access_terminals(df_orig, core_name)

    print('\n=== 终端接入-分析统计 ===')
    print(summary.to_string(index=False))

    print('\n异常类型分布:')
    if anom.empty:
        print('  (无异常)')
    else:
        print(anom.groupby('异常类型').size().sort_values(ascending=False).to_string())
        top_port_anom = anom[anom['异常类型'].astype(str).str.contains('同一物理口挂多终端', na=False)]
        if not top_port_anom.empty:
            print('\n同一物理口多终端 TOP 端口:')
            grp = top_port_anom.groupby(['所在交换机', '所在端口']).size().sort_values(ascending=False).head(20)
            print(grp.to_string())

    # 写回 Excel
    base_cols = ['设备类型', '设备名称', 'IP地址', 'MAC地址', '所在交换机', '所在端口', 'VLAN', '备注']
    out_kept = kept.copy()
    for c in base_cols:
        if c not in out_kept.columns:
            out_kept[c] = ''
    out_kept = out_kept[base_cols]

    existing_sheets = []
    try:
        wb = load_workbook(xlsx_path, read_only=True)
        existing_sheets = list(wb.sheetnames)
        wb.close()
    except Exception:
        pass

    # 保留除我们要替换以外的所有 sheet 原样 —— 使用 openpyxl 先读再重写太麻烦，
    # 这里直接用 openpyxl load + pandas replace 三张 sheet；因为 if_sheet_exists='replace'
    try:
        with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            summary.to_excel(writer, sheet_name='终端接入-分析统计', index=False)
            out_kept.to_excel(writer, sheet_name='终端接入表', index=False)
            anom.to_excel(writer, sheet_name='终端接入-异常明细', index=False)
    except PermissionError as e:
        print(f'\n[FAIL] 写入失败，文件仍被占用: {e}')
        print('请先在 Excel 中关闭该文件后再执行，或直接用备份文件对比。')
        return 2

    print(f'\n[OK] 结果已写回: {xlsx_path}')
    print(f'  - 终端接入-分析统计: {len(summary)} 行')
    print(f'  - 终端接入表(已过滤): {len(out_kept)} 行')
    print(f'  - 终端接入-异常明细: {len(anom)} 行')
    return 0


if __name__ == '__main__':
    sys.exit(main())
