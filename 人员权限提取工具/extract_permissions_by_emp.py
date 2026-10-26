import os
import sys
import glob
import shutil
import argparse
import traceback
import datetime
from typing import List, Tuple, Optional

import pandas as pd


DEFAULT_EMP_FILE = '在职人员.xlsx'
DEFAULT_PERM_DIR = '权限表'
DEFAULT_OUTPUT_FILE = '在职人员权限提取结果.xlsx'
DEFAULT_UNMATCHED_FILE = '在职人员-未匹配权限表.xlsx'
DEFAULT_MATCHED_PERM_DIR = '结果'
DEFAULT_UNMATCHED_PERM_DIR = '未匹配权限表'
PERM_FILE_EXT = '.xls'


def find_emp_id_column(df_columns) -> Optional[str]:
    candidate_keywords = ['人员编号', '员工编号', '工号', '人员工号', '员工号', 'person_id', 'emp_id', 'empno']
    cols = list(df_columns)
    lower_to_original = {str(c).strip().lower(): c for c in cols}
    for kw in candidate_keywords:
        if kw in cols:
            return kw
        low = kw.lower()
        if low in lower_to_original:
            return lower_to_original[low]
    for c in cols:
        s = str(c)
        if '编号' in s and ('人员' in s or '员工' in s or '工号' in s):
            return c
    return None


def read_employee_list(emp_path: str) -> Tuple[pd.DataFrame, str]:
    if not os.path.isfile(emp_path):
        raise FileNotFoundError(f'找不到在职人员表: {emp_path}')
    ext = os.path.splitext(emp_path)[1].lower()
    # 读全部 sheet，找第一列包含人员编号的
    if ext in ('.xlsx', '.xlsm'):
        engine = 'openpyxl'
    elif ext == '.xls':
        engine = 'xlrd'
    else:
        raise ValueError(f'不支持的在职人员表格式: {ext}')

    xl = pd.ExcelFile(emp_path, engine=engine)
    last_err = None
    for sheet_name in xl.sheet_names:
        try:
            df_raw = pd.read_excel(emp_path, sheet_name=sheet_name, engine=engine, dtype=str)
        except Exception as e:
            last_err = e
            continue
        id_col = find_emp_id_column(df_raw.columns)
        if id_col:
            # 去掉完全空的行
            df_raw = df_raw.dropna(axis=0, how='all').copy()
            df_raw[id_col] = df_raw[id_col].astype(str).str.strip()
            # 去掉空的人员编号
            df_raw = df_raw[
                ~df_raw[id_col].isin(['', 'nan', 'None', '人员编号', '员工编号', '工号'])
            ].copy()
            # 人员编号前后去空白字符(含非断空格)
            df_raw[id_col] = df_raw[id_col].astype(str).str.replace('\u3000', '', regex=False)
            print(f'[OK] 读取在职人员表: {emp_path}  sheet={sheet_name}  人员编号列={id_col}  行数={len(df_raw)}')
            return df_raw, id_col
    if last_err is not None:
        raise RuntimeError(f'未能在职人员表 [{emp_path}] 中找到“人员编号”列，最近异常：{last_err}')
    raise RuntimeError(f'未能在职人员表 [{emp_path}] 中找到“人员编号”列，请确认表头是否包含：人员编号 / 员工编号 / 工号')


def build_perm_file_index(perm_dir: str) -> dict:
    if not os.path.isdir(perm_dir):
        raise FileNotFoundError(f'找不到权限表目录: {perm_dir}')
    files = []
    for ext in (PERM_FILE_EXT, PERM_FILE_EXT.upper()):
        files.extend(glob.glob(os.path.join(glob.escape(perm_dir), f'*{ext}')))
    index = {}
    dup = []
    for f in files:
        base = os.path.splitext(os.path.basename(f))[0].strip()
        key = base.upper()
        if key in index:
            dup.append((base, f, index[key]))
            continue
        index[key] = f
    if dup:
        print(f'[WARN] 权限表存在同编号重复文件（保留第一个）:')
        for b, f1, f0 in dup[:10]:
            print(f'   编号 {b} 已取 {f0}  忽略 {f1}')
        if len(dup) > 10:
            print(f'   ... 另有 {len(dup)-10} 条')
    print(f'[OK] 权限表目录: {perm_dir}  共匹配到 {len(index)} 个 {PERM_FILE_EXT} 文件')
    return index


def _read_xls_with_xlrd(perm_file_path: str) -> List[Tuple[str, pd.DataFrame]]:
    import xlrd
    book = xlrd.open_workbook(perm_file_path, encoding_override='utf-8', formatting_info=False, on_demand=True)
    result: List[Tuple[str, pd.DataFrame]] = []
    for sheet_name in book.sheet_names():
        sh = book.sheet_by_name(sheet_name)
        rows = []
        for r in range(sh.nrows):
            row_vals = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = cell.value
                if v is None:
                    row_vals.append('')
                    continue
                if cell.ctype == 2:
                    if abs(v - int(v)) < 1e-9:
                        v = str(int(v))
                    else:
                        v = str(v)
                elif cell.ctype == 3:
                    try:
                        t = xlrd.xldate_as_datetime(v, book.datemode)
                        v = t.strftime('%Y-%m-%d %H:%M:%S')
                    except Exception:
                        v = str(v)
                elif cell.ctype == 4:
                    v = 'TRUE' if bool(v) else 'FALSE'
                else:
                    v = str(v) if str(v).strip().lower() != 'nan' else ''
                row_vals.append(v)
            rows.append(row_vals)
        book.unload_sheet(sheet_name)
        if not rows:
            continue

        scan_end = min(15, len(rows))

        def nonempty_count(row):
            return sum(1 for x in row if str(x).strip() not in ('', 'nan', 'None'))
        header_idx = max(range(scan_end), key=lambda i: nonempty_count(rows[i])) if scan_end else 0
        header_row = rows[header_idx] if header_idx < len(rows) else []
        if not header_row:
            continue
        data_rows = rows[header_idx + 1:]

        ncols = max(len(header_row), max((len(r) for r in data_rows), default=0))

        def pad(r):
            if len(r) < ncols:
                return list(r) + [''] * (ncols - len(r))
            return list(r[:ncols])
        header_row = pad(header_row)
        data_rows = [pad(r) for r in data_rows]

        cleaned_cols = []
        seen = {}
        for i, col in enumerate(header_row):
            c = str(col).strip()
            if c in ('', 'nan', 'None'):
                c = f'列{i + 1}'
            key = c
            count = seen.get(key, 0) + 1
            seen[key] = count
            if count > 1:
                c = f'{key}_{count}'
            cleaned_cols.append(c)

        filtered_rows = []
        for r in data_rows:
            if any(str(x).strip() not in ('', 'nan', 'None') for x in r):
                filtered_rows.append(r)
        if not filtered_rows:
            continue
        df = pd.DataFrame(filtered_rows, columns=cleaned_cols, dtype=str)
        result.append((sheet_name, df))
    book.release_resources()
    return result


def read_permission_file(perm_file_path: str) -> List[Tuple[str, pd.DataFrame]]:
    """读取单个权限文件，返回 list[(sheet_name, data_df)]。
    - 对 .xls 直接用 xlrd 原生解析（避开 pandas 2.x 与 xlrd 版本冲突）
    - 自动跳过空 sheet
    - 自动识别真正的表头行（前 15 行中，找“非空列数最多”的那一行当作表头）
    """
    ext = os.path.splitext(perm_file_path)[1].lower()
    if ext == '.xls':
        return _read_xls_with_xlrd(perm_file_path)

    engine = 'openpyxl'
    xl = pd.ExcelFile(perm_file_path, engine=engine)
    result: List[Tuple[str, pd.DataFrame]] = []
    for sheet_name in xl.sheet_names:
        try:
            raw = pd.read_excel(
                perm_file_path, sheet_name=sheet_name, engine=engine, dtype=str, header=None
            )
        except Exception as e:
            print(f'   [WARN] 读取失败 sheet={sheet_name}: {e}')
            continue
        if raw is None or raw.empty:
            continue
        scan_rows = min(15, raw.shape[0])
        nonempty_counts = raw.iloc[:scan_rows].apply(
            lambda r: sum(1 for x in r.astype(str).str.strip().tolist() if x not in ('', 'nan', 'None')),
            axis=1,
        ).tolist()
        if not nonempty_counts:
            continue
        header_idx = int(max(range(len(nonempty_counts)), key=lambda i: nonempty_counts[i]))
        header_row = raw.iloc[header_idx].astype(str).str.strip().tolist()
        data_only = raw.iloc[header_idx + 1:].copy()
        if data_only.empty:
            continue
        data_only.columns = list(range(len(header_row)))
        cleaned_cols = []
        seen = {}
        for i, col in enumerate(header_row):
            c = str(col)
            if c in ('', 'nan', 'None'):
                c = f'列{i + 1}'
            key = c
            count = seen.get(key, 0) + 1
            seen[key] = count
            if count > 1:
                c = f'{key}_{count}'
            cleaned_cols.append(c)
        data_only.columns = cleaned_cols
        mask_empty = data_only.apply(
            lambda r: all(str(v).strip() in ('', 'nan', 'None') for v in r.tolist()),
            axis=1,
        )
        data_only = data_only[~mask_empty].copy()
        if data_only.empty:
            continue
        data_only = data_only.reset_index(drop=True)
        result.append((sheet_name, data_only))
    return result


def write_ensure_closed(target_path: str, writer_factory):
    """避免 Excel 占用时的友好报错；如果占用自动加时间戳副本。"""
    backup_path = None
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    if os.path.isfile(target_path):
        try:
            os.rename(target_path, target_path + '.tmp_lock_check')
            os.rename(target_path + '.tmp_lock_check', target_path)
        except OSError:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                try:
                    import win32com.client as win32
                    try:
                        excel = win32.GetActiveObject('Excel.Application')
                    except Exception:
                        excel = None
                    if excel is not None:
                        full = os.path.abspath(target_path)
                        for wb in list(excel.Workbooks):
                            try:
                                wb_full = os.path.abspath(wb.FullName)
                            except Exception:
                                continue
                            if os.path.normcase(wb_full) == os.path.normcase(full):
                                print(f'[WARN] 检测到 Excel 占用 {target_path}，尝试自动关闭...')
                                try:
                                    wb.Close(SaveChanges=False)
                                except Exception:
                                    pass
                finally:
                    pythoncom.CoUninitialize()
            except Exception:
                pass
            base, ext = os.path.splitext(target_path)
            backup_path = f'{base}_{ts}{ext}'
            print(f'[WARN] 原文件仍被占用，改为输出到：{backup_path}')
            target_path = backup_path
    wb = writer_factory(target_path)
    return target_path, backup_path


def prepare_output_dir(dir_path: str) -> str:
    if not dir_path:
        return dir_path
    os.makedirs(dir_path, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    # 清掉旧内容 (只清空目标目录下直接子文件，避免删错上级目录)
    try:
        for name in os.listdir(dir_path):
            p = os.path.join(dir_path, name)
            if os.path.isfile(p):
                try:
                    os.remove(p)
                except Exception:
                    # 如果删不掉就改名
                    try:
                        os.rename(p, p + f'.old_{ts}')
                    except Exception:
                        pass
    except Exception:
        pass
    return dir_path


def safe_copy_file(src: str, dst_dir: str, new_name: Optional[str] = None) -> Optional[str]:
    try:
        os.makedirs(dst_dir, exist_ok=True)
        base_name = new_name if new_name else os.path.basename(src)
        dst = os.path.join(dst_dir, base_name)
        shutil.copy2(src, dst)
        return dst
    except Exception:
        return None


def main():
    parser = argparse.ArgumentParser(description='根据“在职人员表”的人员编号匹配权限表文件夹中的 人员编号.xls 权限，汇总提取。')
    parser.add_argument('--emp', default=DEFAULT_EMP_FILE, help=f'在职人员表路径（默认: {DEFAULT_EMP_FILE}）')
    parser.add_argument('--perm-dir', default=DEFAULT_PERM_DIR, help=f'权限表目录（默认: {DEFAULT_PERM_DIR}）')
    parser.add_argument('--output', default=DEFAULT_OUTPUT_FILE, help=f'汇总权限输出文件（默认: {DEFAULT_OUTPUT_FILE}）')
    parser.add_argument('--unmatched', default=DEFAULT_UNMATCHED_FILE, help=f'未匹配人员输出（默认: {DEFAULT_UNMATCHED_FILE}）')
    parser.add_argument('--matched-perm-dir', default=DEFAULT_MATCHED_PERM_DIR,
                        help=f'匹配上的权限表拷贝到这个文件夹（默认: {DEFAULT_MATCHED_PERM_DIR}）')
    parser.add_argument('--unmatched-perm-dir', default=DEFAULT_UNMATCHED_PERM_DIR,
                        help=f'未被任何在职人员匹配到的权限表单独存到这个文件夹（默认: {DEFAULT_UNMATCHED_PERM_DIR}）')
    args = parser.parse_args()

    print('\n' + '=' * 70)
    print('权限提取程序 启动')
    print('=' * 70)
    print(f'在职人员表         : {os.path.abspath(args.emp)}')
    print(f'权限表原目录       : {os.path.abspath(args.perm_dir)}')
    print(f'汇总权限结果xlsx     : {os.path.abspath(args.output)}')
    print(f'未匹配人员xlsx     : {os.path.abspath(args.unmatched)}')
    print(f'已匹配权限表目录  : {os.path.abspath(args.matched_perm_dir)}')
    print(f'未匹配权限表目录  : {os.path.abspath(args.unmatched_perm_dir)}')

    prepare_output_dir(args.matched_perm_dir)
    prepare_output_dir(args.unmatched_perm_dir)

    emp_df, id_col = read_employee_list(args.emp)
    perm_index = build_perm_file_index(args.perm_dir)

    emp_df['人员编号_匹配键'] = emp_df[id_col].astype(str).str.upper().str.strip()
    # 在职人员表 里唯一编号集合
    unique_emp_keys = list(emp_df['人员编号_匹配键'].unique())

    matched_rows_long = []          # 汇总-长表（每条权限一行，含人员信息+来源sheet+来源行号+权限行列）
    unmatched_rows = []             # 未匹配人员
    error_rows = []                 # 读取权限表报错的人

    matched_keys = set()                    # 真正“匹配上 + 有有效权限条目”的人员编号键
    perm_keys_touched_files_ok = set()      # 匹配上并成功从权限表中读取到条目(含空但存在)的权限文件键
    perm_keys_touched_any = set()           # 命中在职人员(含读错/空文件)的权限文件键

    total_emp = len(unique_emp_keys)
    for i, key in enumerate(unique_emp_keys, 1):
        perm_file = perm_index.get(key)
        subset = emp_df[emp_df['人员编号_匹配键'] == key]

        if perm_file is None:
            for _, r in subset.iterrows():
                d = r.to_dict()
                d.pop('人员编号_匹配键', None)
                d['匹配状态'] = '未找到权限文件'
                d['权限表文件名'] = ''
                d['是否无权限(需补)'] = '是'
                d['无权限原因'] = '在职人员表里没有对应人员编号.xls'
                unmatched_rows.append(d)
            print(f'[{i}/{total_emp}] {key}  未找到权限表')
            continue

        perm_keys_touched_any.add(key)
        base_name = os.path.basename(perm_file)
        try:
            sheets = read_permission_file(perm_file)
        except Exception as e:
            err = f'{type(e).__name__}: {str(e).strip()[:200]}'
            for _, r in subset.iterrows():
                d = r.to_dict()
                d.pop('人员编号_匹配键', None)
                d['匹配状态'] = '读取权限表异常'
                d['权限表文件名'] = base_name
                d['异常信息'] = err
                d['是否无权限(需补)'] = '是'
                d['无权限原因'] = '权限表读取失败，请检查文件是否损坏/加密'
                error_rows.append(d)
            print(f'[{i}/{total_emp}] {key}  读取失败: {err[:80]}')
            continue

        if not sheets:
            for _, r in subset.iterrows():
                d = r.to_dict()
                d.pop('人员编号_匹配键', None)
                d['匹配状态'] = '权限表为空（无数据）'
                d['权限表文件名'] = base_name
                d['是否无权限(需补)'] = '是'
                d['无权限原因'] = '权限表文件存在，但所有 sheet 均为空（无权限条目）'
                unmatched_rows.append(d)
            print(f'[{i}/{total_emp}] {key}  权限表文件为空')
            continue

        # 走到这里 = 匹配到权限表文件 + 读到了权限条目
        matched_keys.add(key)
        perm_keys_touched_files_ok.add(key)

        print(f'[{i}/{total_emp}] {key}  读取 {len(sheets)} 个sheet  权限行数合计={sum(s.shape[0] for _, s in sheets)}')
        for sheet_name, perm_df in sheets:
            perm_df = perm_df.copy()
            perm_df['__来源Sheet'] = sheet_name
            perm_df['__来源文件名'] = base_name
            perm_df['__来源行号(权限表内)'] = list(range(1, len(perm_df) + 1))
            for _, emp_row in subset.iterrows():
                emp_dict = emp_row.to_dict()
                emp_dict.pop('人员编号_匹配键', None)
                for _, p_row in perm_df.iterrows():
                    merged = {}
                    merged.update(emp_dict)
                    p_dict = p_row.to_dict()
                    merged['权限来源文件'] = p_dict.pop('__来源文件名')
                    merged['权限来源Sheet'] = p_dict.pop('__来源Sheet')
                    merged['权限来源行号'] = p_dict.pop('__来源行号(权限表内)')
                    for k, v in p_dict.items():
                        merged[f'权限::{k}'] = v
                    matched_rows_long.append(merged)

    # ===== 权限表文件：匹配上的拷到“结果”、未匹配任何在职人员的单独存到“未匹配权限表” =====
    copied_matched_cnt = 0
    copied_unmatched_cnt = 0
    for key, perm_file in perm_index.items():
        base = os.path.basename(perm_file)
        if key in perm_keys_touched_any:
            # 与在职人员编号匹配过的（不管空不空/错不错，都算匹配上的权限表）
            dst = safe_copy_file(perm_file, args.matched_perm_dir)
            if dst:
                copied_matched_cnt += 1
        else:
            # 未被任何在职人员命中
            dst = safe_copy_file(perm_file, args.unmatched_perm_dir)
            if dst:
                copied_unmatched_cnt += 1

    # ===== 汇总-按人展开（宽表摘要） =====
    summary_rows = []
    for _, emp_row in emp_df.iterrows():
        key = emp_row['人员编号_匹配键']
        d = emp_row.to_dict()
        d.pop('人员编号_匹配键', None)
        perm_hit = key in perm_keys_touched_any
        perm_ok = key in perm_keys_touched_files_ok
        if perm_ok:
            d['匹配状态'] = '已匹配'
            d['是否无权限(需补)'] = '否'
            d['无权限原因'] = ''
            d['权限表文件名'] = os.path.basename(perm_index[key]) if key in perm_index else ''
        else:
            # 从 unmatched/error 中捞一下状态和原因，保持一致
            sub_um = [r for r in unmatched_rows if str(r.get(id_col)).strip().upper() == key]
            sub_err = [r for r in error_rows if str(r.get(id_col)).strip().upper() == key]
            if sub_err:
                ref = sub_err[0]
                d['匹配状态'] = ref.get('匹配状态', '读取权限表异常')
                d['是否无权限(需补)'] = ref.get('是否无权限(需补)', '是')
                d['无权限原因'] = ref.get('无权限原因', '权限表读取失败')
                d['权限表文件名'] = ref.get('权限表文件名', '')
            elif sub_um:
                ref = sub_um[0]
                d['匹配状态'] = ref.get('匹配状态', '未找到权限文件')
                d['是否无权限(需补)'] = ref.get('是否无权限(需补)', '是')
                d['无权限原因'] = ref.get('无权限原因', '')
                d['权限表文件名'] = ref.get('权限表文件名', '')
            else:
                d['匹配状态'] = '未找到权限文件'
                d['是否无权限(需补)'] = '是'
                d['无权限原因'] = '在职人员表里没有对应人员编号.xls'
                d['权限表文件名'] = ''
        # 汇总统计（只有 真正读到了权限条目 的人才有数量）
        perms_same_key = [r for r in matched_rows_long if str(r.get(id_col)).strip().upper() == key]
        d['权限条目数合计'] = len(perms_same_key)
        sheets_set = []
        for r in perms_same_key:
            s = r.get('权限来源Sheet')
            if s and s not in sheets_set:
                sheets_set.append(s)
        d['涉及Sheet数'] = len(sheets_set)
        d['涉及Sheet列表'] = ';'.join(sheets_set)
        summary_rows.append(d)

    summary_df = pd.DataFrame(summary_rows)
    long_df = pd.DataFrame(matched_rows_long)
    unmatched_df = pd.DataFrame(unmatched_rows)
    error_df = pd.DataFrame(error_rows)

    # ===== 写汇总结果 =====
    def write_output(target):
        with pd.ExcelWriter(target, engine='openpyxl') as w:
            summary_df.to_excel(w, sheet_name='人员权限汇总', index=False)
            if not long_df.empty:
                # 为避免 sheet 名过长/非法，统一一下
                sheet = '权限明细(长表)'
                long_df.to_excel(w, sheet_name=sheet[:31], index=False)
            else:
                pd.DataFrame(columns=['无匹配权限条目']).to_excel(w, sheet_name='权限明细(长表)', index=False)
            # 未匹配+错误
            cols = list(summary_df.columns)
            cols_um = list(unmatched_df.columns) if not unmatched_df.empty else cols
            cols_err = list(error_df.columns) if not error_df.empty else (cols + ['异常信息'])
            pd.DataFrame(columns=cols_um).to_excel(w, sheet_name='未匹配人员', index=False) \
                if unmatched_df.empty else unmatched_df.to_excel(w, sheet_name='未匹配人员', index=False)
            pd.DataFrame(columns=cols_err).to_excel(w, sheet_name='权限读取错误', index=False) \
                if error_df.empty else error_df.to_excel(w, sheet_name='权限读取错误', index=False)
            # 未匹配详细：你要求的“单独一份表格”也在 unmatched 文件里写入了更纯版本
            # 这里再给一个“在职人员清单(保留原结构)”方便比对
            emp_out = emp_df.drop(columns=['人员编号_匹配键'], errors='ignore').copy()
            emp_out.to_excel(w, sheet_name='在职人员清单(原始)', index=False)

    final_output, used_alt = write_ensure_closed(args.output, lambda p: (write_output(p), p)[1])
    if not os.path.isfile(final_output):
        write_output(final_output)  # fallback，直接写一次

    # ===== 写“未匹配人员”单独一份表格 =====
    def write_unmatched(target):
        with pd.ExcelWriter(target, engine='openpyxl') as w:
            sheet_title_um_err = '未匹配-错误'
            if unmatched_df.empty and error_df.empty:
                pd.DataFrame(columns=[id_col, '匹配状态']).to_excel(w, sheet_name=sheet_title_um_err, index=False)
            else:
                all_um = []
                if not unmatched_df.empty:
                    all_um.append(unmatched_df.copy())
                if not error_df.empty:
                    all_um.append(error_df.copy())
                pd.concat(all_um, ignore_index=True, sort=False).to_excel(w, sheet_name=sheet_title_um_err, index=False)
            cols_keep = [id_col]
            for other in ['姓名', '部门名称', '部门编号', '手机号', '邮箱']:
                if other in summary_df.columns:
                    cols_keep.append(other)
            um_ids = []
            for r in unmatched_rows:
                um_ids.append({k: r.get(k, '') for k in cols_keep})
            for r in error_rows:
                um_ids.append({k: r.get(k, '') for k in cols_keep})
            if um_ids:
                pd.DataFrame(um_ids).drop_duplicates().to_excel(w, sheet_name='需要补权限的人员', index=False)
            else:
                pd.DataFrame(columns=cols_keep).to_excel(w, sheet_name='需要补权限的人员', index=False)

    final_unmatch, _ = write_ensure_closed(args.unmatched, lambda p: (write_unmatched(p), p)[1])
    if not os.path.isfile(final_unmatch):
        write_unmatched(final_unmatch)

    # ===== 控制台汇总 =====
    print('\n' + '=' * 70)
    print('提取完成')
    print('=' * 70)
    print(f'在职人员编号数: {total_emp}')
    print(f'成功匹配（含权限条目）: {len(matched_keys)}')
    print(f'未匹配/无权限 : {len(unmatched_df)}')
    print(f'权限读取异常  : {len(error_df)}')
    print(f'权限条目合计  : {len(long_df)}')
    print(f'汇总文件                 : {os.path.abspath(final_output)}')
    print(f'未匹配人员单独xlsx       : {os.path.abspath(final_unmatch)}')
    print(f'匹配上的权限表（拷贝到） : {os.path.abspath(args.matched_perm_dir)}   文件数: {copied_matched_cnt}')
    print(f'未匹配权限表（单独保存） : {os.path.abspath(args.unmatched_perm_dir)}  文件数: {copied_unmatched_cnt}')
    if used_alt:
        print(f'[WARN] 原汇总文件被 Excel 占用，已改为输出版本: {used_alt}')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        print('\n[FATAL] 程序异常退出：')
        traceback.print_exc()
        sys.exit(2)
