import os
import sys
import json
import datetime
import glob
import pandas as pd

ROOT = os.path.abspath(os.path.dirname(__file__))
OUTFILE = os.path.join(ROOT, '在职人员权限提取结果.xlsx')
UNMATCHED_FILE = os.path.join(ROOT, '在职人员-未匹配权限表.xlsx')
MATCHED_DIR = os.path.join(ROOT, '结果')
UNMATCHED_PERM_DIR = os.path.join(ROOT, '未匹配权限表')
PERM_DIR = os.path.join(ROOT, '权限表')

rpt = []


def P(s=''):
    rpt.append(str(s))


P('== 权限提取结果验证 ==')
P('时间: ' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
P('')

for name, p in [('汇总结果xlsx', OUTFILE), ('未匹配人员xlsx', UNMATCHED_FILE),
                ('匹配权限表目录', MATCHED_DIR), ('未匹配权限表目录', UNMATCHED_PERM_DIR)]:
    exists = os.path.exists(p)
    if os.path.isfile(p):
        P(f'{name}: EXISTS  size={os.path.getsize(p)} bytes  path={p}')
    elif os.path.isdir(p):
        n_files = len([x for x in os.listdir(p) if os.path.isfile(os.path.join(p, x))])
        P(f'{name}: EXISTS  files={n_files}  path={p}')
    else:
        P(f'{name}: MISSING  path={p}')
P('')


def report_excel(path, tag):
    if not os.path.isfile(path):
        return
    P(f'-- [{tag}] sheet统计 --')
    try:
        xl = pd.ExcelFile(path, engine='openpyxl')
        total_rows = 0
        P(f'  sheet_names: {xl.sheet_names}')
        for s in xl.sheet_names:
            df = pd.read_excel(path, sheet_name=s, engine='openpyxl', dtype=str)
            P(f'  sheet={s!r}  shape={df.shape[0]}行 x {df.shape[1]}列')
            total_rows += df.shape[0]
            if s == '人员权限汇总':
                col_set = list(df.columns)
                cols_to_show = ['匹配状态', '是否无权限(需补)', '无权限原因', '权限条目数合计', '涉及Sheet数']
                for cc in cols_to_show:
                    if cc in col_set:
                        vc = df[cc].fillna('').astype(str).value_counts().to_dict()
                        P(f'    [{cc}] 分布={vc}')
        P(f'  合计行数(所有sheet): {total_rows}')
    except Exception as e:
        P(f'  ERR: {type(e).__name__}: {e}')
    P('')


report_excel(OUTFILE, '汇总结果')
report_excel(UNMATCHED_FILE, '未匹配人员单独表')


def n_xls(d):
    if not os.path.isdir(d):
        return 0
    cnt = 0
    for name in os.listdir(d):
        if name.lower().endswith('.xls'):
            cnt += 1
    return cnt


perm_total = 0
if os.path.isdir(PERM_DIR):
    perm_total = n_xls(PERM_DIR)
P(f'原权限目录 xls 文件数: {perm_total}')
P(f'结果目录(匹配上的) xls 文件数: {n_xls(MATCHED_DIR)}')
P(f'未匹配权限表目录 xls 文件数: {n_xls(UNMATCHED_PERM_DIR)}')
if perm_total > 0:
    P(f'  合并校验: {n_xls(MATCHED_DIR)} + {n_xls(UNMATCHED_PERM_DIR)} = {n_xls(MATCHED_DIR)+n_xls(UNMATCHED_PERM_DIR)} (原={perm_total})')
P('')

if os.path.isdir(UNMATCHED_PERM_DIR):
    P(f'-- 未匹配权限表目录 典型样例 --')
    items = sorted([x for x in os.listdir(UNMATCHED_PERM_DIR) if x.lower().endswith('.xls')])
    for x in items[:20]:
        P(f'  {x}')
    if len(items) > 20:
        P(f'  ... 另有 {len(items)-20} 个')

out_path = os.path.join(ROOT, 'perm_verify_report.txt')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(rpt))
print('验证报告写入:', out_path)
