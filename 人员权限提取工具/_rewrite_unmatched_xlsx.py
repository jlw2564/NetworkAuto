import os
import sys
import pandas as pd
ROOT = os.path.abspath(os.path.dirname(__file__))
OUT = os.path.join(ROOT, '在职人员权限提取结果.xlsx')
TARGET = os.path.join(ROOT, '在职人员-未匹配权限表.xlsx')
sheet1 = '未匹配-错误'
sheet2 = '需要补权限的人员'
xl = pd.ExcelFile(OUT, engine='openpyxl')
df_um = pd.read_excel(OUT, sheet_name='未匹配人员', engine='openpyxl',
                      dtype=str) if '未匹配人员' in xl.sheet_names else pd.DataFrame()
df_err = pd.read_excel(OUT, sheet_name='权限读取错误', engine='openpyxl',
                       dtype=str) if '权限读取错误' in xl.sheet_names else pd.DataFrame()
summary = pd.read_excel(OUT, sheet_name='人员权限汇总', engine='openpyxl',
                        dtype=str) if '人员权限汇总' in xl.sheet_names else pd.DataFrame()
if os.path.isfile(TARGET):
    try:
        os.remove(TARGET)
    except Exception:
        ts = __import__('datetime').datetime.now().strftime('%Y%m%d_%H%M%S')
        TARGET = os.path.join(ROOT, f'在职人员-未匹配权限表_{ts}.xlsx')
with pd.ExcelWriter(TARGET, engine='openpyxl') as w:
    all_um = []
    if not df_um.empty:
        all_um.append(df_um.copy())
    if not df_err.empty:
        all_um.append(df_err.copy())
    if all_um:
        pd.concat(all_um, ignore_index=True, sort=False).to_excel(w, sheet_name=sheet1, index=False)
    else:
        pd.DataFrame(columns=['人员编号', '匹配状态']).to_excel(w, sheet_name=sheet1, index=False)
    cols_keep = [c for c in ['人员编号', '姓名', '部门名称', '部门编号', '手机号', '邮箱'] if c in summary.columns]
    frames = []
    if not df_um.empty and set(cols_keep).issubset(df_um.columns):
        frames.append(df_um[cols_keep].copy())
    elif not df_um.empty:
        frames.append(df_um.copy())
    if not df_err.empty and set(cols_keep).issubset(df_err.columns):
        frames.append(df_err[cols_keep].copy())
    elif not df_err.empty:
        frames.append(df_err.copy())
    if frames:
        pd.concat(frames, ignore_index=True, sort=False).drop_duplicates().to_excel(w, sheet_name=sheet2, index=False)
    else:
        pd.DataFrame(columns=cols_keep).to_excel(w, sheet_name=sheet2, index=False)
print('OUTPUT', TARGET, os.path.getsize(TARGET))
