import pandas as pd
import re
import os
import glob

# ==========================================
# 配置
# ==========================================
INPUT_FOLDER = 'to_process'       # 存放待处理结果表的文件夹
OUTPUT_FILE = 'final_port_map.xlsx'  # 最终合并生成的端口示意图
TOTAL_PORTS = 48                  # 默认交换机总端口数


def get_port_number(port_str):
    """从端口字符串中提取数字，如 GE1/0/1 -> 1"""
    if pd.isna(port_str):
        return None
    match = re.search(r'/(\d+)$', str(port_str))
    if match:
        return int(match.group(1))
    # 处理直接是数字的情况
    digit_match = re.search(r'^(\d+)$', str(port_str))
    if digit_match:
        return int(digit_match.group(1))
    return None


def get_floor_id(sw_name):
    """从交换机名称中提取楼层标识，如 EBG-1FJ01-SA-SWITCH-01 -> 1FJ01"""
    if pd.isna(sw_name):
        return "Other"
    # 匹配名称中的楼层部分 (如 1FJ01, 2FJ03 等)
    match = re.search(r'(\d+FJ\d+)', str(sw_name))
    if match:
        return match.group(1)
    return "Other"


def main():
    # 确保路径是绝对路径，避免在不同环境下找不到文件夹
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(base_dir, INPUT_FOLDER)

    if not os.path.exists(input_path):
        os.makedirs(input_path)
        print(f"已创建文件夹 '{input_path}'，请将待处理的 Excel 文件放入其中。")
        return

    # 1. 扫描文件夹下所有的 Excel 文件 (忽略大小写)
    files = []
    for f in os.listdir(input_path):
        if f.lower().endswith('.xlsx') and not f.startswith('~$'):
            files.append(os.path.join(input_path, f))

    if not files:
        print(f"在 '{input_path}' 文件夹下未找到有效的 .xlsx 文件。")
        return

    print(f"发现 {len(files)} 个待处理文件: {[os.path.basename(f) for f in files]}")

    all_raw_data = []
    for f in files:
        try:
            print(f"正在读取: {os.path.basename(f)} ...")
            df = pd.read_excel(f)
            # 统一清理数据中的前后空格
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            all_raw_data.append(df)
        except Exception as e:
            print(f"读取文件 {f} 失败: {e}")

    if not all_raw_data:
        print("没有读取到任何有效数据。")
        return

    # 合并所有扫描到的原始数据
    combined_df = pd.concat(all_raw_data, ignore_index=True)
    print(f"原始数据合并完成，共计 {len(combined_df)} 条记录")

    # 2. 获取所有不重复的交换机名称
    switches = combined_df['交换机'].unique()
    switches = [sw for sw in switches if sw and str(sw) != '未找到' and str(sw) != 'nan']

    # 按楼层分组存储结果: { floor_id: [data_rows] }
    floor_groups = {}

    # 3. 核心逻辑：按交换机生成完整端口图
    for sw in sorted(switches):
        floor_id = get_floor_id(sw)
        if floor_id not in floor_groups:
            floor_groups[floor_id] = []

        print(f"正在处理交换机: {sw} (所属楼层: {floor_id})")
        sw_data = combined_df[combined_df['交换机'] == sw].copy()

        occupied_ports = {}
        for _, row in sw_data.iterrows():
            p_num = get_port_number(row.get('端口'))
            if p_num is not None:
                if p_num not in occupied_ports:
                    occupied_ports[p_num] = []
                occupied_ports[p_num].append(row)

        for i in range(1, TOTAL_PORTS + 1):
            if i in occupied_ports:
                rows = occupied_ports[i]
                for data in rows:
                    remark = str(data.get('备注', ''))
                    if len(occupied_ports[i]) > 1:
                        remark = f"【注意：该端口存在重复记录】 " + remark

                    floor_groups[floor_id].append({
                        '交换机': sw,
                        '端口': i,
                        '*设备名称': data.get('设备名称', ''),
                        '*IP/域名': data.get('IP', ''),
                        'VLAN': data.get('VLAN', ''),
                        '备注': remark
                    })
            else:
                floor_groups[floor_id].append({
                    '交换机': sw,
                    '端口': i,
                    '*设备名称': '',
                    '*IP/域名': '',
                    'VLAN': '',
                    '备注': ''
                })

    # 4. 分 Sheet 保存结果
    if floor_groups:
        with pd.ExcelWriter(OUTPUT_FILE) as writer:
            for floor_id, results in sorted(floor_groups.items()):
                sheet_df = pd.DataFrame(results)
                # Sheet 名称不能超过 31 个字符
                sheet_name = floor_id[:31]
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  已生成 Sheet: {sheet_name} (包含 {len(sheet_df)} 行)")

        print(f"\n批量处理完成！合并后的多 Sheet 结果已保存至: {OUTPUT_FILE}")
    else:
        print("未生成任何结果。")


if __name__ == "__main__":
    main()
