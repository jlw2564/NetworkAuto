import os
import pandas as pd
import re

ROOT = os.path.abspath(os.path.dirname(__file__))


def resolve_port_data():
    """端口数据.xlsx 优先取本目录，其次取同级“交换机端口管理工具”目录。"""
    candidates = [
        os.path.join(ROOT, '端口数据.xlsx'),
        os.path.join(os.path.dirname(ROOT), '交换机端口管理工具', '端口数据.xlsx'),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    raise FileNotFoundError('未找到 端口数据.xlsx，请放到本目录下：' + ROOT)


def parse_switch_name(sw_name):
    """解析交换机名称，提取楼栋、楼层、区域"""
    parts = sw_name.split('-')

    building = ''
    floor = ''
    area = ''

    if len(parts) >= 2:
        building = parts[0]
        second_part = parts[1]

        # 尝试提取楼层 (格式: 1F, 2F, RF等)
        floor_match = re.search(r'(\d+F|RF)', second_part, re.IGNORECASE)
        if floor_match:
            floor = floor_match.group(1).upper()

        # 尝试提取区域 (格式: J01, J02, J05等)
        area_match = re.search(r'(J\d+)', second_part, re.IGNORECASE)
        if area_match:
            area = area_match.group(1).upper()

        # 如果没有找到，尝试其他可能的格式
        if not floor:
            # 尝试在其他位置找
            for part in parts:
                f_match = re.search(r'(\d+F|RF)', part, re.IGNORECASE)
                if f_match:
                    floor = f_match.group(1).upper()
                    break
        if not area:
            for part in parts:
                j_match = re.search(r'(J\d+)', part, re.IGNORECASE)
                if j_match:
                    area = j_match.group(1).upper()
                    break

    return building, floor, area


def main():
    print("读取端口数据.xlsx...")

    xl = pd.ExcelFile(resolve_port_data())

    # 读取两个sheet
    df_switches = pd.read_excel(xl, sheet_name='交换机')
    df_terminals = pd.read_excel(xl, sheet_name='正常数据')

    # 先修复终端数据的端口列
    print("处理终端数据...")
    # 构建一个字典: (交换机, 完整端口名) -> 终端信息
    terminal_dict = {}

    for idx, row in df_terminals.iterrows():
        sw = str(row['所在交换机'])
        # 重新构建完整端口名
        port_part1 = str(row['所在端口'])
        part2 = str(int(row['Unnamed: 6'])) if pd.notna(row['Unnamed: 6']) else '0'
        part3 = str(int(row['Unnamed: 7'])) if pd.notna(row['Unnamed: 7']) else '0'
        full_port = f"{port_part1}/{part2}/{part3}"

        key = (sw, full_port)
        terminal_dict[key] = {
            '设备名称': str(row['设备名称']) if pd.notna(row['设备名称']) else '',
            'IP地址': str(row['IP地址']) if pd.notna(row['IP地址']) else '',
            'MAC地址': str(row['MAC地址']) if pd.notna(row['MAC地址']) else '',
            'VLAN': str(row['VLAN']) if pd.notna(row['VLAN']) else '',
            '备注': str(row['备注']) if pd.notna(row['备注']) else ''
        }

    print(f"读取到 {len(df_switches)} 台交换机，{len(df_terminals)} 个终端")

    # 生成完整端口表
    all_ports = []

    for idx, row in df_switches.iterrows():
        sw_name = str(row['实际名称'])
        port_count = int(row['端口数量'])

        # 解析楼栋、楼层、区域
        building, floor, area = parse_switch_name(sw_name)

        # 生成1到port_count的端口
        for port_num in range(1, port_count + 1):
            # 构造端口名
            port_name = f"GE1/0/{port_num}"

            # 查找终端
            key = (sw_name, port_name)
            terminal = terminal_dict.get(key, None)

            if terminal:
                device_name = terminal['设备名称']
                ip = terminal['IP地址']
                mac = terminal['MAC地址']
                vlan = terminal['VLAN']
                remark = terminal['备注']
            else:
                device_name = ''
                ip = ''
                mac = ''
                vlan = ''
                remark = ''

            port_info = {
                '楼栋': building,
                '楼层': floor,
                '区域': area,
                '交换机': sw_name,
                '端口': port_num,
                '端口名称': port_name,
                '设备名称': device_name,
                'IP地址': ip,
                'MAC地址': mac,
                'VLAN': vlan,
                '备注': remark
            }

            all_ports.append(port_info)

    # 创建DataFrame
    result_df = pd.DataFrame(all_ports)

    # 排序
    result_df = result_df.sort_values(['楼栋', '楼层', '区域', '交换机', '端口'])

    # 保存
    output_file = os.path.join(ROOT, '交换机端口资料管理表.xlsx')
    result_df.to_excel(output_file, index=False)

    print(f"\n完成！结果已保存至: {output_file}")
    print(f"共生成 {len(result_df)} 个端口记录")

    # 统计已用端口数
    used_count = sum(1 for p in all_ports if p['设备名称'])
    print(f"\n统计:")
    print(f"交换机数: {len(df_switches)}")
    print(f"已用端口数: {used_count}")
    print(f"空闲端口数: {len(result_df) - used_count}")


if __name__ == "__main__":
    main()
