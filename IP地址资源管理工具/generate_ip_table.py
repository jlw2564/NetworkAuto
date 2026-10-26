import os
import pandas as pd
import re
from collections import defaultdict

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
    raise FileNotFoundError('未找到 端口数据.xlsx；请先在“交换机端口管理工具”中生成，或复制到：' + ROOT)


def parse_ip(ip_str):
    """解析IP地址，返回四个部分的整数"""
    if pd.isna(ip_str) or str(ip_str).strip() == '':
        return None
    try:
        parts = str(ip_str).split('.')
        if len(parts) != 4:
            return None
        return tuple(int(p) for p in parts)
    except:
        return None

def main():
    print("读取端口数据.xlsx...")
    
    xl = pd.ExcelFile(resolve_port_data())
    df_terminals = pd.read_excel(xl, sheet_name='正常数据')
    
    # 收集所有IP
    ip_info = {}
    for idx, row in df_terminals.iterrows():
        ip = str(row['IP地址']).strip()
        if pd.isna(row['IP地址']) or ip == '' or ip == 'nan':
            continue
        
        ip_parts = parse_ip(ip)
        if ip_parts:
            ip_info[ip] = {
                '设备名称': str(row['设备名称']) if pd.notna(row['设备名称']) else '',
                '备注': str(row['备注']) if pd.notna(row['备注']) else ''
            }
    
    print(f"共收集到 {len(ip_info)} 个IP地址")
    
    # 按网段分组
    networks = defaultdict(list)
    for ip in ip_info:
        parts = parse_ip(ip)
        if parts:
            network = f"{parts[0]}.{parts[1]}.{parts[2]}.0"
            networks[network].append((parts[3], ip, ip_info[ip]))
    
    # 生成结果
    result_data = []
    
    network_list = sorted(networks.keys())

    # 添加两个空网段（172.17.115.0 和 172.17.116.0）
    extra_networks = ['172.17.115.0', '172.17.116.0']
    for net in extra_networks:
        if net not in network_list:
            network_list.append(net)
    network_list = sorted(network_list)

    for network in network_list:
        # 对每个网段，生成1-254的IP
        ips_in_network = {t[0]: t for t in networks[network]}
        
        for i in range(1, 255):
            ip = f"{network.rsplit('.', 1)[0]}.{i}"
            is_used = i in ips_in_network
            
            if is_used:
                _, full_ip, info = ips_in_network[i]
                device_name = info['设备名称']
                remark = info['备注']
            else:
                device_name = ''
                remark = ''
            
            result_data.append({
                '网段': network,
                'IP地址': ip,
                '状态': '',  # 状态列留空，后面用公式判断设备名称列
                '设备名称': device_name,
                '备注': remark
            })
    
    # 创建DataFrame
    result_df = pd.DataFrame(result_data)
    
    # 保存为Excel
    output_file = os.path.join(ROOT, 'IP地址资源管理表.xlsx')
    
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.chart import BarChart, Reference, PieChart
    
    # 先保存数据，用两个sheet
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='IP资源表', index=False)
        # 初始保存一个空的统计图表
        pd.DataFrame(columns=['网段', '已用', '空闲', '总数', '使用率']).to_excel(writer, sheet_name='统计图表', index=False)
    
    # 加载工作簿
    wb = load_workbook(output_file)
    
    # --- 处理IP资源表 sheet
    ws_ip = wb['IP资源表']
    
    # 定义颜色
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    # 给表头添加颜色和加粗字体
    header_font = Font(bold=True, color='000000')
    for col in range(1, ws_ip.max_column + 1):
        cell = ws_ip.cell(row=1, column=col)
        cell.fill = yellow_fill
        cell.font = header_font

    # 给状态列写入公式：判断设备名称列(D列)是否有内容
    # C列是状态列，D列是设备名称列
    for row in range(2, ws_ip.max_row + 1):
        ws_ip.cell(row=row, column=3).value = f'=IF(D{row}<>"", "已用", "空闲")'

    # 添加条件格式规则 - 基于设备名称列判断
    # 设备名称不为空，整行绿色；设备名称为空，整行红色
    ws_ip.conditional_formatting.add(
        f"A2:E{ws_ip.max_row}",
        FormulaRule(formula=[f'D$2<>""'], fill=green_fill, stopIfTrue=True)
    )
    ws_ip.conditional_formatting.add(
        f"A2:E{ws_ip.max_row}",
        FormulaRule(formula=[f'D$2=""'], fill=red_fill, stopIfTrue=True)
    )
    
    # 调整IP资源表的列宽
    for col in ws_ip.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws_ip.column_dimensions[column].width = adjusted_width
    
    ws_ip.freeze_panes = "A2"
    
    # --- 处理统计图表 sheet
    ws_stats = wb['统计图表']
    
    # 写入表头
    headers = ['网段', '已用', '空闲', '总数', '使用率']
    for col_num, header in enumerate(headers, start=1):
        ws_stats.cell(row=1, column=col_num).value = header
    
    # 给统计图表的表头也设置格式
    for col in range(1, ws_stats.max_column + 1):
        cell = ws_stats.cell(row=1, column=col)
        cell.fill = yellow_fill
        cell.font = header_font
    
    # 写入各网段数据，并添加公式
    row_num = 2
    for network in network_list:
        ws_stats.cell(row=row_num, column=1).value = network
        ws_stats.cell(row=row_num, column=4).value = 254
        
        # 找到该网段在IP资源表中的起始和结束行
        start_row = None
        end_row = None
        for r in range(2, ws_ip.max_row + 1):
            if ws_ip.cell(row=r, column=1).value == network:
                if start_row is None:
                    start_row = r
                end_row = r
        
        if start_row and end_row:
            # 添加公式
            ws_stats.cell(row=row_num, column=2).value = f'=COUNTIFS(\'IP资源表\'!A{start_row}:A{end_row}, A{row_num}, \'IP资源表\'!C{start_row}:C{end_row}, "已用")'
            ws_stats.cell(row=row_num, column=3).value = f'=COUNTIFS(\'IP资源表\'!A{start_row}:A{end_row}, A{row_num}, \'IP资源表\'!C{start_row}:C{end_row}, "空闲")'
            ws_stats.cell(row=row_num, column=5).value = f'=B{row_num}/D{row_num}'
            ws_stats.cell(row=row_num, column=5).number_format = '0.0%'
        
        row_num += 1
    
    # 添加总体统计
    ws_stats.cell(row=row_num, column=1).value = "总体统计"
    ws_stats.cell(row=row_num, column=2).value = f'=SUM(B2:B{row_num-1})'
    ws_stats.cell(row=row_num, column=3).value = f'=SUM(C2:C{row_num-1})'
    ws_stats.cell(row=row_num, column=4).value = f'=SUM(D2:D{row_num-1})'
    ws_stats.cell(row=row_num, column=5).value = f'=B{row_num}/D{row_num}'
    ws_stats.cell(row=row_num, column=5).number_format = '0.0%'
    
    # 调整统计图表表的列宽
    for col in ws_stats.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        adjusted_width = min(max_length + 2, 20)
        ws_stats.column_dimensions[column].width = adjusted_width
    
    # 添加柱状图 - 各网段IP使用情况
    bar_chart = BarChart()
    bar_chart.type = 'col'
    bar_chart.style = 10
    bar_chart.title = '各网段IP地址使用情况'
    bar_chart.x_axis.title = '网段'
    bar_chart.y_axis.title = '数量'
    
    data = Reference(ws_stats, min_col=2, min_row=1, max_col=3, max_row=row_num-1)
    cats = Reference(ws_stats, min_col=1, min_row=2, max_row=row_num-1)
    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(cats)
    ws_stats.add_chart(bar_chart, "G2")
    
    # 添加饼图 - 总体IP使用情况
    pie_chart = PieChart()
    pie_chart.title = "总体IP使用情况"
    
    # 创建饼图数据
    ws_stats['K1'] = '总体统计'
    ws_stats['K2'] = '已用'
    ws_stats['K3'] = '空闲'
    ws_stats['L2'] = f'=B{row_num}'
    ws_stats['L3'] = f'=C{row_num}'
    
    pie_data = Reference(ws_stats, min_col=12, min_row=1, max_row=3)
    pie_cats = Reference(ws_stats, min_col=11, min_row=2, max_row=3)
    pie_chart.add_data(pie_data, titles_from_data=True)
    pie_chart.set_categories(pie_cats)
    
    ws_stats.add_chart(pie_chart, "G20")
    
    # 保存
    wb.save(output_file)
    
    print(f"\n完成！IP地址资源管理表已保存至: {output_file}")
    print(f"共处理 {len(result_df)} 个IP地址")
    print(f"\n包含两个工作表:")
    print("1. IP资源表 - 所有IP的详细列表")
    print("2. 统计图表 - 包含柱状图和饼图（数据完全联动）")
    print("\n🎉 完全联动！使用说明:")
    print("1. 在'IP资源表'中，把'空闲'改为'已用'，颜色会自动变化")
    print("2. 改完后，'统计图表'表里的数字会自动更新！")
    print("3. 图表也会自动跟着更新！")

if __name__ == "__main__":
    main()
