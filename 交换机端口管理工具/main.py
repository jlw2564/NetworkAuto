import pandas as pd
import re
import time
from netmiko import ConnectHandler

# ==========================================
# 1. 交换机登录配置 (针对华三 H3C Comware 设备)
# ==========================================
# 核心交换机配置 (用于查询 ARP)
CORE_SWITCH = {
    'device_type': 'hp_comware',  # 华三设备在 netmiko 中使用 hp_comware
    'host': '172.17.115.254',  # 核心交换机 IP
    'username': 'admin',
    'password': 'Inventec@123',
}
# 接入交换机列表 (用于查询 MAC 地址表)
ACCESS_SWITCHES = [
    {
        'name': 'EBG-2FJ03-SA-SWITCH-01',
        'device_type': 'hp_comware',
        'host': '172.17.115.19',
        'username': 'admin',
        'password': 'Inventec@123',
    },
    {
        'name': 'EBG-2FJ03-SA-SWITCH-02',
        'device_type': 'hp_comware',
        'host': '172.17.115.97',
        'username': 'admin',
        'password': 'Inventec@123',
    },
    # {
    #     'name': 'EBG-4FJ02-SA-SWITCH-03',
    #     'device_type': 'hp_comware',
    #     'host': '172.17.115.87',
    #     'username': 'admin',
    #     'password': 'Inventec@123',
    # },
    # {
    #     'name': 'EBG-3FJ04-SA-SWITCH-02',
    #     'device_type': 'hp_comware',
    #     'host': '172.17.115.83',
    #     'username': 'admin',
    #     'password': 'Inventec@123',
    # },
    #     {
    #         'name': 'SW-SBX05-SA-SWITCH-01',
    #         'device_type': 'hp_comware',
    #         'host': '172.17.115.36',
    #         'username': 'admin',
    #         'password': 'Inventec@123',
    #     },
    #     {
    #         'name': 'SW-SBX06-SA-SWITCH-01',
    #         'device_type': 'hp_comware',
    #         'host': '172.17.115.37',
    #         'username': 'admin',
    #         'password': 'Inventec@123',
    #     },
    #     {
    #         'name': 'SW-RDG01-SA-SWITCH-01',
    #         'device_type': 'hp_comware',
    #         'host': '172.17.115.38',
    #         'username': 'admin',
    #         'password': 'Inventec@123',
    #     },
    #     {
    #         'name': 'SW-RDG01-SA-SWITCH-02',
    #         'device_type': 'hp_comware',
    #         'host': '172.17.115.39',
    #         'username': 'admin',
    #         'password': 'Inventec@123',
    #     },
    #     {
    #         'name': 'SSW-RDG02-SA-SWITCH-01',
    #         'device_type': 'hp_comware',
    #         'host': '172.17.115.40',
    #         'username': 'admin',
    #         'password': 'Inventec@123',
    #     },
]

# ==========================================
# 2. 文件路径配置
# ==========================================
INPUT_EXCEL = '2fj3.xlsx'  # 输入文件名
OUTPUT_EXCEL = '2fj3_results.xlsx'  # 输出文件名


# ==========================================
# 3. 核心功能实现 (针对 H3C 命令优化)
# ==========================================

def get_mac_from_arp(ssh_conn, ip):
    """在华三核心交换机上通过 Ping 激活 ARP 并获取 MAC"""
    print(f"  正在探测 IP: {ip} ...")
    # 华三命令：ping -c 2 <ip>
    ssh_conn.send_command(f"ping -c 2 {ip}")
    time.sleep(0.8)
    # 华三命令：display arp | include <ip>
    output = ssh_conn.send_command(f"display arp | include {ip}")

    # 华三 MAC 格式：xxxx-xxxx-xxxx
    mac_pattern = r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})'
    match = re.search(mac_pattern, output.lower())
    return match.group(1) if match else None


def get_all_mac_arp_table(ssh_conn):
    """获取核心交换机上所有的 ARP 表项，返回 {mac: ip} 字典"""
    print("  正在获取核心交换机完整 ARP 表...")
    output = ssh_conn.send_command("display arp")
    mac_to_ip = {}
    # H3C ARP 输出典型行: 172.17.115.1  000f-e201-0001  10  GE1/0/1  1
    lines = output.splitlines()
    for line in lines:
        line = line.lower()
        # 匹配 IP 和 MAC
        ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', line)
        mac_match = re.search(r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})', line)
        if ip_match and mac_match:
            mac_to_ip[mac_match.group(1)] = ip_match.group(1)
    return mac_to_ip


def get_switch_total_ports(ssh_conn):
    """自动探测华三交换机的接入端口总数 (8/24/48)"""
    try:
        # 使用 display device 识别型号或 display interface brief 统计端口
        output = ssh_conn.send_command("display interface brief")
        # 统计物理端口 (GE/FE/Eth)，排除 BAGG/NULL/VLAN 等
        # 典型接入端口命名格式: GE1/0/1, Eth1/0/1
        ports = re.findall(r'(GE|XGE|Eth|Ten-G)\d+/\d+/(\d+)', output)
        if ports:
            # 获取最大端口序号
            port_indices = [int(p[1]) for p in ports]
            max_idx = max(port_indices)

            # 根据华三常规规格归位
            if max_idx <= 10:
                return 8
            if max_idx <= 28:
                return 24
            if max_idx <= 52:
                return 48
            return max_idx
    except Exception:
        pass
    return 48  # 默认返回 48


def find_port_on_access_sw(access_sw_config, mac):
    """在华三接入交换机上查找 MAC 对应的物理端口 (支持 8/24/48 口自动识别)"""
    try:
        # 获取配置中指定的端口数，如果没有则设为 None 触发自动探测
        specified_ports = access_sw_config.get('total_ports')

        # 为了 ConnectHandler，排除自定义键
        conn_params = {k: v for k, v in access_sw_config.items() if k not in ['name', 'total_ports']}

        with ConnectHandler(**conn_params) as ssh:
            # 如果没指定端口数，进行自动探测
            total_ports = specified_ports if specified_ports else get_switch_total_ports(ssh)

            # 华三命令：display mac-address <mac>
            output = ssh.send_command(f"display mac-address {mac}")
            lines = output.splitlines()

            for line in lines:
                line_lower = line.lower()
                if mac in line_lower:
                    parts = line.split()
                    # 华三典型输出: MAC | VLAN | State | Port
                    if len(parts) >= 4:
                        vlan = parts[1]
                        port = "Unknown"
                        for p in parts:
                            p_up = p.upper()
                            if any(x in p_up for x in ['GE', 'XGE', 'ETH', 'TEN-G']):
                                # 彻底排除聚合口
                                if 'BAGG' in p_up or 'RAGG' in p_up:
                                    continue
                                port = p
                                break

                        if port != "Unknown":
                            # --- 增强过滤逻辑 ---
                            # 1. 检查端口序号是否超过该交换机的接入规格
                            port_index_match = re.search(r'/(\d+)$', port)
                            if port_index_match:
                                port_idx = int(port_index_match.group(1))
                                if port_idx > total_ports:
                                    # print(f"  [跳过] 端口 {port} 为上行口 (超过接入规格 {total_ports})")
                                    return None, None

                            # 2. 检查该端口下的 MAC 密度
                            mac_count_out = ssh.send_command(f"display mac-address interface {port}")
                            mac_count = len(re.findall(r'[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}', mac_count_out.lower()))

                            if mac_count > 3:
                                # 接入端口通常只会有 1-2 个 MAC (PC/电话)
                                return None, None

                            return port, vlan
    except Exception as e:
        print(f"  连接交换机 {access_sw_config.get('host')} 失败: {e}")
    return None, None


def get_all_up_ports_with_mac(access_sw_config):
    """获取接入交换机上所有非上行链路的 Up 端口及其 MAC 地址"""
    results = []
    try:
        specified_ports = access_sw_config.get('total_ports')
        conn_params = {k: v for k, v in access_sw_config.items() if k not in ['name', 'total_ports']}

        with ConnectHandler(**conn_params) as ssh:
            total_ports = specified_ports if specified_ports else get_switch_total_ports(ssh)

            output_brief = ssh.send_command("display interface brief")
            up_ports = []
            for line in output_brief.splitlines():
                if ' UP ' in line:
                    parts = line.split()
                    port_name = parts[0]
                    p_up = port_name.upper()

                    # 基本过滤：排除逻辑口
                    if any(x in p_up for x in ['GE', 'XGE', 'ETH', 'TEN-G']) and not any(x in p_up for x in ['BAGG', 'RAGG', 'NULL', 'VLAN']):
                        # 规格过滤：排除上行序号
                        port_index_match = re.search(r'/(\d+)$', port_name)
                        if port_index_match:
                            if int(port_index_match.group(1)) <= total_ports:
                                up_ports.append(port_name)

            # 2. 获取 MAC 地址表
            output_mac = ssh.send_command("display mac-address")
            port_to_macs = {}
            for line in output_mac.splitlines():
                line_lower = line.lower()
                mac_match = re.search(r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})', line_lower)
                if mac_match:
                    parts = line.split()
                    current_port = None
                    for p in parts:
                        if any(x in p.upper() for x in ['GE', 'XGE', 'ETH', 'TEN-G']):
                            current_port = p
                            break
                    if current_port:
                        if current_port not in port_to_macs:
                            port_to_macs[current_port] = []
                        port_to_macs[current_port].append(mac_match.group(1))

            for port in up_ports:
                macs = port_to_macs.get(port, [])
                # 再次校验 MAC 密度
                if len(macs) > 5:
                    continue

                results.append({
                    'port': port,
                    'macs': macs
                })
    except Exception as e:
        print(f"  [额外记录] 连接交换机 {access_sw_config.get('host')} 失败: {e}")
    return results


def main():
    print("程序已启动...")
    # 读取 Excel
    try:
        df = pd.read_excel(INPUT_EXCEL)
        # 确保表格中有对应的 IP 列
        possible_ip_cols = ['*IP/域名', 'IP地址', 'IP/域名', 'IP']
        ip_column = next((col for col in possible_ip_cols if col in df.columns), None)

        # 设备名称列名处理
        name_column = '*设备名称' if '*设备名称' in df.columns else '设备名称'

        if not ip_column:
            print(f"Excel 中未找到 IP 列，当前列名: {df.columns.tolist()}")
            return
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return

    # 准备结果列表
    results_data = []
    # 用于记录哪些端口已被 Excel 中的设备占用
    occupied_ports = {}  # { sw_name: {port_name} }

    print(f"--- 开始自动化网络探测 (共 {len(df)} 条记录) ---")
    try:
        with ConnectHandler(**CORE_SWITCH) as core_ssh:
            # 预先获取全量 ARP，用于后续匹配额外端口的 IP
            global_mac_to_ip = get_all_mac_arp_table(core_ssh)

            for index, row in df.iterrows():
                device_name = row.get(name_column, 'Unknown')
                ip = str(row.get(ip_column)).strip()

                if not ip or ip == 'nan':
                    continue

                # 步骤 1: 获取 MAC
                mac = get_mac_from_arp(core_ssh, ip)

                found_info = {
                    '设备名称': device_name,
                    'IP': ip,
                    'MAC地址': mac if mac else '未发现',
                    '交换机': '未找到',
                    '端口': '未找到',
                    'VLAN': '未找到',
                    '备注': 'Excel列表设备'
                }

                # 步骤 2: 如果有 MAC，去接入交换机找端口
                if mac:
                    for sw_config in ACCESS_SWITCHES:
                        # 为了 ConnectHandler，排除掉我们自定义的 'name' 键
                        conn_params = {k: v for k, v in sw_config.items() if k != 'name'}
                        port, vlan = find_port_on_access_sw(conn_params, mac)

                        if port:
                            # 过滤掉上行端口
                            if 'BAGG' in port.upper() or 'RAGG' in port.upper():
                                continue

                            found_info['交换机'] = sw_config['name']
                            found_info['端口'] = port
                            found_info['VLAN'] = vlan

                            # 记录占用
                            if sw_config['name'] not in occupied_ports:
                                occupied_ports[sw_config['name']] = set()
                            occupied_ports[sw_config['name']].add(port)

                            print(f"  [成功] 设备 {ip} 位于 {sw_config['name']} 的 {port} 端口")
                            break

                results_data.append(found_info)

            # --- 新增功能：扫描接入交换机其他 Up 端口 ---
            print("\n--- 正在扫描接入交换机额外 Up 端口 ---")
            for sw_config in ACCESS_SWITCHES:
                sw_name = sw_config['name']
                print(f"  扫描交换机: {sw_name} ...")
                conn_params = {k: v for k, v in sw_config.items() if k != 'name'}
                up_ports_info = get_all_up_ports_with_mac(conn_params)

                sw_occupied = occupied_ports.get(sw_name, set())

                for item in up_ports_info:
                    port = item['port']
                    # 如果该端口没在之前的 Excel 匹配中出现过
                    if port not in sw_occupied:
                        macs = item['macs']
                        if not macs:
                            results_data.append({
                                '设备名称': '未知设备',
                                'IP': '未知',
                                'MAC地址': '未知',
                                '交换机': sw_name,
                                '端口': port,
                                'VLAN': '未知',
                                '备注': '额外Up端口 (无MAC)'
                            })
                        else:
                            for m in macs:
                                extra_ip = global_mac_to_ip.get(m, '未知')
                                results_data.append({
                                    '设备名称': '发现设备' if extra_ip == '未知' else f"发现设备({extra_ip})",
                                    'IP': extra_ip,
                                    'MAC地址': m,
                                    '交换机': sw_name,
                                    '端口': port,
                                    'VLAN': '未知',
                                    '备注': '额外Up端口 (有数据)'
                                })
                        print(f"  [额外] 发现未记录端口: {sw_name} -> {port}")

    except Exception as e:
        print(f"连接核心交换机失败: {e}")

    # 保存结果到新 Excel
    new_df = pd.DataFrame(results_data)

    # --- 新增功能：结果排序 ---
    def sort_key(row):
        sw = str(row['交换机'])
        port = str(row['端口'])
        # 提取端口中的数字进行逻辑排序，例如 GE1/0/1 -> (1, 0, 1)
        numbers = [int(n) for n in re.findall(r'\d+', port)]
        return (sw, numbers)

    if not new_df.empty:
        # 使用 apply 生成排序辅助列，排序后再删除
        new_df['sort_helper'] = new_df.apply(sort_key, axis=1)
        new_df = new_df.sort_values('sort_helper').drop(columns=['sort_helper'])

    new_df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\n--- 探测完成！结果已排序并保存至: {OUTPUT_EXCEL} ---")


if __name__ == "__main__":
    main()
