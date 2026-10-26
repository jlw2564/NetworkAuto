import pandas as pd
import re
import time
from netmiko import ConnectHandler

# ==========================================
# 核心交换机配置
# ==========================================
CORE_SWITCH = {
    'device_type': 'hp_comware',
    'host': '172.17.115.254',
    'username': 'admin',
    'password': 'Inventec@123',
}

OUTPUT_FILE = 'core_port_planning.xlsx'


def get_lldp_neighbors(ssh_conn):
    """获取 LLDP 邻居信息，返回 {本地端口: 邻居名称}"""
    print("  正在抓取 LLDP 邻居信息...")
    # 增加延时确保完整读取
    output = ssh_conn.send_command("display lldp neighbor-information list", delay_factor=2)
    neighbors = {}

    # H3C 实际输出列顺序: Local Interface | Chassis ID | Port ID | System Name
    lines = output.splitlines()
    start_parsing = False
    for line in lines:
        if 'Local Interface' in line and 'System Name' in line:
            start_parsing = True
            continue

        if start_parsing and line.strip() and not line.startswith('---'):
            parts = line.split()
            if len(parts) >= 3:
                # 第一列是本地接口
                local_port = parts[0]
                # 最后一列（或从第四个元素开始的所有内容）是系统名称
                # 如果只有 3 列说明系统名为空或为 '-'
                sys_name = ""
                if len(parts) >= 4:
                    sys_name = " ".join(parts[3:])

                # 清理系统名称，如果是 '-' 则设为空
                if sys_name == '-':
                    sys_name = ""

                # 华三端口可能缩写，统一转为大写处理
                port_key = local_port.upper()
                if port_key in neighbors:
                    # 如果该端口已有邻居，且当前邻居名不为空，则追加
                    if sys_name:
                        if neighbors[port_key]:
                            neighbors[port_key] += f", {sys_name}"
                        else:
                            neighbors[port_key] = sys_name
                else:
                    neighbors[port_key] = sys_name
    return neighbors


def get_mac_to_ip(ssh_conn):
    """获取 ARP 表，建立 MAC 到 IP 的映射"""
    print("  正在抓取 ARP 表进行 IP 关联...")
    output = ssh_conn.send_command("display arp")
    mac_to_ip = {}
    lines = output.splitlines()
    for line in lines:
        # 匹配 IP 和 MAC
        ip_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', line)
        mac_match = re.search(r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})', line.lower())
        if ip_match and mac_match:
            mac_to_ip[mac_match.group(1)] = ip_match.group(1)
    return mac_to_ip


def get_port_macs(ssh_conn):
    """获取每个端口下学习到的 MAC 地址"""
    print("  正在分析端口 MAC 地址分布...")
    output = ssh_conn.send_command("display mac-address")
    port_macs = {}
    lines = output.splitlines()
    for line in lines:
        line = line.lower()
        mac_match = re.search(r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})', line)
        if mac_match:
            mac = mac_match.group(1)
            # 寻找端口标识
            port_match = re.search(r'(ge\d+/\d+/\d+|xge\d+/\d+/\d+|eth\d+/\d+/\d+|ten-g\d+/\d+/\d+)', line)
            if port_match:
                port = port_match.group(1).upper()
                if port not in port_macs:
                    port_macs[port] = []
                port_macs[port].append(mac)
    return port_macs


def main():
    print(f"--- 开始核心交换机 [{CORE_SWITCH['host']}] 端口连接情况探测 ---")
    try:
        with ConnectHandler(**CORE_SWITCH) as ssh:
            # 1. 获取 LLDP 信息
            lldp_data = get_lldp_neighbors(ssh)

            # 2. 获取 MAC 地址表
            port_mac_data = get_port_macs(ssh)

            # 3. 获取 ARP 表用于 MAC -> IP 转换
            mac_ip_map = get_mac_to_ip(ssh)

            # 4. 获取所有 Up 的物理端口列表
            print("  正在获取端口列表...")
            brief_out = ssh.send_command("display interface brief | include UP")
            all_up_ports = []
            for line in brief_out.splitlines():
                parts = line.split()
                if parts:
                    p_name = parts[0]
                    if any(x in p_name.upper() for x in ['GE', 'XGE', 'ETH', 'TEN-G']):
                        all_up_ports.append(p_name)

            # 5. 整合数据
            results = []
            for port in all_up_ports:
                # 统一转为大写进行匹配
                port_upper = port.upper()
                # 获取 LLDP 邻居名称
                device_name = lldp_data.get(port_upper, "")

                # 获取该端口下的 IP
                macs = port_mac_data.get(port_upper, [])
                ips = []
                for m in macs:
                    ip = mac_ip_map.get(m)
                    if ip:
                        ips.append(ip)

                # 如果是级联口，可能有多个 IP，用逗号连接
                # 增加排序让输出更美观
                unique_ips = sorted(list(set(ips)))
                ip_str = ", ".join(unique_ips) if unique_ips else ""

                results.append({
                    '核心交换机端口': port,
                    '连接设备名': device_name,
                    '对应IP地址': ip_str
                })

            # 6. 排序并保存
            df = pd.DataFrame(results)
            if not df.empty:
                # 端口自然排序
                def sort_key(p):
                    nums = re.findall(r'\d+', p)
                    return [int(n) for n in nums]
                df['sort'] = df['核心交换机端口'].apply(sort_key)
                df = df.sort_values('sort').drop(columns=['sort'])

                df.to_excel(OUTPUT_FILE, index=False)
                print(f"\n探测完成！核心端口规划表已保存至: {OUTPUT_FILE}")
            else:
                print("未发现任何 UP 状态的物理端口。")

    except Exception as e:
        print(f"执行失败: {e}")


if __name__ == "__main__":
    main()
