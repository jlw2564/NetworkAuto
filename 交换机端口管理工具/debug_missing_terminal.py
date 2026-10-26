import sys
import os
import re
import time
import json
from collections import defaultdict
import pandas as pd
from netmiko import ConnectHandler

try:
    from net_port_tool import config as np_config
except Exception:
    sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
    from net_port_tool import config as np_config


def main():
    print("=" * 70)
    print("终端漏检专项诊断")
    print("=" * 70)
    target_ip = '192.168.1.13'
    target_vlan = '1032'
    access_sw_pattern = '122'
    print(f"目标IP    : {target_ip}")
    print(f"目标VLAN  : {target_vlan}")
    print(f"接入交换机名字含: {access_sw_pattern}")
    print("")

    try:
        devices = np_config.load_switches()
    except Exception as exc:
        print(f"[FATAL] 读取交换机清单失败: {exc}")
        return 1
    core_dev = next((d for d in devices if d.is_core), None)
    if core_dev is None:
        print("[FATAL] 没有核心交换机配置")
        return 2
    core = {'name': core_dev.name, **core_dev.conn_params()}

    access_sw = None
    for dev in devices:
        if access_sw_pattern in dev.name:
            access_sw = {'name': dev.name, **dev.conn_params()}
            break
    if not access_sw:
        print(f"[WARN] 没找到名字含 {access_sw_pattern} 的接入交换机，跳过接入侧诊断")

    lines = []

    def P(s=''):
        print(s)
        lines.append(str(s))

    try:
        core_params = {k: v for k, v in core.items() if k not in ['name', 'is_core']}
        P(f"\n[1/6] 连接核心 {core['name']} ({core['host']}) 抓 ARP / MAC / LLDP")
        with ConnectHandler(**core_params) as ssh:
            def run(cmd):
                try:
                    ssh.send_command("screen-length disable", delay_factor=1, expect_string=r'>')
                    time.sleep(0.25)
                    out = ssh.send_command(cmd, delay_factor=4, expect_string=r'>', read_timeout=20)
                    return out
                except Exception as e:
                    return f"[CMD_ERR] {type(e).__name__}: {e}"

            P("\n--- (1.1) display arp | include %s ---" % target_ip)
            out11 = run(f"display arp | include {target_ip}")
            P(out11[:4000])

            P("\n--- (1.2) display arp all | include %s ---" % target_ip)
            out12 = run(f"display arp all | include {target_ip}")
            P(out12[:4000])

            P("\n--- (1.3) display arp vlan %s (VLAN 维度看能不能拿到) ---" % target_vlan)
            out13 = run(f"display arp vlan {target_vlan}")
            P(out13[:6000])

            # 找MAC
            mac_candidates = set()
            for text in [out11, out12, out13]:
                for m in re.finditer(r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})', text.lower()):
                    mac_candidates.add(m.group(1))
            P(f"\n--- 从以上ARP里提取的候选MAC: {sorted(mac_candidates)} ---")

            if not mac_candidates:
                P("[HINT] 核心ARP完全拿不到目标IP的MAC。可能：网关不在本核心/静态IP不通信/MAC绑定在其它设备/该网段没做三层转发。下面会尝试从接入侧MAC表/核心所有VLAN MAC表直接反查。")

            P("\n--- (1.4) display lldp neighbor-information list（核心LLDP邻居）---")
            lldp_out = run("display lldp neighbor-information list")
            P(lldp_out[:8000])

            for mac in list(mac_candidates) + ['DIRECT']:
                P("\n===== 候选MAC %s =====" % mac)
                if mac == 'DIRECT':
                    P("（只看 VLAN1032 的核心MAC表，反推哪个端口出）")
                    out_x = run(f"display mac-address vlan {target_vlan}")
                    P(out_x[:6000])
                    continue
                P(f"--- (1.5) display mac-address {mac} ---")
                P(run(f"display mac-address {mac}")[:4000])
                P(f"--- (1.6) display mac-address {mac} vlan {target_vlan} ---")
                P(run(f"display mac-address {mac} vlan {target_vlan}")[:4000])
                P(f"--- (1.7) display mac-address {mac} all ---")
                P(run(f"display mac-address {mac} all")[:4000])

        if access_sw:
            P(f"\n[2/6] 连接接入交换机 {access_sw['name']} ({access_sw['host']})")
            acc_params = {k: v for k, v in access_sw.items() if k not in ['name', 'is_core', 'total_ports']}
            with ConnectHandler(**acc_params) as ssh:
                def run(cmd):
                    try:
                        ssh.send_command("screen-length disable", delay_factor=1, expect_string=r'>')
                        time.sleep(0.25)
                        out = ssh.send_command(cmd, delay_factor=4, expect_string=r'>', read_timeout=20)
                        return out
                    except Exception as e:
                        return f"[CMD_ERR] {type(e).__name__}: {e}"

                P("\n--- (2.1) display interface brief ---")
                bif = run("display interface brief")
                P(bif[:10000])

                P("\n--- (2.2) display lldp neighbor-information list ---")
                P(run("display lldp neighbor-information list")[:4000])

                P("\n--- (2.3) display link-aggregation verbose ---")
                P(run("display link-aggregation verbose")[:6000])

                P(f"\n--- (2.4) 接入侧全量 display mac-address vlan {target_vlan} ---")
                mac_vlan = run(f"display mac-address vlan {target_vlan}")
                P(mac_vlan[:8000])

                # 提取该VLAN下所有MAC，逐个查对应端口 + IP
                local_macs = sorted({m.group(1) for m in re.finditer(
                    r'([0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4})', mac_vlan.lower())})
                P(f"[INFO] 在接入交换机 VLAN{target_vlan} 学到的MAC数: {len(local_macs)}")
                P(f"[INFO] 前10个: {local_macs[:10]}")
                for mac in list(mac_candidates) + local_macs[:20]:
                    P(f"\n===== MAC {mac} (在接入侧) =====")
                    P(f"--- (2.5) display mac-address {mac} ---")
                    P(run(f"display mac-address {mac}")[:3000])
                    P(f"--- (2.6) display mac-address {mac} all ---")
                    P(run(f"display mac-address {mac} all")[:3000])
                    P(f"--- (2.7) display arp | include {mac} (接入侧本地arp, 一般没，有就是svi) ---")
                    P(run(f"display arp | include {mac}")[:2000])

        P("\n[3/6] 现有 终端IP表 / 交换机端口资源管理表(如果已有) 对照看看目标IP是否已有记录")
        term_file = os.path.join(os.path.dirname(__file__), '终端IP表.xlsx')
        if os.path.isfile(term_file):
            try:
                td = pd.read_excel(term_file, dtype=str)
                P(f"[OK] {term_file} 存在，共 {len(td)} 条终端IP记录")
                hit = td[td.apply(lambda r: target_ip in ' '.join(str(v) for v in r.values), axis=1)]
                if len(hit) == 0:
                    P(f"[WARN] 目标IP {target_ip} 没出现在终端IP表中！ => 它只会走“额外Up端口扫描”，之前版本会被标‘未知设备’进而在终端接入表被过滤。这就是最常见漏检根因！")
                else:
                    P(f"[OK] 目标IP {target_ip} 在终端IP表里存在: {hit.head(5).to_string(index=False)}")
            except Exception as e:
                P(f"[ERR] 读 {term_file} 失败: {e}")
        else:
            P(f"[WARN] 没找到 {term_file}")

        out_xlsx = os.path.join(os.path.dirname(__file__), '交换机端口资源管理表.xlsx')
        if os.path.isfile(out_xlsx):
            try:
                xl = pd.ExcelFile(out_xlsx, engine='openpyxl')
                P(f"\n[OK] 上一次运行输出 {out_xlsx} 存在，sheets = {xl.sheet_names}")
                if '终端接入表' in xl.sheet_names:
                    df = pd.read_excel(out_xlsx, sheet_name='终端接入表', dtype=str)
                    hit = df[df.apply(lambda r: target_ip in ' '.join(str(v) for v in r.values)
                                      or (target_vlan and target_vlan in str(r.get('VLAN', ''))), axis=1)]
                    P(f"终端接入表总行数 = {len(df)}")
                    P(f"命中IP={target_ip} 或 VLAN={target_vlan} 的行数 = {len(hit)}")
                    if len(hit) > 0:
                        P(hit.head(20).to_string(index=False))
                    else:
                        P("[WARN] 之前运行的终端接入表里完全没有目标IP/VLAN，符合你反馈的‘没有显示’现象。")
            except Exception as e:
                P(f"[ERR] 读 {out_xlsx} 失败: {e}")
        else:
            P("[HINT] 还没有上一次的 交换机端口资源管理表.xlsx，跳过历史对照")

    except Exception as e:
        P(f"[FATAL] 诊断过程异常: {type(e).__name__}: {e}")
        import traceback
        P(traceback.format_exc())

    log_path = os.path.join(os.path.dirname(__file__), f'debug_{target_ip.replace(".", "_")}_vlan{target_vlan}.log')
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    P(f"\n[DONE] 诊断日志已保存: {log_path}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
