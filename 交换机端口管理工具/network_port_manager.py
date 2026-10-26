"""交换机端口资源管理 —— 命令行入口（精简版）。

原 1800 行单文件实现已拆分为可复用包 `net_port_tool`，本文件仅保留 CLI。
Web 界面入口：`python -m net_port_tool.webapp`

用法：
    python network_port_manager.py                      # 使用默认清单与默认输出
    python network_port_manager.py -c 8                 # 并发 8 台
    python network_port_manager.py --no-discover        # 关闭未知交换机发现
    python network_port_manager.py --sections name links usage
    python network_port_manager.py --switch-file a.xlsx --terminal-file b.xlsx --out c.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from net_port_tool.report import ALL_SECTIONS
from net_port_tool.service import RunOptions, run


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description='交换机端口资源管理（终端定位 / 端口使用 / 可疑端口）')
    p.add_argument('--switch-file', default=None, help='交换机清单 Excel 路径，默认 switch_table.xlsx')
    p.add_argument('--terminal-file', default=None, help='终端 IP 表 Excel 路径，默认 终端IP表.xlsx')
    p.add_argument('--out', default=None, help='输出 Excel 路径，默认 交换机端口资源管理表.xlsx')
    p.add_argument('-c', '--concurrency', type=int, default=6, help='采集并发度（默认 6）')
    p.add_argument('--no-discover', action='store_true', help='关闭“发现未知交换机”')
    p.add_argument('--sections', nargs='*', choices=list(ALL_SECTIONS), default=None,
                   help=f'要输出的模块，可选：{", ".join(ALL_SECTIONS)}（默认全部）')
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    options = RunOptions(
        switch_file=Path(args.switch_file) if args.switch_file else None,
        terminal_file=Path(args.terminal_file) if args.terminal_file else None,
        output_file=Path(args.out) if args.out else None,
        concurrency=max(1, min(args.concurrency, 32)),
        discover=not args.no_discover,
        sections=tuple(args.sections) if args.sections else tuple(ALL_SECTIONS),
    )
    result = run(options)
    if not result.ok:
        print(f'\n运行失败：{result.error}')
        return 1
    print(f'\n结果文件：{result.output_file}')
    for key, value in result.stats.items():
        print(f'  - {key}: {value}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
