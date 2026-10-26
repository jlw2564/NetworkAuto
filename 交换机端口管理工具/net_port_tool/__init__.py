"""交换机端口资源管理工具包。

模块划分：
    config      配置与设备清单加载
    parsers     Comware 输出纯解析函数（无 IO，可单测）
    ssh_client  SSH 会话封装（一台设备一次登录，批量下发命令）
    collector   并发采集，产出 SwitchSnapshot
    analyzer    基于快照做连线/终端定位/端口使用/可疑端口分析
    report      生成 Excel 报告（Sheet 与列名保持原有约定）
    service     任务编排（带进度回调），供 CLI 与 Web 共用
"""

__all__ = ['config', 'parsers', 'ssh_client', 'collector', 'analyzer', 'report', 'service']
