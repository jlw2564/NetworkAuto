# NetworkAuto — 网络运维自动化工具集

一套针对**网络运维场景**的自动化工具集合，覆盖交换机端口管理、健康采集、人员权限提取、摄像头批量配置、IP 地址资源管理等常见工作。

> 项目按「工具」分类组织，每个工具一个文件夹，各自带有使用说明。

## 工具一览

| 工具 | 用途 | 入口 |
| --- | --- | --- |
| `交换机端口管理工具/` | 交换机端口资源管理：终端定位 / 端口使用 / 可疑端口，支持命令行 + Web 界面 | `network_port_manager.py`、`python -m net_port_tool.webapp` |
| `交换机健康采集工具/` | H3C 交换机配置 / 健康状态 / 拓扑图采集 | `h3c_switch_health_topology.py` |
| `人员权限提取工具/` | 按人员编号匹配并汇总权限表 | `extract_permissions_by_emp.py` |
| `摄像头批量配置工具/` | 宇视摄像头 Excel 批量配置（浏览器自动化） | `batch_configure_unv_cameras.py` |
| `IP地址资源管理工具/` | 由端口数据生成 IP 地址资源管理表 | `generate_ip_table.py` |

## 目录结构

```
NetworkAuto/
├── 交换机端口管理工具/       # 端口资源管理（含 net_port_tool 包、Web 界面）
├── 交换机健康采集工具/       # H3C 交换机健康状态 / 拓扑采集
├── 人员权限提取工具/         # 权限表汇总
├── 摄像头批量配置工具/       # 摄像头 Excel 批量配置
├── IP地址资源管理工具/       # IP 地址资源管理表生成
├── requirements.txt          # 依赖清单
└── 使用总览.md               # 目录与运行总览
```

## 环境依赖

```
netmiko
pandas
openpyxl
xlrd
fastapi
uvicorn[standard]
python-multipart
```

安装：

```powershell
pip install -r requirements.txt
```

## 运行方式

项目统一使用虚拟环境解释器 `venv\Scripts\python.exe`。运行前建议先清掉可能冲突的环境变量：

```powershell
Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue
```

### 示例：启动「交换机端口管理工具」Web 界面

```powershell
cd 交换机端口管理工具
& "..\venv\Scripts\python.exe" -m net_port_tool.webapp
# 浏览器打开 http://127.0.0.1:8000
```

### 示例：命令行运行「交换机端口管理工具」

```powershell
cd 交换机端口管理工具
& "..\venv\Scripts\python.exe" network_port_manager.py -c 6
```

> 各工具的详细用法见各自目录下的 `使用说明.md`。

## 工具间数据流

```
交换机端口管理工具  ──生成──>  端口数据.xlsx  ──>  IP地址资源管理工具  ──>  IP地址资源管理表.xlsx
        │
        └──输出──>  交换机端口资源管理表.xlsx（终端接入表 / 端口使用表 / 可疑端口清单 …）

交换机健康采集工具  ──>  采集结果\（台账 / 健康总表 / 拓扑图.html / 配置文件）
```

## 版本说明

- `main`：稳定发布分支
- `dev`：日常开发分支
- 版本 tag：`v1.0`

## 注意事项

- 本仓库**不包含**交换机配置文件、Excel 数据表、采集结果等敏感数据（详见 `.gitignore`），这些数据请保留在本地。
- 涉及真实网络设备操作时，请先在测试环境验证。
