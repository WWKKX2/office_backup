# Office Auto Backup

[![License: MIT](https://img.shields.io/badge/license-MIT-2563EB.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-2563EB.svg)](https://www.python.org/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-2563EB.svg)](#环境要求)
[![Zero dependencies](https://img.shields.io/badge/dependencies-0-0F9D58.svg)](requirements.txt)

自动备份正在编辑的 Office / WPS 文档。定时扫描当前打开的 Word、Excel、PowerPoint
和 WPS 文件，把带时间戳的副本存到指定目录，并自动清理旧版本。

适合的场景：写长文档时忘了保存、误覆盖、断电崩溃、改错了想回滚。

> 纯标准库实现，装好 Python 即可运行；同时提供图形界面与命令行两种用法。

## 目录

- [特点](#特点)
- [环境要求](#环境要求)
- [快速开始](#快速开始)
- [图形界面](#图形界面)
- [开机自启](#开机自启)
- [命令行参数](#命令行参数)
- [配置说明](#配置说明)
- [备份目录结构](#备份目录结构)
- [工作原理](#工作原理)
- [自检](#自检)
- [常见问题](#常见问题)
- [已知限制](#已知限制)
- [贡献与许可](#贡献与许可)

## 特点

- **零第三方依赖** —— 主引擎是纯 `ctypes` 实现，装好 Python 就能跑，不用装 pywin32
- **同时支持 MS Office 与 WPS** —— 直接读系统句柄表，不依赖 COM 接口是否可用
- **同名文件不打架** —— 按源目录分桶存放，两个「总结.docx」互不影响
- **不会刷爆版本** —— 内容没变不备份，且有冷却期，避免自动保存产生大量重复副本
- **单实例保护** —— 重复启动会自动退出，不会出现两个进程互相覆盖

## 环境要求

- Windows 10 / 11
- Python 3.8+
- **建议以管理员身份运行**（否则看不到其他用户进程里的文档）

## 快速开始

```bash
git clone https://github.com/WWKKX2/office_backup.git
cd office-auto-backup

# 从模板生成本地配置（config.json 不会被提交）
copy config.example.json config.json
```

`config.example.json` 里的 `backup_root` 是相对路径 `backups`，会落在项目目录下；
想存到别处（例如 `D:\Office_Backups`），改 `config.json` 里的 `backup_root` 即可，
目录不存在时会自动创建。

> **注意**：有些 Windows 上的 `python` 命令是应用商店的占位程序（运行后打开商店，
> 什么都不执行）。如果 `python --version` 没反应，请改用 `py -3`，或直接双击
> 下面的 `run.bat` —— 它会自动找到可用的 Python。

```bash
# 看看现在有哪些文档被打开
py -3 main.py --scan

# 立刻备份一轮
py -3 main.py --once

# 常驻运行，每 30 秒检查一次
py -3 main.py
```

也可以直接双击：

| 文件 | 作用 |
|---|---|
| `run_gui.bat` | 打开图形界面（无黑色命令行窗口），推荐日常使用 |
| `run.bat` | 带命令行窗口运行，能看到日志，适合排查问题 |
| `start_silent.vbs` | 后台静默运行，不弹窗口；会自动寻找 `pythonw.exe`（不要求它在 PATH 里） |

后台静默运行的其他方式：

```bash
# 命令行
start /b pythonw main.py
```

## 图形界面

不想记命令的话，直接双击 **`run_gui.bat`** 打开图形界面（不会弹黑色命令行窗口）。

```bash
# 也可以手动启动；带控制台方便看报错
py -3 gui.py
py -3 gui.py --config 其它配置.json
```

界面采用浅色卡片式布局：顶部是标题与常用操作，左侧为导航，右侧是内容区。

| 页面 | 功能 |
|---|---|
| **监控** | 开始 / 停止监听、立即扫描、立即备份；状态胶囊显示监听状态；统计卡片展示打开的文档数、备份版本数、本轮新增与失败次数；表格实时列出文件名、大小、修改时间、已有备份版本数和最后备份时间；双击某行打开所在文件夹，右键可单独备份该文件或打开它的备份目录；底部是深色运行日志，按级别着色 |
| **设置** | 图形化修改备份目录、轮询间隔、冷却期、保留版本数、扩展名、扫描方式、应用开关、日志级别；一键安装 / 卸载开机自启 |

说明：

- 界面启动后会自动开始监听，等价于运行 `main.py`
- 所有扫描 / 备份都在后台线程执行，不会卡住界面
- 支持高 DPI 缩放，在 125% / 150% 缩放的屏幕上不会显得拥挤
- 和命令行版共用同一把单实例锁：如果后台已经有一个备份进程在跑，
  界面会保持空闲并在日志里提示，不会重复备份，也不会弹错误框
- 在「设置」里保存后立即生效，写入的仍然是同目录的 `config.json`
- 「设置 → 安装开机自启」装的是后台备份进程（`main.py`），登录后静默运行，
  不会自动打开这个界面；两者共用单实例锁，同时只能有一个在备份
- 界面只用 tkinter / ttk，不需要安装任何 UI 框架

## 开机自启

用 `install_autostart.py` 一键装成登录后自动运行：

```bash
py -3 install_autostart.py --install     # 安装
py -3 install_autostart.py --status      # 查看状态
py -3 install_autostart.py --uninstall   # 卸载
```

两种模式：

| 模式 | 命令 | 特点 |
|---|---|---|
| `startup`（默认） | `--install` | 在「启动」文件夹放一个 `.vbs`，**不需要管理员权限**，随时可撤销 |
| `task` | `--install --mode task` | 注册计划任务，**需要管理员权限**，以最高权限启动、不弹 UAC，登录后延迟 30 秒运行 |

一般用默认的 `startup` 就够了。只有当你需要扫描**其他用户**进程里的文档，
或者遇到权限导致扫描不全时，才需要 `--mode task`。

加 `--dry-run` 可以先看看会做什么，不实际改动系统：

```bash
py -3 install_autostart.py --install --dry-run
```

> 装完之后如果改了项目目录，需要重新运行一次 `--install`，因为启动器里存的是绝对路径。


## 命令行参数

| 参数 | 说明 |
|---|---|
| `--scan` | 只列出检测到的打开文档，不备份 |
| `--once` | 只运行一轮后退出 |
| `--dry-run` | 演练模式，只打印不实际复制 |
| `--config PATH` | 指定配置文件（默认同目录 `config.json`） |
| `--interval N` | 覆盖轮询间隔（秒） |
| `--log-level LEVEL` | 覆盖日志级别（DEBUG / INFO / WARNING） |
| `--no-lock` | 跳过单实例锁 |
| `--version` | 打印版本号后退出 |

> `main.py` 启动时会自动把控制台输出切到系统代码页，中文日志在 cp936
> 控制台里不会变成乱码；想强制指定编码可以设置 `PYTHONIOENCODING`。

## 配置说明

首次运行前，从模板复制一份本地配置（`config.json` 已被 `.gitignore` 忽略）：

```bash
copy config.example.json config.json
```

然后按需编辑 `config.json`：

```json
{
  "backup_root": "D:\\Office_Backups",
  "interval_seconds": 30,
  "min_seconds_between_backups": 30,
  "max_versions_per_file": 20,
  "scan_sources": ["handle", "com"],
  "office_apps": { "word": true, "excel": true, "powerpoint": true, "wps": true },
  "file_extensions": [".docx", ".xlsx", ".pptx"],
  "log_level": "INFO",
  "log_max_bytes": 1048576,
  "log_backup_count": 3
}
```

| 字段 | 说明 |
|---|---|
| `backup_root` | 备份根目录，不存在会自动创建 |
| `interval_seconds` | 轮询间隔，默认 30 秒 |
| `min_seconds_between_backups` | 同一文件的备份冷却期，防止版本刷爆 |
| `max_versions_per_file` | 每个文件保留的版本数，超出自动删最旧的 |
| `scan_sources` | 扫描源顺序：`handle`（句柄扫描）/ `com`（COM 接口） |
| `office_apps` | 各应用开关，关掉的应用不会被扫描（对句柄扫描同样生效） |
| `file_extensions` | 需要备份的扩展名；启用 WPS 时自动补上 `.wps/.et/.dps` 等 |
| `log_level` | 日志级别，排错时改成 `DEBUG` |

配置文件支持带 BOM 的 UTF-8（记事本、PowerShell 保存的格式都能直接读），
也支持不带 BOM 的 UTF-8。

## 备份目录结构

```
D:\Office_Backups\
├── 季度报告__1a2b3c4d\          # 文件名 + 父目录哈希，同名文件互不干扰
│   ├── 季度报告_20261031_194200.docx
│   └── 季度报告_20261031_195200.docx
└── 台账__9f8e7d6c\
    └── 台账_20261031_194205.xlsx
```

## 工作原理

1. **发现进程** —— 用 Toolhelp 快照找出 `winword.exe` / `excel.exe` / `powerpnt.exe` /
   `wps.exe` / `et.exe` / `wpp.exe` 等文档进程
2. **扫描句柄** —— 通过 `NtQuerySystemInformation` 拉全系统句柄表，挑出这些进程
   持有的 File 类型句柄，`DuplicateHandle` 到本进程后用 `GetFinalPathNameByHandleW`
   反解真实路径。只有磁盘文件句柄会被反解；遇到管道等可能永久阻塞的句柄会直接跳过，
   单个进程解析超过几秒也会自动放弃并在 5 分钟内不再尝试，保证扫描不会卡死
3. **判断变化** —— 用 `(mtime, size)` 对比，内容没变就跳过
4. **写入备份** —— 复制成 `原文件名_时间戳.扩展名`，失败自动重试
5. **清理旧版** —— 每个文件只保留最近 N 个版本

## 文件说明

| 文件 | 作用 |
|---|---|
| `main.py` | 入口：配置加载、扫描编排、主循环、CLI |
| `gui.py` | 图形界面（tkinter，零第三方依赖） |
| `backup_core.py` | 备份存储：变化检测、分桶、版本清理 |
| `office_handle_scanner.py` | 句柄扫描引擎（零依赖，带阻塞句柄超时保护） |
| `com_source.py` | 可选 COM 扫描源（需 pywin32，没有会自动跳过） |
| `console_util.py` | 控制台编码适配 + `pythonw` 下的安全输出 |
| `selftest.py` | 端到端自检，不需要装 Office 就能验证全链路 |
| `_probe.py` | 性能诊断，定位扫描慢在哪个句柄 |
| `run.bat` | 带窗口启动（自动寻找 Python） |
| `run_gui.bat` | 无窗口启动图形界面（自动寻找 `pythonw.exe`） |
| `install_autostart.py` | 安装 / 卸载 / 查看开机自启 |
| `setup_autostart.py` | 旧命令名，已转发到 `install_autostart.py` |
| `start_silent.vbs` | 无窗口启动脚本（自动寻找 `pythonw.exe`） |
| `config.example.json` | 配置模板，复制为 `config.json` 后按需修改 |

## 自检

改完代码或换机器后，先跑一遍自检确认环境正常：

```bash
py -3 selftest.py
```

它会模拟"某个进程打开着 docx/xlsx"，然后验证句柄扫描、备份判定、版本清理、
同名隔离、演练模式和 `main.py --once` 的完整链路。全部通过会输出
`14/14 项通过`。

## 常见问题

**`python main.py` 没反应，或者弹出了应用商店？**
本机的 `python` 可能是 Windows 自带的占位程序。用 `py -3 main.py`、
双击 `run.bat`，或把 `python.exe` 的完整路径写进 `start_silent.vbs`。

**扫描不到任何文档？**
1. 用 `py -3 main.py --scan` 看是否检测到 Office 进程
2. 没检测到进程 → Office/WPS 没打开，或进程名不在支持列表里
3. 检测到进程但没文件 → 大概率是权限问题，试试以管理员身份运行

**启动时报 `Could not find platform independent libraries`？**
这是 Python 安装本身不完整（`sys.prefix` 指向了错误目录），不是本工具的
问题。功能通常仍能正常使用；建议用 Python 官方安装包修复/重装一次 Python。

**想改用 COM 方式？**
把 `config.json` 的 `scan_sources` 改成 `["com"]`，然后 `pip install pywin32`。
一般不推荐，COM 在 WPS 上经常拿不到对象。

**双击 `start_silent.vbs` 报错 `800A0408 Invalid character` / 未结束的字符串？**
Windows Script Host 只认 ANSI 和 UTF-16。带 BOM 的 UTF-8 会直接报
`800A0408`；含中文的 UTF-8 无 BOM 文件在中文系统上也会解析失败。
本项目的 `.vbs` 已统一保存为 UTF-16 LE（带 BOM，文件头 `FF FE`）。
如果自己改过该文件，请另存为「UTF-16 LE」，或重新执行
`py -3 install_autostart.py --install` 重新生成。

**备份太多占空间？**
调小 `max_versions_per_file`，或调大 `min_seconds_between_backups`。

## 已知限制

- 仅支持 Windows（依赖 NT 句柄表接口）
- 文件必须已被保存到磁盘过一次；从未保存过的"无标题文档"没有路径，无法备份
- 云端文件（OneDrive 按需下载）可能在本地只是占位符

## 贡献与许可

- 想参与开发？请先阅读 [CONTRIBUTING.md](CONTRIBUTING.md)
- 版本变化见 [CHANGELOG.md](CHANGELOG.md)
- 安全问题请按 [SECURITY.md](SECURITY.md) 私下报告
- 本项目以 [MIT License](LICENSE) 开源
