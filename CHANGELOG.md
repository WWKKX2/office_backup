# 更新日志

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)，
格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [Unreleased]

## [1.0.0] - 2026-11-06

首个公开版本。

### 新增

- 句柄扫描引擎：通过 `NtQuerySystemInformation` + `DuplicateHandle` 直接读取
  系统句柄表，零第三方依赖即可发现 Word / Excel / PowerPoint / WPS 正在编辑的文档。
- 备份引擎：按源目录分桶存放，同名文件互不干扰；基于 `(mtime, size)` 判断内容变化，
  内容未变不重复备份；支持冷却期与每文件版本数上限，自动清理最旧版本。
- 命令行入口 `main.py`：支持 `--scan`、`--once`、`--dry-run`、`--config`、
  `--interval`、`--log-level`、`--no-lock`、`--version`。
- 图形界面 `gui.py`：tkinter 实现，左侧导航 + 监控 / 设置两个页面，
  包含状态胶囊、统计卡片、文档表格、深色运行日志、开机自启管理。
- 开机自启：`install_autostart.py` 支持启动文件夹（默认，免管理员）与计划任务
  （`--mode task`，可扫描其他用户进程）两种模式。
- 单实例锁：命令行版与图形界面共用，避免重复备份。
- 端到端自检 `selftest.py`：无需安装 Office 即可验证扫描、备份、版本清理、
  同名隔离、演练模式与 `main.py --once` 全链路。
- 可选 COM 兜底扫描源（需 pywin32，未安装时自动跳过）。

### 修复

- 扫描时跳过非磁盘文件句柄（管道等），单进程解析超时后自动放弃并进入冷却，
  避免扫描线程被永久阻塞。
- 控制台编码适配，cp936 下中文日志不再乱码。
- 配置文件支持带 / 不带 BOM 的 UTF-8。
- `start_silent.vbs` 统一保存为 UTF-16 LE（带 BOM），避免 Windows Script Host
  解析中文路径时报 `800A0408`。
- `office_apps` 应用开关对句柄扫描同样生效。
