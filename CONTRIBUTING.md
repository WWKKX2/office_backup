# 贡献指南

感谢你愿意改进 Office Auto Backup！这份文档说明本地开发、测试和提交改动的约定。

## 开发环境

- Windows 10 / 11
- Python 3.8+（推荐 3.10 及以上）
- 不需要安装任何第三方包；`requirements.txt` 里只有可选的 `pywin32`

```bash
git clone https://github.com/WWKKX2/office_backup.git
cd office-auto-backup
copy config.example.json config.json
```

`config.json` 是本地配置，已被 `.gitignore` 忽略，请勿提交。

检查环境是否正常：

```bash
py -3 selftest.py
```

全部通过时应输出 `14/14 项通过`。

## 项目约定

- **零第三方依赖**：主引擎与 GUI 只能使用 Python 标准库。可选功能（如 COM 扫描源）
  必须做到「没装依赖也能正常运行，只是自动跳过」。
- **仅支持 Windows**：句柄扫描依赖 NT 内部接口，不要为了跨平台引入兼容层。
- **保持扁平结构**：所有模块都在仓库根目录，不引入 `src/` 布局。
- **界面使用 tkinter / ttk**：不引入 Qt、Web 前端等重量级方案；
  `run_gui.bat` 必须保持双击可用。
- **注释与界面文案使用中文**，与现有代码保持一致。

## 代码风格

- 4 空格缩进，行宽尽量不超过 100 字符。
- 公开函数写简短的中文 docstring，说明「做什么」而不是「怎么做」。
- 涉及 Windows API 的调用集中放在 `office_handle_scanner.py`，
  其他模块不要直接 `ctypes` 调内核接口。
- 新增用户可见的行为时，同步更新 `README.md` 和 `CHANGELOG.md`。

编辑器配置见 `.editorconfig`；提交前请至少执行：

```bash
py -3 -m py_compile main.py gui.py backup_core.py office_handle_scanner.py install_autostart.py selftest.py
py -3 selftest.py
```

## 特殊文件

- `start_silent.vbs` 必须保存为 **UTF-16 LE（带 BOM，文件头 `FF FE`）**。
  Windows Script Host 无法正确解析带 BOM 的 UTF-8 或含中文的无 BOM UTF-8。
  `.gitattributes` 已将其标记为二进制，避免换行转换破坏编码。
- 任何 `.bat` / `.vbs` 的换行请保持 CRLF。

## 提交信息

推荐使用 [约定式提交](https://www.conventionalcommits.org/zh-hans/)：

```
feat: 新增按文件大小过滤
fix: 修复扫描线程在管道句柄上卡死
docs: 补充开机自启说明
refactor: 拆分备份桶命名逻辑
```

一个提交只做一件事；格式化、重命名等无关改动请单独提交。

## 提交 Pull Request

1. Fork 仓库并从 `main` 创建分支：`feat/xxx`、`fix/xxx`。
2. 确认 `py_compile` 与 `selftest.py` 均通过。
3. 在 PR 描述里说明：改了什么、为什么改、如何验证。
4. 涉及界面改动的，请附上一张截图（注意遮挡个人路径）。
5. CI（`.github/workflows/ci.yml`）会在 Windows 上跑编译检查与自检，必须为绿色。

## 报告问题

请使用仓库的 Issue 模板，并附上：

- 操作系统版本、Python 版本（`py -3 --version`）
- 是否以管理员身份运行
- 相关日志（`backup.log`，注意删除个人信息）
- 复现步骤与期望行为

安全问题请勿公开提交 Issue，参见 [SECURITY.md](SECURITY.md)。
