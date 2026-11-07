#!/usr/bin/env python3
"""
install_autostart.py —— 把 Office Auto Backup 装成开机自启

两种模式
--------
startup（默认，推荐）
    在「启动」文件夹放一个 .vbs 启动器，登录后静默拉起 pythonw。
    不需要管理员权限，随时可撤销。

task
    注册 Windows 计划任务，登录时以最高权限启动，不会弹 UAC。
    扫描其他用户进程里的文档时更可靠，但注册需要管理员权限。

用法
----
    python install_autostart.py --install            # 安装（启动文件夹）
    python install_autostart.py --install --mode task
    python install_autostart.py --status             # 查看状态
    python install_autostart.py --uninstall          # 卸载
    python install_autostart.py --install --dry-run  # 只看会做什么
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from glob import glob
from pathlib import Path
from xml.sax.saxutils import escape

from console_util import configure_console

APP_DIR = Path(__file__).resolve().parent
MAIN_PY = APP_DIR / "main.py"
DEFAULT_NAME = "OfficeAutoBackup"

STARTUP_DIR = Path(
    os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")
) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


# --- 公共 -------------------------------------------------------------------


def _pythonw_search_patterns() -> list:
    """常见 Python 安装位置的通配符。"""
    patterns = [r"C:\Python3*\pythonw.exe"]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        patterns.insert(
            0,
            str(Path(local) / "Programs" / "Python" / "Python3*" / "pythonw.exe"),
        )
    return patterns


def _pythonw_usable(candidate: Path) -> bool:
    """过滤掉应用商店的 0 字节别名等无法真正运行的文件。"""
    try:
        return candidate.is_file() and candidate.stat().st_size > 0
    except OSError:
        return False


def find_pythonw() -> Path:
    """
    找到可用的 pythonw.exe（无控制台窗口版本）。

    查找顺序：当前解释器目录 → PATH → 用户安装目录 → C:\\Python3*。
    都找不到就退回当前解释器（会是 python.exe，登录时会闪一下窗口）。
    """
    exe = Path(sys.executable)
    candidates = [
        exe if exe.name.lower() == "pythonw.exe" else exe.with_name("pythonw.exe")
    ]
    found = shutil.which("pythonw.exe")
    if found:
        candidates.append(Path(found))

    for pattern in _pythonw_search_patterns():
        candidates.extend(Path(p) for p in sorted(glob(pattern), reverse=True))

    seen = set()
    for candidate in candidates:
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        if _pythonw_usable(candidate):
            return candidate
    return exe


def find_launcher() -> tuple:
    """
    返回 (可执行文件, 额外参数)。

    优先 pythonw.exe；如果没有，用 Windows 自带的 pyw.exe -3 兜底，
    避免依赖 PATH 里的 pythonw.exe。
    """
    pythonw = find_pythonw()
    if _pythonw_usable(pythonw) and pythonw.name.lower() == "pythonw.exe":
        return pythonw, []

    pyw = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "pyw.exe"
    if pyw.is_file():
        return pyw, ["-3"]
    return pythonw, []


def is_elevated() -> bool:
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def say(msg: str = "") -> None:
    print(msg)


def read_script(path: Path) -> tuple:
    """
    读取脚本文件，自动识别编码，返回 (文本, 编码名)。

    Windows Script Host 只认 ANSI 和 UTF-16（带 BOM）。如果写成
    UTF-8 带 BOM，WSH 会把 EF BB BF 当成非法字符，报 800A0408。
    """
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16"), "UTF-16"
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "UTF-8 BOM（WSH 不支持）"
    try:
        return raw.decode("utf-8"), "UTF-8 无 BOM"
    except UnicodeDecodeError:
        return raw.decode("mbcs", errors="replace"), "ANSI"


# --- startup 模式 -----------------------------------------------------------


def startup_path(name: str) -> Path:
    return STARTUP_DIR / f"{name}.vbs"


def vbs_content(launcher: Path, extra_args: list | None = None) -> str:
    """生成启动器脚本内容。路径用变量拼接，避免多层引号转义出错。"""
    args = " ".join(extra_args or [])
    command = f'"{launcher}"' + (f" {args}" if args else "") + f' "{MAIN_PY}"'
    # VBS 字符串里的双引号要写成两个双引号
    vbs_command = command.replace('"', '""')
    return (
        "' Office Auto Backup —— 登录后静默启动\n"
        "' 由 install_autostart.py 生成，删除本文件即可取消自启\n"
        "Set shell = CreateObject(\"WScript.Shell\")\n"
        f"shell.CurrentDirectory = \"{APP_DIR}\"\n"
        f"shell.Run \"{vbs_command}\", 0, False\n"
    )


def install_startup(name: str, dry_run: bool) -> int:
    launcher, extra = find_launcher()
    target = startup_path(name)
    content = vbs_content(launcher, extra)

    say(f"启动器路径：{target}")
    say(f"解释器    ：{launcher}{(' ' + ' '.join(extra)) if extra else ''}")
    say(f"入口脚本  ：{MAIN_PY}")

    if dry_run:
        say("\n[演练] 将写入以下内容：\n")
        say(content)
        return 0

    if launcher.name.lower() not in ("pythonw.exe", "pyw.exe"):
        say("\n[警告] 没找到 pythonw.exe，将用当前解释器启动，登录时会闪一下黑窗口。")

    try:
        STARTUP_DIR.mkdir(parents=True, exist_ok=True)
        # 必须用 UTF-16（带 BOM）。UTF-8 带 BOM 会被 WSH 判为非法字符 800A0408。
        target.write_text(content, encoding="utf-16")
    except OSError as exc:
        say(f"\n[失败] 写入启动文件夹出错：{exc}")
        return 1

    # 写完回读校验，确认 BOM 和内容都没问题
    raw = target.read_bytes()
    if not raw.startswith(b"\xff\xfe"):
        say("\n[警告] 文件头不是 UTF-16 BOM，WSH 可能无法执行，请检查磁盘编码设置。")
    else:
        say("编码校验：UTF-16 LE with BOM（WSH 可识别）")

    say(f"\n[完成] 已安装，下次登录时会自动在后台启动。")
    say(f"       想立即启动：双击 {target.name}，或双击项目里的 start_silent.vbs")
    return 0


def uninstall_startup(name: str, dry_run: bool) -> int:
    target = startup_path(name)
    if not target.exists():
        say(f"[跳过] 未安装：{target} 不存在")
        return 0
    if dry_run:
        say(f"[演练] 将删除 {target}")
        return 0
    try:
        target.unlink()
    except OSError as exc:
        say(f"[失败] 删除失败：{exc}")
        return 1
    say(f"[完成] 已卸载：{target}")
    return 0


def status_startup(name: str) -> int:
    target = startup_path(name)
    if not target.exists():
        say(f"[未安装] {target} 不存在")
        return 1

    content, encoding = read_script(target)
    say(f"[已安装] {target}")
    say(f"文件编码：{encoding}")

    if encoding.startswith("UTF-8"):
        if encoding.startswith("UTF-8 BOM") or not content.isascii():
            say("")
            say("[需要修复] 这个文件是 UTF-8 编码，Windows Script Host 在中文系统上")
            say("           可能报 800A0408 Invalid character 或字符串未结束。")
            say("           重新执行一次 --install 即可修复（会写成 UTF-16）。")

    say("\n内容如下：\n")
    for line in content.splitlines():
        say("  " + line)
    return 0


# --- task 模式 --------------------------------------------------------------


def task_xml(name: str, launcher: Path, extra_args: list | None = None) -> str:
    """生成计划任务 XML（登录触发 + 最高权限 + 延迟 30 秒）。"""
    user = os.environ.get("USERNAME", "")
    domain = os.environ.get("USERDOMAIN", "")
    user_id = f"{domain}\\{user}" if domain else user
    arguments = " ".join([*(extra_args or []), f'"{MAIN_PY}"'])

    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>自动备份正在编辑的 Office / WPS 文档</Description>
    <URI>\\{escape(name)}</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>PT30S</Delay>
      <UserId>{escape(user_id)}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{escape(user_id)}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{escape(str(launcher))}</Command>
      <Arguments>{escape(arguments)}</Arguments>
      <WorkingDirectory>{escape(str(APP_DIR))}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def run_schtasks(args: list) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["schtasks"] + args,
        capture_output=True,
        text=True,
        encoding="mbcs",
        errors="replace",
    )


def install_task(name: str, dry_run: bool) -> int:
    if not is_elevated():
        say("[失败] 注册计划任务需要管理员权限。")
        say("       请以管理员身份重新运行，或改用默认的 --mode startup（不需要管理员）。")
        return 1

    pythonw, extra = find_launcher()
    xml = task_xml(name, pythonw, extra)
    say(f"任务名称：{name}")
    say(f"解释器  ：{pythonw}{(' ' + ' '.join(extra)) if extra else ''}")
    say(f"入口脚本：{MAIN_PY}")

    if dry_run:
        say("\n[演练] 将注册以下任务定义：\n")
        say(xml)
        return 0

    tmp = Path(tempfile.gettempdir()) / f"{name}_task.xml"
    try:
        tmp.write_text(xml, encoding="utf-16")
        result = run_schtasks(["/Create", "/TN", name, "/XML", str(tmp), "/F"])
    finally:
        tmp.unlink(missing_ok=True)

    if result.returncode != 0:
        say(f"\n[失败] 注册计划任务出错：\n{result.stdout}{result.stderr}")
        return 1

    say("\n[完成] 计划任务已注册，下次登录 30 秒后自动启动（最高权限）。")
    return 0


def uninstall_task(name: str, dry_run: bool) -> int:
    if dry_run:
        say(f"[演练] 将删除计划任务 {name}")
        return 0
    result = run_schtasks(["/Delete", "/TN", name, "/F"])
    if result.returncode != 0:
        say(f"[跳过] 任务不存在或删除失败：{result.stdout}{result.stderr}".strip())
        return 0
    say(f"[完成] 已删除计划任务 {name}")
    return 0


def status_task(name: str) -> int:
    result = run_schtasks(["/Query", "/TN", name, "/FO", "LIST"])
    if result.returncode != 0:
        say(f"[未安装] 计划任务 {name} 不存在")
        return 1
    say(f"[已安装] 计划任务 {name}")
    for line in result.stdout.splitlines():
        if line.strip():
            say("         " + line.strip())
    return 0


# --- CLI --------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="install_autostart",
        description="把 Office Auto Backup 安装为开机自启",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python install_autostart.py --install\n"
            "  python install_autostart.py --install --mode task\n"
            "  python install_autostart.py --status\n"
            "  python install_autostart.py --uninstall\n"
        ),
    )
    action = p.add_mutually_exclusive_group(required=True)
    action.add_argument("--install", action="store_true", help="安装开机自启")
    action.add_argument("--uninstall", action="store_true", help="卸载开机自启")
    action.add_argument("--status", action="store_true", help="查看当前状态")

    p.add_argument(
        "--mode", choices=["startup", "task"], default="startup",
        help="startup=启动文件夹（默认，免管理员）；task=计划任务（需管理员，最高权限）",
    )
    p.add_argument("--name", default=DEFAULT_NAME, help=f"自启项名称，默认 {DEFAULT_NAME}")
    p.add_argument("--dry-run", action="store_true", help="只显示将执行的操作，不落盘")
    return p


def main(argv: list | None = None) -> int:
    configure_console()
    args = build_parser().parse_args(argv)

    if sys.platform != "win32":
        say("自启动脚本只适用于 Windows。")
        return 2

    if not MAIN_PY.is_file():
        say(f"[失败] 找不到入口脚本：{MAIN_PY}")
        return 1

    if args.mode == "startup":
        if args.install:
            return install_startup(args.name, args.dry_run)
        if args.uninstall:
            return uninstall_startup(args.name, args.dry_run)
        return status_startup(args.name)

    if args.install:
        return install_task(args.name, args.dry_run)
    if args.uninstall:
        return uninstall_task(args.name, args.dry_run)
    return status_task(args.name)


if __name__ == "__main__":
    sys.exit(main())
