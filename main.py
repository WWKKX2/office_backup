#!/usr/bin/env python3
"""
Office Auto Backup —— 自动备份正在编辑的 Office / WPS 文档。

用法
----
    python main.py              常驻运行，按间隔轮询备份
    python main.py --once       只扫描备份一轮就退出
    python main.py --scan       只列出检测到的打开文档，不备份
    python main.py --dry-run    演练模式，只打印不落盘

常用参数：--config、--interval、--log-level、--no-lock
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Dict, List, Set

from console_util import configure_console, echo

APP_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = APP_DIR / "config.json"
LOG_FILE = APP_DIR / "backup.log"
__version__ = "1.0.0"

log = logging.getLogger("office_backup")


# --- 配置 -------------------------------------------------------------------


def load_config(path: Path) -> Dict:
    if not path.is_file():
        raise SystemExit(f"找不到配置文件：{path}")
    try:
        # utf-8-sig：兼容记事本/PowerShell 保存的带 BOM 的 UTF-8
        with open(path, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        raise SystemExit(f"配置文件 {path} 不是合法 JSON：{exc}") from exc
    except UnicodeDecodeError as exc:
        raise SystemExit(f"配置文件 {path} 编码无法识别（请另存为 UTF-8）：{exc}") from exc


def setup_logging(level: str, max_bytes: int = 1_048_576, backup_count: int = 3) -> None:
    handlers: List[logging.Handler] = []
    # pythonw.exe 下没有控制台，sys.stdout 为 None，不能挂 StreamHandler
    if sys.stdout is not None:
        handlers.append(logging.StreamHandler(sys.stdout))
    handlers.append(
        RotatingFileHandler(
            LOG_FILE, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8"
        )
    )
    logging.basicConfig(
        level=getattr(logging, str(level).upper(), logging.INFO),
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
        force=True,
    )


def resolve_extensions(cfg: Dict) -> Set[str]:
    """配置里的扩展名 +（启用 WPS 时）WPS 专属扩展名。"""
    exts = {e.lower() for e in cfg.get("file_extensions", [])}
    if not exts:
        import office_handle_scanner as scanner
        exts = set(scanner.DEFAULT_OFFICE_EXTS)
    if cfg.get("office_apps", {}).get("wps", True):
        exts.update({".wps", ".wpt", ".et", ".ett", ".dps", ".dpt"})
    return exts


# --- 扫描 -------------------------------------------------------------------


def _dedup_key(path: Path) -> str:
    """
    生成用于去重的路径键。

    句柄扫描返回的是长路径（C:\\Users\\Administrator\\...），而 COM 或某些
    环境可能返回 8.3 短路径（C:\\Users\\ADMINI~1\\...）。不归一化的话，
    同一个文件会被当成两个文件重复备份，所以这里统一 realpath + normcase。
    """
    try:
        return os.path.normcase(os.path.realpath(str(path)))
    except OSError:
        return os.path.normcase(str(path))


def _canonical(path: Path) -> Path:
    """统一成完整长路径，避免短路径（8.3）与长路径混用。"""
    try:
        return Path(os.path.realpath(str(path)))
    except OSError:
        return path


def collect_open_files(cfg: Dict) -> List[Path]:
    """按 scan_sources 顺序收集打开着的文档，去重后返回。"""
    import office_handle_scanner as scanner
    import com_source

    exts = resolve_extensions(cfg)
    sources = cfg.get("scan_sources", ["handle", "com"])
    apps_cfg = cfg.get("office_apps", {})

    found: Dict[str, Path] = {}

    for source in sources:
        if source == "handle":
            try:
                procs = scanner.scan_office_processes()
                # 遵循 office_apps 开关：关掉的应用对应的进程不扫。
                # 未知进程（例如自检里模拟 Office 的 python.exe）不在这里过滤。
                allowed = scanner.enabled_exe_names(apps_cfg)
                known = scanner.ALL_OFFICE_EXES
                pids = {
                    pid for pid, name in procs.items()
                    if name in allowed or name not in known
                }
                for p in scanner.scan_office_files(exts, pids=pids):
                    found.setdefault(_dedup_key(p), p)
            except Exception as exc:
                log.error("句柄扫描失败：%s", exc)
        elif source == "com":
            try:
                for p in com_source.get_open_files(apps_cfg):
                    if p.suffix.lower() in exts:
                        found.setdefault(_dedup_key(p), p)
            except Exception as exc:
                log.error("COM 扫描失败：%s", exc)
        else:
            log.warning("未知的扫描源 %r，已忽略（可选：handle / com）", source)

    return sorted((_canonical(p) for p in found.values()), key=lambda p: str(p).lower())


# --- 单轮 -------------------------------------------------------------------


def run_once(manager, cfg: Dict) -> tuple:
    """扫描并备份一轮，返回 (检测到的文件数, 新产生的备份数)。"""
    files = collect_open_files(cfg)
    log.debug("检测到 %d 个打开着的文档", len(files))
    new_backups = 0
    for path in files:
        if manager.consider(path):
            new_backups += 1
    return len(files), new_backups


# --- 单实例锁 ---------------------------------------------------------------


class SingleInstance:
    """用文件锁保证同一时间只有一个备份进程在跑。"""

    def __init__(self, name: str = "office_auto_backup") -> None:
        self.path = Path(os.environ.get("TEMP", ".")) / f"{name}.lock"
        self._fh = None

    def acquire(self) -> bool:
        try:
            import msvcrt
        except ImportError:
            return True  # 非 Windows，直接放行

        try:
            self._fh = open(self.path, "a+")
            self._fh.seek(0, os.SEEK_END)
            if self._fh.tell() == 0:
                self._fh.write("lock")
                self._fh.flush()
            self._fh.seek(0)
            msvcrt.locking(self._fh.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            if self._fh:
                self._fh.close()
                self._fh = None
            return False

    def release(self) -> None:
        if not self._fh:
            return
        try:
            import msvcrt
            self._fh.seek(0)
            msvcrt.locking(self._fh.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        finally:
            self._fh.close()
            self._fh = None


# --- 主循环 -----------------------------------------------------------------


def run_forever(manager, cfg: Dict) -> None:
    interval = float(cfg.get("interval_seconds", 30))
    log.info("开始监听，轮询间隔 %.0f 秒，备份到 %s", interval, manager.root)

    rounds = 0
    while True:
        rounds += 1
        started = time.time()
        try:
            seen, added = run_once(manager, cfg)
            if added:
                log.info("第 %d 轮：检测 %d 个文件，新增 %d 个备份", rounds, seen, added)
            else:
                log.debug("第 %d 轮：检测 %d 个文件，无变化", rounds, seen)
        except KeyboardInterrupt:
            raise
        except Exception:
            log.exception("第 %d 轮扫描出错，将继续运行", rounds)

        elapsed = time.time() - started
        time.sleep(max(0.0, interval - elapsed))


# --- CLI --------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="office_backup",
        description="自动备份正在编辑的 Office / WPS 文档",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python main.py --scan      看看现在有哪些文档被打开\n"
            "  python main.py --once      立刻备份一轮\n"
            "  python main.py             常驻后台运行\n"
        ),
    )
    p.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="配置文件路径")
    p.add_argument("--once", action="store_true", help="只运行一轮后退出")
    p.add_argument("--scan", action="store_true", help="只列出检测到的打开文档，不备份")
    p.add_argument("--dry-run", action="store_true", help="演练模式：只打印不实际复制")
    p.add_argument("--interval", type=float, help="覆盖配置里的轮询间隔（秒）")
    p.add_argument("--log-level", help="覆盖配置里的日志级别（DEBUG/INFO/WARNING）")
    p.add_argument("--no-lock", action="store_true", help="跳过单实例锁")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def main(argv: List[str] | None = None) -> int:
    configure_console()
    args = build_parser().parse_args(argv)

    if sys.platform != "win32":
        print("这个工具依赖 Windows 的句柄表接口，只能在 Windows 上运行。", file=sys.stderr)
        return 2

    cfg = load_config(args.config)
    setup_logging(
        args.log_level or cfg.get("log_level", "INFO"),
        int(cfg.get("log_max_bytes", 1_048_576)),
        int(cfg.get("log_backup_count", 3)),
    )

    if args.interval is not None:
        cfg["interval_seconds"] = args.interval

    import office_handle_scanner as scanner

    if not scanner.is_elevated():
        log.warning("未以管理员身份运行：其他用户进程里的文档可能扫不到")

    # --scan：只看不备份
    if args.scan:
        files = collect_open_files(cfg)
        procs = scanner.scan_office_processes()
        if procs:
            log.info("检测到 %d 个 Office / WPS 进程", len(procs))
        else:
            log.info("没有检测到正在运行的 Office / WPS 进程")
        if files:
            echo(f"\n找到 {len(files)} 个打开着的文档：")
            for f in files:
                echo(f"   {f}")
        else:
            echo("\n没有找到打开着的文档。")
        return 0

    from backup_core import BackupManager

    manager = BackupManager(
        root=cfg.get("backup_root", str(APP_DIR / "backups")),
        max_versions=cfg.get("max_versions_per_file", 20),
        min_interval=cfg.get("min_seconds_between_backups", 30),
        dry_run=args.dry_run,
    )
    try:
        manager.ensure_ready()
    except RuntimeError as exc:
        log.error("%s", exc)
        return 1

    # --once：跑一轮就退出
    if args.once:
        seen, added = run_once(manager, cfg)
        echo(f"\n检测到 {seen} 个打开着的文档，产生 {added} 个新备份。")
        echo(manager.summary())
        return 0 if manager.failed == 0 else 1

    # 常驻模式
    lock = SingleInstance()
    if not args.no_lock and not lock.acquire():
        log.error("已有一个备份进程在运行（锁文件 %s），本次退出。", lock.path)
        return 1

    try:
        run_forever(manager, cfg)
    except KeyboardInterrupt:
        log.info("收到中断信号，正在退出。%s", manager.summary())
    finally:
        lock.release()
    return 0


if __name__ == "__main__":
    sys.exit(main())
