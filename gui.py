#!/usr/bin/env python3
"""
gui.py —— Office 自动备份的图形界面

只用 Python 标准库（tkinter），不需要额外依赖。

功能
----
* 开始 / 停止自动监听，立即扫描、立即备份
* 实时显示当前打开的 Word / Excel / PowerPoint / WPS 文档
* 显示每个文档已备份的版本数、最后一次备份时间
* 内置运行日志
* 图形化修改 config.json（备份目录、间隔、版本数、应用开关等）
* 一键安装 / 卸载开机自启

用法：
    pythonw gui.py            # 无控制台窗口
    py -3 gui.py              # 带控制台（方便看报错）
    py -3 gui.py --config 路径
"""

from __future__ import annotations

import argparse
import ctypes
import io
import json
import logging
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

import backup_core
import install_autostart
import main as app
import office_handle_scanner as scanner
from backup_core import BackupManager, bucket_name, version_datetime
from console_util import configure_console

APP_TITLE = "Office 自动备份"
APP_DIR = app.APP_DIR
AUTOSTART_NAME = "OfficeAutoBackup"

FALLBACK_CFG: Dict = {
    "backup_root": str(APP_DIR / "backups"),
    "interval_seconds": 30,
    "min_seconds_between_backups": 30,
    "max_versions_per_file": 20,
    "scan_sources": ["handle", "com"],
    "office_apps": {"word": True, "excel": True, "powerpoint": True, "wps": True},
    "file_extensions": sorted(scanner.DEFAULT_OFFICE_EXTS),
    "log_level": "INFO",
    "log_max_bytes": 1_048_576,
    "log_backup_count": 3,
}

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
LOG_DATEFMT = "%Y-%m-%d %H:%M:%S"

APP_VERSION = app.__version__

# --- 视觉主题 ---------------------------------------------------------------

BG = "#F3F5F9"            # 窗口背景
CARD = "#FFFFFF"          # 卡片 / 页头
BORDER = "#E4E8F0"        # 分隔线、卡片描边
TEXT = "#1F2937"          # 主文字
MUTED = "#6B7280"         # 次要文字
FAINT = "#9CA3AF"         # 更弱的提示
ACCENT = "#2563EB"        # 主色
ACCENT_HOVER = "#1D4ED8"
ACCENT_SOFT = "#EAF0FF"   # 选中态浅底
SUCCESS = "#0F9D58"
SUCCESS_SOFT = "#E7F7EF"
WARNING = "#B45309"
DANGER = "#DC2626"
DANGER_SOFT = "#FEECEC"
SIDEBAR = "#FFFFFF"
LOG_BG = "#0F172A"
LOG_FG = "#CBD5E1"
FONT = "Microsoft YaHei UI"
MONO_FONT = "Consolas"


def round_rect(canvas: tk.Canvas, x1: float, y1: float, x2: float, y2: float,
               radius: float, **kwargs) -> int:
    """在 Canvas 上画一个圆角矩形（用平滑多边形近似）。"""
    r = max(0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=24, **kwargs)


class StatusPill(tk.Canvas):
    """页头右侧的状态胶囊：小圆点 + 文字。"""

    PALETTE = {
        "ok": (SUCCESS_SOFT, SUCCESS),
        "warn": ("#FEF3C7", WARNING),
        "muted": ("#F1F5F9", MUTED),
    }

    def __init__(self, master, scale: float = 1.0) -> None:
        super().__init__(master, bg=CARD, highlightthickness=0, bd=0)
        self.ui_scale = scale
        self._text = "已停止"
        self._tone = "muted"
        self._draw()

    def set(self, text: str, tone: str = "muted") -> None:
        self._text = text
        self._tone = tone
        self._draw()

    def _draw(self) -> None:
        self.delete("all")
        s = self.ui_scale
        pad_x, pad_y = int(12 * s), int(5 * s)
        dot_r = max(2.0, 3.0 * s)
        gap = int(6 * s)
        font = (FONT, 9, "bold")

        probe = self.create_text(0, 0, text=self._text, font=font, anchor="nw")
        box = self.bbox(probe)
        text_w = (box[2] - box[0]) if box else int(len(self._text) * 12 * s)
        text_h = (box[3] - box[1]) if box else int(14 * s)
        self.delete(probe)

        width = int(pad_x * 2 + dot_r * 2 + gap + text_w)
        height = int(pad_y * 2 + max(text_h, dot_r * 2))
        self.configure(width=width, height=height)

        bg, fg = self.PALETTE.get(self._tone, self.PALETTE["muted"])
        round_rect(self, 0, 0, width, height, height / 2, fill=bg, outline="")
        cy = height / 2
        self.create_oval(pad_x, cy - dot_r, pad_x + dot_r * 2, cy + dot_r,
                         fill=fg, outline="")
        self.create_text(pad_x + dot_r * 2 + gap, cy, text=self._text,
                         anchor="w", fill=fg, font=font)


class NavItem(tk.Frame):
    """左侧导航项：选中时显示浅色底 + 左侧强调条。"""

    def __init__(self, master, text: str, command, scale: float = 1.0) -> None:
        super().__init__(master, bg=SIDEBAR, cursor="hand2")
        self.command = command
        self.active = False
        bar_w = max(3, int(3 * scale))
        self.bar = tk.Frame(self, bg=SIDEBAR, width=bar_w)
        self.bar.pack(side="left", fill="y")
        self.label = tk.Label(
            self, text=text, bg=SIDEBAR, fg=TEXT, anchor="w",
            font=(FONT, 10), padx=int(14 * scale), pady=int(9 * scale),
            cursor="hand2",
        )
        self.label.pack(side="left", fill="both", expand=True)

        for widget in (self, self.bar, self.label):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)

    def _paint(self, bg: str, fg: str, bar: str) -> None:
        self.configure(bg=bg)
        self.bar.configure(bg=bar)
        self.label.configure(bg=bg, fg=fg)

    def set_active(self, active: bool) -> None:
        self.active = active
        if active:
            self._paint(ACCENT_SOFT, ACCENT, ACCENT)
        else:
            self._paint(SIDEBAR, TEXT, SIDEBAR)

    def _on_click(self, _event=None) -> None:
        self.command()

    def _on_enter(self, _event=None) -> None:
        if not self.active:
            self._paint("#F5F8FC", TEXT, "#F5F8FC")

    def _on_leave(self, _event=None) -> None:
        if not self.active:
            self._paint(SIDEBAR, TEXT, SIDEBAR)


class CheckItem(tk.Frame):
    """自绘复选框：圆角描边 / 蓝底白色对勾，避免 clam 主题的叉号样式。"""

    def __init__(self, master, text: str, variable: tk.BooleanVar,
                 scale: float = 1.0, bg: str = CARD) -> None:
        super().__init__(master, bg=bg, cursor="hand2")
        self.var = variable
        self.ui_scale = scale
        self.box_bg = bg
        size = max(14, int(16 * scale))
        self.box = tk.Canvas(self, width=size, height=size, bg=bg,
                             highlightthickness=0, bd=0, cursor="hand2")
        self.box.pack(side="left")
        self.label = tk.Label(self, text=text, bg=bg, fg=TEXT,
                              font=(FONT, 9), cursor="hand2")
        self.label.pack(side="left", padx=(max(4, int(6 * scale)), 0))

        for widget in (self, self.box, self.label):
            widget.bind("<Button-1>", self._toggle)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)

        self._trace = self.var.trace_add("write", self._on_var_change)
        self._draw()

    def _toggle(self, _event=None) -> None:
        self.var.set(not self.var.get())

    def _on_var_change(self, *_args) -> None:
        self._draw()

    def _on_enter(self, _event=None) -> None:
        self.label.configure(fg=ACCENT)

    def _on_leave(self, _event=None) -> None:
        self.label.configure(fg=TEXT)

    def _draw(self) -> None:
        s = self.ui_scale
        size = max(14, int(16 * s))
        self.box.delete("all")
        if self.var.get():
            round_rect(self.box, 0.5, 0.5, size - 0.5, size - 0.5,
                       max(3, 4 * s), fill=ACCENT, outline="")
            self.box.create_line(
                4 * s, 8.5 * s, 7 * s, 11.5 * s, 12.5 * s, 5 * s,
                fill="#FFFFFF", width=max(2, int(1.8 * s)),
                capstyle="round", joinstyle="round",
            )
        else:
            round_rect(self.box, 0.5, 0.5, size - 0.5, size - 0.5,
                       max(3, 4 * s), fill=CARD, outline="#CBD5E1",
                       width=max(1, int(s)))


# --- 小工具 -----------------------------------------------------------------


def enable_dpi_awareness() -> None:
    """让界面在高分屏上不发虚。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def human_size(num: int) -> str:
    if num < 1024:
        return f"{num} B"
    for unit in ("KB", "MB", "GB", "TB"):
        num /= 1024.0
        if num < 1024:
            return f"{num:.1f} {unit}"
    return f"{num:.1f} PB"


def open_in_explorer(path: Path) -> None:
    """打开文件夹，并尽量选中指定文件。"""
    try:
        if path.is_file():
            import subprocess
            subprocess.Popen(["explorer", "/select,", str(path)])
        else:
            os.startfile(str(path))
    except Exception:
        logging.getLogger("gui").exception("打开资源管理器失败：%s", path)


class UiLogHandler(logging.Handler):
    """把日志丢进 UI 队列，由主线程渲染。"""

    def __init__(self, sink: "queue.Queue") -> None:
        super().__init__()
        self.sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.sink.put(("log", self.format(record)))
        except Exception:
            pass


def setup_ui_logging(sink: "queue.Queue", level: str) -> None:
    app.setup_logging(level)
    handler = UiLogHandler(sink)
    handler.setFormatter(logging.Formatter(LOG_FORMAT, LOG_DATEFMT))
    logging.getLogger().addHandler(handler)


# --- 后台工作线程 -----------------------------------------------------------


class BackupWorker(threading.Thread):
    """
    所有扫描 / 备份操作都在这个线程里串行执行，避免卡住界面。

    通过命令队列控制，通过事件队列（out）把结果发回 UI。
    """

    def __init__(self, config: Dict, out: "queue.Queue") -> None:
        super().__init__(name="backup-worker", daemon=True)
        self.cfg = config
        self.out = out
        self.cmds: "queue.Queue" = queue.Queue()
        self.lock = app.SingleInstance()
        self.manager: Optional[BackupManager] = None
        self.monitoring = False
        self._quit = False
        self._next_run = 0.0
        self.rounds = 0

    # -- 与 UI 的接口 --------------------------------------------------------

    def post(self, action: str, payload=None) -> None:
        self.cmds.put((action, payload))

    # -- 线程主循环 ----------------------------------------------------------

    def run(self) -> None:
        log = logging.getLogger("gui")
        log.info("后台线程已启动")
        while not self._quit:
            try:
                action, payload = self.cmds.get(timeout=0.2)
            except queue.Empty:
                action, payload = None, None

            if action:
                try:
                    self._handle(action, payload)
                except Exception:
                    log.exception("操作失败：%s", action)
                    self.out.put(("error", f"操作失败：{action}，详情见日志"))

            if self.monitoring and time.monotonic() >= self._next_run:
                try:
                    self._do_round()
                except Exception:
                    log.exception("自动备份出错")
                self._next_run = time.monotonic() + self._interval()

        self._release_lock()
        log.info("后台线程已退出")

    def _handle(self, action: str, payload) -> None:
        if action == "start":
            auto = bool(isinstance(payload, dict) and payload.get("auto"))
            self._start_monitor(auto=auto)
        elif action == "stop":
            self._stop_monitor()
        elif action == "scan":
            self._scan_once()
        elif action == "backup":
            self._backup_once()
        elif action == "backup_file":
            self._backup_one_file(Path(str(payload)))
        elif action == "reload":
            self._reload(payload or {})
        elif action == "quit":
            self._quit = True
        else:
            logging.getLogger("gui").warning("未知命令：%s", action)

    # -- 生命周期 ------------------------------------------------------------

    def _interval(self) -> float:
        try:
            return max(1.0, float(self.cfg.get("interval_seconds", 30)))
        except (TypeError, ValueError):
            return 30.0

    def _new_manager(self, ensure_ready: bool = True) -> BackupManager:
        manager = BackupManager(
            root=self.cfg.get("backup_root", str(APP_DIR / "backups")),
            max_versions=self.cfg.get("max_versions_per_file", 20),
            min_interval=self.cfg.get("min_seconds_between_backups", 30),
            dry_run=False,
        )
        if ensure_ready:
            manager.ensure_ready()
        return manager

    def _release_lock(self) -> None:
        if self.monitoring:
            self.monitoring = False
            self.lock.release()

    def _start_monitor(self, auto: bool = False) -> None:
        if self.monitoring:
            return
        if not self.lock.acquire():
            if auto:
                # 开机自启的后台守护（或另一个界面）正在跑：安静地保持空闲，
                # 不弹错误框打扰用户。
                logging.getLogger("gui").info(
                    "检测到已有备份进程在运行，界面保持空闲（锁文件：%s）", self.lock.path
                )
                self.out.put((
                    "info",
                    "检测到已有备份进程在运行（可能是后台守护或开机自启），界面保持空闲。",
                ))
            else:
                self.out.put((
                    "error",
                    "已有一个备份进程在运行（可能是 main.py、计划任务或另一个界面）。\n"
                    "请先关闭它，或在这里点击「停止」后再试。\n"
                    f"锁文件：{self.lock.path}",
                ))
            self.out.put(("mode", "idle"))
            return
        try:
            self.manager = self._new_manager()
        except Exception as exc:
            self.lock.release()
            self.out.put(("error", str(exc)))
            self.out.put(("mode", "idle"))
            return

        self.monitoring = True
        self._next_run = 0.0
        self.out.put(("mode", "monitoring"))
        self.out.put(("info", f"开始监听，每 {self._interval():.0f} 秒检查一次"))

    def _stop_monitor(self) -> None:
        if not self.monitoring:
            self.out.put(("mode", "idle"))
            return
        self._release_lock()
        self.out.put(("mode", "idle"))
        self.out.put(("info", "已停止监听"))

    def _reload(self, cfg: Dict) -> None:
        old_cfg = self.cfg
        self.cfg = cfg
        if self.monitoring:
            try:
                self.manager = self._new_manager()
            except Exception as exc:
                self.cfg = old_cfg
                self.out.put(("error", f"新设置无法应用：{exc}"))
                return
            self._next_run = 0.0
        self.out.put(("info", "设置已生效"))

    # -- 扫描 / 备份 ---------------------------------------------------------

    def _current_manager(self) -> BackupManager:
        if self.manager is None:
            self.manager = self._new_manager(ensure_ready=False)
        return self.manager

    def _rows(self, files: List[Path], manager: Optional[BackupManager] = None) -> List[Dict]:
        manager = manager or self._current_manager()
        rows: List[Dict] = []
        for path in files:
            try:
                st = path.stat()
                size = human_size(st.st_size)
                mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            except OSError:
                size, mtime = "—", "—"

            versions = manager.versions(path)
            last = version_datetime(versions[0]) if versions else None
            rows.append({
                "path": str(path),
                "name": path.name,
                "size": size,
                "mtime": mtime,
                "versions": len(versions),
                "last": last.strftime("%Y-%m-%d %H:%M:%S") if last else "—",
                "folder": str(path.parent),
                "bucket": str(manager.root / bucket_name(path)),
            })
        rows.sort(key=lambda r: r["name"].lower())
        return rows

    def _do_round(self) -> None:
        if self.manager is None:
            self.manager = self._new_manager()

        files = app.collect_open_files(self.cfg)
        added = 0
        for path in files:
            if self.manager.consider(path):
                added += 1

        self.rounds += 1
        logging.getLogger("gui").info(
            "第 %d 轮：检测 %d 个文件，新增 %d 个备份", self.rounds, len(files), added
        )
        self.out.put(("rows", self._rows(files)))
        self.out.put((
            "stats",
            {
                "found": len(files),
                "added": added,
                "failed": self.manager.failed,
                "message": (
                    f"第 {self.rounds} 轮：检测 {len(files)} 个文档，"
                    f"新增 {added} 个备份"
                ),
            },
        ))

    def _scan_once(self) -> None:
        files = app.collect_open_files(self.cfg)
        manager = self._current_manager()
        self.out.put(("rows", self._rows(files, manager)))
        self.out.put((
            "stats",
            {
                "found": len(files),
                "added": 0,
                "failed": 0,
                "message": f"扫描完成：找到 {len(files)} 个打开着的文档（未备份）",
            },
        ))

    def _backup_once(self) -> None:
        if self.monitoring:
            self._next_run = 0.0
            self.out.put(("info", "已安排立即执行一轮备份"))
            return

        if not self.lock.acquire():
            self.out.put((
                "error",
                "已有一个备份进程在运行，无法再启动一次备份。\n"
                f"锁文件：{self.lock.path}",
            ))
            return
        try:
            manager = self._new_manager()
            files = app.collect_open_files(self.cfg)
            added = 0
            for path in files:
                if manager.consider(path):
                    added += 1
            self.manager = manager
            self.out.put(("rows", self._rows(files, manager)))
            self.out.put((
                "stats",
                {
                    "found": len(files),
                    "added": added,
                    "failed": manager.failed,
                    "message": (
                        f"立即备份完成：检测 {len(files)} 个文档，"
                        f"新增 {added} 个备份"
                    ),
                },
            ))
        finally:
            self.lock.release()

    def _backup_one_file(self, path: Path) -> None:
        if not path.is_file():
            self.out.put(("error", f"文件不存在：{path}"))
            return

        own_lock = False
        manager = self.manager if self.monitoring else None
        if manager is None:
            if not self.lock.acquire():
                self.out.put((
                    "error",
                    "已有一个备份进程在运行，无法再启动一次备份。\n"
                    f"锁文件：{self.lock.path}",
                ))
                return
            own_lock = True
            manager = self._new_manager()
        try:
            result = manager.consider(path)
            if result is None:
                self.out.put(("info", f"{path.name}：没有产生新备份（内容未变或仍在冷却期）"))
            else:
                self.out.put(("info", f"已备份 {path.name} → {result.name}"))
            files = app.collect_open_files(self.cfg)
            self.out.put(("rows", self._rows(files, manager)))
        finally:
            if own_lock:
                self.lock.release()


# --- 主窗口 -----------------------------------------------------------------


class BackupApp(tk.Tk):
    def __init__(self, config_path: Path) -> None:
        super().__init__()

        self.cfg_path = Path(config_path)
        self.cfg: Dict = self._load_config()
        self.ui: "queue.Queue" = queue.Queue()
        self._row_data: Dict[str, Dict] = {}
        self._closing = False

        self.title(APP_TITLE)
        self._apply_default_geometry()
        self.configure(bg=BG)
        self._set_app_icon()

        self._setup_theme()
        self._build_ui()
        self._fill_settings()

        setup_ui_logging(self.ui, str(self.cfg.get("log_level", "INFO")))

        self.worker = BackupWorker(self.cfg, self.ui)
        self.worker.start()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(150, self._poll)
        self.after(250, self._auto_start)

        logging.getLogger("gui").info("界面已启动，配置文件：%s", self.cfg_path)
        self._refresh_autostart_status()

    # -- 初始化 --------------------------------------------------------------

    def _apply_default_geometry(self) -> None:
        """
        按系统 DPI 放大默认窗口尺寸。

        Tk 的字体/控件会随 DPI 放大，但 geometry 用的是物理像素；
        不补偿的话，150% 缩放的屏幕上窗口会显得很挤。
        """
        try:
            scale = max(1.0, self.winfo_fpixels("1i") / 96.0)
        except tk.TclError:
            scale = 1.0
        self.ui_scale = scale
        width = min(int(1180 * scale), int(self.winfo_screenwidth() * 0.94))
        height = min(int(780 * scale), int(self.winfo_screenheight() * 0.90))
        self.geometry(f"{width}x{height}+{max(0, (self.winfo_screenwidth() - width) // 2)}+"
                      f"{max(0, (self.winfo_screenheight() - height) // 3)}")
        self.minsize(min(int(980 * scale), width), min(int(640 * scale), height))

    def sp(self, value: float) -> int:
        """按 DPI 缩放像素值。"""
        return int(round(value * getattr(self, "ui_scale", 1.0)))

    def _set_app_icon(self) -> None:
        """用 PhotoImage 画一个简单的应用图标（蓝色圆角底 + 白色文档）。"""
        size = 32
        img = tk.PhotoImage(width=size, height=size)
        img.put(ACCENT, to=(0, 0, size, size))
        img.put("#FFFFFF", to=(9, 6, 23, 26))
        for y in (12, 16, 20, 24):
            img.put(ACCENT, to=(12, y, 20, y + 1))
        self.iconphoto(True, img)
        self._app_icon = img

    def _load_config(self) -> Dict:
        try:
            cfg = app.load_config(self.cfg_path)
            merged = dict(FALLBACK_CFG)
            merged.update(cfg)
            return merged
        except SystemExit as exc:
            messagebox.showwarning(
                APP_TITLE,
                f"读取配置失败，将使用默认设置：\n\n{exc}",
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning(APP_TITLE, f"读取配置失败，将使用默认设置：\n\n{exc}")
        return dict(FALLBACK_CFG)

    def _setup_theme(self) -> None:
        scale = getattr(self, "ui_scale", 1.0)

        def px(value: float) -> int:
            return int(value * scale)

        for name in (
            "TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
            "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont",
        ):
            try:
                tkfont.nametofont(name).configure(family=FONT, size=9)
            except tk.TclError:
                pass

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=BG, foreground=TEXT, font=(FONT, 9))
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT)

        # 普通按钮：白底、细描边、悬浮变浅
        style.configure(
            "TButton", background=CARD, foreground=TEXT, bordercolor=BORDER,
            lightcolor=CARD, darkcolor=CARD, focuscolor=CARD, relief="flat",
            padding=(px(12), px(6)), font=(FONT, 9),
        )
        style.map(
            "TButton",
            background=[("disabled", "#F3F4F6"), ("pressed", "#E8EDF5"), ("active", "#F5F8FC")],
            foreground=[("disabled", FAINT)],
            bordercolor=[("active", "#CBD5E1"), ("pressed", "#CBD5E1")],
        )

        # 主按钮：品牌蓝
        style.configure(
            "Accent.TButton", background=ACCENT, foreground="#FFFFFF",
            bordercolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT, focuscolor=ACCENT,
            relief="flat", padding=(px(16), px(7)), font=(FONT, 9, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", "#E2E8F0"), ("pressed", ACCENT_HOVER), ("active", ACCENT_HOVER)],
            foreground=[("disabled", "#9AA6B8")],
            bordercolor=[("disabled", "#E2E8F0"), ("pressed", ACCENT_HOVER), ("active", ACCENT_HOVER)],
        )

        # 小号按钮（自启操作、日志工具栏）
        style.configure(
            "Small.TButton", background=CARD, foreground=TEXT, bordercolor=BORDER,
            lightcolor=CARD, darkcolor=CARD, focuscolor=CARD, relief="flat",
            padding=(px(9), px(4)), font=(FONT, 8),
        )
        style.map(
            "Small.TButton",
            background=[("pressed", "#E8EDF5"), ("active", "#F5F8FC")],
            bordercolor=[("active", "#CBD5E1")],
        )
        style.configure(
            "Quiet.TButton", background=CARD, foreground=MUTED, bordercolor=CARD,
            lightcolor=CARD, darkcolor=CARD, focuscolor=CARD, relief="flat",
            padding=(px(8), px(3)), font=(FONT, 8),
        )
        style.map(
            "Quiet.TButton",
            background=[("pressed", "#E2E8F0"), ("active", "#F1F5F9")],
            foreground=[("active", TEXT)],
        )

        # 输入控件
        style.configure(
            "TEntry", fieldbackground=CARD, background=CARD, foreground=TEXT,
            bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
            insertcolor=TEXT, relief="flat", padding=px(6),
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", ACCENT)],
            lightcolor=[("focus", ACCENT)],
            darkcolor=[("focus", ACCENT)],
        )
        style.configure(
            "TSpinbox", fieldbackground=CARD, background=CARD, foreground=TEXT,
            bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
            arrowcolor=MUTED, arrowsize=px(12), relief="flat", padding=px(4),
        )
        style.map(
            "TSpinbox",
            bordercolor=[("focus", ACCENT)],
            lightcolor=[("focus", ACCENT)],
            darkcolor=[("focus", ACCENT)],
        )
        style.configure(
            "TCombobox", fieldbackground=CARD, background=CARD, foreground=TEXT,
            bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
            arrowcolor=MUTED, padding=px(4), relief="flat",
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", CARD)],
            foreground=[("readonly", TEXT)],
            bordercolor=[("focus", ACCENT)],
            lightcolor=[("focus", ACCENT)],
            darkcolor=[("focus", ACCENT)],
        )

        # 复选框
        style.configure(
            "TCheckbutton", background=CARD, foreground=TEXT,
            indicatorcolor=ACCENT, indicatorbackground=CARD,
            focuscolor=CARD, padding=(0, px(3)), font=(FONT, 9),
        )
        style.map(
            "TCheckbutton",
            background=[("active", CARD)],
            foreground=[("disabled", FAINT)],
        )

        # 表格
        style.configure(
            "Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
            borderwidth=0, relief="flat", rowheight=px(30), font=(FONT, 9),
        )
        style.map(
            "Treeview",
            background=[("selected", ACCENT_SOFT)],
            foreground=[("selected", TEXT)],
        )
        style.configure(
            "Treeview.Heading", background="#F8FAFC", foreground=MUTED,
            relief="flat", borderwidth=0, padding=(px(8), px(7)),
            font=(FONT, 9, "bold"),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#F1F5F9")],
            foreground=[("active", TEXT)],
        )

        # 滚动条
        style.configure(
            "Vertical.TScrollbar", background="#CBD5E1", troughcolor=CARD,
            bordercolor=CARD, arrowcolor=MUTED, relief="flat", arrowsize=px(12),
        )
        style.map("Vertical.TScrollbar", background=[("active", "#94A3B8")])
        style.configure(
            "Dark.Vertical.TScrollbar", background="#334155", troughcolor=LOG_BG,
            bordercolor=LOG_BG, arrowcolor="#94A3B8", relief="flat", arrowsize=px(12),
        )
        style.map(
            "Dark.Vertical.TScrollbar", background=[("active", "#475569")]
        )
        style.configure("TSeparator", background=BORDER)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self._build_header()
        self._build_body()
        self._build_footer()

    def _build_header(self) -> None:
        s = self.sp
        header = tk.Frame(self, bg=CARD)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        logo = tk.Canvas(header, width=s(40), height=s(40), bg=CARD, highlightthickness=0)
        logo.grid(row=0, column=0, rowspan=2, padx=(s(18), s(12)), pady=(s(12), s(11)))
        round_rect(logo, 0, 0, s(40), s(40), s(11), fill=ACCENT, outline="")
        logo.create_rectangle(s(11), s(8), s(29), s(32), fill="#FFFFFF", outline="")
        for y in (15, 20, 25):
            logo.create_line(s(15), s(y), s(25), s(y),
                             fill=ACCENT, width=max(1, s(1)))

        title_box = tk.Frame(header, bg=CARD)
        title_box.grid(row=0, column=1, rowspan=2, sticky="w", pady=(s(11), s(10)))
        tk.Label(title_box, text=APP_TITLE, bg=CARD, fg=TEXT,
                 font=(FONT, 14, "bold")).pack(anchor="w")
        tk.Label(title_box, text="自动保存正在编辑的 Office / WPS 文档版本",
                 bg=CARD, fg=MUTED, font=(FONT, 9)).pack(anchor="w", pady=(s(1), 0))

        right = tk.Frame(header, bg=CARD)
        right.grid(row=0, column=2, rowspan=2, sticky="e",
                   padx=(s(8), s(18)), pady=(s(11), s(10)))

        self.pill = StatusPill(right, scale=self.ui_scale)
        self.pill.grid(row=0, column=0, padx=(0, s(14)))

        self.btn_open = ttk.Button(right, text="打开备份目录", command=self._open_backup_root)
        self.btn_scan = ttk.Button(right, text="立即扫描", command=self._scan_now)
        self.btn_backup = ttk.Button(right, text="立即备份", command=self._backup_now)
        self.btn_stop = ttk.Button(right, text="停止", command=self._stop_monitor, state="disabled")
        self.btn_start = ttk.Button(right, text="开始监听", style="Accent.TButton",
                                    command=self._start_monitor)

        self.btn_open.grid(row=0, column=1, padx=s(3))
        self.btn_scan.grid(row=0, column=2, padx=s(3))
        self.btn_backup.grid(row=0, column=3, padx=s(3))
        self.btn_stop.grid(row=0, column=4, padx=(s(10), s(3)))
        self.btn_start.grid(row=0, column=5, padx=(s(3), 0))

        tk.Frame(header, bg=BORDER, height=1).grid(
            row=2, column=0, columnspan=3, sticky="ew"
        )

    def _build_body(self) -> None:
        s = self.sp
        body = tk.Frame(self, bg=BG)
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        sidebar = tk.Frame(body, bg=SIDEBAR, width=s(176))
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)

        tk.Label(sidebar, text="工作区", bg=SIDEBAR, fg=FAINT,
                 font=(FONT, 8, "bold")).pack(anchor="w", padx=s(20), pady=(s(18), s(6)))

        self.nav_items: Dict[str, NavItem] = {}
        for key, label, action in (
            ("monitor", "◉   监控", lambda: self._show_page("monitor")),
            ("settings", "⚙   设置", lambda: self._show_page("settings")),
        ):
            item = NavItem(sidebar, label, action, scale=self.ui_scale)
            item.pack(fill="x")
            self.nav_items[key] = item

        tk.Frame(sidebar, bg=SIDEBAR).pack(fill="both", expand=True)
        tk.Label(sidebar, text=f"v{APP_VERSION}", bg=SIDEBAR, fg=FAINT,
                 font=(FONT, 8)).pack(anchor="w", padx=s(21), pady=(0, s(14)))

        self.content = tk.Frame(body, bg=BG)
        self.content.grid(row=0, column=1, sticky="nsew",
                          padx=(s(18), s(18)), pady=(s(16), s(6)))
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        self.pages: Dict[str, tk.Frame] = {}
        self._build_monitor_page()
        self._build_settings_page()
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")
        self._show_page("monitor")

    def _show_page(self, name: str) -> None:
        page = self.pages.get(name)
        if page is None:
            return
        page.tkraise()
        for key, item in self.nav_items.items():
            item.set_active(key == name)

    def _make_card(self, parent) -> tk.Frame:
        return tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)

    def _stat_card(self, parent, title: str, variable: tk.StringVar,
                   color: str, column: int) -> None:
        s = self.sp
        card = self._make_card(parent)
        card.grid(row=0, column=column, sticky="nsew",
                  padx=(0 if column == 0 else s(6), 0))
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="both", expand=True, padx=s(14), pady=s(9))
        tk.Label(inner, text=title, bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(anchor="w")
        tk.Label(inner, textvariable=variable, bg=CARD, fg=color,
                 font=(FONT, 15, "bold")).pack(anchor="w", pady=(s(1), 0))

    def _build_monitor_page(self) -> None:
        s = self.sp
        page = tk.Frame(self.content, bg=BG)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(2, weight=1)
        self.pages["monitor"] = page

        head = tk.Frame(page, bg=BG)
        head.grid(row=0, column=0, sticky="ew", pady=(0, s(10)))
        head.columnconfigure(0, weight=1)
        tk.Label(head, text="监控", bg=BG, fg=TEXT,
                 font=(FONT, 13, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(head, text="双击某一行打开文件所在文件夹 · 右键有更多操作",
                 bg=BG, fg=MUTED, font=(FONT, 8)).grid(row=0, column=1, sticky="e")

        # 统计卡片
        stats = tk.Frame(page, bg=BG)
        stats.grid(row=1, column=0, sticky="ew", pady=(0, s(10)))
        for i in range(4):
            stats.columnconfigure(i, weight=1, uniform="stat")

        self.stat_docs_var = tk.StringVar(value="0")
        self.stat_total_var = tk.StringVar(value="0")
        self.stat_added_var = tk.StringVar(value="0")
        self.stat_failed_var = tk.StringVar(value="0")
        self._stat_card(stats, "打开的文档", self.stat_docs_var, TEXT, 0)
        self._stat_card(stats, "备份版本", self.stat_total_var, TEXT, 1)
        self._stat_card(stats, "本轮新增", self.stat_added_var, ACCENT, 2)
        self._stat_card(stats, "失败次数", self.stat_failed_var, DANGER, 3)

        # 文档表格
        table_card = self._make_card(page)
        table_card.grid(row=2, column=0, sticky="nsew")
        table_card.columnconfigure(0, weight=1)
        table_card.rowconfigure(1, weight=1)

        table_head = tk.Frame(table_card, bg=CARD)
        table_head.grid(row=0, column=0, sticky="ew",
                        padx=s(16), pady=(s(12), s(6)))
        table_head.columnconfigure(1, weight=1)
        tk.Label(table_head, text="打开的文档", bg=CARD, fg=TEXT,
                 font=(FONT, 10, "bold")).grid(row=0, column=0, sticky="w")
        self.summary_var = tk.StringVar(value="尚未扫描")
        tk.Label(table_head, textvariable=self.summary_var, bg=CARD, fg=MUTED,
                 font=(FONT, 8)).grid(row=0, column=1, sticky="e")

        tree_body = tk.Frame(table_card, bg=CARD)
        tree_body.grid(row=1, column=0, sticky="nsew",
                       padx=(s(8), s(10)), pady=(0, s(10)))
        tree_body.columnconfigure(0, weight=1)
        tree_body.rowconfigure(0, weight=1)

        columns = ("name", "size", "mtime", "versions", "last", "folder")
        self.tree = ttk.Treeview(
            tree_body, columns=columns, show="headings", selectmode="browse",
        )
        headings = {
            "name": ("文件名", 215),
            "size": ("大小", 70),
            "mtime": ("修改时间", 145),
            "versions": ("备份版本", 75),
            "last": ("最后备份", 145),
            "folder": ("所在目录", 270),
        }
        for col, (title, width) in headings.items():
            anchor = "center" if col in ("size", "versions") else "w"
            self.tree.heading(col, text=title, anchor=anchor)
            self.tree.column(col, width=s(width), anchor=anchor,
                             stretch=(col in ("name", "folder")))

        vsb = ttk.Scrollbar(tree_body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.tag_configure("odd", background="#FAFBFD")
        self.tree.tag_configure("never", foreground=FAINT)
        self.tree.bind("<Double-1>", lambda _e: self._open_selected())
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        self.empty_label = tk.Label(
            tree_body, text="没有检测到正在编辑的 Office / WPS 文档",
            bg=CARD, fg=FAINT, font=(FONT, 10),
        )
        self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

        self.context_menu = tk.Menu(
            self, tearoff=0, bg=CARD, fg=TEXT, bd=0, relief="flat",
            activebackground=ACCENT_SOFT, activeforeground=ACCENT,
        )
        self.context_menu.add_command(label="打开所在文件夹", command=self._open_selected)
        self.context_menu.add_command(label="打开备份目录", command=self._open_selected_bucket)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="立即备份此文件", command=self._backup_selected)
        self.context_menu.add_command(label="复制文件路径", command=self._copy_selected_path)

        # 日志面板（深色，类似终端）
        log_card = self._make_card(page)
        log_card.grid(row=3, column=0, sticky="ew", pady=(s(10), 0))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        log_head = tk.Frame(log_card, bg=CARD)
        log_head.grid(row=0, column=0, sticky="ew", padx=s(16), pady=(s(10), s(6)))
        log_head.columnconfigure(0, weight=1)
        tk.Label(log_head, text="运行日志", bg=CARD, fg=TEXT,
                 font=(FONT, 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Button(log_head, text="打开日志文件", style="Quiet.TButton",
                   command=self._open_log_file).grid(row=0, column=1, padx=(0, s(2)))
        ttk.Button(log_head, text="清空显示", style="Quiet.TButton",
                   command=self._clear_log_view).grid(row=0, column=2)

        log_body = tk.Frame(log_card, bg=CARD)
        log_body.grid(row=1, column=0, sticky="nsew", padx=s(14), pady=(0, s(12)))
        log_body.columnconfigure(0, weight=1)
        log_body.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_body, height=7, wrap="word", font=(MONO_FONT, 9),
            state="disabled", bg=LOG_BG, fg=LOG_FG, insertbackground=LOG_FG,
            relief="flat", borderwidth=0, highlightthickness=0,
            padx=s(12), pady=s(8), selectbackground="#1E3A8A",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_sb = ttk.Scrollbar(log_body, orient="vertical",
                               style="Dark.Vertical.TScrollbar",
                               command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        log_sb.grid(row=0, column=1, sticky="ns")

        self.log_text.tag_configure("info", foreground=LOG_FG)
        self.log_text.tag_configure("warn", foreground="#FBBF24")
        self.log_text.tag_configure("error", foreground="#F87171")
        self.log_text.tag_configure("debug", foreground="#64748B")

    def _card(self, parent, title: str, hint: str = "") -> tuple:
        s = self.sp
        card = self._make_card(parent)
        card.columnconfigure(0, weight=1)
        head = tk.Frame(card, bg=CARD)
        head.grid(row=0, column=0, sticky="ew", padx=s(16), pady=(s(13), s(3)))
        tk.Label(head, text=title, bg=CARD, fg=TEXT,
                 font=(FONT, 10, "bold")).pack(anchor="w")
        if hint:
            tk.Label(head, text=hint, bg=CARD, fg=MUTED,
                     font=(FONT, 8)).pack(anchor="w", pady=(s(2), 0))
        body = tk.Frame(card, bg=CARD)
        body.grid(row=1, column=0, sticky="nsew", padx=s(16), pady=(s(8), s(13)))
        return card, body

    def _field_label(self, parent, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=CARD, fg=TEXT, font=(FONT, 9))

    def _hint_label(self, parent, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=CARD, fg=MUTED, font=(FONT, 8))

    def _build_settings_page(self) -> None:
        s = self.sp
        page = tk.Frame(self.content, bg=BG)
        page.columnconfigure(0, weight=1, uniform="settings")
        page.columnconfigure(1, weight=1, uniform="settings")
        page.rowconfigure(2, weight=1)
        self.pages["settings"] = page

        head = tk.Frame(page, bg=BG)
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, s(10)))
        head.columnconfigure(0, weight=1)
        tk.Label(head, text="设置", bg=BG, fg=TEXT,
                 font=(FONT, 13, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(head, text="保存后立即生效", bg=BG, fg=MUTED,
                 font=(FONT, 8)).grid(row=0, column=1, sticky="e", padx=(0, s(10)))
        ttk.Button(head, text="重新载入配置",
                   command=self._reload_settings).grid(
            row=0, column=2, padx=(0, s(8))
        )
        ttk.Button(head, text="保存设置", style="Accent.TButton",
                   command=self._save_settings).grid(row=0, column=3)

        self.var_root = tk.StringVar()
        self.var_interval = tk.StringVar()
        self.var_min_interval = tk.StringVar()
        self.var_max_versions = tk.StringVar()
        self.var_exts = tk.StringVar()

        self.var_handle = tk.BooleanVar()
        self.var_com = tk.BooleanVar()
        self.var_word = tk.BooleanVar()
        self.var_excel = tk.BooleanVar()
        self.var_powerpoint = tk.BooleanVar()
        self.var_wps = tk.BooleanVar()
        self.var_level = tk.StringVar()

        # 左：备份
        backup_card, backup_body = self._card(
            page, "备份", "决定备份保存到哪里、多久检查一次"
        )
        backup_card.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=(0, s(8)))
        backup_body.columnconfigure(1, weight=1)

        row = 0
        self._field_label(backup_body, "备份目录").grid(
            row=row, column=0, sticky="w", pady=s(5)
        )
        ttk.Entry(backup_body, textvariable=self.var_root).grid(
            row=row, column=1, sticky="ew", padx=(s(10), s(6)), pady=s(5)
        )
        ttk.Button(backup_body, text="浏览…", style="Small.TButton",
                   command=self._browse_root).grid(row=row, column=2, pady=s(5))

        row += 1
        self._field_label(backup_body, "轮询间隔（秒）").grid(
            row=row, column=0, sticky="w", pady=s(5)
        )
        ttk.Spinbox(
            backup_body, from_=1, to=86400, textvariable=self.var_interval,
            width=10, justify="left",
        ).grid(row=row, column=1, sticky="w", padx=(s(10), s(6)), pady=s(5))
        self._hint_label(backup_body, "最小 1 秒").grid(
            row=row, column=2, sticky="w", pady=s(5)
        )

        row += 1
        self._field_label(backup_body, "同一文件冷却期（秒）").grid(
            row=row, column=0, sticky="w", pady=s(5)
        )
        ttk.Spinbox(
            backup_body, from_=0, to=86400, textvariable=self.var_min_interval,
            width=10, justify="left",
        ).grid(row=row, column=1, sticky="w", padx=(s(10), s(6)), pady=s(5))
        self._hint_label(backup_body, "防止版本刷爆").grid(
            row=row, column=2, sticky="w", pady=s(5)
        )

        row += 1
        self._field_label(backup_body, "每个文件保留版本数").grid(
            row=row, column=0, sticky="w", pady=s(5)
        )
        ttk.Spinbox(
            backup_body, from_=1, to=500, textvariable=self.var_max_versions,
            width=10, justify="left",
        ).grid(row=row, column=1, sticky="w", padx=(s(10), s(6)), pady=s(5))

        row += 1
        self._field_label(backup_body, "备份的扩展名").grid(
            row=row, column=0, sticky="nw", pady=(s(8), s(2))
        )
        ttk.Entry(backup_body, textvariable=self.var_exts).grid(
            row=row, column=1, columnspan=2, sticky="ew",
            padx=(s(10), 0), pady=(s(8), s(2)),
        )
        row += 1
        self._hint_label(
            backup_body, "用逗号或空格分隔，例如 .docx .xlsx .pptx"
        ).grid(row=row, column=1, columnspan=2, sticky="w", padx=(s(10), 0))

        # 右上：扫描
        scan_card, scan_body = self._card(
            page, "扫描", "选择扫描方式和需要监控的应用"
        )
        scan_card.grid(row=1, column=1, sticky="new")
        scan_body.columnconfigure(0, weight=1)

        self._field_label(scan_body, "扫描方式").grid(row=0, column=0, sticky="w")
        sources = tk.Frame(scan_body, bg=CARD)
        sources.grid(row=1, column=0, sticky="w", pady=(s(2), s(6)))
        CheckItem(sources, "句柄扫描（推荐）", self.var_handle,
                  scale=self.ui_scale).pack(side="left")
        CheckItem(sources, "COM 兜底", self.var_com,
                  scale=self.ui_scale).pack(side="left", padx=(s(14), 0))
        self._hint_label(
            scan_body, "COM 需要 pywin32；没装会自动跳过。"
        ).grid(row=2, column=0, sticky="w")

        ttk.Separator(scan_body, orient="horizontal").grid(
            row=3, column=0, sticky="ew", pady=s(9)
        )

        self._field_label(scan_body, "监控哪些应用").grid(
            row=4, column=0, sticky="w"
        )
        apps = tk.Frame(scan_body, bg=CARD)
        apps.grid(row=5, column=0, sticky="w", pady=(s(2), 0))
        CheckItem(apps, "Word", self.var_word, scale=self.ui_scale).grid(
            row=0, column=0, sticky="w"
        )
        CheckItem(apps, "Excel", self.var_excel, scale=self.ui_scale).grid(
            row=0, column=1, sticky="w", padx=s(18)
        )
        CheckItem(apps, "PowerPoint", self.var_powerpoint, scale=self.ui_scale).grid(
            row=1, column=0, sticky="w", pady=(s(4), 0)
        )
        CheckItem(apps, "WPS", self.var_wps, scale=self.ui_scale).grid(
            row=1, column=1, sticky="w", padx=s(18), pady=(s(4), 0)
        )

        # 右下：日志与开机自启
        misc_card, misc_body = self._card(page, "日志与开机自启")
        misc_card.grid(row=2, column=1, sticky="nsew", pady=(s(10), 0))
        misc_body.columnconfigure(1, weight=1)

        self._field_label(misc_body, "日志级别").grid(
            row=0, column=0, sticky="w", pady=s(3)
        )
        ttk.Combobox(
            misc_body, textvariable=self.var_level,
            values=("DEBUG", "INFO", "WARNING", "ERROR"),
            state="readonly", width=12,
        ).grid(row=0, column=1, sticky="w", padx=(s(10), 0), pady=s(3))

        ttk.Separator(misc_body, orient="horizontal").grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=s(8)
        )

        self._field_label(misc_body, "开机自启").grid(
            row=2, column=0, sticky="w"
        )
        self.auto_status = tk.Label(
            misc_body, text="检查中…", bg=CARD, fg=MUTED, font=(FONT, 9)
        )
        self.auto_status.grid(row=2, column=1, sticky="w", padx=(s(10), 0))

        auto_buttons = tk.Frame(misc_body, bg=CARD)
        auto_buttons.grid(row=3, column=0, columnspan=3, sticky="w", pady=(s(8), 0))
        ttk.Button(auto_buttons, text="安装", style="Small.TButton",
                   command=self._install_autostart).pack(side="left")
        ttk.Button(auto_buttons, text="卸载", style="Small.TButton",
                   command=self._uninstall_autostart).pack(side="left", padx=s(6))
        ttk.Button(auto_buttons, text="刷新状态", style="Small.TButton",
                   command=self._refresh_autostart_status).pack(side="left")
        ttk.Button(auto_buttons, text="打开启动文件夹", style="Small.TButton",
                   command=self._open_startup_folder).pack(side="left", padx=(s(6), 0))

        self._hint_label(
            misc_body, "登录后在后台自动备份，不打开本界面；无需管理员权限。"
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(s(8), 0))

    def _build_footer(self) -> None:
        s = self.sp
        footer = tk.Frame(self, bg=BG)
        footer.grid(row=2, column=0, sticky="ew")
        tk.Frame(footer, bg=BORDER, height=1).pack(fill="x")

        inner = tk.Frame(footer, bg=BG)
        inner.pack(fill="x", padx=s(18), pady=s(6))
        inner.columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="就绪")
        tk.Label(inner, textvariable=self.status_var, bg=BG, fg=MUTED,
                 font=(FONT, 8)).grid(row=0, column=0, sticky="w")
        tk.Label(
            inner,
            text=f"v{APP_VERSION}  ·  Python {sys.version.split()[0]}  ·  {APP_DIR}",
            bg=BG, fg=FAINT, font=(FONT, 8),
        ).grid(row=0, column=1, sticky="e")

    # -- 设置同步 ------------------------------------------------------------

    def _fill_settings(self) -> None:
        cfg = self.cfg
        self.var_root.set(str(cfg.get("backup_root", "")))
        self.var_interval.set(str(cfg.get("interval_seconds", 30)))
        self.var_min_interval.set(str(cfg.get("min_seconds_between_backups", 30)))
        self.var_max_versions.set(str(cfg.get("max_versions_per_file", 20)))
        exts = cfg.get("file_extensions") or sorted(scanner.DEFAULT_OFFICE_EXTS)
        self.var_exts.set(" ".join(exts))

        sources = set(cfg.get("scan_sources", ["handle", "com"]))
        self.var_handle.set("handle" in sources)
        self.var_com.set("com" in sources)

        apps = cfg.get("office_apps", {})
        self.var_word.set(bool(apps.get("word", True)))
        self.var_excel.set(bool(apps.get("excel", True)))
        self.var_powerpoint.set(bool(apps.get("powerpoint", True)))
        self.var_wps.set(bool(apps.get("wps", True)))
        self.var_level.set(str(cfg.get("log_level", "INFO")).upper())

    def _collect_settings(self) -> Dict:
        root = self.var_root.get().strip()
        if not root:
            raise ValueError("备份目录不能为空。")

        try:
            interval = float(self.var_interval.get())
        except ValueError as exc:
            raise ValueError("轮询间隔必须是数字。") from exc
        if interval < 1:
            raise ValueError("轮询间隔不能小于 1 秒。")

        try:
            min_interval = float(self.var_min_interval.get())
        except ValueError as exc:
            raise ValueError("冷却期必须是数字。") from exc
        if min_interval < 0:
            raise ValueError("冷却期不能是负数。")

        try:
            max_versions = int(float(self.var_max_versions.get()))
        except ValueError as exc:
            raise ValueError("保留版本数必须是整数。") from exc
        if max_versions < 1:
            raise ValueError("至少保留 1 个版本。")

        raw_exts = re.split(r"[,;\s]+", self.var_exts.get().strip())
        exts: List[str] = []
        for item in raw_exts:
            item = item.strip().lower()
            if not item:
                continue
            if not item.startswith("."):
                item = "." + item
            if item not in exts:
                exts.append(item)
        if not exts:
            raise ValueError("至少填一个扩展名，例如 .docx。")

        sources = []
        if self.var_handle.get():
            sources.append("handle")
        if self.var_com.get():
            sources.append("com")
        if not sources:
            raise ValueError("至少选择一种扫描方式。")

        return {
            "backup_root": root,
            "interval_seconds": interval,
            "min_seconds_between_backups": min_interval,
            "max_versions_per_file": max_versions,
            "scan_sources": sources,
            "office_apps": {
                "word": self.var_word.get(),
                "excel": self.var_excel.get(),
                "powerpoint": self.var_powerpoint.get(),
                "wps": self.var_wps.get(),
            },
            "file_extensions": exts,
            "log_level": self.var_level.get().upper(),
        }

    # -- 设置动作 ------------------------------------------------------------

    def _browse_root(self) -> None:
        initial = self.var_root.get().strip() or str(APP_DIR)
        chosen = filedialog.askdirectory(title="选择备份目录", initialdir=initial, mustexist=False)
        if chosen:
            self.var_root.set(os.path.normpath(chosen))

    def _save_settings(self) -> None:
        try:
            changes = self._collect_settings()
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self)
            return

        merged = dict(self.cfg)
        merged.update(changes)
        try:
            self.cfg_path.write_text(
                json.dumps(merged, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"保存配置文件失败：\n{exc}", parent=self)
            return

        self.cfg = merged
        setup_ui_logging(self.ui, merged["log_level"])
        self.worker.post("reload", merged)
        logging.getLogger("gui").info("设置已保存：%s", self.cfg_path)
        self.status_var.set("设置已保存")
        self._refresh_autostart_status()

    def _reload_settings(self) -> None:
        try:
            cfg = app.load_config(self.cfg_path)
        except SystemExit as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self)
            return

        merged = dict(FALLBACK_CFG)
        merged.update(cfg)
        self.cfg = merged
        self._fill_settings()
        setup_ui_logging(self.ui, str(merged.get("log_level", "INFO")))
        self.worker.post("reload", merged)
        logging.getLogger("gui").info("已重新载入配置")
        self.status_var.set("已重新载入配置")

    # -- 按钮动作 ------------------------------------------------------------

    def _auto_start(self) -> None:
        try:
            backup_root = Path(str(self.cfg.get("backup_root", "")))
            backup_root.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        self.worker.post("start", {"auto": True})

    def _start_monitor(self) -> None:
        self.worker.post("start")

    def _stop_monitor(self) -> None:
        self.worker.post("stop")

    def _scan_now(self) -> None:
        self.worker.post("scan")

    def _backup_now(self) -> None:
        self.worker.post("backup")

    def _open_backup_root(self) -> None:
        root = Path(str(self.cfg.get("backup_root", "")))
        try:
            root.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"无法打开备份目录：\n{exc}", parent=self)
            return
        open_in_explorer(root)

    def _open_log_file(self) -> None:
        log_path = app.LOG_FILE
        if log_path.is_file():
            open_in_explorer(log_path)
        else:
            messagebox.showinfo(APP_TITLE, "日志文件还没有生成。", parent=self)

    def _clear_log_view(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # -- 表格交互 ------------------------------------------------------------

    def _selected_row(self) -> Optional[Dict]:
        selection = self.tree.selection()
        if not selection:
            return None
        return self._row_data.get(selection[0])

    def _on_tree_select(self, _event=None) -> None:
        row = self._selected_row()
        if row:
            self.status_var.set(f"{row['path']}    （备份桶：{Path(row['bucket']).name}）")

    def _on_tree_right_click(self, event) -> None:
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _open_selected(self) -> None:
        row = self._selected_row()
        if row:
            open_in_explorer(Path(row["path"]))

    def _open_selected_bucket(self) -> None:
        row = self._selected_row()
        if not row:
            return
        bucket = Path(row["bucket"])
        if bucket.is_dir():
            open_in_explorer(bucket)
        else:
            messagebox.showinfo(APP_TITLE, "这个文件还没有备份。", parent=self)

    def _backup_selected(self) -> None:
        row = self._selected_row()
        if row:
            self.worker.post("backup_file", row["path"])

    def _copy_selected_path(self) -> None:
        row = self._selected_row()
        if row:
            self.clipboard_clear()
            self.clipboard_append(row["path"])
            self.status_var.set("路径已复制到剪贴板")

    # -- 开机自启 ------------------------------------------------------------

    def _refresh_autostart_status(self) -> None:
        target = install_autostart.startup_path(AUTOSTART_NAME)
        if not target.is_file():
            self.auto_status.configure(text="未安装", fg=MUTED)
            return
        try:
            _content, encoding = install_autostart.read_script(target)
        except OSError:
            self.auto_status.configure(text="已安装（无法读取）", fg=DANGER)
            return
        if encoding == "UTF-16":
            self.auto_status.configure(text=f"已安装（{encoding}）", fg=SUCCESS)
        else:
            self.auto_status.configure(
                text=f"已安装，但编码为 {encoding}，建议重装", fg=DANGER
            )

    def _run_autostart(self, args: List[str]) -> None:
        buf = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = buf
        code = 1
        try:
            code = install_autostart.main(args)
        except Exception as exc:  # noqa: BLE001
            buf.write(f"\n{exc}\n")
        finally:
            sys.stdout = old_stdout

        text = buf.getvalue().strip()
        if text:
            for line in text.splitlines():
                logging.getLogger("gui").info("[自启] %s", line)
        self._refresh_autostart_status()
        if code != 0:
            messagebox.showerror(APP_TITLE, text or "操作失败。", parent=self)
        else:
            messagebox.showinfo(APP_TITLE, text or "操作完成。", parent=self)

    def _install_autostart(self) -> None:
        self._run_autostart(["--install"])

    def _uninstall_autostart(self) -> None:
        self._run_autostart(["--uninstall"])

    def _open_startup_folder(self) -> None:
        open_in_explorer(install_autostart.STARTUP_DIR)

    # -- 事件循环 ------------------------------------------------------------

    def _poll(self) -> None:
        if self._closing:
            return
        try:
            while True:
                kind, payload = self.ui.get_nowait()
                self._handle_event(kind, payload)
        except queue.Empty:
            pass
        self.after(150, self._poll)

    def _handle_event(self, kind: str, payload) -> None:
        if kind == "log":
            self._append_log(str(payload))
        elif kind == "rows":
            self._update_rows(payload)
        elif kind == "stats":
            if isinstance(payload, dict):
                self.stat_added_var.set(str(payload.get("added", 0)))
                self.stat_failed_var.set(str(payload.get("failed", 0)))
                message = str(payload.get("message", ""))
            else:
                message = str(payload)
            if message:
                self.status_var.set(message)
        elif kind == "info":
            logging.getLogger("gui").info("%s", payload)
            self.status_var.set(str(payload))
        elif kind == "mode":
            self._set_mode(str(payload))
        elif kind == "error":
            logging.getLogger("gui").error("%s", payload)
            self.status_var.set(str(payload))
            messagebox.showerror(APP_TITLE, str(payload), parent=self)

    def _set_mode(self, mode: str) -> None:
        monitoring = mode == "monitoring"
        self.pill.set("监听中" if monitoring else "已停止",
                      "ok" if monitoring else "muted")
        self.btn_start.configure(state="disabled" if monitoring else "normal")
        self.btn_stop.configure(state="normal" if monitoring else "disabled")

    def _update_rows(self, rows: List[Dict]) -> None:
        selected = self._selected_row()
        selected_path = selected["path"] if selected else None

        self.tree.delete(*self.tree.get_children())
        self._row_data.clear()

        total_versions = 0
        for index, row in enumerate(rows):
            total_versions += int(row.get("versions", 0))
            tags = ["odd"] if index % 2 else []
            if row.get("versions", 0) == 0:
                tags.append("never")
            item = self.tree.insert(
                "", "end",
                values=(
                    row["name"], row["size"], row["mtime"],
                    row["versions"], row["last"], row["folder"],
                ),
                tags=tuple(tags),
            )
            self._row_data[item] = row
            if selected_path and row["path"] == selected_path:
                self.tree.selection_set(item)
                self.tree.focus(item)

        self.stat_docs_var.set(str(len(rows)))
        self.stat_total_var.set(str(total_versions))
        self.summary_var.set(
            f"{len(rows)} 个打开的文档 · 共 {total_versions} 个备份版本"
        )
        if rows:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def _append_log(self, text: str) -> None:
        # 只有用户已经在底部时才自动滚动，避免打断向上翻看
        try:
            at_bottom = self.log_text.yview()[1] > 0.98
        except tk.TclError:
            at_bottom = True

        line = text.rstrip()
        if "[ERROR]" in line:
            tag = "error"
        elif "[WARNING]" in line:
            tag = "warn"
        elif "[DEBUG]" in line:
            tag = "debug"
        else:
            tag = "info"

        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n", tag)
        # 最多保留 2000 行
        try:
            lines = int(self.log_text.index("end-1c").split(".")[0])
            if lines > 2000:
                self.log_text.delete("1.0", f"{lines - 2000}.0")
        except (tk.TclError, ValueError):
            pass
        if at_bottom:
            self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # -- 关闭 ----------------------------------------------------------------

    def _on_close(self) -> None:
        if self._closing:
            return
        self._closing = True
        try:
            self.worker.post("quit")
            self.worker.join(timeout=3)
        finally:
            self.destroy()


# --- 入口 -------------------------------------------------------------------


def main(argv: List[str] | None = None) -> int:
    configure_console()
    enable_dpi_awareness()

    parser = argparse.ArgumentParser(description="Office 自动备份图形界面")
    parser.add_argument("--config", type=Path, default=app.DEFAULT_CONFIG_PATH,
                        help="配置文件路径（默认同目录 config.json）")
    args = parser.parse_args(argv)

    application = BackupApp(args.config)
    application.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
