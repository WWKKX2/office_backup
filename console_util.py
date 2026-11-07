#!/usr/bin/env python3
"""
console_util.py

控制台相关的小工具。

Python 3.15 起默认用 UTF-8 输出，而中文 Windows 的控制台代码页通常是
cp936（GBK）。两者不一致时，日志里的中文会变成乱码。这里在启动时把
标准输出/错误切回控制台实际使用的代码页。

如果设置了 PYTHONIOENCODING，或控制台已经是 UTF-8（chcp 65001），
则保持原样，尊重用户/环境的设置。
"""

from __future__ import annotations

import os
import sys

_UTF8_NAMES = {"utf8", "utf_8"}


def configure_console() -> None:
    """让标准输出/错误使用 Windows 控制台的代码页，避免中文乱码。"""
    if os.name != "nt" or os.environ.get("PYTHONIOENCODING"):
        return

    try:
        import ctypes

        codepage = int(ctypes.windll.kernel32.GetConsoleOutputCP())
    except Exception:
        return

    if codepage <= 0 or codepage == 65001:
        return

    encoding = f"cp{codepage}"
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        current = (getattr(stream, "encoding", "") or "").lower()
        if current not in _UTF8_NAMES:
            continue
        try:
            stream.reconfigure(encoding=encoding, errors="replace")
        except (AttributeError, OSError, ValueError):
            # 不是 TextIOWrapper 或已关闭，忽略即可
            pass


def echo(message: str = "") -> None:
    """安全打印：pythonw 下 sys.stdout 为 None，直接 print 会抛异常。"""
    if sys.stdout is None:
        return
    try:
        print(message)
    except (OSError, UnicodeError):
        pass
