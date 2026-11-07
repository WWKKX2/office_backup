#!/usr/bin/env python3
"""
setup_autostart.py —— 兼容旧命令名的开机自启脚本

这个文件以前有一套自己的 VBS 生成逻辑，存在两个问题：
  1. 用 UTF-8 写 .vbs，Windows Script Host 在中文系统上可能解析失败；
  2. 把路径里的反斜杠双写（VBS 字符串不做反斜杠转义），生成的路径是错的。

现在统一转发到 install_autostart.py，只保留原来的命令行用法：

    python setup_autostart.py              查看状态（等同 --status）
    python setup_autostart.py --install    安装自启
    python setup_autostart.py --uninstall  卸载自启
"""

from __future__ import annotations

import sys

import install_autostart


def main(argv: list | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        args = ["--status"]
    return install_autostart.main(args)


if __name__ == "__main__":
    sys.exit(main())
