#!/usr/bin/env python3
"""
com_source.py

可选的 COM 扫描源（需要 pywin32）。

句柄扫描（office_handle_scanner）是主引擎，零依赖、覆盖全。COM 只作为兜底：
个别环境下句柄表拿不到时，COM 还能救一下。

没装 pywin32 时 available() 返回 False，主程序会静默跳过，不影响运行。
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Dict, List

log = logging.getLogger(__name__)

# (ProgID, 集合属性名, 对应 config.office_apps 的开关)
_APPS = [
    ("Word.Application", "Documents", "word"),
    ("Excel.Application", "Workbooks", "excel"),
    ("PowerPoint.Application", "Presentations", "powerpoint"),
    ("Kwps.Application", "Documents", "wps"),
    ("Ket.Application", "Workbooks", "wps"),
    ("Kpp.Application", "Presentations", "wps"),
]


def available() -> bool:
    """pywin32 是否可用。"""
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def get_open_files(apps_config: Dict[str, bool] | None = None) -> List[Path]:
    """
    通过 COM 读取各 Office / WPS 应用当前打开的文档路径。

    任何单个应用出错都只记 debug 日志，不影响其他应用。
    """
    if not available():
        log.debug("未安装 pywin32，跳过 COM 扫描源")
        return []

    import pythoncom
    import win32com.client

    apps_config = apps_config or {}
    files: List[Path] = []

    pythoncom.CoInitialize()
    try:
        for progid, collection, toggle in _APPS:
            if not apps_config.get(toggle, True):
                continue
            try:
                app = win32com.client.GetActiveObject(progid)
            except Exception as exc:
                log.debug("%s 未运行或无法连接：%s", progid, exc)
                continue

            try:
                items = getattr(app, collection)
            except Exception as exc:
                log.debug("%s 无法读取 %s：%s", progid, collection, exc)
                continue

            try:
                for item in items:
                    try:
                        raw = item.FullName
                    except Exception:
                        continue
                    if raw and os.path.isfile(raw):
                        files.append(Path(raw))
            except Exception as exc:
                log.debug("%s 遍历失败：%s", progid, exc)
    finally:
        # 无论成败都释放 COM，避免累积
        pythoncom.CoUninitialize()

    return files
