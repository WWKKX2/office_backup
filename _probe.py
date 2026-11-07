#!/usr/bin/env python3
"""
_probe.py —— 性能诊断脚本（开发用）

扫描一次 Office 打开的文件，并逐条统计句柄解析耗时，用来定位
"某轮扫描特别慢"是哪个进程 / 哪个句柄造成的。

用法：
    python _probe.py
"""

from __future__ import annotations

import ctypes
import os
import sys
import time
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import office_handle_scanner as s
from console_util import configure_console

SLOW_THRESHOLD = 0.5  # 秒，超过这个值就单独打印


def main() -> int:
    configure_console()
    if not s.is_elevated():
        print("[警告] 未以管理员身份运行，其他用户进程的文件可能看不到\n")

    procs = s.scan_office_processes()
    print(f"Office / WPS 进程：{procs or '无'}")
    if not procs:
        return 0

    ftidx = s._file_type_index()
    print(f"File 对象类型索引：{ftidx}")

    t0 = time.time()
    _buf, count, entries = s._fetch_handle_table()
    print(f"句柄表共 {count} 条，读取耗时 {time.time() - t0:.2f}s")

    # 只挑出目标进程的 File 句柄
    handles: dict = defaultdict(list)
    for i in range(count):
        e = entries[i]
        pid = int(e.UniqueProcessId or 0)
        if pid in procs and (ftidx is None or int(e.ObjectTypeIndex or 0) == ftidx):
            handles[pid].append(int(e.HandleValue or 0))

    total = sum(len(v) for v in handles.values())
    print(f"待解析的文件句柄：{total} 个\n")

    k32 = s.kernel32
    current = ctypes.c_void_p(int(k32.GetCurrentProcess()))
    resolved = 0
    matched = 0
    slow = []

    for pid, hvals in handles.items():
        hproc = k32.OpenProcess(
            s.PROCESS_QUERY_LIMITED_INFORMATION | s.PROCESS_DUP_HANDLE, False, pid
        )
        if not hproc:
            print(f"  无法打开进程 {pid}（权限不足？）")
            continue
        try:
            for hv in hvals:
                dup = ctypes.c_void_p()
                ok = k32.DuplicateHandle(
                    ctypes.c_void_p(int(hproc)), ctypes.c_void_p(hv), current,
                    ctypes.byref(dup), 0, False, s.DUPLICATE_SAME_ACCESS,
                )
                if not ok:
                    continue
                try:
                    t1 = time.time()
                    path = s._final_path(dup.value)
                    dt = time.time() - t1
                finally:
                    k32.CloseHandle(ctypes.c_void_p(dup.value))

                resolved += 1
                if dt > SLOW_THRESHOLD:
                    slow.append((dt, pid, hv, path))
                if path and os.path.splitext(path)[1].lower() in s.DEFAULT_OFFICE_EXTS:
                    matched += 1
                    print(f"  命中 {path}")
        finally:
            k32.CloseHandle(ctypes.c_void_p(int(hproc)))

    print(f"\n解析 {resolved} 个句柄，命中 {matched} 个 Office 文档")
    if slow:
        print(f"\n慢句柄（>{SLOW_THRESHOLD}s）：")
        for dt, pid, hv, path in sorted(slow, reverse=True)[:10]:
            print(f"  {dt:.2f}s  pid={pid} handle={hv}  {path}")
    else:
        print("没有明显慢的句柄。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
