#!/usr/bin/env python3
"""
office_handle_scanner.py

零依赖扫描器（纯 ctypes，不需要 pywin32 / COM），枚举当前在 Microsoft Office
或 WPS Office 中打开着的文档文件。

工作原理
--------
1. 用 Toolhelp 快照列出进程，筛出 Office / WPS 的可执行文件。
2. 通过 NtQuerySystemInformation(64) 拉取全系统句柄表。
3. 只保留这些进程拥有的、对象类型为 File 的句柄。
4. 把句柄 DuplicateHandle 到本进程，用 GetFinalPathNameByHandleW 反解真实路径。
5. 按扩展名过滤并去重。

建议以管理员权限运行：否则无法打开其他用户/高权限进程的句柄，部分文件会漏掉。

作为库使用：
    import office_handle_scanner as scanner
    files = scanner.scan_office_files()             # 自动发现 Office 进程
    files = scanner.scan_office_files(pids={1234})  # 指定进程（测试用）
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import logging
import os
import tempfile
import threading
import time
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Set

log = logging.getLogger(__name__)

try:
    ntdll = ctypes.WinDLL("ntdll")
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
except OSError:  # 非 Windows 平台
    ntdll = None
    kernel32 = None

# --- 常量 -------------------------------------------------------------------
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_DUP_HANDLE = 0x0040
DUPLICATE_SAME_ACCESS = 0x0002

STATUS_SUCCESS = 0x00000000
STATUS_INFO_LENGTH_MISMATCH = 0xC0000004
STATUS_BUFFER_OVERFLOW = 0x80000005
STATUS_BUFFER_TOO_SMALL = 0xC0000023

SYSTEM_EXTENDED_HANDLE_INFORMATION = 64
TH32CS_SNAPPROCESS = 0x00000002

# 初始缓冲区 16MB，通常一次就够；不够会自动翻倍重试
_INITIAL_BUFFER = 16 * 1024 * 1024

# GetFileType 的返回值：只有磁盘文件才能安全地反解路径。
# 管道 / 通信设备句柄在 GetFinalPathNameByHandleW 上可能永久阻塞。
FILE_TYPE_DISK = 0x0001

# 单个进程句柄解析的等待时间，以及整轮解析的总预算（秒）。
# 正常进程只需要毫秒级；超时通常意味着遇到了会阻塞的内核句柄。
_PID_TIMEOUT = 3.0
_TOTAL_TIMEOUT = 8.0
# 某个进程解析超时后，这段时间内不再扫描它，避免每轮都卡住。
_BLOCKED_TTL = 300.0
_BLOCKED_PIDS: Dict[int, float] = {}
_BLOCKED_LOCK = threading.Lock()

# Office / WPS 常见文档扩展名（小写，含点）
DEFAULT_OFFICE_EXTS: Set[str] = {
    ".doc", ".docx", ".docm", ".dot", ".dotx",
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx",
    ".ppt", ".pptx", ".pptm", ".pot", ".potx",
    ".wps", ".wpt", ".et", ".ett", ".dps", ".dpt",
    ".csv", ".rtf", ".odt", ".ods", ".odp",
}

# Office / WPS 的文档进程名（小写），按 config.json 里的 office_apps 开关分组
APP_EXE_NAMES: Dict[str, Set[str]] = {
    "word": {"winword.exe"},                                  # Word
    "excel": {"excel.exe"},                                   # Excel
    "powerpoint": {"powerpnt.exe"},                           # PowerPoint
    "wps": {"wps.exe", "et.exe", "wpp.exe", "wpsoffice.exe"}, # WPS 文字 / 表格 / 演示 / 整合版
}

# 所有受支持的进程名
OFFICE_EXE_NAMES: Set[str] = set().union(*APP_EXE_NAMES.values())
# 配置里能识别的全部进程名（不会随测试临时增删而改变）
ALL_OFFICE_EXES: Set[str] = set(OFFICE_EXE_NAMES)


def enabled_exe_names(apps_config: Dict[str, bool] | None = None) -> Set[str]:
    """根据 office_apps 开关，返回启用应用对应的进程名集合。"""
    cfg = apps_config or {}
    names: Set[str] = set()
    for app, exes in APP_EXE_NAMES.items():
        if cfg.get(app, True):
            names |= exes
    return names

# Office 打开文档时会在同目录创建一个 ~$xxx.docx 属主文件并保持打开。
# 它的扩展名和正主一样，但只是个锁标记，必须排除，否则备份目录会被垃圾文件塞满。
IGNORED_NAME_PREFIXES = ("~$",)


def is_ignored_name(filename: str) -> bool:
    """是否为 Office 的临时/锁文件（~$ 开头）。"""
    return filename.startswith(IGNORED_NAME_PREFIXES)


# --- 结构体 -----------------------------------------------------------------


class SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX(ctypes.Structure):
    """winternl.h 中 SystemExtendedHandleInformation 的条目结构（64 位）。"""

    _fields_ = [
        ("Object", ctypes.c_void_p),
        ("UniqueProcessId", ctypes.c_void_p),
        ("HandleValue", ctypes.c_void_p),
        ("GrantedAccess", wt.ULONG),
        ("CreatorBackTraceIndex", wt.USHORT),
        ("ObjectTypeIndex", wt.USHORT),
        ("HandleAttributes", wt.ULONG),
        ("Reserved", wt.ULONG),
    ]


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", wt.DWORD),
        ("cntUsage", wt.DWORD),
        ("th32ProcessID", wt.DWORD),
        ("th32DefaultHeapID", ctypes.c_void_p),
        ("th32ModuleID", wt.DWORD),
        ("cntThreads", wt.DWORD),
        ("th32ParentProcessID", wt.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wt.DWORD),
        ("szExeFile", ctypes.c_wchar * 260),
    ]


if ntdll is not None:
    ntdll.NtQuerySystemInformation.restype = ctypes.c_long
    ntdll.NtQuerySystemInformation.argtypes = [
        wt.ULONG, ctypes.c_void_p, wt.ULONG, ctypes.POINTER(wt.ULONG)
    ]


def _require_windows() -> None:
    if kernel32 is None:
        raise RuntimeError("office_handle_scanner 只能在 Windows 上运行")


# --- 句柄表 -----------------------------------------------------------------


def _query_buffer_size(infoclass: int) -> int:
    """先空查询一次，拿到建议缓冲区大小。"""
    needed = wt.ULONG(0)
    status = ntdll.NtQuerySystemInformation(infoclass, None, 0, ctypes.byref(needed))
    if status in (STATUS_INFO_LENGTH_MISMATCH, STATUS_BUFFER_OVERFLOW, STATUS_BUFFER_TOO_SMALL):
        return max(_INITIAL_BUFFER, int(needed.value) + 65536)
    return _INITIAL_BUFFER


def _fetch_handle_table():
    """
    拉取全系统句柄表。

    返回 (buffer, count, entries_ptr)。buffer 必须保持引用，否则 entries_ptr 失效。
    """
    _require_windows()
    needed = wt.ULONG(0)
    size = _query_buffer_size(SYSTEM_EXTENDED_HANDLE_INFORMATION)
    status = STATUS_INFO_LENGTH_MISMATCH

    for _ in range(8):
        buf = ctypes.create_string_buffer(size)
        status = ntdll.NtQuerySystemInformation(
            SYSTEM_EXTENDED_HANDLE_INFORMATION, buf, size, ctypes.byref(needed)
        )
        if status == STATUS_SUCCESS:
            count = ctypes.cast(buf, ctypes.POINTER(wt.ULONG)).contents.value
            # SYSTEM_HANDLE_INFORMATION_EX: { ULONG_PTR NumberOfHandles; ULONG_PTR Reserved; ... }
            entries = ctypes.cast(
                ctypes.addressof(buf) + 2 * ctypes.sizeof(ctypes.c_void_p),
                ctypes.POINTER(SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX),
            )
            return buf, count, entries
        size = max(size * 2, int(needed.value) + 65536)

    raise RuntimeError(
        f"NtQuerySystemInformation(64) 失败，最后状态码 0x{status & 0xFFFFFFFF:08X}"
    )


def _final_path(handle: Optional[int]) -> Optional[str]:
    r"""GetFinalPathNameByHandleW，并去掉 \\?\ 前缀。"""
    if not handle:
        return None
    kernel32.GetFinalPathNameByHandleW.restype = wt.DWORD
    kernel32.GetFinalPathNameByHandleW.argtypes = [wt.HANDLE, wt.LPWSTR, wt.DWORD, wt.DWORD]
    buf = ctypes.create_unicode_buffer(32768)
    n = kernel32.GetFinalPathNameByHandleW(wt.HANDLE(handle), buf, 32768, 0)
    if n == 0 or n >= 32768:
        return None
    s = buf.value
    if s.startswith("\\\\?\\UNC\\"):
        return "\\\\" + s[8:]
    if s.startswith("\\\\?\\"):
        return s[4:]
    return s


def _current_process_handle() -> int:
    """当前进程的伪句柄（-1），用于 DuplicateHandle 的目标进程。"""
    try:
        kernel32.GetCurrentProcess.restype = wt.HANDLE
        value = kernel32.GetCurrentProcess()
        return int(value or 0)
    except (OSError, ValueError, TypeError):
        return -1


def _file_type_of(handle: Optional[int]) -> int:
    """返回句柄类型（FILE_TYPE_DISK / PIPE / ...）；失败返回 0。"""
    if not handle:
        return 0
    kernel32.GetFileType.restype = wt.DWORD
    kernel32.GetFileType.argtypes = [wt.HANDLE]
    try:
        return int(kernel32.GetFileType(wt.HANDLE(handle)))
    except OSError:
        return 0


def _is_pid_blocked(pid: int) -> bool:
    now = time.monotonic()
    with _BLOCKED_LOCK:
        expiry = _BLOCKED_PIDS.get(pid)
        if expiry is None:
            return False
        if expiry <= now:
            del _BLOCKED_PIDS[pid]
            return False
        return True


def _mark_pid_blocked(pid: int) -> None:
    with _BLOCKED_LOCK:
        _BLOCKED_PIDS[pid] = time.monotonic() + _BLOCKED_TTL


_FILE_TYPE_INDEX_CACHE: Optional[int] = None
_FILE_TYPE_INDEX_DONE = False


def _probe_file_type_index() -> Optional[int]:
    """
    探测 "File" 对象的类型索引（Win10/11 上通常是 28 或 37，不同版本会变）。

    做法：自己开一个临时文件句柄，在句柄表里找到它，读回 ObjectTypeIndex。
    只探测一次并缓存；失败也不会反复重试。
    """
    global _FILE_TYPE_INDEX_CACHE, _FILE_TYPE_INDEX_DONE
    if _FILE_TYPE_INDEX_DONE:
        return _FILE_TYPE_INDEX_CACHE

    _FILE_TYPE_INDEX_DONE = True
    _FILE_TYPE_INDEX_CACHE = None

    path = os.path.join(tempfile.gettempdir(), f"_ftidx_probe_{os.getpid()}.tmp")
    handle = None
    try:
        with open(path, "w") as f:
            f.write("x")

        kernel32.CreateFileW.restype = wt.HANDLE
        kernel32.CreateFileW.argtypes = [
            wt.LPCWSTR, wt.DWORD, wt.DWORD, ctypes.c_void_p, wt.DWORD, wt.DWORD, wt.HANDLE,
        ]
        # GENERIC_READ, FILE_SHARE_READ|WRITE, OPEN_EXISTING
        handle = kernel32.CreateFileW(
            path, 0x80000000, 0x00000001 | 0x00000002, None, 3, 0, None
        )
        hval = int(handle or 0)
        if not hval:
            return None

        _buf, count, entries = _fetch_handle_table()
        mypid = os.getpid()
        for i in range(count):
            e = entries[i]
            if int(e.UniqueProcessId or 0) == mypid and int(e.HandleValue or 0) == hval:
                _FILE_TYPE_INDEX_CACHE = int(e.ObjectTypeIndex or 0)
                break
        return _FILE_TYPE_INDEX_CACHE
    except Exception:
        return None
    finally:
        if handle:
            kernel32.CloseHandle(handle)
        try:
            os.remove(path)
        except OSError:
            pass


def _file_type_index() -> Optional[int]:
    """File 对象类型索引；返回 None 表示探测失败（此时不做类型过滤）。"""
    return _probe_file_type_index()


# --- 进程发现 ---------------------------------------------------------------


def _proc_list() -> Dict[int, str]:
    """返回 {pid: exe 文件名小写}。"""
    _require_windows()
    kernel32.CreateToolhelp32Snapshot.restype = wt.HANDLE
    kernel32.CreateToolhelp32Snapshot.argtypes = [wt.ULONG, wt.DWORD]
    kernel32.Process32FirstW.restype = wt.BOOL
    kernel32.Process32NextW.restype = wt.BOOL

    out: Dict[int, str] = {}
    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snap or snap == -1:
        return out
    try:
        pe = PROCESSENTRY32W()
        pe.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        if kernel32.Process32FirstW(snap, ctypes.byref(pe)):
            while True:
                out[int(pe.th32ProcessID)] = pe.szExeFile.lower()
                if not kernel32.Process32NextW(snap, ctypes.byref(pe)):
                    break
    finally:
        kernel32.CloseHandle(snap)
    return out


def is_office_family(exe_name: str) -> bool:
    return exe_name.lower() in OFFICE_EXE_NAMES


def scan_office_processes() -> Dict[int, str]:
    """返回 {pid: exe 名}，仅包含 Office / WPS 文档进程。"""
    procs = _proc_list()
    return {pid: name for pid, name in procs.items() if is_office_family(name)}


# --- 打开文件枚举 -----------------------------------------------------------


def _resolve_pid_handles(
    pid: int,
    hvals: List[int],
    current_proc: int,
    exts: Set[str],
    out: List[str],
) -> None:
    """
    在独立线程里解析一个进程的句柄。

    必须放在线程里：GetFinalPathNameByHandleW 遇到某些管道 / 设备句柄会
    永久阻塞，主线程等不起。线程是 daemon，超时就丢下它继续跑。
    """
    hproc = kernel32.OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_DUP_HANDLE, False, pid
    )
    if not hproc:
        return
    try:
        for hv in hvals:
            dup = wt.HANDLE()
            if not kernel32.DuplicateHandle(
                wt.HANDLE(int(hproc)), wt.HANDLE(hv), wt.HANDLE(current_proc),
                ctypes.byref(dup), 0, False, DUPLICATE_SAME_ACCESS
            ):
                continue
            try:
                # 只处理磁盘文件；管道等句柄反解路径时可能卡死。
                if _file_type_of(dup.value) != FILE_TYPE_DISK:
                    continue
                path = _final_path(dup.value)
            finally:
                kernel32.CloseHandle(dup.value)
            if path and os.path.splitext(path)[1].lower() in exts:
                if is_ignored_name(os.path.basename(path)):
                    continue
                out.append(path)
    except Exception:  # noqa: BLE001 - 单个进程失败不影响其它进程
        log.debug("解析进程 %s 的句柄失败", pid, exc_info=True)
    finally:
        kernel32.CloseHandle(hproc)


def enumerate_open_files(
    pids: Set[int],
    exts: Set[str] = DEFAULT_OFFICE_EXTS,
    file_type_index: Optional[int] = None,
) -> Dict[int, List[str]]:
    """
    对 `pids` 中每个进程，枚举它打开着的、扩展名命中 `exts` 的文件。

    返回 {pid: [文件路径, ...]}。整表扫描需要几秒，每轮只调用一次。
    """
    if not pids:
        return {}

    _require_windows()
    _buf, count, entries = _fetch_handle_table()

    handles: Dict[int, List[int]] = defaultdict(list)
    for i in range(count):
        e = entries[i]
        pid = int(e.UniqueProcessId or 0)
        if pid not in pids:
            continue
        if file_type_index is not None and int(e.ObjectTypeIndex or 0) != file_type_index:
            continue
        handles[pid].append(int(e.HandleValue or 0))

    current_proc = _current_process_handle()
    result: Dict[int, List[str]] = {}
    threads: List[tuple] = []

    for pid, hvals in handles.items():
        if _is_pid_blocked(pid):
            log.debug("进程 %s 之前解析超时，冷却期内跳过", pid)
            continue
        bucket: List[str] = []
        result[pid] = bucket
        thread = threading.Thread(
            target=_resolve_pid_handles,
            args=(pid, hvals, current_proc, exts, bucket),
            name=f"handle-scan-{pid}",
            daemon=True,
        )
        thread.start()
        threads.append((pid, thread))

    deadline = time.monotonic() + _TOTAL_TIMEOUT
    for pid, thread in threads:
        remaining = min(_PID_TIMEOUT, max(0.0, deadline - time.monotonic()))
        thread.join(remaining)

    for pid, thread in threads:
        if thread.is_alive():
            _mark_pid_blocked(pid)
            log.warning(
                "进程 %s 的句柄解析超过 %.0f 秒，本轮跳过；%.0f 秒内不再尝试",
                pid, _PID_TIMEOUT, _BLOCKED_TTL,
            )

    return {pid: list(paths) for pid, paths in result.items() if paths}


def scan_office_files(
    exts: Set[str] = DEFAULT_OFFICE_EXTS,
    pids: Optional[Set[int]] = None,
) -> List[Path]:
    """
    枚举 Office / WPS 当前打开着的文档。

    参数
    ----
    exts : 允许的扩展名集合（小写，含点）
    pids : 指定要扫描的进程；None 表示自动发现 Office / WPS 进程

    返回去重、排序后的绝对路径列表。
    """
    if pids is None:
        pids = set(scan_office_processes())
    if not pids:
        return []

    exts = {e.lower() for e in exts}
    found = enumerate_open_files(pids, exts, file_type_index=_file_type_index())

    paths: Set[str] = set()
    for pid_paths in found.values():
        paths.update(pid_paths)

    # 过滤掉已经不存在的路径（进程刚关闭文件等情况）
    return sorted((Path(p) for p in paths if os.path.isfile(p)), key=lambda p: str(p).lower())


# --- 权限检查 ---------------------------------------------------------------


def is_elevated() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


# --- 命令行自测 -------------------------------------------------------------

if __name__ == "__main__":
    from console_util import configure_console

    configure_console()
    if not is_elevated():
        print("[警告] 未以管理员身份运行，其他用户进程的文件可能看不到")

    procs = scan_office_processes()
    if not procs:
        print("没有检测到正在运行的 Office / WPS 进程。")
    else:
        print(f"检测到 {len(procs)} 个 Office / WPS 进程：")
        for pid, name in procs.items():
            print(f"  {pid:>6}  {name}")

    files = scan_office_files()
    print(f"\n共找到 {len(files)} 个打开着的文档：")
    for f in files:
        print("  ", f)
