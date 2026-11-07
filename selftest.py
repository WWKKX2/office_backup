#!/usr/bin/env python3
"""
selftest.py —— 端到端自检

不需要安装 Office，就能验证整条链路是否正常：
  1. 句柄扫描能不能真的枚举出被进程打开的文件
  2. 备份判定（无变化不重复备份）
  3. 版本清理（只保留最近 N 个）
  4. 同名文件隔离（不同目录下的同名文件不互相覆盖）

用法：
    python selftest.py
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import office_handle_scanner as scanner
from backup_core import BackupManager, bucket_name
from console_util import configure_console

PASS, FAIL = "  [OK]  ", "  [FAIL]"
results = []


def check(name: str, ok: bool, detail: str = "") -> bool:
    results.append((name, ok))
    print(f"{PASS if ok else FAIL} {name}{('  ' + detail) if detail else ''}")
    return ok


# 子进程：打开若干文件并挂住句柄，模拟 Office 正打开着这些文档
HOLDER_CODE = """
import sys, time
handles = [open(p, "rb") for p in sys.argv[1:]]
print("READY", flush=True)
time.sleep(120)
"""


def test_handle_scan(tmp: Path) -> None:
    print("\n[1] 句柄扫描")
    docs_dir = tmp / "docs"
    docs_dir.mkdir(parents=True, exist_ok=True)
    docx = docs_dir / "季度报告.docx"
    xlsx = docs_dir / "台账.xlsx"
    # Office 打开文档时产生的属主文件，扩展名同样是 .docx，必须被排除
    owner_lock = docs_dir / "~$季度报告.docx"
    docx.write_bytes(b"fake docx payload" * 100)
    xlsx.write_bytes(b"fake xlsx payload" * 100)
    owner_lock.write_bytes(b"lock marker")

    proc = subprocess.Popen(
        [sys.executable, "-c", HOLDER_CODE, str(docx), str(xlsx), str(owner_lock)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    try:
        line = proc.stdout.readline().strip()
        if not check("子进程成功打开文件句柄", line == "READY", f"(got {line!r})"):
            return

        t0 = time.time()
        found = scanner.scan_office_files(exts={".docx", ".xlsx"}, pids={proc.pid})
        cost = time.time() - t0

        found_norm = {os.path.normcase(os.path.realpath(str(p))) for p in found}
        want = {os.path.normcase(os.path.realpath(str(docx))),
                os.path.normcase(os.path.realpath(str(xlsx)))}

        check("扫描到全部 2 个打开的文件", want.issubset(found_norm),
              f"(找到 {len(found)} 个，耗时 {cost:.2f}s)")
        check("~$ 属主文件被正确排除",
              os.path.normcase(os.path.realpath(str(owner_lock))) not in found_norm)
        for p in found:
            print(f"          -> {p}")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


def test_backup_and_prune(tmp: Path) -> None:
    print("\n[2] 备份与版本清理")
    src_dir = tmp / "src"
    src_dir.mkdir(parents=True, exist_ok=True)
    doc = src_dir / "笔记.docx"

    root = tmp / "backups"
    mgr = BackupManager(root=root, max_versions=3, min_interval=0)
    mgr.ensure_ready()

    doc.write_text("v1", encoding="utf-8")
    first = mgr.consider(doc)
    check("首次扫描产生备份", first is not None)

    again = mgr.consider(doc)
    check("内容未变时不重复备份", again is None)

    for i in range(2, 7):
        time.sleep(0.02)
        doc.write_text(f"v{i}", encoding="utf-8")
        mgr.consider(doc)

    bucket = root / bucket_name(doc)
    versions = sorted(bucket.glob("*.docx"))
    check("版本数被限制在 3 个以内", len(versions) == 3, f"(实际 {len(versions)} 个)")
    check("确实清理过旧版本", mgr.pruned >= 3, f"(清理 {mgr.pruned} 个)")
    kept = {p.read_text(encoding="utf-8") for p in versions}
    check("保留的是最新版本", "v6" in kept, f"(保留内容 {sorted(kept)})")


def test_same_name_isolation(tmp: Path) -> None:
    print("\n[3] 同名文件隔离")
    a_dir = tmp / "部门A"
    b_dir = tmp / "部门B"
    a_dir.mkdir(parents=True, exist_ok=True)
    b_dir.mkdir(parents=True, exist_ok=True)
    a_file = a_dir / "总结.docx"
    b_file = b_dir / "总结.docx"
    a_file.write_text("A 的内容", encoding="utf-8")
    b_file.write_text("B 的内容", encoding="utf-8")

    root = tmp / "backups2"
    mgr = BackupManager(root=root, max_versions=10, min_interval=0)
    mgr.ensure_ready()

    pa = mgr.consider(a_file)
    pb = mgr.consider(b_file)

    check("两个同名文件都完成备份", pa is not None and pb is not None)
    check("落到了不同的备份桶", bucket_name(a_file) != bucket_name(b_file),
          f"({bucket_name(a_file)} vs {bucket_name(b_file)})")
    if pa and pb:
        check("两份备份内容各自独立",
              pa.read_text(encoding="utf-8") == "A 的内容"
              and pb.read_text(encoding="utf-8") == "B 的内容")


def test_dry_run(tmp: Path) -> None:
    print("\n[4] 演练模式")
    f = tmp / "dry.docx"
    f.write_text("x", encoding="utf-8")
    root = tmp / "backups3"
    mgr = BackupManager(root=root, max_versions=5, min_interval=0, dry_run=True)
    mgr.ensure_ready()
    mgr.consider(f)
    produced = list(root.rglob("*.docx"))
    check("演练模式不落盘", len(produced) == 0, f"(产生了 {len(produced)} 个文件)")


def test_main_end_to_end(tmp: Path) -> None:
    """真正跑一遍 main.py --once，验证 CLI → 扫描 → 备份整条链路。"""
    print("\n[5] 主程序端到端（main.py --once）")
    docs = tmp / "e2e"
    docs.mkdir(parents=True, exist_ok=True)
    doc = docs / "流程测试.docx"
    doc.write_bytes(b"e2e payload " * 50)

    backup_root = tmp / "e2e_backups"
    cfg_path = tmp / "e2e_config.json"
    cfg = {
        "backup_root": str(backup_root),
        "interval_seconds": 1,
        "min_seconds_between_backups": 0,
        "max_versions_per_file": 3,
        "scan_sources": ["handle"],
        "office_apps": {"word": True, "excel": True, "powerpoint": True, "wps": True},
        "file_extensions": [".docx"],
        "log_level": "WARNING",
    }
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")

    proc = subprocess.Popen(
        [sys.executable, "-c", HOLDER_CODE, str(doc)],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
    )
    # 让测试用的 python 进程被当作 Office 进程识别
    scanner.OFFICE_EXE_NAMES.add("python.exe")
    try:
        proc.stdout.readline()
        import main as app
        code = app.main(["--once", "--config", str(cfg_path)])
    finally:
        scanner.OFFICE_EXE_NAMES.discard("python.exe")
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()

    check("--once 正常退出", code == 0, f"(退出码 {code})")
    produced = list(backup_root.rglob("*.docx"))
    # 机器上可能真的开着别的 WPS/Word 文档，会被一起备份；
    # 这里只关心本次测试文档的备份有没有落盘。
    expected = [p for p in produced if p.parent.name.startswith(doc.stem + "__")]
    check("备份文件已落盘", len(expected) == 1, f"(本测试文档 {len(expected)} 个，总计 {len(produced)} 个)")
    if expected:
        print(f"          -> {expected[0]}")


def main() -> int:
    configure_console()
    print("=" * 60)
    print("Office Auto Backup 自检")
    print("=" * 60)
    print(f"Python      : {sys.version.split()[0]}")
    print(f"管理员权限  : {'是' if scanner.is_elevated() else '否（部分文件可能扫不到）'}")
    ftidx = scanner._file_type_index()
    print(f"File 类型索引: {ftidx if ftidx is not None else '探测失败（退化为全类型扫描）'}")

    tmp = Path(tempfile.mkdtemp(prefix="office_backup_selftest_"))
    try:
        test_handle_scan(tmp)
        test_backup_and_prune(tmp)
        test_same_name_isolation(tmp)
        test_dry_run(tmp)
        test_main_end_to_end(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    total = len(results)
    passed = sum(1 for _, ok in results if ok)
    print("\n" + "-" * 60)
    print(f"结果：{passed}/{total} 项通过")
    if passed == total:
        print("全部通过，工具可以正常工作。")
        return 0
    print("存在失败项，请检查上面的输出。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
