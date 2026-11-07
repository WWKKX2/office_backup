#!/usr/bin/env python3
"""
backup_core.py

备份存储层：负责"要不要备份"和"怎么存"。

设计要点
--------
* 按源文件所在目录分桶（bucket），不同目录下的同名文件互不干扰。
* 用 (mtime_ns, size) 判断内容是否变化，避免重复备份。
* 冷却期（min_interval）：防止自动保存把版本数刷爆。
* 复制失败自动重试，应对 Office 保存瞬间的文件占用。
* 每个文件只保留最近 N 个版本。
"""

from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

log = logging.getLogger(__name__)

# Windows 文件名非法字符
_UNSAFE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# 备份文件名：<名称>_YYYYMMDD_HHMMSS[_序号]<扩展名>
_VERSION_NAME = re.compile(r"^(.+)_(\d{8})_(\d{6})(?:_(\d+))?$")


def safe_component(name: str, limit: int = 80) -> str:
    """把任意字符串转成合法的文件/目录名片段。"""
    cleaned = _UNSAFE_CHARS.sub("_", name).strip(" .")
    return (cleaned or "unnamed")[:limit]


def bucket_name(src: Path) -> str:
    """
    为源文件生成稳定的备份桶名：`<文件名主干>__<父目录哈希8位>`。

    带父目录哈希是为了让不同目录下的同名文件落到不同桶里，
    同时目录名保持人类可读。
    """
    parent = os.path.normcase(os.path.realpath(str(src.parent)))
    digest = hashlib.sha1(parent.encode("utf-8", "surrogatepass")).hexdigest()[:8]
    return f"{safe_component(src.stem)}__{digest}"


def _state_key(path: Path) -> str:
    """
    变化检测用的状态键。

    统一成 realpath + normcase，避免同一文件在不同轮次以短路径（8.3）和
    长路径两种形式出现，导致被误判为"新文件"而重复备份。
    """
    try:
        return os.path.normcase(os.path.realpath(str(path)))
    except OSError:
        return os.path.normcase(str(path))


def _version_pattern(stem: str, suffix: str) -> "re.Pattern":
    """匹配某个源文件的备份版本文件名。"""
    return re.compile(
        rf"^{re.escape(stem)}_(\d{{8}})_(\d{{6}})(?:_(\d+))?{re.escape(suffix)}$"
    )


def _version_key(pattern: "re.Pattern", name: str) -> tuple:
    m = pattern.match(name)
    if not m:
        return ("00000000", "000000", 0)
    return (m.group(1), m.group(2), int(m.group(3) or 0))


def version_datetime(path: Path) -> Optional[datetime]:
    """从备份文件名里解析备份时间；解析失败返回 None。"""
    m = _VERSION_NAME.match(path.stem)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(2) + m.group(3), "%Y%m%d%H%M%S")
    except ValueError:
        return None


def _copy_with_retry(src: Path, dst: Path, attempts: int = 5, delay: float = 0.4) -> None:
    """复制文件；遇到占用/权限问题时短暂等待后重试。"""
    last_error: Optional[Exception] = None
    for attempt in range(1, attempts + 1):
        try:
            shutil.copy2(src, dst)
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            if dst.exists():
                try:
                    dst.unlink()
                except OSError:
                    pass
            if attempt < attempts:
                log.debug("复制失败（第 %d/%d 次），%.1fs 后重试：%s", attempt, attempts, delay, exc)
                time.sleep(delay)
    raise last_error if last_error else RuntimeError("未知复制错误")


class BackupManager:
    """管理备份的判定、写入与版本清理。"""

    def __init__(
        self,
        root: Path | str,
        max_versions: int = 20,
        min_interval: float = 30.0,
        dry_run: bool = False,
    ) -> None:
        self.root = Path(root)
        self.max_versions = max(1, int(max_versions))
        self.min_interval = max(0.0, float(min_interval))
        self.dry_run = dry_run

        self._signature: Dict[str, Tuple[int, int]] = {}
        self._last_backup: Dict[str, float] = {}

        self.backed_up = 0
        self.failed = 0
        self.pruned = 0

    # -- 初始化 --------------------------------------------------------------

    def ensure_ready(self) -> None:
        """确保备份根目录可用，失败时抛出带说明的异常。"""
        try:
            self.root.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise RuntimeError(
                f"无法创建备份目录 {self.root}：{exc}\n"
                f"请检查盘符是否存在、是否有写入权限，或修改 config.json 中的 backup_root。"
            ) from exc
        probe = self.root / ".write_probe"
        try:
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
        except OSError as exc:
            raise RuntimeError(f"备份目录 {self.root} 不可写：{exc}") from exc

    # -- 主入口 --------------------------------------------------------------

    def consider(self, src: Path) -> Optional[Path]:
        """
        判断 src 是否需要备份，需要则执行备份。

        返回备份文件路径；未备份（无变化 / 冷却中 / 失败）返回 None。
        """
        key = _state_key(src)

        try:
            st = src.stat()
        except OSError as exc:
            log.debug("无法读取 %s：%s", src, exc)
            return None

        signature = (st.st_mtime_ns, st.st_size)
        if self._signature.get(key) == signature:
            return None  # 内容未变

        now = time.time()
        if now - self._last_backup.get(key, 0.0) < self.min_interval:
            log.debug("冷却中，跳过 %s", src.name)
            return None

        dst = self._write(src, now)
        if dst is not None:
            # 只有成功才更新签名，失败的话下一轮会重试
            self._signature[key] = signature
            self._last_backup[key] = now
        return dst

    # -- 内部实现 ------------------------------------------------------------

    def _write(self, src: Path, now: float) -> Optional[Path]:
        bucket = self.root / bucket_name(src)
        stem = safe_component(src.stem)
        suffix = src.suffix

        try:
            bucket.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            self.failed += 1
            log.error("无法创建备份子目录 %s：%s", bucket, exc)
            return None

        stamp = datetime.fromtimestamp(now).strftime("%Y%m%d_%H%M%S")
        # 同一秒内连续备份时递增序号。这里不能只看 `目标文件名是否存在`：
        # 旧版本被清理后文件名会空出来，重新使用会让序号倒退，按时间清理
        # 时就会误删最新版本。所以取同一秒已有文件的最大序号 + 1。
        same_stamp = re.compile(
            rf"^{re.escape(stem)}_{stamp}(?:_(\d+))?{re.escape(suffix)}$"
        )
        max_serial = -1
        try:
            for existing in bucket.iterdir():
                m = same_stamp.match(existing.name)
                if m:
                    max_serial = max(max_serial, int(m.group(1) or 0))
        except OSError:
            pass

        if max_serial < 0:
            target = bucket / f"{stem}_{stamp}{suffix}"
        else:
            target = bucket / f"{stem}_{stamp}_{max_serial + 1}{suffix}"
        while target.exists():
            max_serial += 1
            target = bucket / f"{stem}_{stamp}_{max_serial + 1}{suffix}"

        if self.dry_run:
            self.backed_up += 1
            log.info("[演练] 将备份 %s -> %s", src, target.relative_to(self.root))
            return target

        try:
            _copy_with_retry(src, target)
        except Exception as exc:
            self.failed += 1
            log.error("备份失败 %s：%s", src, exc)
            return None

        self.backed_up += 1
        log.info("已备份 %s -> %s", src.name, target.relative_to(self.root))
        self._prune(bucket, stem, suffix)
        return target

    def _prune(self, bucket: Path, stem: str, suffix: str) -> None:
        """只保留该文件最新的 max_versions 个版本。"""
        # 只处理本工具生成的 `名称_YYYYMMDD_HHMMSS[_序号].扩展名`，
        # 目录里用户自己放的其他文件不碰。
        versions = self._list_versions(bucket, stem, suffix)
        for old in versions[self.max_versions:]:
            try:
                old.unlink()
                self.pruned += 1
                log.debug("清理旧版本 %s", old.name)
            except OSError as exc:
                log.warning("删除旧版本失败 %s：%s", old, exc)

    # -- 查询 ----------------------------------------------------------------

    def versions(self, src: Path) -> List[Path]:
        """返回某个源文件在备份目录里的版本，最新的在前。"""
        return self._list_versions(
            self.root / bucket_name(src), safe_component(src.stem), src.suffix
        )

    def _list_versions(self, bucket: Path, stem: str, suffix: str) -> List[Path]:
        pattern = _version_pattern(stem, suffix)
        if not bucket.is_dir():
            return []
        try:
            found = [
                p for p in bucket.iterdir()
                if p.is_file() and pattern.match(p.name)
            ]
        except OSError as exc:
            log.debug("读取备份目录失败 %s：%s", bucket, exc)
            return []
        found.sort(key=lambda p: _version_key(pattern, p.name), reverse=True)
        return found

    # -- 统计 ----------------------------------------------------------------

    def summary(self) -> str:
        return (
            f"本次共备份 {self.backed_up} 个版本，清理 {self.pruned} 个旧版本"
            + (f"，失败 {self.failed} 个" if self.failed else "")
        )
