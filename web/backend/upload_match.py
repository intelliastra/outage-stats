"""Content identity check for the two separate daily/external uploads."""

from __future__ import annotations

import hashlib
from pathlib import Path


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for part in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(part)
    return digest.hexdigest()


def latest_matching_daily_file(external: Path, daily_directory: Path) -> Path:
    files = [
        path for path in daily_directory.iterdir()
        if path.is_file() and path.suffix.lower() in {".xlsx", ".xls"}
        and not path.name.startswith("~$")
    ] if daily_directory.is_dir() else []
    if not files:
        raise ValueError("请先在日常统计页上传同一份未经手工删行的停电用户表")
    latest = max(files, key=lambda path: (path.stat().st_mtime_ns, path.name))
    if file_sha256(external) != file_sha256(latest):
        raise ValueError("外力页文件与日常统计最新原表内容不一致，请在两页上传同一文件")
    return latest
