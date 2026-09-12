"""Inventory and replay historical Newdata workbooks into the version store."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime
from pathlib import Path

import database


EXPORT_TIMESTAMP = re.compile(r"(?<!\d)(20\d{15})(?!\d)")
FULL_DATE = re.compile(r"(20\d{2})[.\-](\d{1,2})[.\-](\d{1,2})")
SHORT_RANGE_END = re.compile(r"-(\d{1,2})[.\-](\d{1,2})(?:\D|$)")
COMPACT_RANGE = re.compile(r"(20\d{2})(\d{2})(\d{2})-(\d{2})(\d{2})")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def sort_key(path: Path):
    match = EXPORT_TIMESTAMP.search(path.stem)
    if match:
        try:
            return (0, datetime.strptime(match.group(1), "%Y%m%d%H%M%S%f"), path.name)
        except ValueError:
            pass
    full_dates = FULL_DATE.findall(path.stem)
    try:
        if len(full_dates) >= 2:
            year, month, day = full_dates[-1]
            return (0, datetime(int(year), int(month), int(day), 23, 59, 59), path.name)
        if len(full_dates) == 1:
            year = int(full_dates[0][0])
            suffix = SHORT_RANGE_END.search(path.stem)
            if suffix:
                return (
                    0,
                    datetime(year, int(suffix.group(1)), int(suffix.group(2)), 23, 59, 59),
                    path.name,
                )
        compact = COMPACT_RANGE.search(path.stem)
        if compact:
            return (
                0,
                datetime(int(compact.group(1)), int(compact.group(4)), int(compact.group(5)), 23, 59, 59),
                path.name,
            )
    except ValueError:
        pass
    return (1, datetime.max, path.name)


def inventory(paths: list[Path]) -> list[dict]:
    return [
        {
            "path": str(path),
            "filename": path.name,
            "size": path.stat().st_size,
            "sha256": sha256(path),
            "sortable": sort_key(path)[0] == 0,
        }
        for path in paths
    ]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--stats-base", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--apply", action="store_true", help="Stage and activate sortable valid batches")
    parser.add_argument("--confirmed-by", default="historical-migration")
    args = parser.parse_args()

    stats_base = args.stats_base.resolve()
    source_root = stats_base / "input" / "Newdata"
    output_root = stats_base / "output"
    baseline_candidates = [
        stats_base / "input" / "Annual Summary" / "2025" / "2025年停电用户数据.xlsx",
        stats_base / "2025年停电用户数据.xlsx",
    ]
    baseline = next((path for path in baseline_candidates if path.is_file()), None)
    daily_root = source_root / "daily"
    scan_root = daily_root if daily_root.is_dir() else source_root
    newdata_inputs = [
        p for p in scan_root.glob("*.xlsx") if not p.name.startswith("~$")
    ]
    inputs = ([baseline] if baseline else []) + sorted(newdata_inputs, key=sort_key)
    snapshots = sorted(
        [
            p for p in output_root.rglob("*处理结果.xlsx")
            if "_发出版" not in p.name and not p.name.startswith("~$")
        ],
        key=lambda p: str(p),
    )
    report = {
        "generated_at": datetime.now().isoformat(),
        "inputs": inventory(inputs),
        "snapshots": inventory(snapshots),
        "baseline_2025": str(baseline) if baseline else None,
        "ambiguous_inputs": [
            str(path) for path in newdata_inputs if sort_key(path)[0] != 0
        ],
        "batches": [],
    }

    if args.apply:
        database.ensure_schema()
        for path in inputs:
            if path != baseline and sort_key(path)[0] != 0:
                continue
            staged = database.stage_import(path)
            entry = {"path": str(path), "stage": staged}
            if staged.get("status") == "pending_confirmation":
                preview = database.preview_import(staged["batch_id"])
                entry["preview"] = preview
                entry["activation"] = database.activate_import(
                    staged["batch_id"],
                    datetime.fromisoformat(preview["replace_start_date"]).date(),
                    datetime.fromisoformat(preview["replace_end_date"]).date(),
                    args.confirmed_by,
                )
            report["batches"].append(entry)

    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps({"inputs": len(inputs), "snapshots": len(snapshots), "ambiguous": len(report["ambiguous_inputs"])}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
