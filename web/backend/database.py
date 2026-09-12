"""PostgreSQL batch/version storage for outage data.

The module is deliberately optional while DATA_BACKEND=excel.  In shadow mode
uploads are validated and staged, but only an explicit activation changes the
database's current view; report calculation continues to use Excel.
"""

from __future__ import annotations

import hashlib
import json
import os
import uuid
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterator

import pandas as pd

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:  # Excel-only installations remain usable.
    psycopg = None
    dict_row = None


SCHEMA_PATH = Path(__file__).resolve().parent / "sql" / "001_initial.sql"
DATA_BACKEND = os.environ.get("DATA_BACKEND", "excel").strip().lower()
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
DB_ENABLED = DATA_BACKEND in {"shadow", "postgres"} and bool(DATABASE_URL)

ALIASES: dict[str, tuple[str, ...]] = {
    "source_record_id": ("停电记录id", "停电记录ID", "记录id", "记录ID", "主键"),
    "event_id": ("事件id", "事件ID", "中压运行事件id", "中压运行事件ID", "运行事件id"),
    "user_id": ("用户id", "用户ID"),
    "work_order": ("工单号", "停电工单号"),
    "customer_code": ("用户编码", "用户编号"),
    "customer_name": ("中电联用户名称", "用户名称", "申报用户名称"),
    "customer_nature": ("用户性质", "用户类型"),
    "outage_type": ("停电类型",),
    "feeder_code": ("所属馈线编码", "馈线编码"),
    "feeder_name": ("所属馈线名称", "馈线名称", "线路段名称"),
    "outage_start": (
        "去重后停电开始时间",
        "停电开始时间",
        "停电开始时刻",
        "实际停电开始时间",
        "停电开始日期",
    ),
    "outage_end": (
        "去重后停电结束时间",
        "停电结束时间",
        "停电结束时刻",
        "实际停电结束时间",
        "停电结束日期",
    ),
    "city": ("所属地市", "地市", "供电局"),
    "district": ("所属区局", "区局", "县区局"),
    "station": ("所属供电所", "供电所"),
}


def enabled() -> bool:
    return DB_ENABLED


def _require_database() -> None:
    if not DB_ENABLED:
        raise RuntimeError("数据库未启用；请设置 DATA_BACKEND=shadow|postgres 和 DATABASE_URL")
    if psycopg is None:
        raise RuntimeError("缺少 psycopg 依赖")


@contextmanager
def connection(*, autocommit: bool = False):
    _require_database()
    assert psycopg is not None
    with psycopg.connect(
        DATABASE_URL,
        row_factory=dict_row,
        autocommit=autocommit,
        connect_timeout=3,
    ) as conn:
        yield conn


def ensure_schema() -> None:
    with connection(autocommit=True) as conn:
        conn.execute(SCHEMA_PATH.read_text(encoding="utf-8"))


def recover_interrupted_report_runs() -> int:
    """Mark reports left running by a service restart as interrupted."""
    if not DB_ENABLED:
        return 0
    with connection() as conn:
        cursor = conn.execute(
            "UPDATE report_run SET status='interrupted',finished_at=now(),"
            "error_message='服务重启，任务未正常结束' WHERE status='running'"
        )
        count = cursor.rowcount
        conn.commit()
    return count


@contextmanager
def report_advisory_lock():
    """Hold the cross-process single-report lock for one report execution."""
    if not DB_ENABLED:
        yield True
        return
    with connection(autocommit=True) as conn:
        row = conn.execute(
            "SELECT pg_try_advisory_lock(hashtext('outage-report-run')) AS acquired"
        ).fetchone()
        acquired = bool(row["acquired"])
        try:
            yield acquired
        finally:
            if acquired:
                conn.execute("SELECT pg_advisory_unlock(hashtext('outage-report-run'))")


def health() -> dict[str, Any]:
    result: dict[str, Any] = {
        "mode": DATA_BACKEND,
        "configured": bool(DATABASE_URL),
        "enabled": DB_ENABLED,
        "reachable": False,
    }
    if not DB_ENABLED:
        return result
    try:
        with connection() as conn:
            row = conn.execute(
                "SELECT current_setting('server_version') AS version, "
                "(SELECT count(*) FROM import_batch) AS imports"
            ).fetchone()
        result.update({"reachable": True, **dict(row)})
    except Exception as exc:
        result["error"] = str(exc)[:300]
    return result


def _json_value(value: Any) -> Any:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, float) and (value != value or value in (float("inf"), float("-inf"))):
        return None
    if isinstance(value, Decimal):
        return float(value)
    if hasattr(value, "item"):
        try:
            return value.item()
        except (TypeError, ValueError):
            pass
    return value if isinstance(value, (str, int, float, bool, list, dict)) else str(value)


def _text(value: Any) -> str | None:
    value = _json_value(value)
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text or None


def _pick(payload: dict[str, Any], field: str) -> Any:
    for name in ALIASES[field]:
        if name in payload and _text(payload[name]) is not None:
            return payload[name]
    return None


def _timestamp(value: Any) -> datetime | None:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    result = parsed.to_pydatetime() if isinstance(parsed, pd.Timestamp) else parsed
    # Source workbooks contain local wall-clock values. Keep them timezone-naive
    # so a database/session timezone cannot move records to the previous day.
    if result.tzinfo is not None:
        result = result.replace(tzinfo=None)
    return result


def _record_key(payload: dict[str, Any], outage_start: datetime) -> tuple[str | None, str]:
    source_record_id = _text(_pick(payload, "source_record_id"))
    if source_record_id:
        return f"record:{source_record_id}", "停电记录id"
    event_id = _text(_pick(payload, "event_id"))
    user_id = _text(_pick(payload, "user_id"))
    if event_id and user_id:
        return f"event-user:{event_id}|{user_id}", "事件id+用户id"
    work_order = _text(_pick(payload, "work_order"))
    customer_code = _text(_pick(payload, "customer_code"))
    feeder_code = _text(_pick(payload, "feeder_code"))
    if work_order and customer_code and feeder_code:
        return (
            f"fallback:{work_order}|{customer_code}|{feeder_code}|{outage_start.isoformat()}",
            "工单号+用户编码+馈线编码+停电开始时间",
        )
    return None, ""


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _executemany(conn, statement: str, parameters: list[Any]) -> None:
    with conn.cursor() as cursor:
        cursor.executemany(statement, parameters)


def _iter_workbook_rows(path: Path) -> Iterator[tuple[str, int, dict[str, Any]]]:
    if path.suffix.lower() == ".xlsx":
        import openpyxl

        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            for worksheet in workbook.worksheets:
                # Some South Grid exports declare an incorrect A1:A1 worksheet
                # dimension although thousands of rows exist. read_only mode
                # trusts that metadata unless dimensions are reset explicitly.
                if worksheet.calculate_dimension() == "A1:A1":
                    worksheet.reset_dimensions()
                rows = worksheet.iter_rows(values_only=True)
                try:
                    raw_headers = next(rows)
                except StopIteration:
                    continue
                headers = [str(value).strip() if value is not None else "" for value in raw_headers]
                for row_number, values in enumerate(rows, start=2):
                    if not any(value is not None and str(value).strip() for value in values):
                        continue
                    payload = {
                        header: _json_value(value)
                        for header, value in zip(headers, values)
                        if header
                    }
                    yield worksheet.title, row_number, payload
        finally:
            workbook.close()
        return

    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    for sheet_name, table in sheets.items():
        for index, row in table.iterrows():
            yield str(sheet_name), int(index) + 2, {
                str(column): _json_value(value) for column, value in row.items()
            }


INSERT_RECORD_SQL = """
INSERT INTO outage_record_version (
    batch_id, source_sheet, source_row, record_key, content_hash,
    source_record_id, event_id, user_id, work_order, customer_code, feeder_code,
    outage_start, outage_end, city, district, station, raw_payload
) VALUES (
    %(batch_id)s, %(source_sheet)s, %(source_row)s, %(record_key)s, %(content_hash)s,
    %(source_record_id)s, %(event_id)s, %(user_id)s, %(work_order)s,
    %(customer_code)s, %(feeder_code)s, %(outage_start)s, %(outage_end)s,
    %(city)s, %(district)s, %(station)s, %(raw_payload)s::jsonb
)
"""


def stage_import(path: Path | str, *, synthetic_reconciliation: bool = False) -> dict[str, Any]:
    """Validate and stage an uploaded workbook without changing current data."""

    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    batch_id = uuid.uuid4()
    file_hash = _sha256(source)
    errors: list[dict[str, Any]] = []
    duplicates: list[str] = []
    seen: set[str] = set()
    key_sources: Counter[str] = Counter()
    start_date: date | None = None
    end_date: date | None = None
    row_count = 0
    invalid_count = 0
    duplicate_count = 0
    discovered_columns: set[str] = set()

    with connection() as conn:
        existing = conn.execute(
            "SELECT id, status, inferred_start_date, inferred_end_date, row_count, "
            "validation_result FROM import_batch WHERE file_sha256=%s "
            "ORDER BY uploaded_at DESC LIMIT 1",
            (file_hash,),
        ).fetchone()
        if existing and existing["status"] != "validation_error":
            return {
                "batch_id": str(existing["id"]),
                "duplicate_file": True,
                "status": existing["status"],
                "row_count": existing["row_count"],
                "inferred_start_date": (
                    existing["inferred_start_date"].isoformat()
                    if existing["inferred_start_date"] else None
                ),
                "inferred_end_date": (
                    existing["inferred_end_date"].isoformat()
                    if existing["inferred_end_date"] else None
                ),
                "validation": existing["validation_result"],
                "preview_url": f"/api/imports/{existing['id']}/preview",
            }
        if existing:
            batch_id = existing["id"]
            conn.execute("DELETE FROM outage_record_version WHERE batch_id=%s", (batch_id,))
            conn.execute(
                "UPDATE import_batch SET status='validating',uploaded_at=now(),"
                "validation_result='{}'::jsonb WHERE id=%s",
                (batch_id,),
            )
        else:
            conn.execute(
                "INSERT INTO import_batch "
                "(id, filename, file_sha256, status, original_file_path, synthetic_reconciliation) "
                "VALUES (%s, %s, %s, 'validating', %s, %s)",
                (batch_id, source.name, file_hash, str(source), synthetic_reconciliation),
            )

        buffer: list[dict[str, Any]] = []
        for sheet_name, source_row, payload in _iter_workbook_rows(source):
            discovered_columns.update(payload)
            outage_start = _timestamp(_pick(payload, "outage_start"))
            if outage_start is None:
                invalid_count += 1
                if len(errors) < 100:
                    errors.append({"sheet": sheet_name, "row": source_row, "error": "非法或缺失停电开始时间"})
                continue
            record_key, key_source = _record_key(payload, outage_start)
            if record_key is None:
                invalid_count += 1
                if len(errors) < 100:
                    errors.append({"sheet": sheet_name, "row": source_row, "error": "无法生成记录匹配键"})
                continue
            if record_key in seen:
                duplicate_count += 1
                if len(duplicates) < 100:
                    duplicates.append(record_key)
                continue
            seen.add(record_key)
            key_sources[key_source] += 1
            canonical = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
            content_hash = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
            outage_date = outage_start.date()
            start_date = outage_date if start_date is None else min(start_date, outage_date)
            end_date = outage_date if end_date is None else max(end_date, outage_date)
            row_count += 1
            buffer.append(
                {
                    "batch_id": batch_id,
                    "source_sheet": sheet_name,
                    "source_row": source_row,
                    "record_key": record_key,
                    "content_hash": content_hash,
                    "source_record_id": _text(_pick(payload, "source_record_id")),
                    "event_id": _text(_pick(payload, "event_id")),
                    "user_id": _text(_pick(payload, "user_id")),
                    "work_order": _text(_pick(payload, "work_order")),
                    "customer_code": _text(_pick(payload, "customer_code")),
                    "feeder_code": _text(_pick(payload, "feeder_code")),
                    "outage_start": outage_start,
                    "outage_end": _timestamp(_pick(payload, "outage_end")),
                    "city": _text(_pick(payload, "city")),
                    "district": _text(_pick(payload, "district")),
                    "station": _text(_pick(payload, "station")),
                    "raw_payload": canonical,
                }
            )
            if len(buffer) >= 1000:
                _executemany(conn, INSERT_RECORD_SQL, buffer)
                buffer.clear()
        if buffer:
            _executemany(conn, INSERT_RECORD_SQL, buffer)

        required_fields = (
            "work_order", "customer_code", "customer_name", "customer_nature",
            "outage_type", "outage_start", "outage_end", "feeder_code",
            "feeder_name", "station", "district", "city",
        )
        missing_fields = [
            field for field in required_fields
            if not any(alias in discovered_columns for alias in ALIASES[field])
        ]
        if missing_fields:
            errors.append({"error": "缺少业务必需字段", "fields": missing_fields})

        status = "validation_error" if errors or duplicates or row_count == 0 else "pending_confirmation"
        validation = {
            "errors": errors,
            "duplicate_keys": duplicates,
            "duplicate_key_count": duplicate_count,
            "invalid_row_count": invalid_count,
            "missing_fields": missing_fields,
            "key_sources": dict(key_sources),
            "valid_rows": row_count,
        }
        conn.execute(
            "UPDATE import_batch SET inferred_start_date=%s, inferred_end_date=%s, "
            "row_count=%s, status=%s, validation_result=%s::jsonb WHERE id=%s",
            (start_date, end_date, row_count, status, json.dumps(validation, ensure_ascii=False), batch_id),
        )
        conn.commit()

    return {
        "batch_id": str(batch_id),
        "status": status,
        "row_count": row_count,
        "inferred_start_date": start_date.isoformat() if start_date else None,
        "inferred_end_date": end_date.isoformat() if end_date else None,
        "validation": validation,
        "preview_url": f"/api/imports/{batch_id}/preview",
    }


def stage_mask(path: Path | str, year: int) -> dict[str, Any]:
    """Version the currently uploaded Table-11 mask while retaining old versions."""
    source = Path(path).resolve()
    frame = pd.read_excel(source)
    column = next((name for name in frame.columns if "重大事件日期" in str(name)), None)
    errors: list[str] = []
    dates: set[date] = set()
    if column is None:
        errors.append("未找到重大事件日期列")
    else:
        for value in frame[column]:
            parsed = _timestamp(value)
            if parsed is not None and parsed.year == year:
                dates.add(parsed.date())
    batch_id = uuid.uuid4()
    status = "validation_error" if errors else "active"
    with connection() as conn:
        if not errors:
            conn.execute(
                "UPDATE mask_batch SET status='superseded' WHERE mask_year=%s AND status='active'",
                (year,),
            )
        conn.execute(
            "INSERT INTO mask_batch (id,filename,file_sha256,status,mask_year,activated_at,"
            "validation_result,original_file_path) VALUES (%s,%s,%s,%s,%s,%s,%s::jsonb,%s)",
            (
                batch_id,
                source.name,
                _sha256(source),
                status,
                year,
                datetime.now(timezone.utc) if not errors else None,
                json.dumps({"errors": errors, "date_count": len(dates)}, ensure_ascii=False),
                str(source),
            ),
        )
        if dates:
            _executemany(
                conn,
                "INSERT INTO mask_date (mask_batch_id,mask_date) VALUES (%s,%s)",
                [(batch_id, value) for value in sorted(dates)],
            )
        conn.commit()
    return {
        "mask_batch_id": str(batch_id),
        "mask_status": status,
        "mask_date_count": len(dates),
    }


def _load_batch(batch_id: str, conn) -> dict[str, Any]:
    try:
        parsed = uuid.UUID(batch_id)
    except ValueError as exc:
        raise ValueError("无效的 batch_id") from exc
    row = conn.execute("SELECT * FROM import_batch WHERE id=%s", (parsed,)).fetchone()
    if not row:
        raise KeyError("批次不存在")
    return dict(row)


def preview_import(batch_id: str, start: date | None = None, end: date | None = None) -> dict[str, Any]:
    with connection() as conn:
        batch = _load_batch(batch_id, conn)
        replace_start = start or batch["inferred_start_date"]
        replace_end = end or batch["inferred_end_date"]
        if not replace_start or not replace_end or replace_start > replace_end:
            raise ValueError("替换日期范围无效")

        staged_rows = conn.execute(
            "SELECT record_key, content_hash, outage_start::date AS day, coalesce(city, '') AS city "
            "FROM outage_record_version WHERE batch_id=%s AND outage_start::date BETWEEN %s AND %s",
            (batch["id"], replace_start, replace_end),
        ).fetchall()
        current_rows = conn.execute(
            "SELECT record_key, content_hash, outage_start::date AS day, coalesce(city, '') AS city "
            "FROM current_outage_record WHERE outage_start::date BETWEEN %s AND %s",
            (replace_start, replace_end),
        ).fetchall()

    staged = {row["record_key"]: row for row in staged_rows}
    current = {row["record_key"]: row for row in current_rows}
    staged_keys, current_keys = set(staged), set(current)
    common = staged_keys & current_keys
    unchanged = sum(staged[key]["content_hash"] == current[key]["content_hash"] for key in common)
    modified = len(common) - unchanged
    daily_new = Counter(row["day"].isoformat() for row in staged_rows)
    daily_old = Counter(row["day"].isoformat() for row in current_rows)
    city_new = Counter(row["city"] or "未归属" for row in staged_rows)
    city_old = Counter(row["city"] or "未归属" for row in current_rows)
    date_index = pd.date_range(replace_start, replace_end, freq="D")
    missing_dates = [stamp.date().isoformat() for stamp in date_index if daily_new[stamp.date().isoformat()] == 0]
    daily_changes = [
        {"date": day, "before": daily_old[day], "after": daily_new[day], "delta": daily_new[day] - daily_old[day]}
        for day in sorted(set(daily_new) | set(daily_old))
        if daily_new[day] != daily_old[day]
    ]
    city_impact = [
        {"city": city, "before": city_old[city], "after": city_new[city], "delta": city_new[city] - city_old[city]}
        for city in sorted(set(city_new) | set(city_old))
        if city_new[city] != city_old[city]
    ]
    return {
        "batch_id": batch_id,
        "status": batch["status"],
        "filename": batch["filename"],
        "inferred_start_date": batch["inferred_start_date"].isoformat() if batch["inferred_start_date"] else None,
        "inferred_end_date": batch["inferred_end_date"].isoformat() if batch["inferred_end_date"] else None,
        "replace_start_date": replace_start.isoformat(),
        "replace_end_date": replace_end.isoformat(),
        "added": len(staged_keys - current_keys),
        "deleted": len(current_keys - staged_keys),
        "modified": modified,
        "unchanged": unchanged,
        "missing_dates": missing_dates,
        "span_days": (replace_end - replace_start).days + 1,
        "abnormal_span": (replace_end - replace_start).days + 1 > 120,
        "daily_changes": daily_changes,
        "city_impact": city_impact,
        "validation": batch["validation_result"],
    }


def activate_import(batch_id: str, start: date, end: date, confirmed_by: str) -> dict[str, Any]:
    preview = preview_import(batch_id, start, end)
    activation_id = uuid.uuid4()
    with connection() as conn:
        batch = _load_batch(batch_id, conn)
        if batch["status"] != "pending_confirmation":
            raise ValueError(f"批次状态 {batch['status']} 不允许激活")
        if batch["inferred_start_date"] < start or batch["inferred_end_date"] > end:
            raise ValueError("确认范围必须覆盖批次中全部停电日期")
        conn.execute("SELECT pg_advisory_xact_lock(hashtext('outage-import-activation'))")
        conn.execute(
            "INSERT INTO batch_record_audit "
            "(batch_id,record_key,content_hash,source_sheet,source_row,outage_start,city) "
            "SELECT batch_id,record_key,content_hash,source_sheet,source_row,outage_start,city "
            "FROM outage_record_version WHERE batch_id=%s",
            (batch["id"],),
        )
        conn.execute(
            "CREATE TEMP TABLE unchanged_staged ON COMMIT DROP AS "
            "SELECT staged.id AS staged_id,current.id AS current_id,staged.record_key "
            "FROM outage_record_version staged JOIN current_outage_record current "
            "ON current.record_key=staged.record_key AND current.content_hash=staged.content_hash "
            "WHERE staged.batch_id=%s",
            (batch["id"],),
        )
        conn.execute(
            "UPDATE batch_record_audit audit SET version_id=unchanged.current_id "
            "FROM unchanged_staged unchanged WHERE audit.batch_id=%s "
            "AND audit.record_key=unchanged.record_key",
            (batch["id"],),
        )
        # Unchanged rows retain their existing version. The batch audit points
        # to it, avoiding a full JSON copy for every overlapping upload.
        conn.execute(
            "DELETE FROM outage_record_version WHERE id IN "
            "(SELECT staged_id FROM unchanged_staged)"
        )
        conn.execute(
            "UPDATE outage_record_version SET superseded_at=now(), superseded_by_batch_id=%s "
            "WHERE activated_at IS NOT NULL AND superseded_at IS NULL AND rolled_back_at IS NULL "
            "AND outage_start::date BETWEEN %s AND %s "
            "AND NOT EXISTS (SELECT 1 FROM batch_record_audit audit "
            "WHERE audit.batch_id=%s AND audit.version_id=outage_record_version.id)",
            (batch["id"], start, end, batch["id"]),
        )
        conn.execute(
            "UPDATE outage_record_version SET activated_at=now() WHERE batch_id=%s",
            (batch["id"],),
        )
        conn.execute(
            "UPDATE batch_record_audit audit SET version_id=version.id "
            "FROM outage_record_version version WHERE audit.batch_id=%s "
            "AND version.batch_id=%s AND audit.record_key=version.record_key",
            (batch["id"], batch["id"]),
        )
        conn.execute(
            "INSERT INTO batch_activation "
            "(id, batch_id, confirmed_by, replace_start_date, replace_end_date, added_count, "
            "deleted_count, modified_count, unchanged_count) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (
                activation_id, batch["id"], confirmed_by, start, end,
                preview["added"], preview["deleted"], preview["modified"], preview["unchanged"],
            ),
        )
        conn.execute(
            "UPDATE import_batch SET status='active', confirmed_start_date=%s, "
            "confirmed_end_date=%s WHERE id=%s",
            (start, end, batch["id"]),
        )
        conn.commit()
    return {"activation_id": str(activation_id), "status": "active", **preview}


def reject_import(batch_id: str) -> dict[str, Any]:
    with connection() as conn:
        batch = _load_batch(batch_id, conn)
        if batch["status"] not in {"pending_confirmation", "validation_error"}:
            raise ValueError(f"批次状态 {batch['status']} 不允许拒绝")
        conn.execute(
            "DELETE FROM outage_record_version WHERE batch_id=%s AND activated_at IS NULL",
            (batch["id"],),
        )
        conn.execute("UPDATE import_batch SET status='rejected' WHERE id=%s", (batch["id"],))
        conn.commit()
    return {"batch_id": batch_id, "status": "rejected"}


def discard_invalid_staging(batch_id: str) -> int:
    """Keep validation metadata but release bulky rows for a non-activatable batch."""
    with connection() as conn:
        batch = _load_batch(batch_id, conn)
        if batch["status"] != "validation_error":
            return 0
        cursor = conn.execute(
            "DELETE FROM outage_record_version WHERE batch_id=%s AND activated_at IS NULL",
            (batch["id"],),
        )
        count = cursor.rowcount
        conn.commit()
    return count


def rollback_import(batch_id: str, confirmed_by: str) -> dict[str, Any]:
    with connection() as conn:
        batch = _load_batch(batch_id, conn)
        activation = conn.execute(
            "SELECT * FROM batch_activation WHERE batch_id=%s AND rolled_back_at IS NULL",
            (batch["id"],),
        ).fetchone()
        if not activation:
            raise ValueError("批次没有可回滚的激活记录")
        latest = conn.execute(
            "SELECT id FROM batch_activation WHERE rollback_of IS NULL AND rolled_back_at IS NULL "
            "ORDER BY confirmed_at DESC LIMIT 1"
        ).fetchone()
        if not latest or latest["id"] != activation["id"]:
            raise ValueError("仅支持直接回滚最近一次激活")
        conn.execute("SELECT pg_advisory_xact_lock(hashtext('outage-import-activation'))")
        conn.execute(
            "UPDATE outage_record_version SET rolled_back_at=now() WHERE batch_id=%s AND activated_at IS NOT NULL",
            (batch["id"],),
        )
        conn.execute(
            "UPDATE outage_record_version SET superseded_at=NULL, superseded_by_batch_id=NULL "
            "WHERE superseded_by_batch_id=%s",
            (batch["id"],),
        )
        conn.execute("UPDATE batch_activation SET rolled_back_at=now() WHERE id=%s", (activation["id"],))
        rollback_id = uuid.uuid4()
        conn.execute(
            "INSERT INTO batch_activation (id,batch_id,confirmed_by,replace_start_date,replace_end_date,"
            "added_count,deleted_count,modified_count,unchanged_count,rollback_of,rolled_back_at) "
            "VALUES (%s,%s,%s,%s,%s,0,0,0,0,%s,now())",
            (
                rollback_id, batch["id"], confirmed_by,
                activation["replace_start_date"], activation["replace_end_date"], activation["id"],
            ),
        )
        conn.execute("UPDATE import_batch SET status='rolled_back' WHERE id=%s", (batch["id"],))
        conn.commit()
    return {"batch_id": batch_id, "status": "rolled_back", "rollback_id": str(rollback_id)}


def latest_activation_id() -> str | None:
    if not DB_ENABLED:
        return None
    with connection() as conn:
        row = conn.execute(
            "SELECT id FROM batch_activation WHERE rollback_of IS NULL AND rolled_back_at IS NULL "
            "ORDER BY confirmed_at DESC LIMIT 1"
        ).fetchone()
    return str(row["id"]) if row else None


def latest_mask_batch_id() -> str | None:
    if not DB_ENABLED:
        return None
    with connection() as conn:
        row = conn.execute(
            "SELECT id FROM mask_batch WHERE status='active' ORDER BY activated_at DESC LIMIT 1"
        ).fetchone()
    return str(row["id"]) if row else None


def create_report_run(job_id: str, parameters: dict[str, Any]) -> str | None:
    if not DB_ENABLED:
        return None
    report_id = uuid.uuid4()
    activation = parameters.get("activation_id") or latest_activation_id()
    if activation is not None:
        activation = uuid.UUID(str(activation))
    mask_batch = latest_mask_batch_id()
    with connection() as conn:
        conn.execute(
            "INSERT INTO report_run (id,job_id,activation_id,mask_batch_id,parameters,status) "
            "VALUES (%s,%s,%s,%s,%s::jsonb,'pending')",
            (report_id, job_id, activation, mask_batch, json.dumps(parameters, ensure_ascii=False)),
        )
        conn.commit()
    return str(report_id)


def update_report_run(
    job_id: str,
    status: str,
    *,
    output_files: dict[str, Any] | None = None,
    error_message: str | None = None,
    elapsed_seconds: float | None = None,
) -> None:
    if not DB_ENABLED:
        return
    finished = status in {"success", "failed", "cancelled", "interrupted"}
    with connection() as conn:
        conn.execute(
            "UPDATE report_run SET status=%s, started_at=CASE WHEN %s='running' "
            "THEN coalesce(started_at,now()) ELSE started_at END, "
            "finished_at=CASE WHEN %s THEN now() ELSE finished_at END, "
            "elapsed_seconds=coalesce(%s,elapsed_seconds), "
            "output_files=coalesce(%s::jsonb,output_files), "
            "error_message=coalesce(%s,error_message) WHERE job_id=%s",
            (
                status, status, finished, elapsed_seconds,
                json.dumps(output_files, ensure_ascii=False) if output_files is not None else None,
                error_message, job_id,
            ),
        )
        conn.commit()
