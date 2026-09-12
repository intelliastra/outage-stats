"""Read the active outage dataset from PostgreSQL for the existing pandas rules."""

from __future__ import annotations

import json
import os
from datetime import date

import pandas as pd


def load_active_records(start: date | None = None, end: date | None = None) -> pd.DataFrame:
    try:
        import psycopg
    except ImportError as exc:
        raise RuntimeError("DATA_BACKEND=postgres 需要安装 psycopg") from exc

    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        raise RuntimeError("DATA_BACKEND=postgres 但未配置 DATABASE_URL")

    activation_id = os.environ.get("REPORT_ACTIVATION_ID", "").strip()
    where: list[str] = []
    params: list[object] = []
    target_time = None
    with psycopg.connect(database_url) as lookup_connection:
        if activation_id:
            row = lookup_connection.execute(
                "SELECT confirmed_at FROM batch_activation WHERE id=%s", (activation_id,)
            ).fetchone()
            if not row:
                raise RuntimeError(f"指定的 activation_id 不存在：{activation_id}")
            target_time = row[0]
    if target_time is None:
        where.extend(["activated_at IS NOT NULL", "superseded_at IS NULL", "rolled_back_at IS NULL"])
    else:
        where.extend(
            [
                "activated_at <= %s",
                "(superseded_at IS NULL OR superseded_at > %s)",
                "(rolled_back_at IS NULL OR rolled_back_at > %s)",
            ]
        )
        params.extend([target_time, target_time, target_time])
    if start is not None:
        where.append("outage_start::date >= %s")
        params.append(start)
    if end is not None:
        where.append("outage_start::date <= %s")
        params.append(end)

    rows: list[dict] = []
    query = (
        "SELECT raw_payload FROM outage_record_version WHERE "
        + " AND ".join(where)
        + " ORDER BY outage_start, id"
    )
    with psycopg.connect(database_url) as connection:
        with connection.cursor(name="active_outage_stream") as cursor:
            cursor.itersize = 2000
            cursor.execute(query, params)
            for (payload,) in cursor:
                rows.append(payload if isinstance(payload, dict) else json.loads(payload))
    frame = pd.DataFrame.from_records(rows)
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame.dropna(how="all").reset_index(drop=True)
