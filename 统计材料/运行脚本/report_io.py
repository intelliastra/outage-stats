"""Memory-bounded, atomic Excel report writer shared by statistics scripts."""

from __future__ import annotations

import math
import os
import tempfile
import time
from datetime import date, datetime
from pathlib import Path
from typing import Mapping

import pandas as pd
import xlsxwriter


DEFAULT_TEXT_COLUMNS = {"工单号", "用户编码", "所属馈线编码"}
DEFAULT_NUMBER_COLUMNS = {
    "停电总次数",
    "故障停电次数",
    "预安排停电次数",
    "用户停电总次数",
}


def _peak_memory_mb() -> float | None:
    try:
        import resource

        value = float(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
        # Linux reports KiB, macOS reports bytes.
        return value / (1024 * 1024) if value > 10_000_000 else value / 1024
    except (ImportError, OSError, ValueError):
        return None


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    try:
        result = pd.isna(value)
        return bool(result) if not hasattr(result, "__len__") else False
    except (TypeError, ValueError):
        return False


def _identifier(value: object) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    return text[:-2] if text.endswith(".0") else text


def _number(value: object) -> int | float | None:
    if _is_blank(value) or value == "":
        return None
    try:
        converted = float(value)
        return int(converted) if converted.is_integer() else converted
    except (TypeError, ValueError, OverflowError):
        return None


def _display_width(value: object) -> int:
    if _is_blank(value):
        return 0
    text = str(value)
    # A CJK glyph is roughly two Latin characters wide in Excel.
    return sum(2 if ord(char) > 255 else 1 for char in text)


def _prepare_columns(
    table: pd.DataFrame,
    text_columns: set[str],
    number_columns: set[str],
) -> tuple[list[str], set[int], set[int]]:
    columns = [str(column) for column in table.columns]
    text_indexes = {
        index
        for index, column in enumerate(columns)
        if column in text_columns or column.endswith("编码")
    }
    number_indexes = {
        index for index, column in enumerate(columns) if column in number_columns
    }
    return columns, text_indexes, number_indexes


def write_tables_streaming(
    tables: Mapping[str, pd.DataFrame],
    output_path: Path | str,
    *,
    text_columns: set[str] | None = None,
    number_columns: set[str] | None = None,
    width_sample_rows: int = 2000,
) -> Path:
    """Write tables row-by-row and atomically publish a complete XLSX file.

    XlsxWriter ``constant_memory`` keeps only the current row in memory.  The
    destination is replaced only after the workbook closes successfully, so a
    failed export never exposes a partial report.
    """

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    text_columns = text_columns or DEFAULT_TEXT_COLUMNS
    number_columns = number_columns or DEFAULT_NUMBER_COLUMNS
    started = time.perf_counter()

    handle = tempfile.NamedTemporaryFile(
        prefix=f".{destination.stem}.",
        suffix=".xlsx.tmp",
        dir=destination.parent,
        delete=False,
    )
    temp_path = Path(handle.name)
    handle.close()

    try:
        workbook = xlsxwriter.Workbook(
            str(temp_path),
            {
                "constant_memory": True,
                "strings_to_urls": False,
                "strings_to_formulas": False,
                "nan_inf_to_errors": True,
            },
        )
        header_format = workbook.add_format({"bold": True})
        text_format = workbook.add_format({"num_format": "@"})
        number_format = workbook.add_format({"num_format": "0"})
        datetime_format = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})

        try:
            for sheet_name, table in tables.items():
                sheet_started = time.perf_counter()
                worksheet = workbook.add_worksheet(str(sheet_name)[:31])
                columns, text_indexes, number_indexes = _prepare_columns(
                    table, text_columns, number_columns
                )

                for column_index, column in enumerate(columns):
                    worksheet.write(0, column_index, column, header_format)

                widths = [_display_width(column) for column in columns]
                for row_index, values in enumerate(
                    table.itertuples(index=False, name=None), start=1
                ):
                    for column_index, value in enumerate(values):
                        if (
                            not isinstance(value, (str, bytes, date, datetime, pd.Timestamp))
                            and hasattr(value, "item")
                        ):
                            try:
                                value = value.item()
                            except (TypeError, ValueError):
                                pass
                        if column_index < len(widths) and row_index <= width_sample_rows:
                            widths[column_index] = max(
                                widths[column_index], _display_width(value)
                            )

                        if column_index in text_indexes:
                            worksheet.write_string(
                                row_index, column_index, _identifier(value), text_format
                            )
                        elif column_index in number_indexes:
                            converted = _number(value)
                            if converted is None:
                                worksheet.write_blank(
                                    row_index, column_index, None, number_format
                                )
                            else:
                                worksheet.write_number(
                                    row_index, column_index, converted, number_format
                                )
                        elif _is_blank(value):
                            worksheet.write_blank(row_index, column_index, None)
                        elif isinstance(value, pd.Timestamp):
                            worksheet.write_datetime(
                                row_index, column_index, value.to_pydatetime(), datetime_format
                            )
                        elif isinstance(value, datetime):
                            worksheet.write_datetime(
                                row_index, column_index, value, datetime_format
                            )
                        elif isinstance(value, date):
                            worksheet.write_datetime(
                                row_index,
                                column_index,
                                datetime.combine(value, datetime.min.time()),
                                date_format,
                            )
                        elif isinstance(value, bool):
                            worksheet.write_boolean(row_index, column_index, value)
                        elif isinstance(value, (int, float)) and not (
                            isinstance(value, float) and math.isnan(value)
                        ):
                            worksheet.write_number(row_index, column_index, value)
                        else:
                            worksheet.write(row_index, column_index, value)

                worksheet.freeze_panes(1, 0)
                if columns:
                    worksheet.autofilter(0, 0, len(table), len(columns) - 1)
                for column_index, width in enumerate(widths):
                    column_format = (
                        text_format
                        if column_index in text_indexes
                        else number_format if column_index in number_indexes else None
                    )
                    worksheet.set_column(
                        column_index, column_index, min(max(width + 2, 10), 38), column_format
                    )

                print(
                    f"[PERF] Excel Sheet={sheet_name} rows={len(table)} "
                    f"cols={len(columns)} elapsed={time.perf_counter() - sheet_started:.1f}s",
                    flush=True,
                )
        finally:
            workbook.close()

        os.replace(temp_path, destination)
        size_mb = destination.stat().st_size / (1024 * 1024)
        peak_mb = _peak_memory_mb()
        peak_text = f" peak_memory={peak_mb:.1f}MB" if peak_mb is not None else ""
        print(
            f"[PERF] Excel output={destination.name} rows="
            f"{sum(len(table) for table in tables.values())} size={size_mb:.1f}MB "
            f"elapsed={time.perf_counter() - started:.1f}s{peak_text}",
            flush=True,
        )
        return destination
    except BaseException:
        temp_path.unlink(missing_ok=True)
        raise


def write_text_atomic(path: Path | str, text: str) -> Path:
    """Atomically publish a UTF-8 text artifact."""
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        prefix=f".{destination.stem}.",
        suffix=".txt.tmp",
        dir=destination.parent,
        delete=False,
        mode="w",
        encoding="utf-8",
        newline="\n",
    )
    temp_path = Path(handle.name)
    try:
        with handle:
            handle.write(text)
            if text and not text.endswith("\n"):
                handle.write("\n")
        os.replace(temp_path, destination)
        return destination
    except BaseException:
        temp_path.unlink(missing_ok=True)
        raise
