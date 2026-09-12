"""FastAPI backend for outage statistics web app."""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
import zipfile
from datetime import date
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

import database

from exclude_utils import (
    DEFAULT_EXCLUDE_YEAR,
    get_year_latest_file_info,
    list_exclude_years,
    validate_exclude_year,
)
from historical_utils import (
    HISTORICAL_YEARS,
    annual_2025_target,
    build_historical_config,
    ensure_annual_2025_file,
    resolve_exclude_file,
    resolve_period_dates,
)
from job_manager import JobManager, JobStatus
from log_parser import validate_output_path

# Paths
WEB_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = WEB_DIR / "frontend"
STATS_BASE = Path(os.environ.get("STATS_BASE", str(WEB_DIR.parent / "统计材料"))).resolve()
NEWDATA_DIR = STATS_BASE / "input" / "Newdata"
OUTPUT_BASE = STATS_BASE / "output"
EXCLUDE_BASE = STATS_BASE / "input" / "exclude"

# --- 三个页面独立的输入/输出子目录 ---
DAILY_NEWDATA_DIR = NEWDATA_DIR / "daily"
HISTORICAL_NEWDATA_DIR = NEWDATA_DIR / "historical"
EXCLUDE_NEWDATA_DIR = NEWDATA_DIR / "exclude"

DAILY_OUTPUT_BASE = OUTPUT_BASE / "daily"
HISTORICAL_OUTPUT_BASE = OUTPUT_BASE / "historical"
EXCLUDE_OUTPUT_BASE = OUTPUT_BASE / "exclude"

MAX_UPLOAD_BYTES = 150 * 1024 * 1024
MAX_CHUNK_BYTES = 4 * 1024 * 1024
UPLOAD_TMP = STATS_BASE / ".upload_tmp"
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

app = FastAPI(title="供电可靠性统计材料", version="1.1.0")
job_manager = JobManager(STATS_BASE)


class HistoricalRunRequest(BaseModel):
    year: int = Field(..., description="主数据年份 2025 或 2026")
    start_month: int = Field(..., ge=1, le=12)
    start_day: int = Field(..., ge=1, le=31)
    end_month: int = Field(..., ge=1, le=12)
    end_day: int = Field(..., ge=1, le=31)


class ExcludeRunRequest(BaseModel):
    input_filename: str = Field(..., description="已上传到服务器临时目录的文件名")


class DailyRunRequest(BaseModel):
    activation_id: str | None = Field(default=None, description="可选的数据激活版本")


class ImportActionRequest(BaseModel):
    start_date: date | None = None
    end_date: date | None = None
    confirmed_by: str = Field(default="admin", min_length=1, max_length=100)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def initialize_optional_database() -> None:
    if not database.enabled():
        return
    try:
        database.ensure_schema()
        interrupted = database.recover_interrupted_report_runs()
        if interrupted:
            print(f"[DB] 已标记 {interrupted} 个重启前未结束任务为 interrupted", flush=True)
    except Exception as exc:
        # Shadow database failure must not take down the existing Excel service.
        print(f"[DB] 初始化失败：{exc}", flush=True)


async def _stage_daily_upload(path: Path, dest: str) -> dict[str, object]:
    if dest not in ("daily", "newdata") or not database.enabled():
        return {}
    try:
        result = await run_in_threadpool(database.stage_import, path)
        result["database_mode"] = database.DATA_BACKEND
        return result
    except Exception as exc:
        if database.DATA_BACKEND == "postgres":
            raise HTTPException(status_code=503, detail=f"数据库批次校验失败：{exc}") from exc
        return {"database_status": "shadow_error", "database_error": str(exc)[:300]}


async def _stage_mask_upload(path: Path, year: int) -> dict[str, object]:
    if not database.enabled():
        return {}
    try:
        return await run_in_threadpool(database.stage_mask, path, year)
    except Exception as exc:
        if database.DATA_BACKEND == "postgres":
            raise HTTPException(status_code=503, detail=f"Mask 版本登记失败：{exc}") from exc
        return {"database_status": "shadow_error", "database_error": str(exc)[:300]}


def sanitize_filename(name: str) -> str:
    base = Path(name).name
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base)
    if not base or base in (".", ".."):
        raise HTTPException(status_code=400, detail="无效的文件名")
    suffix = Path(base).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="仅支持 .xlsx / .xls 文件")
    return base


def _resolve_newdata_dir(dest: str) -> Path:
    """根据页面标识返回对应的 Newdata 子目录。"""
    if dest in ("historical",):
        return HISTORICAL_NEWDATA_DIR
    if dest in ("exclude", "exclude_newdata"):
        return EXCLUDE_NEWDATA_DIR
    # "daily", "newdata" 及其他默认值 → 日常统计
    return DAILY_NEWDATA_DIR


def _validate_upload_id(upload_id: str) -> str:
    if not re.fullmatch(r"[0-9a-fA-F-]{8,64}", upload_id or ""):
        raise HTTPException(status_code=400, detail="无效的 upload_id")
    return upload_id


def _chunk_dir(upload_id: str) -> Path:
    root = UPLOAD_TMP.resolve()
    target = (root / _validate_upload_id(upload_id)).resolve()
    if target != root and root not in target.parents:
        raise HTTPException(status_code=400, detail="无效路径")
    target.mkdir(parents=True, exist_ok=True)
    return target


def _cleanup_stale_uploads(max_age_sec: int = 7200) -> None:
    if not UPLOAD_TMP.is_dir():
        return
    now = time.time()
    for item in UPLOAD_TMP.iterdir():
        try:
            if now - item.stat().st_mtime > max_age_sec:
                if item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
                else:
                    item.unlink(missing_ok=True)
        except OSError:
            continue


def _write_assembled_file(parts_dir: Path, total: int, dest: Path) -> int:
    missing = [i for i in range(total) if not (parts_dir / f"{i:08d}.part").is_file()]
    if missing:
        raise HTTPException(status_code=400, detail=f"分片不完整，缺少 {missing[:8]}")
    size = 0
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".uploading")
    with tmp.open("wb") as out:
        for i in range(total):
            data = (parts_dir / f"{i:08d}.part").read_bytes()
            size += len(data)
            if size > MAX_UPLOAD_BYTES:
                tmp.unlink(missing_ok=True)
                raise HTTPException(status_code=400, detail="文件大小超过 150MB 限制")
            out.write(data)
    tmp.replace(dest)
    return size


def exclude_year_dir(year: int) -> Path:
    validate_exclude_year(year)
    target = EXCLUDE_BASE / str(year)
    target.mkdir(parents=True, exist_ok=True)
    return target


def _web_asset_version() -> str:
    path = FRONTEND_DIR / "ASSET_VERSION"
    if not path.is_file():
        return ""
    line = path.read_text(encoding="utf-8").strip().splitlines()
    return line[0].strip() if line else ""


@app.middleware("http")
async def html_no_cache(request: Request, call_next):
    response = await call_next(request)
    path = request.url.path
    if path.endswith(".html") or path in ("/", ""):
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
    return response


@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "stats_base": str(STATS_BASE),
        "daily_newdata_exists": DAILY_NEWDATA_DIR.is_dir(),
        "historical_newdata_exists": HISTORICAL_NEWDATA_DIR.is_dir(),
        "exclude_newdata_exists": EXCLUDE_NEWDATA_DIR.is_dir(),
        "exclude_base_exists": EXCLUDE_BASE.is_dir(),
        "busy": job_manager.is_busy(),
        "script_name": job_manager.script_path.name,
        "script_version": job_manager.script_version,
        "historical_script": job_manager.historical_script_path.name,
        "historical_script_exists": job_manager.historical_script_path.is_file(),
        "exclude_script": job_manager.exclude_script_path.name,
        "exclude_script_exists": job_manager.exclude_script_path.is_file(),
        "web_version": _web_asset_version(),
        "database": database.health(),
    }


@app.get("/api/exclude/years")
def get_exclude_years():
    current: dict[str, dict[str, object] | None] = {}
    for year in list_exclude_years():
        info = get_year_latest_file_info(EXCLUDE_BASE / str(year))
        current[str(year)] = info

    return {
        "default_year": DEFAULT_EXCLUDE_YEAR,
        "years": list_exclude_years(),
        "current": current,
    }


@app.post("/api/upload/exclude")
async def upload_exclude_file(
    file: UploadFile = File(...),
    year: int = Query(DEFAULT_EXCLUDE_YEAR, description="剔除文件所属年份"),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="未选择文件")

    try:
        validate_exclude_year(year)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    safe_name = sanitize_filename(file.filename)
    dest_dir = exclude_year_dir(year)
    dest = dest_dir / safe_name

    content = await file.read()
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=400,
            detail=f"文件大小超过 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB 限制",
        )
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="文件为空")

    dest.write_bytes(content)
    result = {
        "year": year,
        "filename": safe_name,
        "size": len(content),
        "path": str(dest),
        "dir": str(dest_dir),
    }
    result.update(await _stage_mask_upload(dest, year))
    return result


@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    dest: str = Query("daily", description="目标页面: daily / historical / exclude"),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="未选择文件")

    safe_name = sanitize_filename(file.filename)
    # 根据 dest 路由到对应页面的 Newdata 子目录
    dest_dir = _resolve_newdata_dir(dest)
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / safe_name

    content = await file.read()
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail=f"文件大小超过 {MAX_UPLOAD_BYTES // (1024*1024)}MB 限制")
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="文件为空")

    dest_path.write_bytes(content)
    result = {
        "filename": safe_name,
        "size": len(content),
        "path": str(dest_path),
        "dest": dest,
    }
    result.update(await _stage_daily_upload(dest_path, dest))
    return result


@app.post("/api/upload/chunk")
async def upload_chunk(
    file: UploadFile = File(...),
    upload_id: str = Query(...),
    index: int = Query(..., ge=0),
    total: int = Query(..., ge=1, le=800),
    filename: str = Query(...),
    dest: str = Query("newdata"),
    year: int = Query(DEFAULT_EXCLUDE_YEAR),
):
    """Accept one small slice so FRP is not blocked by a 10MB+ POST."""
    _cleanup_stale_uploads()
    if dest not in ("newdata", "exclude", "daily", "historical", "exclude_newdata"):
        raise HTTPException(status_code=400, detail="无效的目标目录")
    if index >= total:
        raise HTTPException(status_code=400, detail="分片序号无效")
    safe_name = sanitize_filename(filename)
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="空分片")
    if len(content) > MAX_CHUNK_BYTES:
        raise HTTPException(status_code=400, detail="单片超过 4MB")
    estimated = total * MAX_CHUNK_BYTES
    if estimated > MAX_UPLOAD_BYTES + MAX_CHUNK_BYTES:
        raise HTTPException(status_code=400, detail="分片总数超过 150MB 限制")

    parts_dir = _chunk_dir(upload_id)
    part_path = parts_dir / f"{index:08d}.part"
    part_path.write_bytes(content)
    received = sum(1 for p in parts_dir.glob("*.part") if p.is_file())
    done = received >= total
    result: dict[str, object] = {
        "upload_id": upload_id,
        "index": index,
        "total": total,
        "received": received,
        "complete": done,
    }
    if not done:
        return result

    if dest == "exclude":
        try:
            validate_exclude_year(year)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        out_path = exclude_year_dir(year) / safe_name
    else:
        # 路由到对应页面的 Newdata 子目录
        target_dir = _resolve_newdata_dir(dest)
        target_dir.mkdir(parents=True, exist_ok=True)
        out_path = target_dir / safe_name

    size = _write_assembled_file(parts_dir, total, out_path)
    shutil.rmtree(parts_dir, ignore_errors=True)
    result.update(
        {
            "filename": safe_name,
            "size": size,
            "path": str(out_path),
        }
    )
    if dest == "exclude":
        result["year"] = year
        result["dir"] = str(out_path.parent)
        result.update(await _stage_mask_upload(out_path, year))
    else:
        result.update(await _stage_daily_upload(out_path, dest))
    return result


@app.get("/api/imports/{batch_id}/preview")
def import_preview(
    batch_id: str,
    start_date: date | None = Query(default=None),
    end_date: date | None = Query(default=None),
):
    try:
        return database.preview_import(batch_id, start_date, end_date)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except (ValueError, RuntimeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/imports/{batch_id}/activate")
def import_activate(batch_id: str, request: ImportActionRequest):
    try:
        preview = database.preview_import(batch_id, request.start_date, request.end_date)
        start = request.start_date or date.fromisoformat(preview["replace_start_date"])
        end = request.end_date or date.fromisoformat(preview["replace_end_date"])
        return database.activate_import(batch_id, start, end, request.confirmed_by)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except (ValueError, RuntimeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/imports/{batch_id}/reject")
def import_reject(batch_id: str):
    try:
        return database.reject_import(batch_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except (ValueError, RuntimeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/imports/{batch_id}/rollback")
def import_rollback(batch_id: str, request: ImportActionRequest):
    try:
        return database.rollback_import(batch_id, request.confirmed_by)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except (ValueError, RuntimeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/newdata/files")
def list_newdata_files(kind: str = Query("daily", description="页面标识: daily / historical / exclude")):
    """List Excel files in the page-specific Newdata subdirectory."""
    scan_dir = _resolve_newdata_dir(kind)
    scan_dir.mkdir(parents=True, exist_ok=True)
    files = []
    for path in sorted(scan_dir.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        if not path.is_file():
            continue
        if path.suffix.lower() not in ALLOWED_EXTENSIONS:
            continue
        if path.name.startswith("~$"):
            continue
        files.append(
            {
                "filename": path.name,
                "size": path.stat().st_size,
                "mtime": path.stat().st_mtime,
            }
        )
    return {"files": files, "count": len(files), "kind": kind, "dir": str(scan_dir)}


@app.post("/api/run")
def run_script(request: DailyRunRequest | None = None):
    activation_id = request.activation_id if request else None
    job, error = job_manager.create_and_start_job(activation_id=activation_id)
    if error:
        raise HTTPException(status_code=409 if job_manager.is_busy() else 400, detail=error)
    return {
        "job_id": job.id,
        "status": job.status.value,
        "kind": job.kind,
        "created_at": job.created_at,
        "expected_seconds": job_manager.expected_seconds(job.kind),
    }


@app.get("/api/historical/config")
def historical_config():
    return build_historical_config(STATS_BASE)


@app.post("/api/historical/run")
def run_historical(req: HistoricalRunRequest):
    if req.year not in HISTORICAL_YEARS:
        raise HTTPException(status_code=400, detail=f"仅支持年份 {list(HISTORICAL_YEARS)}")

    try:
        period_start, period_end = resolve_period_dates(
            req.year,
            req.start_month,
            req.start_day,
            req.end_month,
            req.end_day,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    input_path: Path | None = None
    annual_path: Path | None = None
    exclude_2025: Path | None = None
    exclude_2026: Path | None = None

    try:
        if req.year == 2025:
            annual_path = ensure_annual_2025_file(STATS_BASE)
            input_path = annual_path
            exclude_2025 = resolve_exclude_file(STATS_BASE, 2025)
            if exclude_2025 is None:
                raise FileNotFoundError("未找到 2025 年表11重大事件日剔除文件")
        else:
            from historical_utils import find_latest_newdata_file

            newdata = find_latest_newdata_file(STATS_BASE)
            if newdata is None:
                raise FileNotFoundError("Newdata 中没有可用的 Excel 文件")
            # 2026 走 Newdata 文件夹模式（不传 input），脚本自动选最新
            input_path = None
            exclude_2026 = resolve_exclude_file(STATS_BASE, 2026)
            if exclude_2026 is None:
                raise FileNotFoundError("未找到 2026 年表11重大事件日剔除文件")
            # 仍尽量准备 2025 annual 供统计表 YoY
            try:
                annual_path = ensure_annual_2025_file(STATS_BASE)
            except FileNotFoundError:
                annual_path = annual_2025_target(STATS_BASE) if annual_2025_target(STATS_BASE).is_file() else None
            exclude_2025 = resolve_exclude_file(STATS_BASE, 2025)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    job, error = job_manager.create_historical_job(
        data_year=req.year,
        period_start=period_start,
        period_end=period_end,
        input_path=input_path,
        exclude_2025_file=exclude_2025,
        exclude_2026_file=exclude_2026,
        annual_2025_path=annual_path,
    )
    if error:
        raise HTTPException(status_code=409 if job_manager.is_busy() else 400, detail=error)

    return {
        "job_id": job.id,
        "status": job.status.value,
        "kind": job.kind,
        "created_at": job.created_at,
        "expected_seconds": job_manager.expected_seconds(job.kind),
        "period_start": period_start.isoformat(),
        "period_end": period_end.isoformat(),
        "data_year": req.year,
    }


@app.get("/api/exclude/config")
def exclude_config():
    """Return available input files for the exclude page."""
    EXCLUDE_NEWDATA_DIR.mkdir(parents=True, exist_ok=True)
    files = []
    for path in sorted(EXCLUDE_NEWDATA_DIR.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        if not path.is_file():
            continue
        if path.suffix.lower() not in ALLOWED_EXTENSIONS:
            continue
        if path.name.startswith("~$"):
            continue
        files.append({
            "filename": path.name,
            "size": path.stat().st_size,
            "mtime": path.stat().st_mtime,
        })
    return {
        "files": files,
        "count": len(files),
        "exclude_script": job_manager.exclude_script_path.name,
        "exclude_script_exists": job_manager.exclude_script_path.is_file(),
    }


@app.post("/api/exclude/run")
def run_exclude(req: ExcludeRunRequest):
    """Run the exclude-external-damage script on an uploaded file."""
    safe_name = sanitize_filename(req.input_filename)
    # Look for the file in exclude Newdata dir first, then upload tmp
    input_path = EXCLUDE_NEWDATA_DIR / safe_name
    if not input_path.is_file():
        input_path = UPLOAD_TMP / safe_name
    if not input_path.is_file():
        raise HTTPException(status_code=400, detail=f"文件不存在：{safe_name}，请先上传")

    job, error = job_manager.create_exclude_job(input_path=input_path)
    if error:
        raise HTTPException(status_code=409 if job_manager.is_busy() else 400, detail=error)
    return {
        "job_id": job.id,
        "status": job.status.value,
        "kind": job.kind,
        "created_at": job.created_at,
        "expected_seconds": job_manager.expected_seconds(job.kind),
    }


def _job_payload(job, *, include_logs: bool = False) -> dict:
    duration = job_manager.job_duration_seconds(job)
    payload = {
        "job_id": job.id,
        "kind": job.kind,
        "status": job.status.value,
        "exit_code": job.exit_code,
        "summary_simple": job.result.summary_simple,
        "summary_full": job.result.summary_full,
        "files": job.result.output_files,
        "created_at": job.created_at,
        "finished_at": job.finished_at,
        "duration_seconds": duration,
        "expected_seconds": job_manager.expected_seconds(job.kind),
        "busy": job_manager.is_busy(),
    }
    if include_logs:
        payload["logs"] = job.get_log_text()
    return payload


def _persisted_job_payload(row: dict, *, include_logs: bool = False) -> dict:
    outputs = row.get("output_files") or {}

    def read_summary(key: str) -> str:
        path_text = outputs.get(key)
        if not path_text:
            return ""
        path = validate_output_path(path_text, OUTPUT_BASE)
        if not path or not path.is_file() or path.suffix.lower() != ".txt":
            return ""
        try:
            return path.read_text(encoding="utf-8")
        except OSError:
            return ""

    parameters = row.get("parameters") or {}
    payload = {
        "job_id": row["job_id"],
        "status": row["status"],
        "kind": parameters.get("kind", "daily"),
        "exit_code": 0 if row["status"] == "success" else 1,
        "summary_simple": read_summary("summary_simple"),
        "summary_full": read_summary("summary_full"),
        "files": outputs,
        "created_at": row["created_at"].timestamp() if row.get("created_at") else None,
        "finished_at": row["finished_at"].timestamp() if row.get("finished_at") else None,
        "duration_seconds": float(row["elapsed_seconds"]) if row.get("elapsed_seconds") is not None else None,
        "expected_seconds": job_manager.expected_seconds(parameters.get("kind", "daily")),
        "busy": job_manager.is_busy(),
        "persisted": True,
    }
    if include_logs:
        payload["logs"] = row.get("error_message") or ""
    return payload


@app.get("/api/timing")
def get_timing():
    """Last successful run durations used for progress ETA."""
    return job_manager.get_timing()


@app.get("/api/jobs/latest")
def get_latest_job(
    include_logs: bool = Query(False),
    kind: str | None = Query(None, description="daily 或 historical；不传则返回全局最近任务"),
):
    """Return the running or most recent job for cross-page state restore."""
    if kind is not None and kind not in ("daily", "historical", "exclude"):
        raise HTTPException(status_code=400, detail="kind 仅支持 daily、historical 或 exclude")
    job = job_manager.get_latest_job(kind=kind)
    running = job_manager.get_latest_job()
    persisted = None
    if job is None:
        try:
            persisted = database.get_latest_report_run(kind)
        except Exception:
            persisted = None
    return {
        "job": (
            _job_payload(job, include_logs=include_logs)
            if job else _persisted_job_payload(persisted, include_logs=include_logs)
            if persisted else None
        ),
        "busy": job_manager.is_busy(),
        "running_kind": running.kind if running and running.status.value in ("running", "pending") else None,
    }


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str, include_logs: bool = Query(False)):
    job = job_manager.get_job(job_id)
    if not job:
        try:
            persisted = database.get_report_run(job_id)
        except Exception:
            persisted = None
        if not persisted:
            raise HTTPException(status_code=404, detail="任务不存在")
        return _persisted_job_payload(persisted, include_logs=include_logs)
    return _job_payload(job, include_logs=include_logs)


@app.get("/api/jobs/{job_id}/stream")
async def stream_job_logs(job_id: str):
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在")

    return StreamingResponse(
        job_manager.stream_logs(job_id),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@app.post("/api/jobs/{job_id}/stop")
def stop_job(job_id: str):
    """Terminate a specific running job process."""
    job, error = job_manager.stop_job(job_id)
    if error and job is None:
        raise HTTPException(status_code=404 if "不存在" in error else 400, detail=error)
    if error and job is not None and "无法停止" in error:
        raise HTTPException(status_code=409, detail=error)
    if error and job is not None and "终止进程" in error:
        raise HTTPException(status_code=500, detail=error)
    return _job_payload(job)


@app.post("/api/stop")
def stop_current_job():
    """Terminate the currently running job (any page can call this)."""
    job, error = job_manager.stop_job(None)
    if error and job is None:
        raise HTTPException(status_code=409, detail=error)
    if error and job is not None and "无法停止" in error:
        raise HTTPException(status_code=409, detail=error)
    if error and job is not None and "终止进程" in error:
        raise HTTPException(status_code=500, detail=error)
    return _job_payload(job)


@app.get("/api/download/{job_id}/{file_key}")
def download_file(job_id: str, file_key: str):
    if file_key not in ("result", "publish", "stats", "summary_simple", "summary_full", "exclude_output", "all"):
        raise HTTPException(status_code=400, detail="无效的文件类型")

    job = job_manager.get_job(job_id)
    persisted = None
    if not job:
        try:
            persisted = database.get_report_run(job_id)
        except Exception:
            persisted = None
        if not persisted:
            raise HTTPException(status_code=404, detail="任务不存在")
    status = job.status.value if job else persisted["status"]
    if status != JobStatus.SUCCESS.value:
        raise HTTPException(status_code=400, detail="任务未成功完成，无法下载")

    output_files = job.result.output_files if job else (persisted.get("output_files") or {})

    if file_key == "all":
        return _download_all_zip(job_id, output_files)

    file_path_str = output_files.get(file_key)
    if not file_path_str:
        raise HTTPException(status_code=404, detail="未找到输出文件路径")

    validated = validate_output_path(file_path_str, OUTPUT_BASE)
    if not validated:
        raise HTTPException(status_code=403, detail="文件路径无效")

    media = (
        "text/plain; charset=utf-8"
        if validated.suffix.lower() == ".txt"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return FileResponse(
        path=str(validated),
        filename=validated.name,
        media_type=media,
    )


def _download_all_zip(job_id: str, output_files: dict[str, str]) -> FileResponse:
    files: list[tuple[str, Path]] = []
    for key in ("result", "publish", "stats", "summary_simple", "summary_full"):
        path_str = output_files.get(key)
        if not path_str:
            continue
        validated = validate_output_path(path_str, OUTPUT_BASE)
        if validated and validated.is_file():
            files.append((validated.name, validated))

    if not files:
        raise HTTPException(status_code=404, detail="没有可打包的输出文件")

    download_tmp = STATS_BASE / ".download_tmp"
    download_tmp.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        prefix=f"outage-stats-{job_id[:8]}-",
        suffix=".zip",
        dir=download_tmp,
        delete=False,
    )
    zip_path = Path(handle.name)
    handle.close()
    try:
        # XLSX files are ZIP containers already; storing them avoids expensive
        # recompression. FileResponse streams from disk and supports ranges.
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_STORED) as zf:
            for name, path in files:
                zf.write(path, arcname=name)
    except BaseException:
        zip_path.unlink(missing_ok=True)
        raise

    zip_name = f"outage-stats-{job_id[:8]}.zip"
    return FileResponse(
        path=str(zip_path),
        filename=zip_name,
        media_type="application/zip",
        background=BackgroundTask(zip_path.unlink, missing_ok=True),
    )


if FRONTEND_DIR.is_dir():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")

