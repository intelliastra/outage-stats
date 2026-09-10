"""Aliyun-local upload receiver.

Browser uploads stay on Aliyun. NAS then downloads the assembled file over
the public Internet (not Tailscale/DERP). SSH is used only to start curl.
"""

from __future__ import annotations

import asyncio
import base64
import json
import os
import re
import secrets
import shlex
import shutil
import subprocess
import threading
import time
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

MAX_UPLOAD_BYTES = 150 * 1024 * 1024
MAX_CHUNK_BYTES = 4 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MIN_EXCLUDE_YEAR = 2024
MAX_EXCLUDE_YEAR = 2030
DEFAULT_EXCLUDE_YEAR = 2026

UPLOAD_TMP = Path(os.environ.get("UPLOAD_TMP", "/var/lib/outage-upload")).resolve()
NAS_USER = os.environ.get("NAS_SFTP_USER", "zhitu")
NAS_KEY = os.environ.get("NAS_SFTP_KEY", "/root/.ssh/outage_nas_ed25519")
NAS_NEWDATA = os.environ.get(
    "NAS_NEWDATA",
    "/volume1/docker/outage-stats/统计材料/input/Newdata",
)
NAS_EXCLUDE = os.environ.get(
    "NAS_EXCLUDE",
    "/volume1/docker/outage-stats/统计材料/input/exclude",
)
PULL_BASE_URL = os.environ.get(
    "PULL_BASE_URL",
    "https://outage-stats.demo.intelliastra.com",
)
PULL_RESOLVE = os.environ.get(
    "PULL_RESOLVE",
    "outage-stats.demo.intelliastra.com:443:47.121.130.229",
)
NAS_TARGETS = [
    (
        os.environ.get("NAS_SFTP_HOST", "100.72.193.8"),
        int(os.environ.get("NAS_SFTP_PORT", "6330")),
    ),
]

app = FastAPI(title="outage-upload-gateway", version="1.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

_lock = threading.Lock()
_pulls: dict[str, dict[str, object]] = {}


def sanitize_filename(name: str) -> str:
    base = Path(name).name
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base)
    if not base or base in (".", ".."):
        raise HTTPException(status_code=400, detail="无效的文件名")
    if Path(base).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="仅支持 .xlsx / .xls 文件")
    return base


def validate_exclude_year(year: int) -> int:
    if year < MIN_EXCLUDE_YEAR or year > MAX_EXCLUDE_YEAR:
        raise HTTPException(
            status_code=400,
            detail=f"年份须在 {MIN_EXCLUDE_YEAR}–{MAX_EXCLUDE_YEAR} 之间",
        )
    return year


def _validate_upload_id(upload_id: str) -> str:
    if not re.fullmatch(r"[0-9a-fA-F-]{8,64}", upload_id or ""):
        raise HTTPException(status_code=400, detail="无效的 upload_id")
    return upload_id


def _chunk_dir(upload_id: str) -> Path:
    root = UPLOAD_TMP
    root.mkdir(parents=True, exist_ok=True)
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
    with _lock:
        expired = [k for k, v in _pulls.items() if now > float(v["expires"])]
        for k in expired:
            _pulls.pop(k, None)


def _write_assembled_file(parts_dir: Path, total: int, dest: Path) -> int:
    missing = [i for i in range(total) if not (parts_dir / f"{i:08d}.part").is_file()]
    if missing:
        raise HTTPException(status_code=400, detail=f"分片不完整，缺少 {missing[:8]}")
    size = 0
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".assembling")
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


def _ssh_base(host: str, port: int) -> list[str]:
    return [
        "-i",
        NAS_KEY,
        "-o",
        "BatchMode=yes",
        "-o",
        "StrictHostKeyChecking=accept-new",
        "-o",
        "ConnectTimeout=20",
        "-o",
        "ServerAliveInterval=15",
        "-p",
        str(port),
        f"{NAS_USER}@{host}",
    ]


def _run(cmd: list[str], timeout: int) -> str:
    proc = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout,
        check=False,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()[:500]
        raise RuntimeError(err or f"exit {proc.returncode}")
    return (proc.stdout or "").strip()


def nas_newdata_path(filename: str) -> str:
    return f"{NAS_NEWDATA.rstrip('/')}/{filename}"


def nas_exclude_path(year: int, filename: str) -> str:
    return f"{NAS_EXCLUDE.rstrip('/')}/{year}/{filename}"


def push_via_nas_pull(local: Path, remote_path: str) -> str:
    """NAS curls the file from Aliyun public HTTPS. Bulk bytes skip Tailscale."""
    token = secrets.token_hex(16)
    with _lock:
        _pulls[token] = {
            "path": str(local.resolve()),
            "filename": local.name,
            "expires": time.time() + 900,
        }
    url = f"{PULL_BASE_URL.rstrip('/')}/api/upload/pull/{token}"
    staging = f"/tmp/outage_pull_{os.getpid()}_{int(time.time())}.bin"
    payload = base64.b64encode(
        json.dumps(
            {
                "url": url,
                "resolve": PULL_RESOLVE,
                "src": staging,
                "dst": remote_path,
            },
            ensure_ascii=False,
        ).encode("utf-8")
    ).decode("ascii")
    py = (
        "import base64,json,os,subprocess,sys;"
        f"d=json.loads(base64.b64decode('{payload}'));"
        "os.makedirs(os.path.dirname(d['dst']), exist_ok=True);"
        "tmp=d['dst']+'.downloading';"
        "cmd=['curl','-fsSL','--max-time','90','-k','--resolve',d['resolve'],"
        "'-o',tmp,d['url']];"
        "r=subprocess.run(cmd,capture_output=True,text=True);"
        "sys.stderr.write(r.stderr or '');"
        "r.check_returncode();"
        "sz=os.path.getsize(tmp);"
        "assert sz>0, 'empty download';"
        "os.replace(tmp, d['dst']);"
        "print(sz)"
    )
    t0 = time.time()
    last = "unknown"
    for host, port in NAS_TARGETS:
        try:
            ssh = ["ssh"] + _ssh_base(host, port)
            out = _run(ssh[:-1] + [ssh[-1], "python3 -c " + shlex.quote(py)], timeout=120)
            dt = time.time() - t0
            print(f"nas-pull ok via {host}:{port} {dt:.1f}s size={out}", flush=True)
            with _lock:
                _pulls.pop(token, None)
            return f"nas-pull:{host}:{port}"
        except Exception as exc:  # noqa: BLE001
            last = f"{host}:{port} {exc}"
            print(f"nas-pull fail {last}", flush=True)
    with _lock:
        _pulls.pop(token, None)
    raise HTTPException(status_code=503, detail="写入服务器失败，请稍后重试")


async def _push_assembled(local: Path, remote: str) -> str:
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, push_via_nas_pull, local, remote)


def _require_local(request: Request) -> None:
    if request.headers.get("x-forwarded-for"):
        raise HTTPException(status_code=403, detail="forbidden")
    host = request.client.host if request.client else ""
    if host not in ("127.0.0.1", "::1"):
        raise HTTPException(status_code=403, detail="forbidden")


@app.post("/api/upload/internal-push")
async def internal_push(
    request: Request,
    path: str = Query(...),
    dest: str = Query("newdata"),
    year: int = Query(DEFAULT_EXCLUDE_YEAR),
):
    """Localhost-only: push an already-assembled file to NAS via public pull."""
    _require_local(request)
    src = Path(path).resolve()
    if UPLOAD_TMP not in src.parents and src.parent != UPLOAD_TMP:
        raise HTTPException(status_code=400, detail="路径不允许")
    if not src.is_file():
        raise HTTPException(status_code=400, detail="文件不存在")
    safe_name = sanitize_filename(src.name)
    if dest == "exclude":
        validate_exclude_year(year)
        remote = nas_exclude_path(year, safe_name)
    else:
        remote = nas_newdata_path(safe_name)
    via = await _push_assembled(src, remote)
    return {"filename": safe_name, "size": src.stat().st_size, "path": remote, "via": via}


@app.get("/api/upload/gateway-health")
def gateway_health():
    return {
        "status": "ok",
        "via": "aliyun-local",
        "push": "nas-pull",
        "key_exists": Path(NAS_KEY).is_file(),
        "tmp": str(UPLOAD_TMP),
    }


@app.get("/api/upload/pull/{token}")
def pull_file(token: str, request: Request):
    if not re.fullmatch(r"[0-9a-f]{32}", token or ""):
        raise HTTPException(status_code=404, detail="not found")
    with _lock:
        rec = _pulls.get(token)
        if not rec or time.time() > float(rec["expires"]):
            raise HTTPException(status_code=404, detail="not found")
        path = Path(str(rec["path"]))
        filename = str(rec["filename"])
    if not path.is_file():
        raise HTTPException(status_code=404, detail="not found")
    print(f"pull GET token={token[:8]} from {request.client.host if request.client else '?'}", flush=True)
    return FileResponse(path, filename=filename, media_type="application/octet-stream")


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="未选择文件")
    safe_name = sanitize_filename(file.filename)
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="文件为空")
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail="文件大小超过 150MB 限制")
    UPLOAD_TMP.mkdir(parents=True, exist_ok=True)
    tmp = UPLOAD_TMP / f"whole_{os.getpid()}_{safe_name}"
    tmp.write_bytes(content)
    try:
        remote = nas_newdata_path(safe_name)
        via = await _push_assembled(tmp, remote)
    finally:
        tmp.unlink(missing_ok=True)
    return {"filename": safe_name, "size": len(content), "path": remote, "via": via}


@app.post("/api/upload/exclude")
async def upload_exclude_file(
    file: UploadFile = File(...),
    year: int = Query(DEFAULT_EXCLUDE_YEAR),
):
    if not file.filename:
        raise HTTPException(status_code=400, detail="未选择文件")
    validate_exclude_year(year)
    safe_name = sanitize_filename(file.filename)
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="文件为空")
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail="文件大小超过 150MB 限制")
    UPLOAD_TMP.mkdir(parents=True, exist_ok=True)
    tmp = UPLOAD_TMP / f"excl_{os.getpid()}_{safe_name}"
    tmp.write_bytes(content)
    try:
        remote = nas_exclude_path(year, safe_name)
        via = await _push_assembled(tmp, remote)
    finally:
        tmp.unlink(missing_ok=True)
    return {
        "year": year,
        "filename": safe_name,
        "size": len(content),
        "path": remote,
        "via": via,
    }


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
    _cleanup_stale_uploads()
    if dest not in ("newdata", "exclude"):
        raise HTTPException(status_code=400, detail="无效的目标目录")
    if index >= total:
        raise HTTPException(status_code=400, detail="分片序号无效")
    safe_name = sanitize_filename(filename)
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="空分片")
    if len(content) > MAX_CHUNK_BYTES:
        raise HTTPException(status_code=400, detail="单片超过 4MB")
    if total * MAX_CHUNK_BYTES > MAX_UPLOAD_BYTES + MAX_CHUNK_BYTES:
        raise HTTPException(status_code=400, detail="分片总数超过 150MB 限制")

    parts_dir = _chunk_dir(upload_id)
    (parts_dir / f"{index:08d}.part").write_bytes(content)
    received = sum(1 for p in parts_dir.glob("*.part") if p.is_file())
    result: dict[str, object] = {
        "upload_id": upload_id,
        "index": index,
        "total": total,
        "received": received,
        "complete": received >= total,
        "via": "aliyun-local",
    }
    if received < total:
        return result

    assembled = parts_dir / safe_name
    size = _write_assembled_file(parts_dir, total, assembled)
    if dest == "exclude":
        validate_exclude_year(year)
        remote = nas_exclude_path(year, safe_name)
    else:
        remote = nas_newdata_path(safe_name)
    try:
        via = await _push_assembled(assembled, remote)
    except HTTPException:
        assembled.unlink(missing_ok=True)
        raise
    shutil.rmtree(parts_dir, ignore_errors=True)
    result.update(
        {
            "filename": safe_name,
            "size": size,
            "path": remote,
            "via": via,
            "complete": True,
        }
    )
    if dest == "exclude":
        result["year"] = year
    return result
