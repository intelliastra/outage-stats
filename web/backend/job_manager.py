"""In-memory job manager with SSE log streaming and stop support."""

from __future__ import annotations

import asyncio
import json
import os
import re
import signal
import subprocess
import sys
import threading
import time
import uuid
from collections import deque
from contextlib import nullcontext
from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from pathlib import Path
from typing import AsyncIterator

from log_parser import ParsedResult, parse_exclude_output, parse_script_output

import database

DEFAULT_EXPECTED_SECONDS = 1800  # 30 minutes fallback for ETA
_FILENAME_VERSION_RE = re.compile(r"_v(\d+\.\d+\.\d+)", re.I)
_DOC_VERSION_RE = re.compile(r"^Version:\s*v?(\d+\.\d+\.\d+)", re.M)


def parse_script_version(script_path: Path) -> str:
    """Read analysis-script version from filename, then docstring."""
    match = _FILENAME_VERSION_RE.search(script_path.name)
    if match:
        return match.group(1)
    try:
        head = script_path.read_text(encoding="utf-8", errors="replace")[:4000]
    except OSError:
        return ""
    match = _DOC_VERSION_RE.search(head)
    return match.group(1) if match else ""


class JobStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class Job:
    id: str
    status: JobStatus = JobStatus.PENDING
    logs: deque[str] = field(default_factory=lambda: deque(maxlen=50000))
    log_lock: threading.Lock = field(default_factory=threading.Lock)
    exit_code: int | None = None
    result: ParsedResult = field(default_factory=ParsedResult)
    created_at: float = field(default_factory=time.time)
    finished_at: float | None = None
    kind: str = "daily"
    process: subprocess.Popen | None = field(default=None, repr=False)
    stop_requested: bool = False
    parameters: dict = field(default_factory=dict)
    _subscribers: list[asyncio.Queue] = field(default_factory=list)
    _subscribers_lock: threading.Lock = field(default_factory=threading.Lock)

    def append_log(self, line: str) -> None:
        with self.log_lock:
            self.logs.append(line)
        with self._subscribers_lock:
            for queue in self._subscribers:
                try:
                    queue.put_nowait(line)
                except asyncio.QueueFull:
                    pass

    def subscribe(self) -> asyncio.Queue:
        queue: asyncio.Queue = asyncio.Queue(maxsize=1000)
        with self._subscribers_lock:
            self._subscribers.append(queue)
        return queue

    def unsubscribe(self, queue: asyncio.Queue) -> None:
        with self._subscribers_lock:
            if queue in self._subscribers:
                self._subscribers.remove(queue)

    def get_log_text(self) -> str:
        with self.log_lock:
            return "".join(self.logs)


class JobManager:
    JOB_TTL_SECONDS = 1800  # 30 minutes

    def __init__(self, stats_base: Path, script_name: str = "process_outage_tables_v1.0.8.py") -> None:
        self.stats_base = stats_base.resolve()
        # Scripts live under 统计材料/运行脚本/
        script_dir = self.stats_base / "运行脚本"

        # --- 日常统计脚本 ---
        self.script_path = script_dir / script_name
        if not self.script_path.is_file():
            for fallback_name in (
                "process_outage_tables_v1.0.7.py",
                "process_outage_tables_v1.0.6.py",
                "process_outage_tables_v1.0.5.py",
            ):
                fallback = script_dir / fallback_name
                if fallback.is_file():
                    self.script_path = fallback
                    break
        if not self.script_path.is_file():
            self.script_path = self.stats_base / script_name
        self.script_version = parse_script_version(self.script_path)

        # --- 历史区间统计脚本（基于 v1.0.8 独立副本） ---
        self.historical_script_path = script_dir / "process_outage_tables_historical_v1.0.0.py"
        if not self.historical_script_path.is_file():
            for fb in ("process_outage_tables_historical_v1.0.0.py",):
                fallback = script_dir / fb
                if fallback.is_file():
                    self.historical_script_path = fallback
                    break
        if not self.historical_script_path.is_file():
            self.historical_script_path = self.script_path  # fallback to daily

        # --- 剔除外力破坏脚本 ---
        self.exclude_script_path = script_dir / "exclude_external_damage_v1.0.0.py"

        # --- 三个页面独立的输入/输出目录 ---
        newdata_base = self.stats_base / "input" / "Newdata"
        output_root = self.stats_base / "output"

        self.daily_newdata_dir = newdata_base / "daily"
        self.daily_output_base = output_root / "daily"

        self.historical_newdata_dir = newdata_base / "historical"
        self.historical_output_base = output_root / "historical"

        self.exclude_newdata_dir = newdata_base / "exclude"
        self.exclude_output_base = output_root / "exclude"

        # 兼容旧属性名（日常统计）
        self.newdata_dir = self.daily_newdata_dir
        self.output_base = self.daily_output_base

        self.timing_file = self.stats_base / ".last_run_durations.json"
        self.jobs: dict[str, Job] = {}
        self._lock = threading.Lock()
        self._running_job_id: str | None = None
        self._last_durations: dict[str, float] = self._load_durations()

    def _load_durations(self) -> dict[str, float]:
        try:
            if self.timing_file.is_file():
                data = json.loads(self.timing_file.read_text(encoding="utf-8"))
                return {
                    "daily": float(data.get("daily") or 0) or 0.0,
                    "historical": float(data.get("historical") or 0) or 0.0,
                    "exclude": float(data.get("exclude") or 0) or 0.0,
                }
        except Exception:
            pass
        return {"daily": 0.0, "historical": 0.0, "exclude": 0.0}

    def _save_duration(self, kind: str, seconds: float) -> None:
        if seconds <= 0:
            return
        self._last_durations[kind] = seconds
        try:
            self.timing_file.write_text(
                json.dumps(self._last_durations, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def get_timing(self) -> dict:
        return {
            "daily": self._last_durations.get("daily") or None,
            "historical": self._last_durations.get("historical") or None,
            "exclude": self._last_durations.get("exclude") or None,
            "default_expected_seconds": DEFAULT_EXPECTED_SECONDS,
        }

    def expected_seconds(self, kind: str) -> float:
        saved = self._last_durations.get(kind) or 0
        return float(saved) if saved > 0 else float(DEFAULT_EXPECTED_SECONDS)

    def job_duration_seconds(self, job: Job) -> float | None:
        if job.finished_at is None:
            return max(0.0, time.time() - job.created_at)
        return max(0.0, job.finished_at - job.created_at)

    def _cleanup_old_jobs(self) -> None:
        now = time.time()
        expired = [
            jid
            for jid, job in self.jobs.items()
            if job.finished_at and now - job.finished_at > self.JOB_TTL_SECONDS
        ]
        for jid in expired:
            if jid != self._running_job_id:
                self.jobs.pop(jid, None)

    def get_job(self, job_id: str) -> Job | None:
        return self.jobs.get(job_id)

    def get_latest_job(self, kind: str | None = None) -> Job | None:
        """Return latest job, optionally filtered by kind (daily / historical)."""
        jobs = list(self.jobs.values())
        if kind:
            jobs = [job for job in jobs if job.kind == kind]
        if not jobs:
            return None

        if self._running_job_id and self._running_job_id in self.jobs:
            running = self.jobs[self._running_job_id]
            if kind is None or running.kind == kind:
                return running

        return max(jobs, key=lambda job: job.created_at)

    def _fill_missing_outputs(self, result: ParsedResult, kind: str = "daily") -> None:
        if all(
            key in result.output_files
            for key in ("result", "publish", "stats", "summary_simple", "summary_full")
        ):
            return
        # 按 kind 选择对应的输出根目录
        if kind == "historical":
            out_base = self.historical_output_base
        elif kind == "exclude":
            out_base = self.exclude_output_base
        else:
            out_base = self.daily_output_base
        if not out_base.is_dir():
            return

        subdirs = [path for path in out_base.iterdir() if path.is_dir()]
        if not subdirs:
            return

        latest = max(subdirs, key=lambda path: path.stat().st_mtime)
        discovered: dict[str, str] = {}
        for file_path in latest.iterdir():
            if not file_path.is_file() or file_path.stat().st_size == 0:
                continue
            name = file_path.name
            if name.endswith("统计表.xlsx"):
                discovered["stats"] = str(file_path)
            elif name.endswith("停电摘要（简版）.txt"):
                discovered["summary_simple"] = str(file_path)
            elif name.endswith("停电摘要（全量版）.txt"):
                discovered["summary_full"] = str(file_path)
            elif "_发出版" in name:
                discovered["publish"] = str(file_path)
            elif "处理结果" in name:
                discovered["result"] = str(file_path)

        for key, path in discovered.items():
            result.output_files.setdefault(key, path)

    def is_busy(self) -> bool:
        return self._running_job_id is not None

    def create_and_start_job(self, *, activation_id: str | None = None) -> tuple[Job | None, str | None]:
        with self._lock:
            self._cleanup_old_jobs()
            if self._running_job_id is not None:
                return None, "已有任务正在运行，请等待完成后再试"

            if not self.script_path.is_file():
                return None, f"脚本不存在：{self.script_path}"

            if database.DATA_BACKEND == "postgres":
                try:
                    if not (activation_id or database.latest_activation_id()):
                        return None, "数据库中没有已激活的数据批次，请先确认并激活上传批次"
                except Exception as exc:
                    return None, f"无法读取数据库激活版本：{exc}"
            else:
                excel_files = list(self.newdata_dir.glob("*.xlsx")) + list(self.newdata_dir.glob("*.xls"))
                if not excel_files:
                    return None, "Newdata 文件夹中没有 Excel 文件，请先上传"

            job_id = str(uuid.uuid4())
            try:
                effective_activation = activation_id or database.latest_activation_id()
            except Exception as exc:
                if database.DATA_BACKEND == "postgres":
                    return None, f"无法读取数据库激活版本：{exc}"
                effective_activation = activation_id
            job = Job(
                id=job_id,
                kind="daily",
                parameters={"activation_id": effective_activation, "data_backend": database.DATA_BACKEND},
            )
            self.jobs[job_id] = job
            self._running_job_id = job_id
            cmd = [sys.executable, "-u", str(self.script_path), "--non-interactive"]

            try:
                database.create_report_run(job_id, job.parameters)
            except Exception as exc:
                if database.DATA_BACKEND == "postgres":
                    self.jobs.pop(job_id, None)
                    self._running_job_id = None
                    return None, f"无法创建数据库任务记录：{exc}"
                job.append_log(f"[DB] 影子任务记录失败，继续使用 Excel：{exc}\n")

        thread = threading.Thread(target=self._run_script, args=(job, cmd), daemon=True)
        thread.start()
        return job, None

    def create_historical_job(
        self,
        *,
        data_year: int,
        period_start: date,
        period_end: date,
        input_path: Path | None,
        exclude_2025_file: Path | None,
        exclude_2026_file: Path | None,
        annual_2025_path: Path | None,
    ) -> tuple[Job | None, str | None]:
        with self._lock:
            self._cleanup_old_jobs()
            if self._running_job_id is not None:
                return None, "已有任务正在运行，请等待完成后再试"

            if not self.historical_script_path.is_file():
                return None, f"脚本不存在：{self.historical_script_path}"

            job_id = str(uuid.uuid4())
            job = Job(id=job_id, kind="historical")
            self.jobs[job_id] = job
            self._running_job_id = job_id

            cmd = [
                sys.executable,
                "-u",
                str(self.historical_script_path),
                "--non-interactive",
                "--period-start",
                period_start.isoformat(),
                "--period-end",
                period_end.isoformat(),
                "--data-year",
                str(data_year),
            ]
            if input_path is not None:
                cmd.append(str(input_path))
            if annual_2025_path is not None:
                cmd.extend(["--annual-2025-path", str(annual_2025_path)])
            if exclude_2025_file is not None:
                cmd.extend(["--exclude-2025-file", str(exclude_2025_file)])
            if exclude_2026_file is not None:
                cmd.extend(["--exclude-2026-file", str(exclude_2026_file)])

        thread = threading.Thread(target=self._run_script, args=(job, cmd), daemon=True)
        thread.start()
        return job, None

    def create_exclude_job(self, input_path: Path) -> tuple[Job | None, str | None]:
        """Create a job to run the exclude-external-damage script."""
        with self._lock:
            self._cleanup_old_jobs()
            if self._running_job_id is not None:
                return None, "已有任务正在运行，请等待完成后再试"

            if not self.exclude_script_path.is_file():
                return None, f"剔除脚本不存在：{self.exclude_script_path}"

            if not input_path.is_file():
                return None, f"输入文件不存在：{input_path}"

            job_id = str(uuid.uuid4())
            job = Job(id=job_id, kind="exclude")
            self.jobs[job_id] = job
            self._running_job_id = job_id
            cmd = [
                sys.executable,
                "-u",
                str(self.exclude_script_path),
                str(input_path),
                "--non-interactive",
            ]

        thread = threading.Thread(target=self._run_script, args=(job, cmd), daemon=True)
        thread.start()
        return job, None

    def stop_job(self, job_id: str | None = None) -> tuple[Job | None, str | None]:
        """Terminate a running job process. If job_id is None, stop the current running job."""
        with self._lock:
            target_id = job_id or self._running_job_id
            if not target_id:
                return None, "当前没有正在运行的任务"
            job = self.jobs.get(target_id)
            if not job:
                return None, "任务不存在"
            if job.status != JobStatus.RUNNING:
                return None, f"任务当前状态为 {job.status.value}，无法停止"
            job.stop_requested = True
            proc = job.process

        if proc is None or proc.poll() is not None:
            return job, None

        job.append_log("\n[WEB] 收到停止请求，正在终止脚本进程...\n")
        try:
            if sys.platform == "win32":
                proc.terminate()
            else:
                # Kill process group if possible
                try:
                    os.killpg(os.getpgid(proc.pid), signal.SIGTERM)
                except (ProcessLookupError, PermissionError, AttributeError):
                    proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                if sys.platform == "win32":
                    proc.kill()
                else:
                    try:
                        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
                    except (ProcessLookupError, PermissionError, AttributeError):
                        proc.kill()
                proc.wait(timeout=3)
        except Exception as exc:
            job.append_log(f"[WEB] 终止进程时出错：{exc}\n")
            return job, f"终止进程时出错：{exc}"

        job.append_log("[WEB] 脚本进程已终止。\n")
        return job, None

    def _run_script(self, job: Job, cmd: list[str]) -> None:
        job.status = JobStatus.RUNNING
        job.append_log(f"[WEB] 启动命令：{' '.join(cmd)}\n")
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        if job.parameters.get("activation_id"):
            env["REPORT_ACTIVATION_ID"] = str(job.parameters["activation_id"])

        popen_kwargs: dict = {
            "cwd": str(self.stats_base),
            "stdout": subprocess.PIPE,
            "stderr": subprocess.STDOUT,
            "env": env,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "bufsize": 1,
        }
        if sys.platform != "win32":
            popen_kwargs["start_new_session"] = True

        try:
            lock_context = nullcontext(True)
            if database.enabled():
                state = database.health()
                if state.get("reachable"):
                    lock_context = database.report_advisory_lock()
                elif database.DATA_BACKEND == "postgres":
                    raise RuntimeError(f"数据库不可用：{state.get('error', '连接失败')}")
                else:
                    job.append_log("[DB] 影子数据库不可用，本次仍由进程内锁保护。\n")

            try:
                database.update_report_run(job.id, "running")
            except Exception as exc:
                if database.DATA_BACKEND == "postgres":
                    raise
                job.append_log(f"[DB] 影子任务状态写入失败：{exc}\n")

            with lock_context as acquired:
                if not acquired:
                    raise RuntimeError("数据库中已有统计任务正在运行")
                proc = subprocess.Popen(cmd, **popen_kwargs)
                job.process = proc
                assert proc.stdout is not None
                for line in proc.stdout:
                    job.append_log(line)
                proc.wait()
                job.exit_code = proc.returncode
                log_text = job.get_log_text()
                if job.kind == "exclude":
                    job.result = parse_exclude_output(log_text)
                else:
                    job.result = parse_script_output(log_text)
                if job.stop_requested:
                    job.status = JobStatus.CANCELLED
                    if job.exit_code is None or job.exit_code == 0:
                        job.exit_code = -15
                elif proc.returncode == 0:
                    self._fill_missing_outputs(job.result, job.kind)
                    job.status = JobStatus.SUCCESS
                else:
                    job.status = JobStatus.FAILED
        except Exception as exc:
            job.append_log(f"\n[WEB] 启动脚本失败：{exc}\n")
            job.exit_code = 1
            job.status = JobStatus.CANCELLED if job.stop_requested else JobStatus.FAILED
        finally:
            job.process = None
            job.finished_at = time.time()
            if job.status == JobStatus.SUCCESS:
                duration = self.job_duration_seconds(job) or 0.0
                self._save_duration(job.kind, duration)
            try:
                database.update_report_run(
                    job.id,
                    job.status.value,
                    output_files=job.result.output_files,
                    error_message=None if job.status == JobStatus.SUCCESS else job.get_log_text()[-2000:],
                    elapsed_seconds=self.job_duration_seconds(job),
                )
            except Exception as exc:
                job.append_log(f"[DB] 任务终态写入失败：{exc}\n")
            with self._lock:
                if self._running_job_id == job.id:
                    self._running_job_id = None
            with job._subscribers_lock:
                for queue in job._subscribers:
                    try:
                        queue.put_nowait(None)
                    except asyncio.QueueFull:
                        pass

    async def stream_logs(self, job_id: str) -> AsyncIterator[str]:
        job = self.get_job(job_id)
        if not job:
            yield "event: error\ndata: {\"message\":\"任务不存在\"}\n\n"
            return

        with job.log_lock:
            for line in job.logs:
                yield f"data: {self._sse_data(line)}\n\n"

        if job.status in (JobStatus.SUCCESS, JobStatus.FAILED, JobStatus.CANCELLED):
            yield self._done_event(job)
            return

        queue = job.subscribe()
        try:
            while True:
                try:
                    line = await asyncio.wait_for(queue.get(), timeout=30.0)
                except asyncio.TimeoutError:
                    yield ": keepalive\n\n"
                    continue
                if line is None:
                    break
                yield f"data: {self._sse_data(line)}\n\n"
            yield self._done_event(job)
        finally:
            job.unsubscribe(queue)

    @staticmethod
    def _sse_data(text: str) -> str:
        import json

        return json.dumps({"line": text}, ensure_ascii=False)

    def _done_event(self, job: Job) -> str:
        duration = self.job_duration_seconds(job)
        payload = {
            "status": job.status.value,
            "exit_code": job.exit_code,
            "summary_simple": job.result.summary_simple,
            "summary_full": job.result.summary_full,
            "files": job.result.output_files,
            "kind": job.kind,
            "created_at": job.created_at,
            "finished_at": job.finished_at,
            "duration_seconds": duration,
            "expected_seconds": self.expected_seconds(job.kind),
        }
        return f"event: done\ndata: {json.dumps(payload, ensure_ascii=False)}\n\n"
