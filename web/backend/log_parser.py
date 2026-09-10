"""Parse script stdout for summaries and output file paths."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ParsedResult:
    summary_simple: str = ""
    summary_full: str = ""
    output_files: dict[str, str] = field(default_factory=dict)


SUMMARY_SIMPLE_TITLE = "停电摘要（简版）"
SUMMARY_FULL_TITLE = "停电摘要（全量版）"

# Cut trailing runner/step logs that follow summary text
_TRAILING_CUT_RE = re.compile(
    r"(?:"
    r"\[OK\]\s*步骤\s*7[：:].*"
    r"|>>\s*步骤\s*8\b"
    r"|步骤\s*8[：:]"
    r"|\[OK\]\s*全部步骤执行完毕"
    r")",
    re.MULTILINE,
)

OUTPUT_PATTERNS = {
    "result": re.compile(r"处理结果输出\s*：\s*(.+)$", re.MULTILINE),
    "publish": re.compile(r"发出版输出\s*：\s*(.+)$", re.MULTILINE),
    "stats": re.compile(r"统计表输出\s*：\s*(.+)$", re.MULTILINE),
    "summary_simple": re.compile(r"简版摘要输出\s*：\s*(.+)$", re.MULTILINE),
    "summary_full": re.compile(r"全量摘要输出\s*：\s*(.+)$", re.MULTILINE),
    "exclude_output": re.compile(r"已写入[：:]\s*(.+\.xlsx)\s*$", re.MULTILINE),
}

# exclude 脚本每月累计频繁停电用户数
_EXCLUDE_MONTHLY_RE = re.compile(r"截至(\d+)月累计频繁停电用户[：:]\s*(\d+)")


def _strip_log_prefix(text: str) -> str:
    lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped:
            lines.append(stripped)
    return "\n".join(lines)


def _trim_trailing_runner_logs(text: str) -> str:
    """Remove step-7 completion and anything after from summary text."""
    if not text:
        return ""
    match = _TRAILING_CUT_RE.search(text)
    if match:
        text = text[: match.start()]
    return text.rstrip()


def _extract_between(text: str, start_marker: str, end_marker: str | None) -> str:
    start_idx = text.find(start_marker)
    if start_idx < 0:
        return ""
    content_start = start_idx + len(start_marker)
    if end_marker:
        end_idx = text.find(end_marker, content_start)
        if end_idx < 0:
            chunk = text[content_start:]
        else:
            chunk = text[content_start:end_idx]
    else:
        chunk = text[content_start:]
    return _strip_log_prefix(chunk)


def _extract_summary_section(text: str, title: str, end_title: str | None = None) -> str:
    """Extract summary text from inline or multi-line banner formats."""
    inline = f"── {title}──"
    if inline in text:
        end_inline = f"── {end_title}──" if end_title else None
        chunk = _extract_between(text, inline, end_inline)
        return _trim_trailing_runner_logs(chunk)

    if end_title:
        pattern = re.compile(
            rf"─{{2,}}\s*\n\s*{re.escape(title)}\s*\n\s*─{{2,}}\s*\n(.*?)"
            rf"(?=\n\s*─{{2,}}\s*\n\s*{re.escape(end_title)}\s*\n\s*─{{2,}}"
            rf"|\n\s*\[OK\]\s*步骤\s*7"
            rf"|\n\s*>>\s*步骤\s*8"
            rf"|\Z)",
            re.DOTALL,
        )
    else:
        pattern = re.compile(
            rf"─{{2,}}\s*\n\s*{re.escape(title)}\s*\n\s*─{{2,}}\s*\n(.*?)"
            rf"(?=\n\s*\[OK\]\s*步骤\s*7"
            rf"|\n\s*>>\s*步骤\s*8"
            rf"|\n\s*\[OK\]\s*全部步骤执行完毕"
            rf"|\Z)",
            re.DOTALL,
        )
    match = pattern.search(text)
    if match:
        return _trim_trailing_runner_logs(_strip_log_prefix(match.group(1)))
    return ""


def parse_exclude_output(log_text: str) -> ParsedResult:
    """Parse exclude script stdout for monthly cumulative frequent-outage user counts."""
    result = ParsedResult()

    # Extract monthly counts
    monthly: list[tuple[int, int]] = []
    for match in _EXCLUDE_MONTHLY_RE.finditer(log_text):
        month = int(match.group(1))
        count = int(match.group(2))
        monthly.append((month, count))

    if monthly:
        lines = ["按月累计频繁停电用户统计："]
        for month, count in sorted(monthly, key=lambda t: t[0]):
            lines.append(f"  截至{month}月：{count} 户")
        total_months = len(monthly)
        max_users = max(c for _, c in monthly)
        lines.append(f"共统计 {total_months} 个月，最多检出 {max_users} 户。")
        result.summary_simple = "\n".join(lines)
    else:
        result.summary_simple = "未解析到每月累计频繁停电用户数据。"

    # Extract output file path
    match = OUTPUT_PATTERNS["exclude_output"].search(log_text)
    if match:
        result.output_files["exclude_output"] = match.group(1).strip()

    return result


def parse_script_output(log_text: str) -> ParsedResult:
    result = ParsedResult()
    result.summary_simple = _extract_summary_section(
        log_text, SUMMARY_SIMPLE_TITLE, SUMMARY_FULL_TITLE
    )
    result.summary_full = _extract_summary_section(log_text, SUMMARY_FULL_TITLE)

    for key, pattern in OUTPUT_PATTERNS.items():
        match = pattern.search(log_text)
        if match:
            result.output_files[key] = match.group(1).strip()

    return result


def validate_output_path(file_path: str, output_base: Path) -> Path | None:
    """Ensure resolved path is under output_base."""
    try:
        resolved = Path(file_path).resolve()
        base = output_base.resolve()
        resolved.relative_to(base)
        if resolved.is_file():
            return resolved
    except (ValueError, OSError):
        pass
    return None
