"""Retail policy: finished pricing only (site-wide default)."""

from __future__ import annotations

import re

UNFINISHED_NOTE_RE = re.compile(
    r"\d+\s*%\s*(less|off|deduct)\s+.*unfinished|"
    r"deduct\s+\d+\s*%.*unfinished|"
    r"unfinished.*\d+\s*%|"
    r"less\s+for\s+unfinished|"
    r"for\s+unfinished\s+prices",
    re.I,
)
UNFINISHED_HEADER_RE = re.compile(r"^\s*unfinished\s*$", re.I)


def drop_unfinished_note_text(text: str) -> str:
    """Return empty if the line is only an unfinished discount note."""
    if not text:
        return ""
    if UNFINISHED_NOTE_RE.search(text):
        return ""
    t = re.sub(r"\s*•\s*finished\s+and\s+unfinished\s*•\s*", " ", text, flags=re.I)
    t = re.sub(r"\bfinished\s+and\s+unfinished\b", "Finished", t, flags=re.I)
    return t.strip()


def is_unfinished_header_label(label: str) -> bool:
    return bool(UNFINISHED_HEADER_RE.match((label or "").strip()))


def finished_price_column_indices(
    row: list, header_rows: list[list], *, default_finished_start: int | None = None
) -> set[int]:
    """Infer finished wholesale columns from a header row containing Unfinished / Finished."""
    finished: set[int] = set()
    unfinished: set[int] = set()
    for h in header_rows:
        for j, c in enumerate(h):
            if not isinstance(c, str):
                continue
            s = c.strip().lower()
            if s == "unfinished":
                unfinished.add(j)
            if "finished" in s and "unfinished" not in s:
                finished.add(j)
    if finished:
        span = max(finished) - min(finished)
        cols: set[int] = set()
        for h in header_rows:
            for j in range(min(finished), min(finished) + span + 6):
                if j < len(h) and isinstance(h[j], (int, float)):
                    continue
                if j in finished or j > max(finished):
                    cols.add(j)
        # price columns after "Standard Finished" label — next 5 numeric cols in data rows
        start = min(finished) if finished else (default_finished_start or 0)
        return set(range(start, start + 5))
    nums = [
        j
        for j, c in enumerate(row)
        if isinstance(c, (int, float)) and not isinstance(c, bool) and c > 20
    ]
    if len(nums) >= 16:
        half = len(nums) // 2
        return set(nums[half:])
    if default_finished_start is not None:
        return {j for j in nums if j >= default_finished_start}
    return set(nums)
