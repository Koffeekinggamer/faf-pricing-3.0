"""Read/write builders.json with stable alphabetical ordering."""

from __future__ import annotations

import json
import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
BUILDERS_PATH = REPO_ROOT / "data" / "builders.json"


def sort_key(name: str) -> str:
    """Case-insensitive sort key; ignore leading punctuation."""
    trimmed = re.sub(r"^[^a-zA-Z0-9]+", "", name or "")
    return trimmed.casefold()


FORBIDDEN_INDEX_NOTE_RE = re.compile(
    r"pass[- ]?through|\bmultiplier\b",
    re.I,
)


def validate_index_note(note: str | None) -> None:
    """Reject index labels that expose multiplier / pass-through wording to sales."""
    if not note:
        return
    if FORBIDDEN_INDEX_NOTE_RE.search(note):
        raise ValueError(
            "Index notes must not mention pass-through or multiplier "
            "(multipliers stay in data/multipliers.json only)"
        )


def slugify(name: str) -> str:
    base = re.sub(r"[^a-zA-Z0-9]+", "_", name.strip()).strip("_")
    return base or "builder"


def load_builders(path: Path | None = None) -> list[dict]:
    path = path or BUILDERS_PATH
    if not path.exists():
        return []
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError(f"{path} must contain a JSON array")
    return data


def find_builder_by_name(name: str, path: Path | None = None) -> dict | None:
    """Return catalog entry by display name (case-insensitive), or None."""
    key = sort_key(name)
    for b in load_builders(path):
        if sort_key(b.get("name", "")) == key:
            return b
    return None


def save_builders(builders: list[dict], path: Path | None = None) -> None:
    path = path or BUILDERS_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    sorted_builders = sorted(builders, key=lambda b: sort_key(b.get("name", "")))
    with path.open("w", encoding="utf-8") as f:
        json.dump(sorted_builders, f, indent=2, ensure_ascii=False)
        f.write("\n")


def upsert_builder(
    name: str,
    pdf_path: str,
    updated: str,
    note: str | None = None,
    *,
    price_list_date: str | None = None,
    path: Path | None = None,
) -> tuple[dict, bool]:
    """Insert or replace builder by name (case-insensitive). Returns (entry, replaced)."""
    validate_index_note(note)
    builders = load_builders(path)
    key = sort_key(name)
    replaced = False
    entry = {
        "name": name,
        "pdf": pdf_path,
        "updated": updated,
    }
    if note:
        entry["note"] = note
    if price_list_date:
        entry["price_list_date"] = price_list_date

    new_list: list[dict] = []
    found = False
    for b in builders:
        if sort_key(b.get("name", "")) == key:
            new_list.append(entry)
            found = True
            replaced = True
        else:
            new_list.append(b)
    if not found:
        new_list.append(entry)

    save_builders(new_list, path)
    return entry, replaced
