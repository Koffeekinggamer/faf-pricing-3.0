"""Smith Brothers source vs catalog SKU/group inventories."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pdfplumber

SKU_TOKEN_RE = re.compile(r"\b(\d{3}-\d{2}(?:/\d{2})?(?:C)?)\b")
GROUP_HEADING_RE = re.compile(r"\b(\d{3})\s+Group\b", re.I)
SPLIT_HEADER_RE = re.compile(
    r"\b(?:Lovesea\s+t|Ottoma\s+n|Sof\s+a|Lovese\s+at)\b", re.I
)

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INVENTORY = REPO_ROOT / "data" / "smith_brothers_source_inventory.json"


def _sku_group(sku: str) -> str:
    return sku.split("-")[0]


def source_inventory_from_parts(paths: list[Path]) -> tuple[set[str], set[str]]:
    groups: set[str] = set()
    skus: set[str] = set()
    for path in paths:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if not re.search(r"Grade\s+10\b", text):
                    continue
                for m in SKU_TOKEN_RE.finditer(text):
                    sk = m.group(1)
                    skus.add(sk)
                    groups.add(_sku_group(sk))
    return groups, skus


def load_expected_inventory(path: Path | None = None) -> tuple[set[str], set[str]]:
    p = path or DEFAULT_INVENTORY
    data = json.loads(p.read_text(encoding="utf-8"))
    return set(data["groups"]), set(data["skus"])


def catalog_inventory(pdf_path: Path) -> tuple[set[str], set[str]]:
    groups: set[str] = set()
    skus: set[str] = set()
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    for m in GROUP_HEADING_RE.finditer(text):
        groups.add(m.group(1))
    for line in text.splitlines():
        if line.strip().startswith("Item #") or "Item #" in line:
            for sk in SKU_TOKEN_RE.findall(line):
                skus.add(sk)
                groups.add(_sku_group(sk))
    if not skus:
        for sk in SKU_TOKEN_RE.findall(text):
            skus.add(sk)
            groups.add(_sku_group(sk))
    return groups, skus


def catalog_has_split_headers(pdf_path: Path) -> list[str]:
    hits: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    for line in text.splitlines():
        if SPLIT_HEADER_RE.search(line):
            hits.append(line.strip()[:80])
    return hits


def coverage_report(
    source_groups: set[str],
    source_skus: set[str],
    site_groups: set[str],
    site_skus: set[str],
) -> dict:
    return {
        "source_groups": len(source_groups),
        "site_groups": len(site_groups),
        "missing_groups": sorted(source_groups - site_groups, key=int),
        "extra_groups": sorted(site_groups - source_groups, key=int),
        "source_skus": len(source_skus),
        "site_skus": len(site_skus),
        "missing_skus": sorted(source_skus - site_skus),
        "extra_skus": sorted(site_skus - source_skus),
    }