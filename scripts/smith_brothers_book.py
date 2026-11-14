"""Smith Brothers FAF retail book PDF → structured retail catalog (1.0× pass-through)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pdfplumber

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.dealer_terms import is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.pricing_lib import retail_whole_dollar
from scripts.retail_text import forbidden_phrases, sanitize_retail_text

BUILDER = "Smith Brothers"
DOC_TITLE = "Smith Brothers Retail Price List · Fall 2025 (Jan 12, 2026)"
MULTIPLIER = 1.0

FOOTPRINT_RE = re.compile(
    r"Foothills\s+Amish\s+Furniture\s+Retail\s+Prices|^\s*\d+\s+Foothills",
    re.I,
)
PAGE_NUM_ONLY_RE = re.compile(r"^\s*\d{1,3}\s+\d{1,3}\s*$")
GROUP_RE = re.compile(r"^(\d{3})\s+Group\b", re.I)
FRAME_SKU_RE = re.compile(r"^Frame\s*-\s*SKU", re.I)
SKU_HEADER_RE = re.compile(
    r"^((?:\d{3}-\d{2}(?:/\d{2})?(?:C)?\s+)+)\s*$"
)
SKU_TOKEN_RE = re.compile(r"\b(\d{3}-\d{2}(?:/\d{2})?(?:C)?)\b")
DIM_CHUNK_RE = re.compile(r"\d+''?\s*x\s*\d+''?\s*x\s*\d+''?", re.I)
SKIP_PROSE_RE = re.compile(
    r"Wood Trim Finishes|Nailhead Colors|PRICE the Frame|PER NUMBER OF FABRICS|"
    r"See P\d+ for help|Standard nailhead is removable|AVERAGE ALL FRAME|"
    r"Leather frames do not come",
    re.I,
)
DESC_RE = re.compile(r"^Description\s+(.+)$", re.I)
GRADE_ROW_RE = re.compile(r"^Grade\s+", re.I)
OPTION_PRICE_RE = re.compile(
    r"^(Firmer Seat Upgrade|Foam Back Option|Fabric Contrast|Contrast Leather|"
    r"Use Multiple Fabrics|Contrast Fabric|Multiple Fabrics|Per Grade|Down Option)",
    re.I,
)
DROP_ATTR_RE = re.compile(
    r"^(Wood Trim|Back Style|Nailhead|Standard Pillows|Seat Firmness|\(F=)",
    re.I,
)
DROP_BOILERPLATE_RE = re.compile(
    r"drawings represented|not drawn to scale|diagrams are not|"
    r"please see catalog page for detail|shaded area represents",
    re.I,
)
DIAGRAM_GARBAGE_RE = re.compile(
    r"^(serutaeF|snoisnemiD|nommoC|swolliP|gnicirP|snoitpO|lenaP|mrA|edoC|"
    r"gni\w*|noit|ruF|tin|ruO|wolliP|snoitpO|wolliP)",
    re.I,
)
PRICE_TAIL_RE = re.compile(
    r"(\$?\d[\d,]*(?:\.\d{2})?|Not Available|\u2014|-)(?:\s+(\$?\d[\d,]*(?:\.\d{2})?|Not Available|\u2014|-))*$",
    re.I,
)

KNOWN_PRODUCTS = (
    "LARGE SOFA",
    "LOVESEAT",
    "CHAIR",
    "OTTOMAN",
    "MID SOFA",
    "CONVERTIBLE SOFA",
    "CHAIR AND A HALF",
    "OTTOMAN AND A HALF",
    "WEDGE",
    "CHASE",
    "CORNER CHAIR",
    "CORNER SOFA",
    "SWIVEL CHAIR",
    "SWIVEL GLIDER",
    "GLIDER CHAIR",
    "ARMLESS SOFA",
    "ARMLESS CHAIR",
    "SOFA",
    "LAF / RAF SOFA",
    "LAF / RAF LOVESEAT",
    "LAF / RAF CHAIR",
    "LAF / RAF CORNER SOFA",
    "LAF / RAF CHAISE",
)

REVERSED_ARTIFACTS = ("serutaeF", "snoisnemiD", "nommoC", "swolliP", "gnicirP")

# Fixed spot checks (source book → retail PDF at 1.0×)
SPOT_CHECKS: list[tuple[str, str, str]] = [
    ("253-13 Large Sofa", "Grade 10", "$3,943"),
    ("253-20 Loveseat", "Grade 10", "$3,843"),
    ("212-40C Ottoman w/ Casters", "Grade 10", "$720"),
    ("212-21/22", "Grade 10", "$1,843"),
    ("505-30", "Grade 10", "$1,500"),
    ("505-40", "Grade 10", "$648"),
    ("704-10", "Grade 10", "$2,748"),
    ("770-10", "Grade 10", "$3,443"),
    ("900-60", "Grade 10", "$1,688"),
    ("919-30", "Grade 10", "$1,900"),
    ("932-47", "Grade 10", "$748"),
    ("260-10 Sofa", "Grade 10", "$3,443"),
    ("260-56 Swivel Chair", "Grade 10", "$1,860"),
]


@dataclass
class ParseStats:
    source_pages: int
    pages_kept: int
    pages_skipped: int
    wide_rows: int


def _looks_reversed(line: str) -> bool:
    if len(line) < 8:
        return False
    letters = re.sub(r"[^a-zA-Z]", "", line)
    if len(letters) < 8:
        return False
    return letters == letters[::-1]


def _is_garbage_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    if _looks_reversed(s) or DIAGRAM_GARBAGE_RE.match(s):
        return True
    if DROP_BOILERPLATE_RE.search(s):
        return True
    if FOOTPRINT_RE.search(s) or PAGE_NUM_ONLY_RE.match(s):
        return True
    if is_dealer_terms_text(s):
        return True
    if re.search(r"\bwholesale\b|preferred\s+dealer|charleston|surcharge", s, re.I):
        return True
    return False


def _fmt_retail_price(token: str) -> str:
    t = token.strip()
    if not t or t in ("-", "\u2014"):
        return "\u2014"
    if re.match(r"not\s+available", t, re.I):
        return "Not Available"
    if t.startswith("$"):
        inner = re.sub(r"[^\d.]", "", t)
        if not inner or inner == ".":
            return t
        val = float(inner)
    else:
        if t.startswith("+") or t.startswith("-"):
            sign = t[0]
            inner = re.sub(r"[^\d.]", "", t[1:])
            if not inner:
                return t
            val = float(inner)
            if sign == "-":
                return f"-${int(retail_whole_dollar(val, MULTIPLIER)):,}"
            return f"+${int(retail_whole_dollar(val, MULTIPLIER)):,}"
        inner = re.sub(r"[^\d.]", "", t)
        if not inner or inner == ".":
            return t
        val = float(inner)
    whole = retail_whole_dollar(val, MULTIPLIER)
    if t.startswith("$"):
        return f"${whole:,}"
    return f"${whole:,}"


def _split_product_headers(rest: str) -> list[str]:
    rest = rest.strip()
    if not rest:
        return []
    upper = rest.upper()
    found: list[tuple[int, str]] = []
    for prod in sorted(KNOWN_PRODUCTS, key=len, reverse=True):
        idx = 0
        while True:
            pos = upper.find(prod, idx)
            if pos < 0:
                break
            found.append((pos, prod.title() if prod != "SOFA" else "Sofa"))
            idx = pos + len(prod)
    if not found:
        return re.split(r"\s{2,}", rest)
    found.sort(key=lambda x: x[0])
    return [p for _, p in found]


def _split_row_values(rest: str) -> list[str]:
    if not rest.strip():
        return []
    parts = re.split(r"\s{2,}", rest.strip())
    if len(parts) <= 1:
        parts = rest.split()
    out: list[str] = []
    i = 0
    while i < len(parts):
        tok = parts[i]
        if tok.lower() == "not" and i + 1 < len(parts) and parts[i + 1].lower() == "available":
            out.append("Not Available")
            i += 2
            continue
        if re.match(r"^-?\$?\d", tok) or tok in ("-", "\u2014"):
            out.append(_fmt_retail_price(tok))
        elif re.match(r"^\d[\d,]*$", tok):
            out.append(_fmt_retail_price(tok))
        else:
            out.append(tok)
        i += 1
    return out


def _collect_skus(lines: list[str]) -> list[str]:
    best: list[str] = []
    for line in lines:
        found = SKU_TOKEN_RE.findall(line)
        if not found:
            continue
        if len(found) > len(best):
            best = found
        elif len(found) == len(best) and FRAME_SKU_RE.match(line):
            best = found
    seen: set[str] = set()
    ordered: list[str] = []
    for sku in best:
        if sku not in seen:
            seen.add(sku)
            ordered.append(sku)
    return ordered


def _group_sku_blocks(skus: list[str]) -> list[tuple[str, list[int]]]:
    from collections import OrderedDict

    blocks: OrderedDict[str, list[int]] = OrderedDict()
    for i, sku in enumerate(skus):
        prefix = sku.split("-")[0]
        blocks.setdefault(prefix, []).append(i)
    return [(f"{prefix} Group", idxs) for prefix, idxs in blocks.items()]


def _infer_group(skus: list[str], lines: list[str]) -> str:
    for line in lines:
        gm = GROUP_RE.search(line)
        if gm:
            return f"{gm.group(1)} Group"
    if skus:
        prefix = skus[0].split("-")[0]
        if prefix.isdigit():
            return f"{prefix} Group"
    return ""


def _normalize_dim(chunk: str) -> str:
    return chunk.replace("''", '"').strip()


def _parse_dimensions(line: str) -> list[str]:
    m = re.search(r"Dimensions.*?\)\s*(.+)$", line, re.I)
    rest = m.group(1).strip() if m else line
    chunks = DIM_CHUNK_RE.findall(rest)
    if chunks:
        return [_normalize_dim(c) for c in chunks]
    return re.split(r"\s{2,}", rest.strip())


def _product_headers_from_line(line: str) -> list[str] | None:
    dm = DESC_RE.match(line)
    if dm:
        return _split_product_headers(dm.group(1))
    stripped = line.strip()
    if not stripped or "$" in stripped or re.search(r"Grade\s+\d", stripped, re.I):
        return None
    if any(k in stripped.upper() for k in ("SOFA", "LOVESEAT", "CHAIR", "OTTOMAN", "WEDGE", "CHASE")):
        headers = _split_product_headers(stripped)
        if len(headers) >= 2:
            return headers
    return None


def _parse_grade_tail(line: str) -> list[str] | None:
    m = re.search(r"(Grade\s+.+)$", line, re.I)
    if not m:
        return None
    tail = m.group(1)
    m2 = re.match(r"(Grade\s+[^$]+?)\s+(.+)$", tail, re.I)
    if not m2:
        gm = re.match(r"(Grade\s+\S+(?:/\S+)?)", tail.strip(), re.I)
        if not gm:
            return None
        return None
    label = m2.group(1).strip()
    prices = _split_row_values(m2.group(2))
    if not prices:
        return None
    return [label, *prices]


def _parse_grade_line(line: str) -> list[str] | None:
    stripped = line.strip()
    if not stripped or "Wood Trim Finishes" in stripped:
        return None
    if re.match(r"^Grade\s+", stripped, re.I):
        row = _parse_grade_tail(stripped)
        if row:
            return row
    m = re.match(
        r"^(.+?)\s+(Grade\s+(?:\d+|>\d+).*?)$",
        stripped,
        re.I,
    )
    if m and re.search(r"[\d,$]", m.group(1)):
        label_m = re.match(r"(Grade\s+\S+(?:/\S+)?(?:\s+Add:)?[^$]*)", m.group(2).strip(), re.I)
        if not label_m:
            return None
        label = label_m.group(1).strip()
        if label.endswith("Add:"):
            label = label[:-4].strip()
        prices = _split_row_values(m.group(1))
        if prices:
            return [label, *prices]
    return None


def _parse_labeled_prices(line: str) -> list[str] | None:
    m = re.match(r"^(.+?)\s+((?:\$|-)?\d[\d\s,$+-]+.*)$", line)
    if not m:
        return None
    label = m.group(1).strip()
    if DROP_ATTR_RE.match(label):
        return None
    rest = m.group(2)
    if not re.search(r"\d", rest):
        return None
    prices = []
    for tok in rest.split():
        if tok.lower() == "not" and prices:
            prices[-1] = "Not Available"
            continue
        if tok.lower() == "available":
            continue
        prices.append(_fmt_retail_price(tok))
    if len(prices) < 2 and not label.lower().startswith("grade"):
        return None
    return [label, *prices]


def _truncate_before_diagram(lines: list[str]) -> list[str]:
    i = 0
    while i < len(lines) and (
        FOOTPRINT_RE.search(lines[i]) or PAGE_NUM_ONLY_RE.match(lines[i])
    ):
        i += 1
    out: list[str] = []
    for line in lines[i:]:
        s = line.strip()
        if DIAGRAM_GARBAGE_RE.match(s):
            continue
        if _looks_reversed(s) and not re.search(r"Grade\s+\d", s, re.I):
            continue
        if re.search(
            r"drawings represented|not drawn to scale|please see catalog page",
            s,
            re.I,
        ):
            gm = GROUP_RE.search(s)
            if gm:
                out.append(f"{gm.group(1)} Group")
            continue
        if _is_garbage_line(line):
            if re.search(r"Grade\s+\d+", line, re.I):
                out.append(line)
            elif FRAME_SKU_RE.match(line) or DESC_RE.match(line) or line.startswith(
                "Dimensions"
            ):
                out.append(line)
            continue
        out.append(line)
    return out


def _page_lines(page: pdfplumber.page.Page) -> list[str]:
    text = page.extract_text() or ""
    lines = [clean_pdf_plain_text(ln).strip() for ln in text.splitlines()]
    return _truncate_before_diagram([ln for ln in lines if ln])


def _slice_row(cells: list[str], indices: list[int]) -> list[str]:
    if not cells:
        return []
    label = cells[0]
    vals = cells[1:]
    picked = [vals[i] if i < len(vals) else "\u2014" for i in indices]
    return [label, *picked]


def _parse_grade_page(lines: list[str]) -> list[ProcessedRow]:
    skus = _collect_skus(lines)
    if not skus:
        return []

    headers: list[str] = []
    shared_rows: list[list[str]] = []

    for line in lines:
        if SKIP_PROSE_RE.search(line):
            continue
        if GROUP_RE.search(line) and "Drawings" not in line:
            continue
        if FRAME_SKU_RE.match(line) or SKU_HEADER_RE.match(line.strip()):
            continue
        hdrs = _product_headers_from_line(line)
        if hdrs and len(hdrs) >= 2:
            headers = hdrs
            continue
        if line.startswith("Dimensions"):
            dims = _parse_dimensions(line)
            if dims:
                shared_rows.append(["Dimensions (W x D x H)", *dims])
            continue
        if DROP_ATTR_RE.match(line) and not re.search(r"\$\d", line):
            continue
        grade_cells = _parse_grade_line(line)
        if grade_cells:
            if len(grade_cells) - 1 > len(skus):
                grade_cells = [grade_cells[0], *grade_cells[1 : len(skus) + 1]]
            shared_rows.append(grade_cells)
            continue
        if OPTION_PRICE_RE.match(line) or re.match(
            r"^(Contrast Fabric|Per Grade|Wall Clearance|Down Option)", line, re.I
        ):
            wide = _parse_labeled_prices(line)
            if wide:
                if len(wide) - 1 > len(skus):
                    wide = [wide[0], *wide[1 : len(skus) + 1]]
                shared_rows.append(wide)
            continue

    grade_rows = [r for r in shared_rows if r[0].startswith("Grade")]
    if not grade_rows:
        return []

    out: list[ProcessedRow] = []
    for group_name, indices in _group_sku_blocks(skus):
        ncol = len(indices)
        out.append(ProcessedRow(RowKind.SECTION, [group_name]))
        if headers:
            hdr_cells = ["", *[headers[i] if i < len(headers) else "" for i in indices]]
            out.append(ProcessedRow(RowKind.TABLE_HEADER, hdr_cells))
        out.append(
            ProcessedRow(
                RowKind.TABLE_HEADER,
                ["Item #", *[skus[i] for i in indices]],
            )
        )
        for row in shared_rows:
            out.append(ProcessedRow(RowKind.WIDE, _slice_row(row, indices)))
    return out


def _parse_options_page(lines: list[str]) -> list[ProcessedRow]:
    if not lines:
        return []
    title = sanitize_retail_text(lines[0], MULTIPLIER) or lines[0]
    if re.search(r"warranty details|standard features", title, re.I):
        return []
    rows: list[ProcessedRow] = [ProcessedRow(RowKind.SECTION, [title])]
    for line in lines[1:]:
        if _is_garbage_line(line):
            break
        if re.search(r"drawings represented", line, re.I):
            break
        wide = _parse_labeled_prices(line)
        if wide and (wide[0].lower().startswith("grade") or len(wide) >= 3):
            rows.append(ProcessedRow(RowKind.WIDE, wide))
            continue
        if re.search(r"\$\d", line):
            m = re.match(r"^(.+?)\s+((?:\$|-)?\d[\d\s,$+-]+.*)$", line.strip())
            if m:
                label = m.group(1).strip()
                if len(label) > 4 and not DROP_ATTR_RE.match(label):
                    prices = _split_row_values(m.group(2))
                    if prices:
                        rows.append(ProcessedRow(RowKind.WIDE, [label, *prices]))
                        continue
        if len(line.strip()) > 12 and "$" not in line:
            rows.append(ProcessedRow(RowKind.PROSE, [line.strip()]))

    if len(rows) <= 1:
        return []
    return rows


def _page_kind(lines: list[str]) -> str:
    text = "\n".join(lines)
    has_skus = bool(_collect_skus(lines))
    has_grade_row = any(_parse_grade_line(ln) for ln in lines)
    if has_skus and has_grade_row:
        return "grade"
    if any(
        k in text
        for k in (
            "Fringe & cord",
            "Fabric / leather collages",
            "Arm caps",
            "Rechargeable battery",
            "Decorative Tape",
            "Contrast and bias",
        )
    ):
        return "options"
    return "skip"


def process_smith_brothers_pdfs(
    paths: list[Path],
    *,
    multiplier: float = MULTIPLIER,
) -> tuple[list[ProcessedRow], dict]:
    ordered = sorted(paths, key=lambda p: p.name.casefold())
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]
    stats = ParseStats(source_pages=0, pages_kept=0, pages_skipped=0, wide_rows=0)

    for path in ordered:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                stats.source_pages += 1
                lines = _page_lines(page)
                kind = _page_kind(lines)
                if kind == "skip":
                    stats.pages_skipped += 1
                    continue
                if kind == "grade":
                    block = _parse_grade_page(lines)
                else:
                    block = _parse_options_page(lines)
                if not block:
                    stats.pages_skipped += 1
                    continue
                stats.pages_kept += 1
                stats.wide_rows += sum(1 for r in block if r.kind == RowKind.WIDE)
                out.extend(block)

    return collapse_blank_rows(out), {
        "source_pages": stats.source_pages,
        "pages_kept": stats.pages_kept,
        "pages_skipped": stats.pages_skipped,
        "wide_rows": stats.wide_rows,
        "parts": len(ordered),
    }


def smith_brothers_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if is_dealer_terms_text(str(c)):
                bad.append(str(c))
    return bad


def pdf_contains_prices(pdf_path: Path, needles: list[str]) -> dict[str, bool]:
    with pdfplumber.open(pdf_path) as pdf:
        full = "\n".join((p.extract_text() or "") for p in pdf.pages)
    return {n: n in full for n in needles}


def pdf_has_reversed_artifacts(pdf_path: Path) -> list[str]:
    hits: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    for tok in REVERSED_ARTIFACTS:
        if tok in text:
            hits.append(tok)
    for line in text.splitlines():
        if _looks_reversed(line.strip()):
            hits.append(line.strip()[:40])
            break
    return hits