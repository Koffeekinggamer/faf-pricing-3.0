"""Guards for RH Yoder extraction (no workbook / wholesale numbers in repo)."""

from __future__ import annotations

import re

from scripts.dealer_book import RowKind
from scripts.rh_yoder_book import (
    _parse_table_groups,
    held_items_for_pr,
    printed_whole_dollar,
    process_rh_yoder_workbook,
    retail_from_raw,
)


def test_printed_half_up_then_ceil_retail():
    # Acadia Oak raw 254.69783 → printed $255 → ceil(255×2.7)=$689
    assert printed_whole_dollar(254.69783) == 255
    assert retail_from_raw(254.69783, 2.7) == "$689"
    # Half-up boundary
    assert printed_whole_dollar(254.5) == 255
    assert printed_whole_dollar(254.4) == 254


def test_keelan_label_shift_groups_by_solid_order():
    title = [
        "Keelan Table",
        "Oak",
        None,
        None,
        None,
        "B. Maple, Wormy Maple,Sap Cherry",
        None,
        None,
        None,
        None,
        "Cherry, H. Maple",
    ]
    size = [
        "Size",
        "Solid",
        "1 Leaf",
        "2 Leaf",
        None,
        None,
        "Solid",
        "1 Leaf",
        "2 Leaf",
        None,
        "Solid",
        "1 Leaf",
        "2 Leaf",
    ]
    groups = _parse_table_groups(title, size)
    assert len(groups) == 3
    assert groups[0][1] == 1  # Oak prices at col B (0-based 1)
    assert groups[1][1] == 6  # B.Maple prices at G despite label at F
    assert groups[2][1] == 10


def test_synthetic_acadia_row_and_held_albany_rustic():
    """Minimal sheet fragment: Acadia + Albany rustic held + options dollar."""
    # Build a fake workbook via process path is heavy; unit-test helpers + mini rows
    # through private chair logic by constructing a tiny xlsx in tmp is OK but
    # keep wholesale out of the repo — use in-memory openpyxl.
    import openpyxl
    from pathlib import Path
    import tempfile

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Finished Prices From RH Yoder"
    # Hide padding rows 2–27
    for r in range(2, 28):
        ws.row_dimensions[r].hidden = True
    ws["A28"] = "Acadia"
    ws["B28"] = "Oak"
    ws["C28"] = "B. Maple"
    ws["D28"] = "Cherry"
    ws["E28"] = "QSWO"
    ws["F28"] = "Walnut"
    ws["G28"] = "Rustic"
    ws["H28"] = "Heartland Standard Grade Fabric"
    ws["I28"] = "Premium"
    ws["J28"] = "Leather"
    ws["A29"] = "Side Chair"
    ws["B29"] = 254.69783
    ws["C29"] = 271.55
    ws["D29"] = 296.84
    ws["E29"] = 313.69
    ws["F29"] = 404.29
    ws["G29"] = 296.84
    ws["H29"] = "Included"
    ws["I29"] = 32
    ws["J29"] = 98
    ws["A34"] = "Albany"
    ws["A35"] = "Side Chair"
    ws["B35"] = 171.65
    ws["C35"] = 187.86
    ws["D35"] = 199.08
    ws["E35"] = 219.02
    ws["F35"] = 265.14
    ws["G35"] = 149.58  # held rustic
    ws["H35"] = "-"
    # Keelan label-shift block
    ws["A1301"] = "Keelan Table"
    ws["B1301"] = "Oak"
    ws["F1301"] = "B. Maple, Wormy Maple,Sap Cherry"
    ws["K1301"] = "Cherry, H. Maple"
    ws["A1302"] = "Size"
    ws["B1302"] = "Solid"
    ws["C1302"] = "1 Leaf"
    ws["D1302"] = "2 Leaf"
    ws["G1302"] = "Solid"
    ws["H1302"] = "1 Leaf"
    ws["I1302"] = "2 Leaf"
    ws["K1302"] = "Solid"
    ws["L1302"] = "1 Leaf"
    ws["M1302"] = "2 Leaf"
    ws["A1303"] = 42
    ws["B1303"] = 859.54
    ws["C1303"] = 987.54
    ws["D1303"] = 1069.54
    ws["G1303"] = 947.36
    ws["H1303"] = 1088.88
    ws["I1303"] = 1178.42
    ws["K1303"] = 1116.24
    ws["L1303"] = 1283.76
    ws["M1303"] = 1387.80
    # Held Sinclair-style benches
    ws["A607"] = '13" x 96" Bench'
    ws["B607"] = 451.24
    ws["C607"] = 498.85
    ws["A608"] = '13" x 108" Bench'
    ws["B608"] = 500.36
    # Crayton held 1-leaf
    ws["A952"] = "Crayton Table"
    ws["B952"] = "Oak"
    ws["A953"] = "Size"
    ws["B953"] = "Solid"
    ws["C953"] = "1 Leaf"
    ws["D953"] = "2 Leaf"
    ws["A954"] = "42x60"
    ws["B954"] = 768.22
    ws["C954"] = 853.22  # held
    ws["D954"] = 1020.22
    # Options
    ws["A720"] = "For Oak and Brown Maple chairs with Cherry seats"
    ws["K720"] = 10
    ws["A2264"] = "Leaf Locks  (cost/leaf)"
    ws["K2264"] = 18
    ws["A2265"] = "Magnetic Skirts Available"
    ws["K2265"] = "$45 per leaf"
    ws["A2267"] = "Thickness upgrade"
    ws["K2267"] = "add 22%"
    # Specialty
    ws["A2361"] = "Specialty Finishes"
    ws["A2364"] = "Chairs"
    ws["B2364"] = 16
    ws["C2364"] = 40
    ws["E2364"] = 30
    ws["G2364"] = 54

    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / "rh.xlsx"
        wb.save(path)
        res = process_rh_yoder_workbook(path, multiplier=2.7)

    wide = [r for r in res.rows if r.kind == RowKind.WIDE]
    acadia = next(r for r in wide if r.cells[0] == "Side Chair" and "$689" in r.cells)
    assert acadia.cells[1] == "$689"
    assert len([c for c in acadia.cells[1:7] if c.startswith("$")]) == 6
    assert "Included" in acadia.cells
    assert "$87" in acadia.cells  # 32×2.7 ceil
    assert "$265" in acadia.cells  # 98×2.7 ceil

    albany = next(
        r
        for r in wide
        if r.cells[0] == "Side Chair" and r.cells[1] == "$465"
    )
    # Rustic held — at most 5 wood $
    wood_dollars = [c for c in albany.cells[1:7] if c.startswith("$")]
    assert len(wood_dollars) == 5
    assert "$405" not in albany.cells  # 149.58 printed 150 → would be $405

    blob = "\n".join(" ".join(r.cells) for r in res.rows)
    assert '13" x 96" Bench' not in blob or "$1,218" not in blob
    # Crayton 1-leaf blank between solid and 2-leaf
    crayton = next(r for r in wide if r.cells[0] == "42x60")
    assert crayton.cells[1].startswith("$")
    assert crayton.cells[2] == ""  # held 1-leaf
    assert crayton.cells[3].startswith("$")

    keelan = next(r for r in wide if r.cells[0] == "42")
    assert len([c for c in keelan.cells[1:] if c.startswith("$")]) == 9

    # Options: K720=10→$27; leaf locks 18→$49/leaf; magnetic $45→$122/leaf
    assert any("$27" in " ".join(r.cells) for r in res.rows)
    assert any("$49 per leaf" in " ".join(r.cells) for r in res.rows)
    assert any("$122 per leaf" in " ".join(r.cells) for r in res.rows)
    assert any("add 22%" in " ".join(r.cells) for r in res.rows)

    # No bare yearish / non-$ wholesale leftovers on price rows
    for r in wide:
        joined = " ".join(r.cells)
        assert not re.search(r"(?<!\$)(?<![,\d])(20[1-3]\d)(?!\d)", joined)

    assert len(held_items_for_pr()) == 6
