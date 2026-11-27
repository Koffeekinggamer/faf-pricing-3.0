"""Guards for RH Yoder extraction (no workbook / wholesale numbers in repo)."""

from __future__ import annotations

import re
import tempfile
from pathlib import Path

import openpyxl

from scripts.dealer_book import RowKind
from scripts.rh_yoder_book import (
    _drop_empty_wood_group,
    _parse_table_groups,
    held_items_for_pr,
    printed_whole_dollar,
    process_rh_yoder_workbook,
    retail_from_raw,
)

# Synthetic look-alike wholesales (not RH Yoder list values)
_W1 = 100.0
_W2 = 110.0
_W3 = 120.0
_W4 = 115.0
_W5 = 130.0
_W6 = 125.0


def test_printed_half_up_then_ceil_retail():
    # Rule check: printed whole ×2.7, ceil (no real list numbers)
    assert printed_whole_dollar(100.49) == 100
    assert retail_from_raw(100.49, 2.7) == "$270"
    assert printed_whole_dollar(100.5) == 101
    assert printed_whole_dollar(99.4) == 99


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
    assert groups[0][1] == 1
    assert groups[1][1] == 6
    assert groups[2][1] == 10


def _build_synthetic_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Finished Prices From RH Yoder"
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
    ws["B29"] = _W1
    ws["C29"] = _W2
    ws["D29"] = _W3
    ws["E29"] = _W4
    ws["F29"] = _W5
    ws["G29"] = _W6
    ws["H29"] = "Included"
    ws["I29"] = 29.0
    ws["J29"] = 96.0
    ws["A34"] = "Albany"
    ws["A35"] = "Side Chair"
    ws["B35"] = 150.0
    ws["C35"] = 160.0
    ws["D35"] = 170.0
    ws["E35"] = 180.0
    ws["F35"] = 200.0
    ws["G35"] = 130.0  # held rustic
    ws["H35"] = "-"
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
    kvals = [320.0, 360.0, 380.0, 340.0, 370.0, 390.0, 350.0, 380.0, 400.0]
    ws["A1303"] = 42
    for col, val in zip(["B", "C", "D", "G", "H", "I", "K", "L", "M"], kvals):
        ws[f"{col}1303"] = val
    ws["A607"] = '13" x 96" Bench'
    ws["B607"] = 450.0
    ws["C607"] = 500.0
    ws["A608"] = '13" x 108" Bench'
    ws["B608"] = 520.0
    ws["A952"] = "Crayton Table"
    ws["B952"] = "Oak"
    ws["A953"] = "Size"
    ws["B953"] = "Solid"
    ws["C953"] = "1 Leaf"
    ws["D953"] = "2 Leaf"
    ws["A954"] = "42x60"
    ws["B954"] = 300.0
    ws["C954"] = 330.0  # held
    ws["D954"] = 380.0
    ws["A764"] = "Mini Table"
    ws["B764"] = "Oak"
    ws["A765"] = "Size"
    ws["B765"] = "Solid"
    ws["A766"] = "42x60"
    ws["B766"] = 400.0
    ws["M766"] = "Deduct 10% for Rustic Species"
    ws["A720"] = "For Oak and Brown Maple chairs with Cherry seats"
    ws["K720"] = 11.0
    ws["A2264"] = "Leaf Locks  (cost/leaf)"
    ws["K2264"] = 17.0
    ws["A2265"] = "Magnetic Skirts Available"
    ws["K2265"] = "$45 per leaf"
    ws["A2267"] = 'Thickness-Upgrade from 1" thick top to 1 1/4". Cost for all wood species. (unfinished price)'
    ws["K2267"] = "add 22%"
    ws["A2361"] = "Specialty Finishes"
    ws["A2364"] = "Chairs"
    ws["B2364"] = 15.0
    ws["C2364"] = 39.0
    ws["E2364"] = 31.0
    ws["G2364"] = 53.0
    # Iron Forge: Live Edge walnut held (J1288–J1291); without Live Edge keeps walnut
    ws["A1286"] = "Iron Forge Table-Live Edge"
    ws["J1286"] = "Walnut"
    ws["A1287"] = "Size"
    ws["J1287"] = "Solid Top"
    ws["A1288"] = "42x66"
    ws["J1288"] = 600.0
    ws["A1294"] = "Iron Forge Table without Live Edge"
    ws["J1294"] = "Walnut"
    ws["A1295"] = "Size"
    ws["J1295"] = "Solid Top"
    ws["A1296"] = "42x66"
    ws["J1296"] = 500.0
    return wb


def test_synthetic_workbook_paths_and_notes():
    wb = _build_synthetic_workbook()
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / "rh_synth.xlsx"
        wb.save(path)
        res = process_rh_yoder_workbook(path, multiplier=2.7)

    wide = [r for r in res.rows if r.kind == RowKind.WIDE]
    acadia = next(r for r in wide if r.cells[0] == "Side Chair" and r.cells[1] == retail_from_raw(_W1, 2.7))
    assert len([c for c in acadia.cells[1:7] if c.startswith("$")]) == 6
    assert "Included" in acadia.cells
    assert retail_from_raw(29.0, 2.7) in acadia.cells
    assert retail_from_raw(96.0, 2.7) in acadia.cells

    albany = next(
        r for r in wide if r.cells[0] == "Side Chair" and r.cells[1] == retail_from_raw(150.0, 2.7)
    )
    assert len([c for c in albany.cells[1:7] if c.startswith("$")]) == 5
    assert retail_from_raw(130.0, 2.7) not in albany.cells

    crayton = None
    after_crayton = False
    for pr in res.rows:
        if pr.kind == RowKind.SUBHEAD and pr.cells[0] == "Crayton Table":
            after_crayton = True
            continue
        if after_crayton and pr.kind == RowKind.WIDE and pr.cells[0] == "42x60":
            crayton = pr
            break
    assert crayton is not None
    assert crayton.cells[1].startswith("$")
    assert crayton.cells[2] == ""
    assert crayton.cells[3].startswith("$")

    keelan = None
    after_keelan = False
    for pr in res.rows:
        if pr.kind == RowKind.SUBHEAD and pr.cells[0] == "Keelan Table":
            after_keelan = True
            continue
        if after_keelan and pr.kind == RowKind.WIDE and pr.cells[0] == "42":
            keelan = pr
            break
    assert keelan is not None
    assert keelan.cells[1] == retail_from_raw(320.0, 2.7)
    assert len([c for c in keelan.cells[1:] if c.startswith("$")]) == 9
    assert any("\nSolid" in c or "\n1 Leaf" in c for r in wide for c in r.cells)

    blob = "\n".join(" ".join(r.cells) for r in res.rows)
    assert '13" x 96" Bench' not in blob or retail_from_raw(450.0, 2.7) not in blob
    assert any(retail_from_raw(11.0, 2.7) in " ".join(r.cells) for r in res.rows)
    assert any(retail_from_raw(17.0, 2.7) + " per leaf" in " ".join(r.cells) for r in res.rows)
    assert any("$122 per leaf" in " ".join(r.cells) for r in res.rows)
    assert any("add 22%" in " ".join(r.cells) for r in res.rows)
    assert "Deduct 10% for Rustic Species" in blob

    for r in wide:
        joined = " ".join(r.cells)
        assert not re.search(r"(?<!\$)(?<![,\d])(20[1-3]\d)(?!\d)", joined)

    assert len(held_items_for_pr()) == 6
    assert not any(r.kind == RowKind.TABLE_HEADER for r in res.rows)

    in_wle = False
    wle_header: list[str] = []
    for pr in res.rows:
        if pr.kind == RowKind.SUBHEAD and pr.cells[0] == "Iron Forge Table without Live Edge":
            in_wle = True
            continue
        if in_wle and pr.kind == RowKind.SUBHEAD:
            break
        if in_wle and pr.kind == RowKind.WIDE and not any(c.startswith("$") for c in pr.cells):
            wle_header = pr.cells
        if in_wle and pr.kind == RowKind.WIDE and pr.cells[0] == "42x66" and wle_header:
            wi = next(i for i, h in enumerate(wle_header) if "Walnut" in h)
            assert pr.cells[wi] == retail_from_raw(500.0, 2.7)
            break


def test_iron_forge_non_held_walnut_survives_empty_group_filter():
    """Only Live Edge J1288–J1291 is held; without-live-edge walnut must not be dropped."""
    rows: list[list] = [[None] * 11 for _ in range(1300)]
    rows[1295][0] = "42x66"
    rows[1295][9] = 510.0
    hidden: set[int] = set()
    assert not _drop_empty_wood_group(
        "Iron Forge Table without Live Edge",
        "Walnut",
        rows,
        hidden,
        1295,
        1296,
        9,
        9,
    )
    rows[1287][0] = "42x66"
    rows[1287][9] = 520.0
    assert _drop_empty_wood_group(
        "Iron Forge Table-Live Edge",
        "Walnut",
        rows,
        hidden,
        1287,
        1288,
        9,
        9,
    )