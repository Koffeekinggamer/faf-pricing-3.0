from math import ceil

from scripts.meadow_wood_book import process_meadow_wood_workbook
from scripts.pricing_lib import retail_whole_dollar
from scripts.red_barn_book import _secondary_block, process_red_barn_workbook
from pathlib import Path


def test_meadow_retail_rounding():
    mult = 2.7
    assert retail_whole_dollar(559, mult) == ceil(559 * mult)
    assert retail_whole_dollar(219, mult) == ceil(219 * mult)


def test_englehart_secondary_desk_row():
    row = [
        "EH8030 Bed/King ",
        "HB 66",
        "",
        "",
        1151.0,
        1266.0,
        1554.0,
        2129.0,
        "",
        "EH8001",
        "Double Pedestal Desk",
        "",
        "",
        1151.0,
        1266.0,
        1554.0,
        2129.0,
    ]
    primary = {"EH8030", "EH8031"}
    sec = _secondary_block(row, primary)
    assert sec is not None
    assert sec[0] == "EH8001"


def _upload_fixture(path: Path) -> bool:
    try:
        return path.is_file()
    except OSError:
        return False


def test_red_barn_and_meadow_forbidden(tmp_path_factory):
    rb = Path("/home/ubuntu/.cursor/projects/workspace/uploads/Red_Barn_739a.zip")
    mw = Path("/home/ubuntu/.cursor/projects/workspace/uploads/Meadow_Wood_fe17.xlsx")
    if not _upload_fixture(rb) or not _upload_fixture(mw):
        return
    for proc, forbidden in (
        (process_red_barn_workbook(rb, multiplier=2.7), "wholesale"),
        (process_meadow_wood_workbook(mw, multiplier=2.7), "wholesale"),
    ):
        text = " ".join(c for pr in proc for c in pr.cells).lower()
        assert forbidden not in text
