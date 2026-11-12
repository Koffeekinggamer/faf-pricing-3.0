from pathlib import Path

from scripts.builders_data import BUILDERS_PATH, load_builders
from scripts import builders_data
from scripts.add_sheet import main as add_main
import sys


def test_add_sheet_look_alike_run(tmp_path, monkeypatch):
    """End-to-end with look-alike fixture; does not touch live builders.json."""
    monkeypatch.setattr(builders_data, "BUILDERS_PATH", tmp_path / "builders.json")
    monkeypatch.chdir(tmp_path)
    sheets = tmp_path / "sheets"
    sheets.mkdir()
    monkeypatch.setattr(
        "scripts.add_sheet.SHEETS_DIR",
        sheets,
    )

    fixture = Path(__file__).resolve().parent / "fixtures" / "look_alike_wholesale.csv"
    argv = [
        "add_sheet.py",
        "--builder",
        "Look-alike Builder",
        "--multiplier",
        "2.7",
        "--file",
        str(fixture),
        "--date",
        "2026-11-12",
    ]
    monkeypatch.setattr(sys, "argv", argv)
    assert add_main() == 0

    builders = load_builders(tmp_path / "builders.json")
    assert len(builders) == 1
    assert builders[0]["name"] == "Look-alike Builder"
    pdf = tmp_path / builders[0]["pdf"]
    assert pdf.is_file()
    text = pdf.read_bytes()
    assert b"approx" not in text.lower()
    assert b"Look-alike Builder" in text
