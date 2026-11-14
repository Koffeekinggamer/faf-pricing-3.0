"""Public builder index must not expose multiplier / pass-through labels."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from scripts.builders_data import validate_index_note

REPO = Path(__file__).resolve().parents[1]
FORBIDDEN_LABEL = re.compile(r"pass[- ]?through|\bmultiplier\b", re.I)
DISPLAY_KEYS = frozenset({"name", "pdf", "updated", "note", "price_list_date"})


def test_builders_json_no_pass_through_or_multiplier_in_display_fields():
    builders = json.loads((REPO / "data" / "builders.json").read_text(encoding="utf-8"))
    for b in builders:
        for key, value in b.items():
            if key not in DISPLAY_KEYS or not isinstance(value, str):
                continue
            assert not FORBIDDEN_LABEL.search(value), (
                f"{b.get('name')}: {key!r} contains forbidden label: {value!r}"
            )


def test_app_js_does_not_render_builder_notes_or_multipliers():
    js = (REPO / "assets" / "app.js").read_text(encoding="utf-8")
    assert "b.note" not in js
    assert "multiplier" not in js.lower()


def test_validate_index_note_rejects_pass_through_wording():
    with pytest.raises(ValueError, match="pass-through|multiplier"):
        validate_index_note("FAF retail pass-through (multiplier 1.0)")