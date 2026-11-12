import pytest

from scripts.multipliers_data import (
    get_stored_multiplier,
    load_multipliers,
    resolve_multiplier,
    set_stored_multiplier,
)


def test_resolve_new_builder_requires_multiplier(tmp_path):
    path = tmp_path / "multipliers.json"
    with pytest.raises(ValueError, match="new builder"):
        resolve_multiplier("New Builder", None, override=False, path=path)


def test_resolve_existing_uses_saved(tmp_path):
    path = tmp_path / "multipliers.json"
    set_stored_multiplier("Acme", 2.7, path=path)
    mult, persist = resolve_multiplier("acme", None, override=False, path=path)
    assert mult == 2.7
    assert persist is False


def test_resolve_mismatch_requires_override(tmp_path):
    path = tmp_path / "multipliers.json"
    set_stored_multiplier("Acme", 2.7, path=path)
    with pytest.raises(ValueError, match="override-multiplier"):
        resolve_multiplier("Acme", 3.0, override=False, path=path)
    mult, persist = resolve_multiplier("Acme", 3.0, override=True, path=path)
    assert mult == 3.0
    assert persist is True


def test_first_add_persists_multiplier(tmp_path):
    path = tmp_path / "multipliers.json"
    mult, persist = resolve_multiplier("First Co", 2.7, override=False, path=path)
    assert mult == 2.7
    assert persist is True
    set_stored_multiplier("First Co", mult, path=path)
    assert get_stored_multiplier("First Co", path=path) == 2.7


def test_backfill_matches_builders_json():
    from scripts.builders_data import BUILDERS_PATH, load_builders
    from scripts.multipliers_data import MULTIPLIERS_PATH

    builders = load_builders(BUILDERS_PATH)
    m = load_multipliers(MULTIPLIERS_PATH)
    names = {b["name"] for b in builders}
    assert set(m.keys()) == names
    assert all(v == 2.7 for v in m.values())
