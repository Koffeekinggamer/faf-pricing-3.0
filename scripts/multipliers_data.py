"""Per-builder retail multipliers (staff tooling only; not shown on the public site)."""

from __future__ import annotations

import json
from pathlib import Path

from scripts.builders_data import REPO_ROOT, sort_key

MULTIPLIERS_PATH = REPO_ROOT / "data" / "multipliers.json"


def load_multipliers(path: Path | None = None) -> dict[str, float]:
    path = path or MULTIPLIERS_PATH
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError(f"{path} must contain a JSON object")
    out: dict[str, float] = {}
    for k, v in data.items():
        out[str(k)] = float(v)
    return out


def save_multipliers(multipliers: dict[str, float], path: Path | None = None) -> None:
    path = path or MULTIPLIERS_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    ordered = dict(sorted(multipliers.items(), key=lambda kv: sort_key(kv[0])))
    with path.open("w", encoding="utf-8") as f:
        json.dump(ordered, f, indent=2, ensure_ascii=False)
        f.write("\n")


def _canonical_key(multipliers: dict[str, float], name: str) -> str | None:
    key = sort_key(name)
    for k in multipliers:
        if sort_key(k) == key:
            return k
    return None


def get_stored_multiplier(name: str, path: Path | None = None) -> float | None:
    multipliers = load_multipliers(path)
    canon = _canonical_key(multipliers, name)
    if canon is None:
        return None
    return multipliers[canon]


def set_stored_multiplier(
    name: str, multiplier: float, *, path: Path | None = None
) -> None:
    multipliers = load_multipliers(path)
    canon = _canonical_key(multipliers, name)
    store_as = canon if canon is not None else name
    multipliers[store_as] = float(multiplier)
    save_multipliers(multipliers, path)


def resolve_multiplier(
    builder_name: str,
    multiplier: float | None,
    *,
    override: bool,
    path: Path | None = None,
) -> tuple[float, bool]:
    """Return (multiplier to use, should_persist). Raises ValueError on policy violations."""
    stored = get_stored_multiplier(builder_name, path)

    if multiplier is None:
        if stored is None:
            raise ValueError(
                f"No saved multiplier for {builder_name!r}. "
                "Pass --multiplier for a new builder."
            )
        return stored, False

    if multiplier <= 0:
        raise ValueError("Multiplier must be positive.")

    if stored is None:
        return multiplier, True

    if abs(stored - multiplier) > 1e-9 and not override:
        raise ValueError(
            f"Saved multiplier for {builder_name!r} is {stored}; "
            f"you passed {multiplier}. Use --override-multiplier to change it."
        )

    if abs(stored - multiplier) > 1e-9 and override:
        return multiplier, True

    return multiplier, False
