"""
Base de données VRTF : propriétés des gaz/combustibles (via combustion)
et propriétés de l'acier BISRA (données spécifiques VRTF).
"""

import csv
from pathlib import Path

# Gaz, combustibles, comburants → package combustion
from combustion.database import (  # noqa: F401
    fuel_names, fuel_composition,
    comburant_names, comburant_composition,
    gas_names, gas_props,
)

_DATA_DIR = Path(__file__).parent.parent / "data"


# ---------------------------------------------------------------------------
# Parseurs CSV internes (acier BISRA)
# ---------------------------------------------------------------------------

def _read_csv(filename: str) -> list[list[str]]:
    path = _DATA_DIR / filename
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def _safe(val: str) -> float:
    s = val.strip()
    return float(s) if s else 0.0


def _parse_bisra_table(filename: str) -> dict[str, tuple]:
    rows = _read_csv(filename)
    grades = [c.strip() for c in rows[1][1:] if c.strip()]
    T_dict = {g: [] for g in grades}
    v_dict = {g: [] for g in grades}
    for row in rows[2:]:
        if not row or not row[0].strip():
            break
        try:
            T = float(row[0])
        except ValueError:
            break
        for i, g in enumerate(grades):
            raw = row[i + 1].strip() if i + 1 < len(row) else ""
            if raw:
                T_dict[g].append(T)
                v_dict[g].append(float(raw))
    return {g: (T_dict[g], v_dict[g]) for g in grades}


def _parse_bisra_scalar(filename: str) -> dict[str, float]:
    rows = _read_csv(filename)
    grades = [c.strip() for c in rows[1][1:] if c.strip()]
    vals = rows[2][1:]
    return {g: float(vals[i]) for i, g in enumerate(grades) if i < len(vals) and vals[i].strip()}


# ---------------------------------------------------------------------------
# Cache BISRA
# ---------------------------------------------------------------------------

_cache: dict[str, object] = {}


def _get(key: str, loader):
    if key not in _cache:
        _cache[key] = loader()
    return _cache[key]


def _bisra_density_data():
    return _get("bisra_rho", lambda: _parse_bisra_scalar("BisraRo.csv"))

def _bisra_cp_data():
    return _get("bisra_cp",  lambda: _parse_bisra_table("BisraCp.csv"))

def _bisra_co_data():
    return _get("bisra_co",  lambda: _parse_bisra_table("BisraCo.csv"))

def _bisra_eps_data():
    return _get("bisra_eps", lambda: _parse_bisra_table("BisraEps.csv"))

def _bisra_h_data():
    return _get("bisra_h",   lambda: _parse_bisra_table("Bisrah.csv"))


# ---------------------------------------------------------------------------
# API publique — acier BISRA
# ---------------------------------------------------------------------------

def bisra_grades() -> list[str]:
    return list(_bisra_density_data().keys())


def bisra_density(grade: str) -> float:
    data = _bisra_density_data()
    if grade not in data:
        raise KeyError(f"Grade BISRA inconnu : {grade!r}. Disponibles : {list(data)}")
    return data[grade]


def bisra_cp_table(grade: str) -> tuple:
    return _bisra_cp_data()[grade]


def bisra_co_table(grade: str) -> tuple:
    return _bisra_co_data()[grade]


def bisra_eps_table(grade: str) -> tuple:
    return _bisra_eps_data()[grade]


def bisra_enthalpy_table(grade: str) -> tuple:
    return _bisra_h_data()[grade]
