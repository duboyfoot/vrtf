"""
Base de données combustibles, comburants et propriétés des gaz élémentaires.
Lecture des CSV stockés dans combustion/data/.

Format CSV combustibles / comburants (Fuels.csv, Comburants.csv) :
  Ligne 0 : [n_especes, ...]
  Ligne 1 : [n_combustibles, nom_1, nom_2, ...]
  Lignes 2+ : [formule_gaz, %_1, %_2, ...]

Format basdo_gaz.csv :
  Ligne 0 : titre
  Ligne 1 : descriptions colonnes
  Ligne 2 : en-têtes
  Lignes 3+ : [NOM, FORMULE, C, H, O, N, S, M, Hf, PCI, Cp273, a, b, c, d, e, f, g]
"""

import csv
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"


# ---------------------------------------------------------------------------
# Parseurs CSV internes
# ---------------------------------------------------------------------------

def _read_csv(filename: str) -> list[list[str]]:
    path = DATA_DIR / filename
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def _safe(val: str) -> float:
    s = val.strip()
    return float(s) if s else 0.0


def _parse_composition_table(filename: str) -> dict[str, dict[str, float]]:
    """Retourne dict{ nom -> dict{ formule_gaz -> % vol } }."""
    rows = _read_csv(filename)
    names = [c.strip() for c in rows[1][1:] if c.strip()]
    compo: dict[str, dict[str, float]] = {n: {} for n in names}
    for row in rows[2:]:
        if not row or not row[0].strip():
            continue
        gas = row[0].strip()
        for i, name in enumerate(names):
            raw = row[i + 1].strip() if i + 1 < len(row) else ""
            if raw:
                compo[name][gas] = float(raw)
    return compo


def _parse_gas_props() -> dict[str, dict]:
    """
    Retourne dict{ formule -> propriétés } indexé aussi par nom français.
    Colonnes : NOM, FORMULE, C, H, O, N, S, M, Hf, PCI, Cp273, a..g
    """
    rows = _read_csv("basdo_gaz.csv")
    props: dict[str, dict] = {}
    for row in rows[3:]:
        if not row or not row[0].strip():
            continue
        entry = {
            "nom":        row[0].strip(),
            "formula":    row[1].strip(),
            "C":          _safe(row[2]),
            "H":          _safe(row[3]),
            "O":          _safe(row[4]),
            "N":          _safe(row[5]),
            "S":          _safe(row[6]),
            "molar_mass": _safe(row[7]),
            "hf":         _safe(row[8]),
            "pci":        _safe(row[9]),
            "cp_273":     _safe(row[10]),
            "cp_coeffs":  [_safe(row[j]) for j in range(11, 18) if j < len(row)],
        }
        props[entry["formula"]] = entry
        props[entry["nom"].lower()] = entry
    return props


# ---------------------------------------------------------------------------
# Cache module-level (chargement paresseux)
# ---------------------------------------------------------------------------

_cache: dict[str, object] = {}


def _get(key: str, loader):
    if key not in _cache:
        _cache[key] = loader()
    return _cache[key]


def _fuels_data() -> dict[str, dict[str, float]]:
    return _get("fuels", lambda: _parse_composition_table("Fuels.csv"))


def _comburants_data() -> dict[str, dict[str, float]]:
    return _get("comburants", lambda: _parse_composition_table("Comburants.csv"))


def _gas_props_data() -> dict[str, dict]:
    return _get("gas_props", _parse_gas_props)


# ---------------------------------------------------------------------------
# API publique — combustibles
# ---------------------------------------------------------------------------

def fuel_names() -> list[str]:
    """Liste de tous les combustibles disponibles."""
    return list(_fuels_data().keys())


def fuel_composition(name: str) -> dict[str, float]:
    """Composition volumique [%] du combustible {formule_gaz: %}."""
    data = _fuels_data()
    if name not in data:
        raise KeyError(f"Combustible inconnu : {name!r}. Disponibles : {list(data)}")
    return dict(data[name])


# ---------------------------------------------------------------------------
# API publique — comburants
# ---------------------------------------------------------------------------

def comburant_names() -> list[str]:
    """Liste de tous les comburants disponibles."""
    return list(_comburants_data().keys())


def comburant_composition(name: str) -> dict[str, float]:
    """Composition volumique [%] du comburant {formule_gaz: %}."""
    data = _comburants_data()
    if name not in data:
        raise KeyError(f"Comburant inconnu : {name!r}. Disponibles : {list(data)}")
    return dict(data[name])


# ---------------------------------------------------------------------------
# API publique — propriétés des gaz élémentaires
# ---------------------------------------------------------------------------

def gas_names() -> list[str]:
    """Liste des formules de gaz disponibles."""
    return [k for k, v in _gas_props_data().items() if k == v["formula"]]


def gas_props(key: str) -> dict:
    """
    Propriétés d'un gaz par formule (ex. 'CH4') ou nom français.
    Retourne dict avec : nom, formula, C, H, O, N, S, molar_mass, hf, pci,
                         cp_273, cp_coeffs.
    """
    data = _gas_props_data()
    if key not in data:
        raise KeyError(f"Gaz inconnu : {key!r}. Formules disponibles : {gas_names()}")
    return data[key]
