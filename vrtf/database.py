"""
Chargement des données depuis les CSV et fonctions de lookup.
Remplace ModuleDatabase.bas et la lecture des feuilles Excel de référence.

Format des CSV BISRA (tableaux T-dépendants, ex. BisraCo.csv) :
  Ligne 0 : [n_points, titre, ...]
  Ligne 1 : [n_grades, "BISRA 01", ..., "BISRA 22"]
  Lignes 2+ : [T_K, val_01, ..., val_22]

Format CSV BISRA scalaire (BisraRo.csv) :
  Ligne 0 : ["", titre, ...]
  Ligne 1 : [n_grades, "BISRA 01", ..., "BISRA 22"]
  Ligne 2 : ["", rho_01, ..., rho_22]

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

DATA_DIR = Path(__file__).parent.parent / "data"


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


def _parse_bisra_table(filename: str) -> dict[str, tuple]:
    """
    Retourne dict{ grade -> (T_array, val_array) }.
    S'arrête à la première ligne vide (certains CSV ont une seconde section
    de polynômes après la table principale).
    """
    rows = _read_csv(filename)
    grades = [c.strip() for c in rows[1][1:] if c.strip()]
    T_dict = {g: [] for g in grades}
    v_dict = {g: [] for g in grades}
    for row in rows[2:]:
        # Ligne vide ou col[0] non numérique → fin de la table de données
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
    """Retourne dict{ grade -> valeur scalaire } (ex. densité)."""
    rows = _read_csv(filename)
    grades = [c.strip() for c in rows[1][1:] if c.strip()]
    vals = rows[2][1:]
    return {g: float(vals[i]) for i, g in enumerate(grades) if i < len(vals) and vals[i].strip()}


def _parse_composition_table(filename: str) -> dict[str, dict[str, float]]:
    """Retourne dict{ nom_combustible -> dict{ formule_gaz -> % vol } }."""
    rows = _read_csv(filename)
    names = [c.strip() for c in rows[1][1:] if c.strip()]
    compo: dict[str, dict[str, float]] = {n: {} for n in names}
    for row in rows[2:]:
        if not row or not row[0].strip():
            continue
        gas = row[0].strip()
        for i, fuel in enumerate(names):
            raw = row[i + 1].strip() if i + 1 < len(row) else ""
            if raw:
                compo[fuel][gas] = float(raw)
    return compo


def _parse_gas_props() -> dict[str, dict]:
    """
    Retourne dict{ formule -> propriétés } et dict{ nom_français -> propriétés }.
    Colonnes (0-based) :
      0=NOM  1=FORMULE  2=C  3=H  4=O  5=N  6=S
      7=M[g/mol]  8=Hf[J/kmol]  9=PCI[J/kmol]  10=Cp_273
      11..17 = a, b, c, d, e, f, g  (coefficients Cp(T) polynomial)
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
            "molar_mass": _safe(row[7]),   # g/mol = kg/kmol
            "hf":         _safe(row[8]),   # J/kmol
            "pci":        _safe(row[9]),   # J/kmol (PCI, valeur positive pour combustibles)
            "cp_273":     _safe(row[10]),  # J/kg.K à 273 K
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


def _bisra_density_data() -> dict[str, float]:
    return _get("bisra_rho", lambda: _parse_bisra_scalar("BisraRo.csv"))

def _bisra_cp_data() -> dict[str, tuple]:
    return _get("bisra_cp", lambda: _parse_bisra_table("BisraCp.csv"))

def _bisra_co_data() -> dict[str, tuple]:
    return _get("bisra_co", lambda: _parse_bisra_table("BisraCo.csv"))

def _bisra_eps_data() -> dict[str, tuple]:
    return _get("bisra_eps", lambda: _parse_bisra_table("BisraEps.csv"))

def _bisra_h_data() -> dict[str, tuple]:
    return _get("bisra_h", lambda: _parse_bisra_table("Bisrah.csv"))

def _fuels_data() -> dict[str, dict[str, float]]:
    return _get("fuels", lambda: _parse_composition_table("Fuels.csv"))

def _comburants_data() -> dict[str, dict[str, float]]:
    return _get("comburants", lambda: _parse_composition_table("Comburants.csv"))

def _gas_props_data() -> dict[str, dict]:
    return _get("gas_props", _parse_gas_props)


# ---------------------------------------------------------------------------
# API publique — acier BISRA
# ---------------------------------------------------------------------------

def bisra_grades() -> list[str]:
    """Liste de tous les grades BISRA disponibles."""
    return list(_bisra_density_data().keys())


def bisra_density(grade: str) -> float:
    """Masse volumique [kg/m³] du grade BISRA (constante)."""
    data = _bisra_density_data()
    if grade not in data:
        raise KeyError(f"Grade BISRA inconnu : {grade!r}. Disponibles : {list(data)}")
    return data[grade]


def bisra_cp_table(grade: str) -> tuple:
    """Retourne (T_array [K], Cp_array [J/kg.K]) pour le grade BISRA."""
    return _bisra_cp_data()[grade]


def bisra_co_table(grade: str) -> tuple:
    """Retourne (T_array [K], lambda_array [W/m.K]) pour le grade BISRA."""
    return _bisra_co_data()[grade]


def bisra_eps_table(grade: str) -> tuple:
    """Retourne (T_array [K], eps_array [-]) pour le grade BISRA."""
    return _bisra_eps_data()[grade]


def bisra_enthalpy_table(grade: str) -> tuple:
    """Retourne (T_array [K], h_array [kJ/kg]) pour le grade BISRA."""
    return _bisra_h_data()[grade]


# ---------------------------------------------------------------------------
# API publique — combustibles et comburants
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
    """Liste des formules de gaz disponibles dans basdo_gaz."""
    return [k for k, v in _gas_props_data().items() if k == v["formula"]]


def gas_props(key: str) -> dict:
    """
    Propriétés d'un gaz par formule (ex. 'CH4') ou nom français (ex. 'méthane').
    Retourne dict avec : nom, formula, C, H, O, N, S, molar_mass, hf, pci,
                         cp_273, cp_coeffs.
    """
    data = _gas_props_data()
    if key not in data:
        raise KeyError(f"Gaz inconnu : {key!r}. Formules disponibles : {gas_names()}")
    return data[key]
