"""
Post-traitement des résultats Thermette.
Remplace ModulePosttreatment.bas.
"""

import re
from pathlib import Path

from . import database as db
from . import thermo
from .combustion import flue_gas_composition, stoich_air_vol
from .thermo import lhv_vol, temperature_from_enthalpy

T_REF = 273.0

# Ligne de résultat Thermette : "NomSortie : valeur"
_OUTPUT_RE = re.compile(
    r"^(\S+)\s*:\s*([-+]?[0-9]+\.?[0-9]*(?:[eE][+-]?[0-9]+)?)"
)


# ---------------------------------------------------------------------------
# Lecture du fichier résultat
# ---------------------------------------------------------------------------

def parse_results(results_file: str | Path) -> dict[str, float]:
    """
    Lit et analyse le fichier de résultats VRTF_resu produit par Thermette.
    Lève RuntimeError si le calcul n'a pas convergé.
    Retourne un dict {nom_sortie: valeur}.
    """
    text = Path(results_file).read_text(encoding="latin-1")

    if "CALCUL TERMINE NORMALEMENT" not in text:
        raise RuntimeError(
            "Thermette n'a pas convergé — 'CALCUL TERMINE NORMALEMENT' absent du fichier résultat."
        )

    start = text.find("Valeurs des sorties")
    if start == -1:
        raise ValueError("Section 'Valeurs des sorties' introuvable dans le fichier résultat.")

    values: dict[str, float] = {}
    for line in text[start:].splitlines():
        m = _OUTPUT_RE.match(line.strip())
        if m:
            values[m.group(1)] = float(m.group(2))

    return values


# ---------------------------------------------------------------------------
# Helpers d'extraction
# ---------------------------------------------------------------------------

def _extract_ij(values: dict, prefix: str) -> dict[tuple, float]:
    """
    Extrait les entrées du dict de la forme prefix{i}_R{j} et
    retourne {(i, j): valeur}.
    """
    pat = re.compile(re.escape(prefix) + r"(\d+)_R(\d+)$")
    result: dict[tuple, float] = {}
    for key, val in values.items():
        m = pat.match(key)
        if m:
            result[(int(m.group(1)), int(m.group(2)))] = val
    return result


# ---------------------------------------------------------------------------
# Post-traitement complet
# ---------------------------------------------------------------------------

def postprocess(
    results_file: str | Path,
    cfg: dict,
    fuel_name: str,
    air_name: str = "Air_sec",
    excess_air_pct: float = 0.0,
) -> dict:
    """
    Post-traite les résultats Thermette.
    Remplace ThPostTreatment VBA (avec correction des bugs d'accumulation).

    Paramètres
    ----------
    results_file     : chemin vers VRTF_resu
    cfg              : même dict que pour write_thermette_files ; chaque entrée
                       de cfg["tubes"] peut contenir :
                         real_zone    int   numéro de zone réelle (défaut = section)
                         tube_eff_pct float rendement tube [%] (défaut = 100)
    fuel_name        : nom du combustible (base de données)
    air_name         : nom du comburant (défaut "Air_sec")
    excess_air_pct   : excès d'air [%] (défaut 0)

    Retour
    ------
    dict contenant :
      strip_temp_K        {(i,j): K}   température peau bande
      flux_mur_W          {(i,j): W}   flux conduction paroi
      flux_strip_W        {(i,j): W}   flux chaleur bande
      flux_tube_W         {(i,j): W}   flux chaleur tube (positif = vers tube)
      E_strip_in_W        W            puissance enthalpique bande entrée
      E_strip_out_W       W            puissance enthalpique bande sortie
      deb_gaz_Nm3h        {(i,j): Nm³/h}
      deb_air_Nm3h        {(i,j): Nm³/h}
      deb_fum_Nm3h        {(i,j): Nm³/h}
      flux_fum_W          {(i,j): W}   puissance sensible fumées
      T_fum_K             {(i,j): K}   température fumées sortie tube
      flux_tube_zone_W    {zone: W}    agrégé par zone réelle
      flux_mur_zone_W     {zone: W}
      flux_fum_zone_W     {zone: W}
      flux_tube_four_W    W            total four
      flux_mur_four_W     W
      flux_fum_four_W     W
    """
    raw = parse_results(results_file)

    # ── Extraire valeurs Thermette par (section, rangée) ─────────────────────
    strip_temp_K = _extract_ij(raw, "SStrip_S")

    # La dernière rangée de chaque zone n'a pas de branche bande dans Thermette
    # (n-1 branches pour n rangées). On utilise le nœud inter-zones, qui est
    # physiquement la température de bande à l'entrée de cette rangée.
    n_zones       = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]
    for i in range(1, n_zones + 1):
        last_j = rows_per_zone[i - 1]
        ij_last = (i, last_j)
        ij_prev = (i, last_j - 1)
        if last_j > 1 and ij_last not in strip_temp_K and ij_prev in strip_temp_K:
            strip_temp_K[ij_last] = strip_temp_K[ij_prev]
    flux_mur_W   = _extract_ij(raw, "SFluxMur_S")
    flux_strip_W = _extract_ij(raw, "SFluxStrip_S")

    # SFluxTube : signe Thermette inversé (négatif = chaleur cédée au tube)
    flux_tube_W  = {ij: -v for ij, v in _extract_ij(raw, "SFluxTube_S").items()}

    E_strip_in_W  = raw.get("S_Prod_IN",  0.0)
    E_strip_out_W = raw.get("S_Prod_OUT", 0.0)

    # ── Propriétés combustion (globales) ─────────────────────────────────────
    afr    = 1.0 + excess_air_pct / 100.0
    fuel_c = db.fuel_composition(fuel_name) if isinstance(fuel_name, str) else fuel_name
    air_c  = db.comburant_composition(air_name) if isinstance(air_name, str) else air_name
    flue_c = flue_gas_composition(fuel_c, air_c, afr)

    pci_Nm3  = lhv_vol(fuel_name) if isinstance(fuel_name, str) else thermo.lhv_vol_from_compo(fuel_c)
    va0      = stoich_air_vol(fuel_c, air_c)               # Nm³_air/Nm³_fuel
    rho_fuel = thermo.mixture_density(fuel_c, T_REF)       # kg/Nm³
    rho_air  = thermo.mixture_density(air_c,  T_REF)       # kg/Nm³
    rho_flue = thermo.mixture_density(flue_c, T_REF)       # kg/Nm³

    # ── Table tube : (section, rangée) → {real_zone, tube_eff_pct} ───────────
    tube_map: dict[tuple, dict] = {(t["zone"], t["row"]): t for t in cfg.get("tubes", [])}

    # ── Calculs dérivés par (section, rangée) ────────────────────────────────
    deb_gaz_Nm3h: dict[tuple, float] = {}
    deb_air_Nm3h: dict[tuple, float] = {}
    deb_fum_Nm3h: dict[tuple, float] = {}
    flux_fum_W:   dict[tuple, float] = {}
    T_fum_K:      dict[tuple, float] = {}

    flux_tube_zone_W: dict[int, float] = {}
    flux_mur_zone_W:  dict[int, float] = {}
    flux_fum_zone_W:  dict[int, float] = {}

    n_zones       = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]

    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1] + 1):
            ij        = (i, j)
            t         = tube_map.get(ij, {})
            real_zone = t.get("real_zone", i)
            eff       = t.get("tube_eff_pct", 100.0) / 100.0

            q_tube = flux_tube_W.get(ij, 0.0)
            q_mur  = flux_mur_W.get(ij, 0.0)

            # ── Débits [Nm³/s] ────────────────────────────────────────────
            if pci_Nm3 > 0.0 and eff > 0.0:
                deb_gaz = q_tube / pci_Nm3 / eff                          # Nm³/s
            else:
                deb_gaz = 0.0
            deb_air = deb_gaz * va0 * afr                                  # Nm³/s
            deb_fum = (
                (deb_gaz * rho_fuel + deb_air * rho_air) / rho_flue
                if rho_flue else 0.0
            )                                                              # Nm³/s

            deb_gaz_Nm3h[ij] = deb_gaz * 3600.0
            deb_air_Nm3h[ij] = deb_air * 3600.0
            deb_fum_Nm3h[ij] = deb_fum * 3600.0

            # ── Flux fumées et température ─────────────────────────────────
            q_fum = (1.0 - eff) / eff * q_tube if eff > 0.0 else 0.0
            flux_fum_W[ij] = q_fum

            if q_fum > 0.0 and deb_fum > 0.0:
                # h_vol [J/Nm³] = puissance sensible / débit volumique normal
                h_vol = q_fum / deb_fum
                T_fum_K[ij] = temperature_from_enthalpy(flue_c, h_vol)
            else:
                T_fum_K[ij] = T_REF

            # ── Agrégation par zone réelle ─────────────────────────────────
            flux_tube_zone_W[real_zone] = flux_tube_zone_W.get(real_zone, 0.0) + q_tube
            flux_mur_zone_W[real_zone]  = flux_mur_zone_W.get(real_zone, 0.0)  + q_mur
            flux_fum_zone_W[real_zone]  = flux_fum_zone_W.get(real_zone, 0.0)  + q_fum

    return {
        "strip_temp_K":     strip_temp_K,
        "flux_mur_W":       flux_mur_W,
        "flux_strip_W":     flux_strip_W,
        "flux_tube_W":      flux_tube_W,
        "E_strip_in_W":     E_strip_in_W,
        "E_strip_out_W":    E_strip_out_W,
        "deb_gaz_Nm3h":     deb_gaz_Nm3h,
        "deb_air_Nm3h":     deb_air_Nm3h,
        "deb_fum_Nm3h":     deb_fum_Nm3h,
        "flux_fum_W":       flux_fum_W,
        "T_fum_K":          T_fum_K,
        "flux_tube_zone_W": flux_tube_zone_W,
        "flux_mur_zone_W":  flux_mur_zone_W,
        "flux_fum_zone_W":  flux_fum_zone_W,
        "flux_tube_four_W": sum(flux_tube_zone_W.values()),
        "flux_mur_four_W":  sum(flux_mur_zone_W.values()),
        "flux_fum_four_W":  sum(flux_fum_zone_W.values()),
    }
