"""
Calculs thermodynamiques pour les mélanges gazeux :
enthalpie, Cp, densité, PCI, échangeur.
"""

from . import database as db
from .math_utils import poly_cp, poly_enthalpy, bisect

AIR_DEFAULT = "Air_sec"
T_REF = 273.0
R_UNIVERSAL = 8314.0


# ---------------------------------------------------------------------------
# Gaz purs
# ---------------------------------------------------------------------------

def gas_cp(formula: str, T_K: float) -> float:
    """Cp [J/kg.K] d'un gaz pur à T_K."""
    return poly_cp(db.gas_props(formula)["cp_coeffs"], T_K)


def gas_enthalpy(formula: str, T_K: float, T_ref: float = T_REF) -> float:
    """Enthalpie spécifique [J/kg] d'un gaz pur de T_ref à T_K."""
    return poly_enthalpy(db.gas_props(formula)["cp_coeffs"], T_K, T_ref)


# ---------------------------------------------------------------------------
# Mélanges de gaz
# ---------------------------------------------------------------------------

def _mass_fractions(composition: dict[str, float]) -> dict[str, float]:
    """Convertit une composition volumique [%] en fractions massiques."""
    total_M = sum(
        pct / 100.0 * db.gas_props(f)["molar_mass"]
        for f, pct in composition.items() if pct
    )
    if total_M == 0:
        return {}
    return {
        f: (pct / 100.0 * db.gas_props(f)["molar_mass"]) / total_M
        for f, pct in composition.items() if pct
    }


def mixture_molar_mass(composition: dict[str, float]) -> float:
    """Masse molaire [g/mol] d'un mélange. composition = {formule: % vol}."""
    return sum(
        pct / 100.0 * db.gas_props(f)["molar_mass"]
        for f, pct in composition.items() if pct
    )


def mixture_density(composition: dict[str, float], T_K: float, P: float = 101325.0) -> float:
    """Masse volumique [kg/m³] d'un mélange gazeux (gaz idéal)."""
    M = mixture_molar_mass(composition)
    return (M / 22.4136) * (273.0 / T_K) * (P / 101325.0)


def mixture_cp(composition: dict[str, float], T_K: float) -> float:
    """Cp [J/kg.K] d'un mélange gazeux à T_K."""
    mf = _mass_fractions(composition)
    return sum(frac * gas_cp(f, T_K) for f, frac in mf.items())


def mixture_enthalpy(composition: dict[str, float], T_K: float, T_ref: float = T_REF) -> float:
    """Enthalpie spécifique [J/kg] d'un mélange gazeux de T_ref à T_K."""
    mf = _mass_fractions(composition)
    return sum(frac * gas_enthalpy(f, T_K, T_ref) for f, frac in mf.items())


def mixture_enthalpy_vol(composition: dict[str, float], T_K: float, T_ref: float = T_REF) -> float:
    """Enthalpie volumique [J/Nm³] d'un mélange (référence 0°C)."""
    rho_n = mixture_density(composition, T_ref)
    return mixture_enthalpy(composition, T_K, T_ref) * rho_n


# ---------------------------------------------------------------------------
# Combustibles et comburants nommés
# ---------------------------------------------------------------------------

def fuel_enthalpy_vol(fuel_name: str, T_K: float) -> float:
    """Enthalpie × densité [J/Nm³] d'un combustible nommé."""
    return mixture_enthalpy_vol(db.fuel_composition(fuel_name), T_K)


def fuel_enthalpy_kg(fuel_name: str, T_K: float) -> float:
    """Enthalpie spécifique [J/kg] d'un combustible nommé."""
    return mixture_enthalpy(db.fuel_composition(fuel_name), T_K)


def air_enthalpy_vol(T_K: float, air_name: str = AIR_DEFAULT) -> float:
    """Enthalpie × densité [J/Nm³] de l'air."""
    return mixture_enthalpy_vol(db.comburant_composition(air_name), T_K)


def air_density(T_K: float, air_name: str = AIR_DEFAULT, P: float = 101325.0) -> float:
    """Masse volumique [kg/m³] de l'air à T_K."""
    return mixture_density(db.comburant_composition(air_name), T_K, P)


# ---------------------------------------------------------------------------
# PCI (Pouvoir Calorifique Inférieur)
# ---------------------------------------------------------------------------

def lhv_vol_from_compo(compo: dict) -> float:
    """PCI [J/Nm³] d'une composition volumique {formule: % vol}."""
    total = 0.0
    for formula, pct in compo.items():
        if not pct:
            continue
        pci = db.gas_props(formula).get("pci") or 0.0
        total += (pct / 100.0) * pci / 22.4136
    return total


def lhv_vol(fuel_name: str) -> float:
    """PCI [J/Nm³] d'un combustible nommé."""
    return lhv_vol_from_compo(db.fuel_composition(fuel_name))


def lhv_kg(fuel_name: str) -> float:
    """PCI [J/kg] d'un combustible nommé."""
    compo = db.fuel_composition(fuel_name)
    rho_n = mixture_density(compo, T_REF)
    return lhv_vol(fuel_name) / rho_n if rho_n else 0.0


def wobbe_index(fuel_name: str, air_name: str = AIR_DEFAULT) -> float:
    """Indice de Wobbe [J/Nm³] = PCI_vol / sqrt(densité_relative)."""
    rho_fuel = mixture_density(db.fuel_composition(fuel_name), T_REF)
    rho_air  = mixture_density(db.comburant_composition(air_name), T_REF)
    d = rho_fuel / rho_air
    return lhv_vol(fuel_name) / (d ** 0.5)


# ---------------------------------------------------------------------------
# Résolution inverse : température depuis l'enthalpie
# ---------------------------------------------------------------------------

def temperature_from_enthalpy(
    composition: dict[str, float],
    target_h_vol: float,
    T_min: float = 273.0,
    T_max: float = 2273.0,
) -> float:
    """Trouve T [K] tel que mixture_enthalpy_vol(composition, T) = target_h_vol."""
    def f(T):
        return mixture_enthalpy_vol(composition, T) - target_h_vol
    return bisect(f, T_min, T_max, tol=abs(target_h_vol) * 1e-4 or 1.0)


# ---------------------------------------------------------------------------
# Échangeur de chaleur (récupérateur)
# ---------------------------------------------------------------------------

def heat_exchanger(
    hot_compo: dict[str, float], m_hot_kgs: float, T_hot_in_K: float,
    cold_compo: dict[str, float], m_cold_kgs: float, T_cold_in_K: float,
    effectiveness_pct: float, rendement_pct: float = 100.0,
) -> tuple:
    """
    Températures de sortie d'un échangeur (régénérateur).
    Retourne (T_hot_out_K, T_cold_out_K).
    """
    eff  = effectiveness_pct / 100.0
    rend = rendement_pct     / 100.0

    def _dh(compo, T_high, T_low):
        return mixture_enthalpy(compo, T_high, T_low)

    P_hot_max  = m_hot_kgs  * _dh(hot_compo,  T_hot_in_K, T_cold_in_K)
    P_cold_max = m_cold_kgs * _dh(cold_compo, T_hot_in_K, T_cold_in_K)
    P_target   = eff * min(P_hot_max, P_cold_max)

    if P_hot_max <= 0.0 or P_target >= P_hot_max * (1.0 - 1e-9):
        T_hot_out = T_cold_in_K
    else:
        def f_hot(T_out):
            return m_hot_kgs * _dh(hot_compo, T_hot_in_K, T_out) - P_target
        T_hot_out = bisect(f_hot, T_cold_in_K, T_hot_in_K, tol=1.0)

    P1 = rend * m_hot_kgs * _dh(hot_compo, T_hot_in_K, T_hot_out)
    if P1 <= 0.0:
        T_cold_out = T_cold_in_K
    elif P1 >= P_cold_max * (1.0 - 1e-9):
        T_cold_out = T_hot_in_K
    else:
        def f_cold(T_out):
            return m_cold_kgs * _dh(cold_compo, T_out, T_cold_in_K) - P1
        T_cold_out = bisect(f_cold, T_cold_in_K, T_hot_in_K, tol=1.0)

    return T_hot_out, T_cold_out
