"""
Calculs thermodynamiques : enthalpie, Cp, densité des gaz et de l'acier.
Remplace Enthalpyfluide.bas et les fonctions de ModuleCombustion.bas.
"""

from . import database as db
from .math_utils import poly_cp, poly_enthalpy, bisect, interp

# Comburant par défaut (air sec standard)
AIR_DEFAULT = "Air_sec"

# Température de référence [K]
T_REF = 273.0

# Constante des gaz parfaits [J/(kmol.K)]  — utile pour densité
R_UNIVERSAL = 8314.0


# ---------------------------------------------------------------------------
# Gaz purs
# ---------------------------------------------------------------------------

def gas_cp(formula: str, T_K: float) -> float:
    """Cp [J/kg.K] d'un gaz pur à T_K, évalué par le polynôme de basdo_gaz."""
    return poly_cp(db.gas_props(formula)["cp_coeffs"], T_K)


def gas_enthalpy(formula: str, T_K: float, T_ref: float = T_REF) -> float:
    """Enthalpie spécifique [J/kg] d'un gaz pur, intégrée de T_ref à T_K."""
    return poly_enthalpy(db.gas_props(formula)["cp_coeffs"], T_K, T_ref)


# ---------------------------------------------------------------------------
# Mélanges de gaz
# ---------------------------------------------------------------------------

def _mass_fractions(composition: dict[str, float]) -> dict[str, float]:
    """
    Convertit une composition volumique [%] en fractions massiques.
    composition = {formule: % vol}.
    """
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
    """
    Masse volumique [kg/m³] d'un mélange (gaz idéal).
    composition = {formule: % vol}.
    """
    M = mixture_molar_mass(composition)          # g/mol
    # rho_0 à 0°C, 101325 Pa = M/22.4136 [kg/m³]
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
    """Enthalpie volumique [J/m³_n] d'un mélange (référence : conditions normales 0°C)."""
    rho_n = mixture_density(composition, T_ref)
    return mixture_enthalpy(composition, T_K, T_ref) * rho_n


# ---------------------------------------------------------------------------
# Combustibles et comburants nommés
# ---------------------------------------------------------------------------

def fuel_enthalpy_vol(fuel_name: str, T_K: float) -> float:
    """Enthalpie × densité [J/Nm³] d'un combustible nommé."""
    compo = db.fuel_composition(fuel_name)
    return mixture_enthalpy_vol(compo, T_K)


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
    rho_n = mixture_density(compo, T_REF)   # kg/Nm³ à conditions normales
    return lhv_vol(fuel_name) / rho_n if rho_n else 0.0


def wobbe_index(fuel_name: str, air_name: str = AIR_DEFAULT) -> float:
    """Indice de Wobbe [J/Nm³] = PCI_vol / sqrt(densité_relative)."""
    compo = db.fuel_composition(fuel_name)
    air_compo = db.comburant_composition(air_name)
    rho_fuel = mixture_density(compo, T_REF)
    rho_air  = mixture_density(air_compo, T_REF)
    d = rho_fuel / rho_air   # densité relative
    return lhv_vol(fuel_name) / (d ** 0.5)


# ---------------------------------------------------------------------------
# Acier BISRA — propriétés interpolées
# ---------------------------------------------------------------------------

def steel_density(grade: str) -> float:
    """Masse volumique [kg/m³] de l'acier (constante)."""
    return db.bisra_density(grade)


def steel_cp(grade: str, T_K: float) -> float:
    """Chaleur massique [J/kg.K] de l'acier au grade BISRA à T_K."""
    T_arr, cp_arr = db.bisra_cp_table(grade)
    return interp(T_K, T_arr, cp_arr)


def steel_conductivity(grade: str, T_K: float) -> float:
    """Conductivité thermique [W/m.K] de l'acier BISRA à T_K."""
    T_arr, co_arr = db.bisra_co_table(grade)
    return interp(T_K, T_arr, co_arr)


def steel_emissivity(grade: str, T_K: float) -> float:
    """Émissivité [-] de l'acier BISRA à T_K."""
    T_arr, eps_arr = db.bisra_eps_table(grade)
    return interp(T_K, T_arr, eps_arr)


def steel_enthalpy(grade: str, T_K: float) -> float:
    """Enthalpie [kJ/kg] de l'acier BISRA à T_K (référence table = 273 K)."""
    T_arr, h_arr = db.bisra_enthalpy_table(grade)
    return interp(T_K, T_arr, h_arr)


# ---------------------------------------------------------------------------
# Résolution inverse : température depuis l'enthalpie
# ---------------------------------------------------------------------------

def temperature_from_enthalpy(
    composition: dict[str, float],
    target_h_vol: float,
    T_min: float = 273.0,
    T_max: float = 2273.0,
) -> float:
    """
    Inverse de mixture_enthalpy_vol : trouve T [K] tel que H_vol(T) = target_h_vol.
    Remplace TWasteGasEnthalp / GasTemp VBA (dichotomie).
    """
    def f(T):
        return mixture_enthalpy_vol(composition, T) - target_h_vol
    return bisect(f, T_min, T_max, tol=abs(target_h_vol) * 1e-4 or 1.0)


def heat_exchanger(
    hot_compo: dict[str, float], m_hot_kgs: float, T_hot_in_K: float,
    cold_compo: dict[str, float], m_cold_kgs: float, T_cold_in_K: float,
    effectiveness_pct: float, rendement_pct: float = 100.0,
) -> tuple:
    """
    Températures de sortie d'un échangeur (régénérateur).
    Remplace CalculEchangeur VBA.

    Toutes les températures en [K], débits en [kg/s].
    effectiveness_pct : efficacité [%] — fraction de l'échange maximal réalisé.
    rendement_pct     : rendement thermique [%] — fraction cédée au fluide froid.

    Retourne : (T_hot_out_K, T_cold_out_K)
    """
    eff  = effectiveness_pct / 100.0
    rend = rendement_pct     / 100.0

    def _dh(compo, T_high, T_low):
        return mixture_enthalpy(compo, T_high, T_low)

    P_hot_max  = m_hot_kgs  * _dh(hot_compo,  T_hot_in_K, T_cold_in_K)
    P_cold_max = m_cold_kgs * _dh(cold_compo, T_hot_in_K, T_cold_in_K)
    P_target   = eff * min(P_hot_max, P_cold_max)

    # Etape 1 : trouver T_hot_out tel que Pchaud = P_target
    # Cas limite : P_target >= P_hot_max → chaud refroidi jusqu'à T_cold_in
    if P_hot_max <= 0.0 or P_target >= P_hot_max * (1.0 - 1e-9):
        T_hot_out = T_cold_in_K
    else:
        def f_hot(T_out):
            return m_hot_kgs * _dh(hot_compo, T_hot_in_K, T_out) - P_target
        T_hot_out = bisect(f_hot, T_cold_in_K, T_hot_in_K, tol=1.0)

    # Etape 2 : trouver T_cold_out tel que Pabsorbee = rend x Pchaud
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
