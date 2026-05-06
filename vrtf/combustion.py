"""
Calculs de combustion : stœchiométrie, fumées, température adiabatique.
Remplace combustion_pat.bas, ModuleCombustion.bas et Enthalpyfluide.bas.
"""

import math

from . import database as db
from . import thermo
from .math_utils import bisect

AIR_DEFAULT = "Air_sec"
T_REF = 273.0


# ---------------------------------------------------------------------------
# Helpers internes
# ---------------------------------------------------------------------------

def _compo(name_or_dict, kind: str = "fuel") -> dict[str, float]:
    """Accepte un nom ou un dict de composition."""
    if isinstance(name_or_dict, dict):
        return name_or_dict
    if kind == "fuel":
        return db.fuel_composition(name_or_dict)
    return db.comburant_composition(name_or_dict)


def _vo2_gas(formula: str) -> float:
    """
    Volume d'O2 stœchiométrique [Nm³/Nm³] pour combustion complète du gaz pur.
    Formule : VO2 = C + H/4 - O/2 + S  (atomes par molécule).
    """
    p = db.gas_props(formula)
    return p["C"] + p["H"] / 4.0 - p["O"] / 2.0 + p["S"]


def _vco2_gas(formula: str) -> float:
    """CO2 produit [Nm³/Nm³] lors de la combustion complète."""
    return db.gas_props(formula)["C"]


def _vh2o_gas(formula: str) -> float:
    """H2O produit [Nm³/Nm³] lors de la combustion complète."""
    return db.gas_props(formula)["H"] / 2.0


def _vso2_gas(formula: str) -> float:
    """SO2 produit [Nm³/Nm³] lors de la combustion complète."""
    return db.gas_props(formula)["S"]


def _vn2_gas(formula: str) -> float:
    """N2 produit [Nm³/Nm³] lors de la combustion complète."""
    return db.gas_props(formula)["N"] / 2.0


# ---------------------------------------------------------------------------
# Stœchiométrie
# ---------------------------------------------------------------------------

def stoich_air_vol(fuel, air=AIR_DEFAULT) -> float:
    """
    Volume d'air stœchiométrique Va [Nm³_air / Nm³_fuel].
    Remplace GasVa0 / CalculVa VBA.
    """
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")
    sum_vo2 = sum(pct / 100.0 * _vo2_gas(f) for f, pct in fuel_c.items() if pct)
    o2_air = air_c.get("O2", 20.8) / 100.0
    return sum_vo2 / o2_air if o2_air else 0.0


def stoich_air_kg(fuel, air=AIR_DEFAULT) -> float:
    """Va [kg_air / kg_fuel]."""
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")
    va_vol = stoich_air_vol(fuel_c, air_c)
    rho_f  = thermo.mixture_density(fuel_c, T_REF)
    rho_a  = thermo.mixture_density(air_c,  T_REF)
    return va_vol * rho_a / rho_f if rho_f else 0.0


def flue_gas_volume_wet(fuel, air=AIR_DEFAULT, air_fuel_ratio: float = 1.0) -> float:
    """
    Volume de fumées humides Vfh [Nm³_fumées / Nm³_fuel].
    Remplace GasVf0 / CalculVfh VBA.
    air_fuel_ratio = excès_air (1.0 = stœchiométrie, 1.1 = 10 % excès).
    """
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")
    va = stoich_air_vol(fuel_c, air_c)

    vco2 = sum(pct / 100.0 * _vco2_gas(f) for f, pct in fuel_c.items() if pct)
    vh2o = sum(pct / 100.0 * _vh2o_gas(f) for f, pct in fuel_c.items() if pct)
    vso2 = sum(pct / 100.0 * _vso2_gas(f) for f, pct in fuel_c.items() if pct)
    vn2  = (
        sum(pct / 100.0 * _vn2_gas(f) for f, pct in fuel_c.items() if pct)
        + air_fuel_ratio * va * (air_c.get("N2", 79.2) / 100.0)
    )
    vo2_exc = (air_fuel_ratio - 1.0) * va * (air_c.get("O2", 20.8) / 100.0)
    return vco2 + vh2o + vso2 + vn2 + max(0.0, vo2_exc)


# ---------------------------------------------------------------------------
# Composition des fumées
# ---------------------------------------------------------------------------

def flue_gas_composition(fuel, air=AIR_DEFAULT, air_fuel_ratio: float = 1.0) -> dict[str, float]:
    """
    Composition volumique [%] des fumées pour un ratio air/fuel donné.
    Remplace CalculFracVolCO2/H2O/N2/O2/SO2 VBA.
    Retourne dict{ 'CO2', 'H2O', 'O2', 'N2', 'SO2', 'CO', 'H2' }.
    """
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")
    va = stoich_air_vol(fuel_c, air_c)
    vfh = flue_gas_volume_wet(fuel_c, air_c, air_fuel_ratio)

    vco2 = sum(pct / 100.0 * _vco2_gas(f) for f, pct in fuel_c.items() if pct)
    vh2o = sum(pct / 100.0 * _vh2o_gas(f) for f, pct in fuel_c.items() if pct)
    vso2 = sum(pct / 100.0 * _vso2_gas(f) for f, pct in fuel_c.items() if pct)
    vn2  = (
        sum(pct / 100.0 * _vn2_gas(f) for f, pct in fuel_c.items() if pct)
        + air_fuel_ratio * va * (air_c.get("N2", 79.2) / 100.0)
    )
    vo2 = max(0.0, (air_fuel_ratio - 1.0) * va * (air_c.get("O2", 20.8) / 100.0))

    if vfh == 0:
        return {"CO2": 0, "H2O": 0, "O2": 0, "N2": 0, "SO2": 0, "CO": 0, "H2": 0}
    return {
        "CO2": 100.0 * vco2 / vfh,
        "H2O": 100.0 * vh2o / vfh,
        "O2":  100.0 * vo2  / vfh,
        "N2":  100.0 * vn2  / vfh,
        "SO2": 100.0 * vso2 / vfh,
        "CO":  0.0,
        "H2":  0.0,
    }


def waste_gas_composition(
    fuel,
    air=AIR_DEFAULT,
    air_fuel_ratio: float = 1.0,
    T_equil: float = 20.0,
) -> dict[str, float]:
    """
    Composition des fumées [%] avec équilibre eau-gaz CO2+H2 ⇌ CO+H2O.
    Remplace WasteGasCompo VBA.

    K(T) = exp(3.75 - 4075 / (T_equil + 273))   [direction : CO2+H2 → CO+H2O]

    Pour air_fuel_ratio ≥ 1 : pas de CO ni H2, O2 en excès.
    Pour air_fuel_ratio < 1  : résolution d'une équation du second degré.
    """
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")
    va = stoich_air_vol(fuel_c, air_c)

    # Produits de combustion complète [Nm³/Nm³_fuel]
    a = sum(pct / 100.0 * _vco2_gas(f) for f, pct in fuel_c.items() if pct)   # CO2 pot.
    b = sum(pct / 100.0 * _vh2o_gas(f) for f, pct in fuel_c.items() if pct)   # H2O pot.
    vso2 = sum(pct / 100.0 * _vso2_gas(f) for f, pct in fuel_c.items() if pct)
    vn2  = (
        sum(pct / 100.0 * _vn2_gas(f) for f, pct in fuel_c.items() if pct)
        + air_fuel_ratio * va * (air_c.get("N2", 79.2) / 100.0)
    )

    o2_st   = va * (air_c.get("O2", 20.8) / 100.0)
    o2_avail = air_fuel_ratio * o2_st
    o2_needed = sum(pct / 100.0 * _vo2_gas(f) for f, pct in fuel_c.items() if pct)

    if air_fuel_ratio >= 1.0:
        # Excès d'air : combustion complète + O2 résiduel
        vo2_exc = o2_avail - o2_needed
        total = a + b + vso2 + vn2 + vo2_exc
        if total == 0:
            return {"CO2": 0, "H2O": 0, "O2": 0, "N2": 0, "SO2": 0, "CO": 0, "H2": 0}
        return {
            "CO2": 100.0 * a      / total,
            "H2O": 100.0 * b      / total,
            "O2":  100.0 * vo2_exc / total,
            "N2":  100.0 * vn2    / total,
            "SO2": 100.0 * vso2   / total,
            "CO":  0.0,
            "H2":  0.0,
        }

    # --- Sous-stœchiométrique : résolution de l'équilibre ---
    # déficit O2 : delta = o2_needed - o2_avail
    # bilan O2 : x/2 + y/2 = delta  →  x + y = 2*delta  (x=CO, y=H2)
    # équilibre : K*(a-x)*(b-y) = x*(b-y) impossible — on pose l'équilibre sur
    # les quantités finales CO, H2O, CO2, H2 :
    #   K = CO * H2O / (CO2 * H2)  avec CO2=(a-x), H2O=(b-y), CO=x, H2=y
    #   y = 2*delta - x
    #   K*(a-x)*(b-(2*delta-x)) = x*(b-(2*delta-x))  ... développé ci-dessous
    delta = o2_needed - o2_avail
    K = math.exp(3.75 - 4075.0 / (T_equil + 273.0))

    # Développement de K*(a-x)*(2*delta-x) = x*(b-2*delta+x) avec y=2*delta-x:
    #   (K-1)*x² + (-(K*a) - 2*delta*(K-1) - b)*x + K*2*a*delta = 0
    A_coef = K - 1.0
    B_coef = -(K * a) - 2.0 * delta * (K - 1.0) - b
    C_coef = K * 2.0 * a * delta

    if abs(A_coef) < 1e-12:
        # K ≈ 1 : équation linéaire
        x_co = -C_coef / B_coef if abs(B_coef) > 1e-12 else delta
    else:
        discriminant = B_coef ** 2 - 4.0 * A_coef * C_coef
        discriminant = max(0.0, discriminant)
        sqrt_d = math.sqrt(discriminant)
        # Deux racines ; on prend celle dans [0, min(a, 2*delta)]
        x1 = (-B_coef + sqrt_d) / (2.0 * A_coef)
        x2 = (-B_coef - sqrt_d) / (2.0 * A_coef)
        valid = [
            x for x in (x1, x2)
            if 0.0 <= x <= min(a, 2.0 * delta) and 0.0 <= (2.0 * delta - x) <= b
        ]
        x_co = valid[0] if valid else delta   # fallback : tout le déficit en CO

    x_co = max(0.0, min(x_co, min(a, 2.0 * delta)))
    x_h2 = 2.0 * delta - x_co
    x_h2 = max(0.0, min(x_h2, b))

    vco2 = a - x_co
    vh2o = b - x_h2
    total = vco2 + vh2o + vso2 + vn2 + x_co + x_h2
    if total == 0:
        return {"CO2": 0, "H2O": 0, "O2": 0, "N2": 0, "SO2": 0, "CO": 0, "H2": 0}
    return {
        "CO2": 100.0 * vco2 / total,
        "H2O": 100.0 * vh2o / total,
        "O2":  0.0,
        "N2":  100.0 * vn2  / total,
        "SO2": 100.0 * vso2 / total,
        "CO":  100.0 * x_co / total,
        "H2":  100.0 * x_h2 / total,
    }


# ---------------------------------------------------------------------------
# Débits
# ---------------------------------------------------------------------------

def flow_fuel(power_W: float, fuel, air=AIR_DEFAULT, air_fuel_ratio: float = 1.0,
              T_fuel: float = T_REF) -> dict[str, float]:
    """
    Débits à partir de la puissance thermique.
    Retourne dict{ 'fuel_Nm3h', 'air_Nm3h', 'air_kgs', 'fuel_kgs', 'flue_kgs' }.
    Remplace CalculDebitFuel/Air/Fumees VBA.
    """
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")

    pci = thermo.lhv_vol(fuel if isinstance(fuel, str) else "__dict__")
    if isinstance(fuel, dict):
        # Calcule le PCI depuis la composition directement
        pci = sum(
            pct / 100.0 * db.gas_props(f)["pci"] / 22.4136
            for f, pct in fuel_c.items() if pct and db.gas_props(f)["pci"]
        )

    q_fuel_Nm3s = power_W / pci if pci else 0.0     # [Nm³/s]
    va = stoich_air_vol(fuel_c, air_c)
    q_air_Nm3s  = q_fuel_Nm3s * air_fuel_ratio * va
    q_flue_Nm3s = q_fuel_Nm3s * flue_gas_volume_wet(fuel_c, air_c, air_fuel_ratio)

    rho_fuel = thermo.mixture_density(fuel_c, T_fuel)
    rho_air  = thermo.mixture_density(air_c,  T_REF)
    vfh = flue_gas_volume_wet(fuel_c, air_c, air_fuel_ratio)
    flue_c = waste_gas_composition(fuel_c, air_c, air_fuel_ratio)
    rho_flue = thermo.mixture_density(flue_c, T_REF) if any(flue_c.values()) else 1.3

    return {
        "fuel_Nm3h":  q_fuel_Nm3s * 3600.0,
        "air_Nm3h":   q_air_Nm3s  * 3600.0,
        "fuel_kgs":   q_fuel_Nm3s * rho_fuel,
        "air_kgs":    q_air_Nm3s  * rho_air,
        "flue_kgs":   q_flue_Nm3s * rho_flue,
    }


# ---------------------------------------------------------------------------
# Températures adiabatique et équivalente
# ---------------------------------------------------------------------------

def adiabatic_temperature(
    fuel, air=AIR_DEFAULT, air_fuel_ratio: float = 1.0,
    T_fuel: float = T_REF, T_air: float = T_REF,
) -> float:
    """
    Température de flamme adiabatique [K].
    Remplace CalculTadiab VBA (dichotomie sur bilan enthalpique).
    """
    fuel_c = _compo(fuel, "fuel")
    air_c  = _compo(air,  "air")
    va = stoich_air_vol(fuel_c, air_c)
    flue_c = waste_gas_composition(fuel_c, air_c, air_fuel_ratio)

    # Enthalpie entrante = PCI + sensible fuel + sensible air
    pci = sum(
        pct / 100.0 * db.gas_props(f)["pci"] / 22.4136
        for f, pct in fuel_c.items() if pct and db.gas_props(f)["pci"]
    )   # J/Nm³
    h_fuel_in = thermo.mixture_enthalpy_vol(fuel_c, T_fuel)
    h_air_in  = thermo.mixture_enthalpy_vol(air_c,  T_air) * air_fuel_ratio * va

    rho_flue = thermo.mixture_density(flue_c, T_REF) if any(flue_c.values()) else 1.3
    vfh = flue_gas_volume_wet(fuel_c, air_c, air_fuel_ratio)

    h_target_vol = (pci + h_fuel_in + h_air_in) / vfh if vfh else 0.0

    def f(T):
        return thermo.mixture_enthalpy(flue_c, T) * rho_flue - h_target_vol

    return bisect(f, 300.0, 2500.0, tol=h_target_vol * 1e-4 or 1.0)


def equivalent_temperature(
    compo_A: dict[str, float], flow_A: float,
    compo_B: dict[str, float], flow_B: float,
    T_A: float, T_B: float,
) -> float:
    """
    Température équivalente de deux débits mélangés [K].
    Remplace CalculTeqAB VBA.
    flow_A, flow_B en [kg/s] ou [Nm³/s] (même unité).
    """
    h_A = thermo.mixture_enthalpy(compo_A, T_A) * flow_A
    h_B = thermo.mixture_enthalpy(compo_B, T_B) * flow_B
    total_flow = flow_A + flow_B
    if total_flow == 0:
        return (T_A + T_B) / 2.0
    # Composition mélangée (pondération volumique/massique approximative)
    all_gases = set(compo_A) | set(compo_B)
    compo_mix = {}
    for g in all_gases:
        xa = compo_A.get(g, 0.0) * flow_A
        xb = compo_B.get(g, 0.0) * flow_B
        compo_mix[g] = (xa + xb) / total_flow
    h_mix = (h_A + h_B) / total_flow

    def f(T):
        return thermo.mixture_enthalpy(compo_mix, T) - h_mix

    return bisect(f, min(T_A, T_B) - 10, max(T_A, T_B) + 10, tol=abs(h_mix) * 1e-4 or 1.0)
