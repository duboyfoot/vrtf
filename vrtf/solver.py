"""
Prétraitement de la combustion par zone : débits, enthalpies, polynômes.
Remplace ModuleSover.bas.
"""

from . import database as db
from . import thermo
from .combustion import flue_gas_composition, stoich_air_vol, flue_gas_volume_wet
from .math_utils import polint3, bisect

T_REF = 273.0

# 4 températures de référence pour les polynômes d'enthalpie [K]
_T4 = [273.0, 673.0, 1073.0, 1473.0]

# 4 températures de zone pour les polynômes régénérateur [K] : 500, 800, 1100, 1400°C
_T4_ZONE = [773.0, 1073.0, 1373.0, 1673.0]


def _teq_ab(compo_a, m_a, T_a_K, compo_b, m_b):
    """
    Température équivalente de B telle que m_b*h_b(T_eq) = m_a*h_a(T_a).
    Remplace CalculTeqAB VBA.
    """
    h_target = (m_a / m_b) * thermo.mixture_enthalpy(compo_a, T_a_K)
    def f(T):
        return thermo.mixture_enthalpy(compo_b, T) - h_target
    return bisect(f, T_REF, 2500.0, tol=max(abs(h_target) * 1e-4, 1.0))


def solve_zone(zone: dict) -> dict:
    """
    Prétraitement d'une zone brûleur.

    Entrée (dict) :
      fuel               str | dict  combustible (nom ou {formule: %})
      air                str | dict  comburant   (nom ou {formule: %})
      excess_air_pct     float  excès d'air [%]
      power_W            float  puissance installée [W]
      T_fuel_K           float  température combustible entrée [K]
      T_air_K            float  température comburant entrée [K]
      regen_air_eff_pct  float  efficacité régénérateur air [%]  (0 = sans)
      regen_air_aspi_pct float  taux d'aspiration fumées rég. air [%]

    Retourne un dict de résultats par zone.
    """
    fuel_name = zone["fuel"]
    air_name  = zone.get("air", "Air_sec")
    excess_air_pct   = zone.get("excess_air_pct", 0.0)
    power_W          = zone.get("power_W", 0.0)
    T_fuel_K         = zone.get("T_fuel_K", T_REF)
    T_air_K          = zone.get("T_air_K",  T_REF)
    regen_air_eff    = zone.get("regen_air_eff_pct",  0.0)
    regen_air_aspi   = zone.get("regen_air_aspi_pct", 0.0)

    fuel_c = db.fuel_composition(fuel_name) if isinstance(fuel_name, str) else fuel_name
    air_c  = db.comburant_composition(air_name) if isinstance(air_name, str) else air_name
    afr = 1.0 + 0.01 * excess_air_pct

    # ── Débits ──────────────────────────────────────────────────────────────
    pci_Nm3 = sum(
        pct / 100.0 * (db.gas_props(f).get("pci") or 0.0) / 22.4136
        for f, pct in fuel_c.items() if pct
    )
    q_fuel_Nm3s = power_W / pci_Nm3 if pci_Nm3 else 0.0

    va  = stoich_air_vol(fuel_c, air_c)
    vfh = flue_gas_volume_wet(fuel_c, air_c, afr)
    q_air_Nm3s  = q_fuel_Nm3s * afr * va
    q_flue_Nm3s = q_fuel_Nm3s * vfh

    rho_fuel = thermo.mixture_density(fuel_c, T_fuel_K)
    rho_air  = thermo.mixture_density(air_c,  T_REF)
    flue_c   = flue_gas_composition(fuel_c, air_c, afr)
    rho_flue = thermo.mixture_density(flue_c, T_REF) if any(flue_c.values()) else 1.3

    m_fuel_kgs = q_fuel_Nm3s * rho_fuel
    m_air_kgs  = q_air_Nm3s  * rho_air
    flue_kgs   = q_flue_Nm3s * rho_flue
    fuel_Nm3h  = q_fuel_Nm3s * 3600.0
    air_Nm3h   = q_air_Nm3s  * 3600.0

    # ── Polynômes d'enthalpie aux 4 températures de référence ───────────────
    def _h4(compo):
        return [thermo.mixture_enthalpy(compo, T) for T in _T4]

    poly_enthalp_air  = polint3(_T4, _h4(air_c))
    poly_enthalp_fuel = polint3(_T4, _h4(fuel_c))
    poly_enthalp_flue = polint3(_T4, _h4(flue_c))

    # ── Polynôme Cp des fumées ───────────────────────────────────────────────
    poly_cp_flue = polint3(_T4, [thermo.mixture_cp(flue_c, T) for T in _T4])

    # ── Masse volumique fumées à 0°C [kg/Nm³] ───────────────────────────────
    density_flue = rho_flue

    # ── Débits post-traitement ───────────────────────────────────────────────
    m_flue_rege_air_kgs = 0.0
    m_air_rege_kgs      = 0.0
    m_air_recu_kgs      = m_air_kgs
    m_fuel_recu_kgs     = m_fuel_kgs
    Q_flue_recu_Nm3h    = q_flue_Nm3s * 3600.0
    m_flue_recu_kgs     = flue_kgs
    Q_air_recu_Nm3h     = air_Nm3h

    if regen_air_aspi > 0.0:
        m_air_rege_kgs      = m_air_kgs
        m_flue_rege_air_kgs = regen_air_aspi / 100.0 * flue_kgs
        frac_recu           = 1.0 - regen_air_aspi / 100.0
        Q_flue_recu_Nm3h    = frac_recu * q_flue_Nm3s * 3600.0
        m_flue_recu_kgs     = frac_recu * flue_kgs
        m_air_recu_kgs      = 0.0
        Q_air_recu_Nm3h     = 0.0

    # ── Polynômes de températures (initialisation cas sans régénérateur) ─────
    poly_T_air_hot       = [T_air_K,  0.0, 0.0, 0.0]
    poly_T_fuel_hot      = [T_fuel_K, 0.0, 0.0, 0.0]
    poly_T_flue_cold_air = [0.0, 0.0, 0.0, 0.0]
    poly_T_eq_air        = [0.0, 0.0, 0.0, 0.0]
    poly_T_eq_fuel       = [0.0, 0.0, 0.0, 0.0]

    if m_fuel_kgs > 0.0:
        T_eq_fuel = _teq_ab(fuel_c, m_fuel_kgs, T_fuel_K, flue_c, m_fuel_kgs)
        poly_T_eq_fuel = [T_eq_fuel, 0.0, 0.0, 0.0]

    if regen_air_aspi == 0.0 and m_fuel_kgs > 0.0:
        T_eq_air = _teq_ab(air_c, m_air_kgs, T_air_K, flue_c, m_air_kgs)
        poly_T_eq_air = [T_eq_air, 0.0, 0.0, 0.0]

    # ── Régénérateur sur l'air ───────────────────────────────────────────────
    if regen_air_aspi > 0.0:
        pts_air_hot   = []
        pts_flue_cold = []
        pts_eq_air    = []

        for T_zone in _T4_ZONE:
            T_flue_out, T_air_out = thermo.heat_exchanger(
                flue_c, m_flue_rege_air_kgs, T_zone,
                air_c,  m_air_rege_kgs,      T_air_K,
                regen_air_eff, rendement_pct=100.0,
            )
            pts_air_hot.append(T_air_out)
            pts_flue_cold.append(T_flue_out)
            T_eq = _teq_ab(air_c, m_air_rege_kgs, T_air_out, flue_c, m_air_rege_kgs)
            pts_eq_air.append(T_eq)

        poly_T_air_hot       = polint3(_T4_ZONE, pts_air_hot)
        poly_T_flue_cold_air = polint3(_T4_ZONE, pts_flue_cold)
        poly_T_eq_air        = polint3(_T4_ZONE, pts_eq_air)

    return {
        "fuel": fuel_name,
        "air": air_name,
        "excess_air_pct": excess_air_pct,
        "power_W": power_W,
        "T_fuel_K": T_fuel_K,
        "T_air_K": T_air_K,
        "regen_air_eff_pct": regen_air_eff,
        "regen_air_aspi_pct": regen_air_aspi,
        "fuel_Nm3h": fuel_Nm3h,
        "air_Nm3h": air_Nm3h,
        "flue_kgs": flue_kgs,
        "m_air_kgs": m_air_kgs,
        "m_fuel_kgs": m_fuel_kgs,
        "flue_compo": flue_c,
        "density_flue": density_flue,
        "poly_enthalp_air": poly_enthalp_air,
        "poly_enthalp_fuel": poly_enthalp_fuel,
        "poly_enthalp_flue": poly_enthalp_flue,
        "poly_cp_flue": poly_cp_flue,
        "poly_T_air_hot": poly_T_air_hot,
        "poly_T_fuel_hot": poly_T_fuel_hot,
        "poly_T_flue_cold_air": poly_T_flue_cold_air,
        "poly_T_eq_air": poly_T_eq_air,
        "poly_T_eq_fuel": poly_T_eq_fuel,
        "Q_air_recu_Nm3h": Q_air_recu_Nm3h,
        "m_air_recu_kgs": m_air_recu_kgs,
        "m_fuel_recu_kgs": m_fuel_recu_kgs,
        "m_air_rege_kgs": m_air_rege_kgs,
        "m_flue_recu_kgs": m_flue_recu_kgs,
        "Q_flue_recu_Nm3h": Q_flue_recu_Nm3h,
        "m_flue_rege_air_kgs": m_flue_rege_air_kgs,
    }


def solve_combustion(zones: list) -> list:
    """
    Prétraitement de combustion pour une liste de zones.
    Remplace SolveCombustion VBA.
    """
    return [solve_zone(z) for z in zones]
