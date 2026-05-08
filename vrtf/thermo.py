"""
Thermodynamique VRTF : gaz (via combustion) + acier BISRA.
"""

# Fonctions gaz → package combustion
from combustion.thermo import (  # noqa: F401
    gas_cp, gas_enthalpy,
    mixture_molar_mass, mixture_density, mixture_cp,
    mixture_enthalpy, mixture_enthalpy_vol,
    fuel_enthalpy_vol, fuel_enthalpy_kg,
    air_enthalpy_vol, air_density,
    lhv_vol_from_compo, lhv_vol, lhv_kg, wobbe_index,
    temperature_from_enthalpy,
    heat_exchanger,
    AIR_DEFAULT, T_REF,
)

from . import database as db
from .math_utils import interp


# ---------------------------------------------------------------------------
# Acier BISRA — propriétés interpolées
# ---------------------------------------------------------------------------

def steel_density(grade: str) -> float:
    return db.bisra_density(grade)


def steel_cp(grade: str, T_K: float) -> float:
    T_arr, cp_arr = db.bisra_cp_table(grade)
    return interp(T_K, T_arr, cp_arr)


def steel_conductivity(grade: str, T_K: float) -> float:
    T_arr, co_arr = db.bisra_co_table(grade)
    return interp(T_K, T_arr, co_arr)


def steel_emissivity(grade: str, T_K: float) -> float:
    T_arr, eps_arr = db.bisra_eps_table(grade)
    return interp(T_K, T_arr, eps_arr)


def steel_enthalpy(grade: str, T_K: float) -> float:
    T_arr, h_arr = db.bisra_enthalpy_table(grade)
    return interp(T_K, T_arr, h_arr)
