"""
Bibliothèque de calculs de combustion — Kappa Dubois.

Usage rapide
------------
    from combustion import stoich_air_vol, adiabatic_temperature, lhv_vol

    va   = stoich_air_vol("Gaz_naturel")
    Tad  = adiabatic_temperature("Gaz_naturel", air_fuel_ratio=1.1)
    pci  = lhv_vol("Gaz_naturel")
"""

from .database import (
    fuel_names, fuel_composition,
    comburant_names, comburant_composition,
    gas_names, gas_props,
)
from .thermo import (
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
from .core import (
    stoich_air_vol, stoich_air_kg,
    flue_gas_volume_wet,
    flue_gas_composition, waste_gas_composition,
    flow_fuel,
    adiabatic_temperature, equivalent_temperature,
)
from .math_utils import interp, poly_cp, poly_enthalpy, bisect, linear_interp, polint3

__all__ = [
    # database
    "fuel_names", "fuel_composition",
    "comburant_names", "comburant_composition",
    "gas_names", "gas_props",
    # thermo
    "gas_cp", "gas_enthalpy",
    "mixture_molar_mass", "mixture_density", "mixture_cp",
    "mixture_enthalpy", "mixture_enthalpy_vol",
    "fuel_enthalpy_vol", "fuel_enthalpy_kg",
    "air_enthalpy_vol", "air_density",
    "lhv_vol_from_compo", "lhv_vol", "lhv_kg", "wobbe_index",
    "temperature_from_enthalpy", "heat_exchanger",
    "AIR_DEFAULT", "T_REF",
    # core (combustion)
    "stoich_air_vol", "stoich_air_kg",
    "flue_gas_volume_wet",
    "flue_gas_composition", "waste_gas_composition",
    "flow_fuel",
    "adiabatic_temperature", "equivalent_temperature",
    # math
    "interp", "poly_cp", "poly_enthalpy", "bisect", "linear_interp", "polint3",
]
