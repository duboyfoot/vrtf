"""Délègue au package combustion."""
from combustion.core import (  # noqa: F401
    AIR_DEFAULT, T_REF,
    stoich_air_vol, stoich_air_kg,
    flue_gas_volume_wet,
    flue_gas_composition, waste_gas_composition,
    flow_fuel,
    adiabatic_temperature, equivalent_temperature,
)
