"""
Calcul des propriétés clés d'un couple combustible/comburant.
Remplace ModuleUpdate.bas (UpdateFuelAir VBA).

La version VBA lisait / écrivait dans des cellules Excel ; ici les résultats
sont retournés dans un dict.
"""

from . import database as db
from . import thermo
from .combustion import stoich_air_vol, flue_gas_volume_wet
from .thermo import lhv_vol, wobbe_index

T_REF = 273.0


def fuel_air_summary(
    fuel_name: str,
    air_name: str = "Air_sec",
    excess_air_pct: float = 0.0,
) -> dict:
    """
    Calcule les propriétés clés d'un couple combustible/comburant.
    Remplace UpdateFuelAir VBA.

    Retourne un dict contenant :
      fuel_name       str          nom du combustible
      air_name        str          nom du comburant
      fuel_compo      dict         composition vol. [%] du combustible
      air_compo       dict         composition vol. [%] du comburant
      density_kgNm3   float        masse volumique combustible [kg/Nm³] à 0°C
      lhv_J_Nm3       float        PCI [J/Nm³]
      lhv_kJ_Nm3      float        PCI [kJ/Nm³]
      wobbe_J_Nm3     float        indice de Wobbe [J/Nm³]  (base PCI)
      wobbe_MJ_Nm3    float        indice de Wobbe [MJ/Nm³]
      va_Nm3_Nm3      float        volume d'air stœchiométrique [Nm³/Nm³]
      vf_Nm3_Nm3      float        volume fumées humides [Nm³/Nm³]
    """
    fuel_c = db.fuel_composition(fuel_name)
    air_c  = db.comburant_composition(air_name)
    afr    = 1.0 + excess_air_pct / 100.0

    density = thermo.mixture_density(fuel_c, T_REF)    # kg/Nm³
    lhv     = lhv_vol(fuel_name)                        # J/Nm³
    wobbe   = wobbe_index(fuel_name, air_name)          # J/Nm³
    va      = stoich_air_vol(fuel_c, air_c)             # Nm³_air/Nm³_fuel
    vf      = flue_gas_volume_wet(fuel_c, air_c, afr)  # Nm³_fumées/Nm³_fuel

    return {
        "fuel_name":     fuel_name,
        "air_name":      air_name,
        "fuel_compo":    fuel_c,
        "air_compo":     air_c,
        "density_kgNm3": density,
        "lhv_J_Nm3":     lhv,
        "lhv_kJ_Nm3":    lhv / 1e3,
        "wobbe_J_Nm3":   wobbe,
        "wobbe_MJ_Nm3":  wobbe / 1e6,
        "va_Nm3_Nm3":    va,
        "vf_Nm3_Nm3":    vf,
    }
