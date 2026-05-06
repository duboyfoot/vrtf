"""
Package vrtf — Vertical Radiant Tube Furnace
Remplace les lookups Excel/VBA de BLD VRTF 1.1.xlsm en Python pur.
"""

from .database import (
    bisra_grades, bisra_density,
    bisra_cp_table, bisra_co_table, bisra_eps_table, bisra_enthalpy_table,
    fuel_names, fuel_composition,
    comburant_names, comburant_composition,
    gas_names, gas_props,
)
from .thermo import (
    gas_cp, gas_enthalpy,
    mixture_molar_mass, mixture_density, mixture_cp, mixture_enthalpy,
    mixture_enthalpy_vol,
    fuel_enthalpy_vol, fuel_enthalpy_kg,
    air_enthalpy_vol, air_density,
    lhv_vol, lhv_kg, wobbe_index,
    steel_density, steel_cp, steel_conductivity, steel_emissivity, steel_enthalpy,
    temperature_from_enthalpy,
    heat_exchanger,
)
from .combustion import (
    stoich_air_vol, stoich_air_kg,
    flue_gas_volume_wet,
    flue_gas_composition, waste_gas_composition,
    flow_fuel,
    adiabatic_temperature, equivalent_temperature,
)
from .math_utils import interp, poly_cp, poly_enthalpy, bisect, linear_interp, polint3
from .hottel import gas_emissivity, absorption_coefficient
from .solver import solve_zone, solve_combustion
from .thermette import (
    write_steel_tables, write_fluides_prp,
    write_reseau, write_calc, write_thermette_files,
)
from .posttreatment import parse_results, postprocess
from .update import fuel_air_summary
from .constants import (
    SIGMA, PI, R, T_REF, VERSION,
    SOLVER_SIM_TIME, SOLVER_GAIN, SOLVER_INTEGRAL, SOLVER_DERIVATIVE,
    file_exists, shell_wait,
)
from .modray import (
    write_modray_section, run_modray_section,
    read_radex_lines, generate_radex_lines,
)
from .visualisation import run_viewer, generate_furnace_section_files, visualize
from .fichiers import save_project, open_project

__all__ = [
    # database
    "bisra_grades", "bisra_density",
    "bisra_cp_table", "bisra_co_table", "bisra_eps_table", "bisra_enthalpy_table",
    "fuel_names", "fuel_composition",
    "comburant_names", "comburant_composition",
    "gas_names", "gas_props",
    # thermo
    "gas_cp", "gas_enthalpy",
    "mixture_molar_mass", "mixture_density", "mixture_cp",
    "mixture_enthalpy", "mixture_enthalpy_vol",
    "fuel_enthalpy_vol", "fuel_enthalpy_kg",
    "air_enthalpy_vol", "air_density",
    "lhv_vol", "lhv_kg", "wobbe_index",
    "steel_density", "steel_cp", "steel_conductivity", "steel_emissivity", "steel_enthalpy",
    "temperature_from_enthalpy", "heat_exchanger",
    # combustion
    "stoich_air_vol", "stoich_air_kg",
    "flue_gas_volume_wet",
    "flue_gas_composition", "waste_gas_composition",
    "flow_fuel", "adiabatic_temperature", "equivalent_temperature",
    # math
    "interp", "poly_cp", "poly_enthalpy", "bisect", "linear_interp", "polint3",
    # hottel
    "gas_emissivity", "absorption_coefficient",
    # solver
    "solve_zone", "solve_combustion",
    # thermette
    "write_steel_tables", "write_fluides_prp",
    "write_reseau", "write_calc", "write_thermette_files",
    # update
    "fuel_air_summary",
    # constants
    "SIGMA", "PI", "R", "T_REF", "VERSION",
    "SOLVER_SIM_TIME", "SOLVER_GAIN", "SOLVER_INTEGRAL", "SOLVER_DERIVATIVE",
    "file_exists", "shell_wait",
    # posttreatment
    "parse_results", "postprocess",
    # modray
    "write_modray_section", "run_modray_section",
    "read_radex_lines", "generate_radex_lines",
    # visualisation
    "run_viewer", "generate_furnace_section_files", "visualize",
    # fichiers
    "save_project", "open_project",
]
