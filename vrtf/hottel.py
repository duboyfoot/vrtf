"""
Calcul de l'émissivité et du coefficient d'absorption des gaz (CO2, H2O)
par les tables de Hottel. Remplace ModuleHottel.bas.

Référence : W.H. McADAMS, "Transmission de la Chaleur", Dunod 1961, tab. 4.3 p.100.
            L. Ferrand, 8/02/2006 (version VBA originale).
"""

import csv
import math
from functools import lru_cache
from pathlib import Path

from .math_utils import interp

_DATA_DIR = Path(__file__).parent.parent / "data"

# Facteurs de correction du rayon équivalent de la masse gazeuse (McAdams 1961)
_PL_AXIS = [0.01, 0.013, 0.02, 0.05, 0.1,   0.2,   0.3,   0.5,   0.7,   1.0]
_ERR_H2O = [0.97, 0.964, 0.958, 0.942, 0.93, 0.906, 0.892, 0.874, 0.862, 0.85]
_ERR_CO2 = [0.85, 0.844, 0.835, 0.815, 0.80, 0.791, 0.786, 0.779, 0.775, 0.77]

# Températures de référence des tables de correction de superposition CO2/H2O [°C]
_T_OVERLAP = [126.0, 535.0, 940.0]


def _bracket(y: float, data: list) -> tuple:
    """Indices (j1, j2) tels que data[j1] <= y <= data[j2], saturé aux bords."""
    if y <= data[0]:
        return 0, 1
    if y >= data[-1]:
        return len(data) - 2, len(data) - 1
    for i in range(len(data) - 1):
        if data[i] <= y <= data[i + 1]:
            return i, i + 1
    return len(data) - 2, len(data) - 1


def _interp_2d(table: list, axis1: list, axis2: list, x1: float, x2: float) -> float:
    """
    Interpolation bilinéaire sur table[axis1_idx][axis2_idx].
    x1 est interpolé sur axis1 ; x2 est interpolé sur axis2.
    """
    j1, j2 = _bracket(x1, axis1)
    dx = axis1[j2] - axis1[j1]
    if dx == 0.0:
        curve = table[j1]
    else:
        t = (x1 - axis1[j1]) / dx
        curve = [table[j1][k] + t * (table[j2][k] - table[j1][k])
                 for k in range(len(axis2))]
    return interp(x2, axis2, curve)


@lru_cache(maxsize=1)
def _load_tables() -> dict:
    """Charge et met en cache les tables Hottel depuis data/Hottel.csv."""
    path = _DATA_DIR / "Hottel.csv"
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))

    def find_section(name: str) -> int:
        for i, row in enumerate(rows):
            if row and row[0].strip() == name:
                return i
        raise KeyError(f"Section '{name}' introuvable dans Hottel.csv")

    def read_1d(name: str) -> list:
        return [float(v) for v in rows[find_section(name) + 1] if v.strip()]

    def read_2d(name: str, n1: int, n2: int) -> list:
        """
        Lit n2 lignes × n1 colonnes depuis le CSV et retourne table[n1][n2].
        Dans le CSV : lignes = axis2, colonnes = axis1.
        """
        start = find_section(name) + 1
        block = [
            [float(v) for v in rows[start + j] if v.strip()]
            for j in range(n2)
        ]
        return [[block[j][i] for j in range(n2)] for i in range(n1)]

    return {
        "temperature":  read_1d("Températures[25]"),
        "pcl18":        read_1d("pCO2H2O_L[18]"),
        "emiss_co2":    read_2d("emiss_CO2[18][25]",  18, 25),
        "emiss_h2o":    read_2d("emiss_H20[18][25]",  18, 25),   # typo CSV : H20 ≠ H2O
        "t_correc_h2o": read_1d("T_correcH2O[5]"),
        "ph2o_l":       read_1d("pH2O_L[8]"),
        "correc_h2o":   read_2d("correc_H2O[8][5]",    8,  5),
        "pcl_correc":   read_1d("pCO2H2O_L_correc[10]"),
        "ratio_correc": read_1d("pH20_pH2OCO2_correc[11]"),
        "correc126":    read_2d("correc126[10][11]",   10, 11),
        "correc535":    read_2d("correc535[10][11]",   10, 11),
        "correc940":    read_2d("correc940[10][11]",   10, 11),
    }


def gas_emissivity(pCO2: float, pH2O: float, T_celsius: float, beam_length: float) -> float:
    """
    Émissivité totale des fumées (CO2 + H2O) par les tables de Hottel.
    Remplace EmissiviteHottel VBA.

    pCO2        : pression partielle CO2 [atm]
    pH2O        : pression partielle H2O [atm]
    T_celsius   : température de la zone [°C]
    beam_length : longueur moyenne de rayon (mean beam length) [m]

    Retourne : émissivité [-]
    """
    tbl = _load_tables()
    T = T_celsius
    L = beam_length

    # 1. Correction des rayons équivalents de la masse gazeuse
    R_CO2 = L * interp(pCO2 * L, _PL_AXIS, _ERR_CO2)
    R_H2O = L * interp(pH2O * L, _PL_AXIS, _ERR_H2O)

    # 2. Émissivité CO2 (interpolation bilinéaire pL × T)
    eps_co2 = _interp_2d(
        tbl["emiss_co2"], tbl["pcl18"], tbl["temperature"],
        pCO2 * R_CO2, T,
    )

    # 3. Émissivité H2O sans correction de pression partielle
    eps_h2o_prime = _interp_2d(
        tbl["emiss_h2o"], tbl["pcl18"], tbl["temperature"],
        pH2O * R_H2O, T,
    )

    # 4. Correction H2O par auto-élargissement (pression partielle)
    correc_factor = _interp_2d(
        tbl["correc_h2o"], tbl["ph2o_l"], tbl["t_correc_h2o"],
        pH2O * R_H2O, pH2O,
    )
    eps_h2o = eps_h2o_prime * correc_factor

    # 5. Correction de recouvrement de bandes CO2/H2O (Delta_eps)
    if T <= _T_OVERLAP[0]:
        correc = tbl["correc126"]
    elif T >= _T_OVERLAP[-1]:
        correc = tbl["correc940"]
    else:
        n1 = len(tbl["pcl_correc"])
        n2 = len(tbl["ratio_correc"])
        correc = [
            [interp(T, _T_OVERLAP,
                    [tbl["correc126"][i][j],
                     tbl["correc535"][i][j],
                     tbl["correc940"][i][j]])
             for j in range(n2)]
            for i in range(n1)
        ]

    pL_total = pH2O * R_H2O + pCO2 * R_CO2
    ratio = pH2O / (pH2O + pCO2) if (pH2O + pCO2) > 0.0 else 0.0
    delta_eps = _interp_2d(
        correc, tbl["pcl_correc"], tbl["ratio_correc"],
        pL_total, ratio,
    )

    return eps_co2 + eps_h2o - delta_eps


def absorption_coefficient(pCO2: float, pH2O: float, T_celsius: float, beam_length: float) -> float:
    """
    Coefficient d'absorption des fumées [1/(atm·m)] par loi de Beer-Lambert.
    Remplace CoeffAbsorptionHottel VBA.

    pCO2, pH2O  : pressions partielles [atm]
    T_celsius   : température [°C]
    beam_length : longueur de rayon [m]

    Retourne : coefficient d'absorption [1/(atm·m)]
    """
    p_atm = 1.013
    eps = gas_emissivity(pCO2, pH2O, T_celsius, beam_length)
    eps = min(eps, 0.99)
    return -1.0 / (p_atm * beam_length) * math.log(1.0 - eps)
