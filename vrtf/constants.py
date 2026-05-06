"""
Constantes physiques, paramètres par défaut et utilitaires système.
Remplace Parametres.bas et ModuleParametres.bas.

Les constantes de feuilles Excel (SGeom, SFiles, ...) et d'adresses de
cellules (CelNbSect, CelHautSect, ...) sont sans objet en Python : les
données sont passées directement via dict (cfg).
"""

import math
import subprocess
from pathlib import Path

# ---------------------------------------------------------------------------
# Constantes physiques
# ---------------------------------------------------------------------------

SIGMA = 5.67e-8     # Constante de Stefan-Boltzmann [W/m².K⁴]
PI    = math.pi     # π (valeur Python exacte, > précision VBA)
R     = 8314.0      # Constante des gaz parfaits [J/(kmol.K)]
T_REF = 273.0       # Température de référence [K] (0 °C)

# ---------------------------------------------------------------------------
# Métadonnées du projet
# ---------------------------------------------------------------------------

VERSION = "1.0"

# ---------------------------------------------------------------------------
# Paramètres par défaut du solveur PID
# (SIM_TIME, GAIN, INTEGRAL, DERIVATIVE de ModuleParametres.bas)
# ---------------------------------------------------------------------------

SOLVER_SIM_TIME   = 240.0   # durée de simulation [s]
SOLVER_GAIN       = 10.0    # gain proportionnel
SOLVER_INTEGRAL   = 20.0    # constante intégrale
SOLVER_DERIVATIVE = 0.0     # constante dérivée

# ---------------------------------------------------------------------------
# Utilitaires système
# ---------------------------------------------------------------------------

def file_exists(path: str | Path) -> bool:
    """
    Vérifie si un fichier existe.
    Remplace FileExists VBA (GetAttr + test vbDirectory).
    """
    p = Path(path)
    return p.exists() and p.is_file()


def shell_wait(
    command: str | list,
    cwd: str | Path | None = None,
    timeout: int = 300,
) -> int:
    """
    Lance un processus et attend sa terminaison.
    Remplace ShellWait VBA (Shell + boucle GetExitCodeProcess / Sleep).

    command : chaîne shell ou liste d'arguments
    cwd     : répertoire de travail (optionnel)
    timeout : délai maximal en secondes (défaut 300 s)

    Retourne le code de retour du processus.
    """
    result = subprocess.run(
        command if isinstance(command, list) else command,
        shell=isinstance(command, str),
        cwd=str(cwd) if cwd else None,
        timeout=timeout,
        stdin=subprocess.DEVNULL,
    )
    return result.returncode
