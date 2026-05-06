"""
Visualisation du four : génération maillage Modray + lancement viewer.
Remplace modulevisualisation.bas.

Note : ModrayFour() n'est pas dans les sources VBA disponibles
(fonction définie dans une bibliothèque externe). La fonction
generate_furnace_section_files() est l'équivalent Python basé sur
le module modray.
"""

import subprocess
from pathlib import Path

from .modray import (
    write_modray_section,
    run_modray_section,
    read_radex_lines,
    _TOP_ZONES_DEFAULT,
)


def run_viewer(mesh_file: str | Path, medit_exe: str | Path) -> None:
    """
    Lance le viewer medit sur un fichier .mesh.
    Remplace l'appel ShellWait(medit + fourmesh_m1.mesh) de ModrayVisu VBA.
    """
    subprocess.Popen(
        [str(Path(medit_exe)), str(Path(mesh_file))],
        cwd=str(Path(medit_exe).parent),
    )


def generate_furnace_section_files(
    output_dir: str | Path,
    cfg: dict,
    top_zones: frozenset | None = None,
) -> list[Path]:
    """
    Génère les fichiers d'entrée Modray1 pour toutes les sections du four.
    Retourne la liste des fichiers créés.

    Équivalent Python de ModrayFour/ModrayVrtf VBA pour la partie génération
    de fichiers (sans exécution des solveurs).
    """
    if top_zones is None:
        top_zones = _TOP_ZONES_DEFAULT

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    files: list[Path] = []
    n_zones       = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]

    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1] + 1):
            path = out / f"VRTF_S{i}_{j}"
            write_modray_section(path, cfg, i, j, top_zones)
            files.append(path)

    return files


def visualize(
    output_dir: str | Path,
    cfg: dict,
    modray1_exe: str | Path,
    modray2_exe: str | Path,
    medit_exe: str | Path | None = None,
    top_zones: frozenset | None = None,
    timeout: int = 300,
) -> list[str]:
    """
    Pipeline complet de visualisation :
    1. Génère les fichiers VRTF_S pour toutes les sections
    2. Lance Modray1/Modray2 sur chaque section
    3. Collecte les lignes d'échanges radiatifs (pour thermette)
    4. Lance medit sur la dernière section (si medit_exe fourni)

    Remplace ModrayVisu VBA (avec ModrayFour reconstitué).

    Retourne la liste de lignes radex pour write_reseau(radex_lines=...).
    """
    if top_zones is None:
        top_zones = _TOP_ZONES_DEFAULT

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    all_radex: list[str] = []
    n_zones       = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]
    last_mesh: Path | None = None

    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1] + 1):
            sec_file = out / f"VRTF_S{i}_{j}"
            write_modray_section(sec_file, cfg, i, j, top_zones)
            run_modray_section(sec_file, modray1_exe, modray2_exe, timeout)

            radex = out / "Res_VRTF_R"
            if radex.exists():
                all_radex.extend(read_radex_lines(radex))

            mesh = Path(str(sec_file) + "_m1.mesh")
            if mesh.exists():
                last_mesh = mesh

    if medit_exe and last_mesh:
        run_viewer(last_mesh, medit_exe)

    n = len(all_radex)
    return [str(n)] + all_radex
