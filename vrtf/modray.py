"""
Génération des fichiers d'entrée pour le solveur Modray1/Modray2.
Remplace Module_modray.bas.
"""

from pathlib import Path
import subprocess

_TOP_ZONES_DEFAULT = frozenset({2, 3, 6, 7, 10})


# ---------------------------------------------------------------------------
# Comptage des objets (en-tête)
# ---------------------------------------------------------------------------

def _n_tube_parts(tube_type: str, n_tube: int) -> int:
    """Nombre de parties géométriques par type de tube."""
    if tube_type == "W":
        return 7 * n_tube
    if tube_type == "I":
        return 1 * n_tube
    return 3 * n_tube   # U (défaut)


def _n_wall_strip_objects(n_rows: int, j: int) -> int:
    """
    Nombre d'objets parois + bandes pour la rangée j.
    Fidèle au VBA : 7 si rangée de bord, 9 sinon.
    """
    if j == 1 or j == n_rows:
        return 7
    return 9


# ---------------------------------------------------------------------------
# Sections internes : chaque MdRes_* du VBA
# ---------------------------------------------------------------------------

def _write_header(f, n_rows: int, j: int, n_tube: int, tube_type: str) -> None:
    """Remplace MdRes_objet VBA."""
    n_obj = _n_tube_parts(tube_type, n_tube) + _n_wall_strip_objects(n_rows, j)

    f.write("**Nbandes 1\n")
    f.write("0 100\n")
    f.write("\n")
    f.write("**Calculs groupe 1\n")
    f.write("1 diffus Res_VRTF_R\n")
    f.write("**Ordre_rotations XYZ\n")
    f.write("\n")
    f.write(f"**nobjets {n_obj}\n")
    f.write("\n")


def _write_walls(f, n_rows: int, i: int, j: int,
                 H: float, L: float, gap: float, top_zones: frozenset) -> None:
    """Remplace MdRes_parois VBA."""
    ihaut = i in top_zones

    def _paral(emis, faces_only=None, faces_none=None, enceinte=True):
        f.write(f"*geom 0 0 0 0 0 0 paral {gap} {L} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write(f"*proprad {emis} 0\n")
        if enceinte:
            f.write("*enceinte\n")
        if faces_only:
            f.write(f"*seulrayon {faces_only}\n")
        if faces_none:
            for face in faces_none:
                f.write(f"*norayon {face}\n")
        f.write("*noeudthermette 0\n")
        f.write("*fin\n")

    # Mur gauche (première rangée seulement)
    if j == 1:
        f.write("\n")
        f.write(f"**objet mur_S{i}_R{j}Gauche_0\n")
        _paral(0.85, faces_only="gauche")

    # Mur droit (dernière rangée seulement)
    if j == n_rows:
        f.write("\n")
        f.write(f"**objet mur_S{i}_R{j}Droit_0\n")
        _paral(0.85, faces_only="droite")

    # Mur latéral (toujours)
    f.write("\n")
    f.write(f"**objet mur_S{i}_R{j}Lat_0\n")
    f.write(f"*geom 0 0 0 0 0 0 paral {gap} {L} {H}\n")
    f.write("*maill  1 5 1 5 1 5 1\n")
    f.write("*proprad 0.85 0\n")
    f.write("*enceinte\n")
    f.write("*norayon droite\n")
    f.write("*norayon gauche\n")
    f.write("*norayon base\n")
    f.write("*norayon sommet\n")
    f.write("*noeudthermette 0\n")
    f.write("*fin\n")

    # Mur haut (toujours présent, proprad dépend de ihaut)
    f.write("\n")
    f.write(f"**objet mur_S{i}_R{j}Haut_0\n")
    f.write(f"*geom 0 0 0 0 0 0 paral {gap} {L} {H}\n")
    f.write("*maill  1 5 1 5 1 5 1\n")
    f.write(f"*proprad {0.85 if ihaut else 0} 0\n")
    f.write("*enceinte\n")
    f.write("*seulrayon sommet\n")
    f.write("*noeudthermette 0\n")
    f.write("*fin\n")

    # Mur bas (toujours présent, proprad inverse de haut)
    f.write("\n")
    f.write(f"**objet mur_S{i}_R{j}Bas_0\n")
    f.write(f"*geom 0 0 0 0 0 0 paral {gap} {L} {H}\n")
    f.write("*maill  1 5 1 5 1 5 1\n")
    f.write(f"*proprad {0 if ihaut else 0.85} 0\n")
    f.write("*enceinte\n")
    f.write("*seulrayon base\n")
    f.write("*noeudthermette 0\n")
    f.write("*fin\n")


def _write_strips(f, n_rows: int, i: int, j: int,
                  H: float, L: float, gap: float,
                  ep: float, larg: float, emis: float) -> None:
    """Remplace MdRes_bande VBA (fidèle au VBA incluant le bug de nommage Sepa2)."""
    delta = (L - larg) / 2.0

    # Bande droite (j > 1) — strip entre rangée j-1 (côté droit du gap)
    if j > 1:
        f.write("\n")
        f.write(f"**objet B_S{i}_R{j - 1}\n")
        f.write(f"*geom 0 {delta} 0 0 0 0 paral {ep} {larg} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write(f"*proprad {emis} 0\n")
        f.write("*seulrayon droite\n")
        f.write(f"*branchethermette {larg * H}\n")
        f.write("*fin\n")

        f.write("\n")
        f.write(f"**objet B_S{i}_R{j - 1}_Sepa\n")
        f.write(f"*geom 0 0 0 0 0 0 paral {ep} {delta} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write("*proprad 0 0\n")
        f.write("*seulrayon droite\n")
        f.write("*fin\n")

        f.write("\n")
        f.write(f"**objet B_S{i}_R{j - 1}_Sepa2\n")  # bug VBA reproduit : j-1
        f.write(f"*geom 0 {larg + delta} 0 0 0 0 paral {ep} {delta} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write("*proprad 0 0\n")
        f.write("*seulrayon droite\n")
        f.write("*fin\n")

    # Bande gauche (j < n_rows) — strip entre rangée j (côté gauche du gap)
    if j < n_rows:
        f.write("\n")
        f.write(f"**objet B_S{i}_R{j}\n")
        f.write(f"*geom {gap} {delta} 0 0 0 0 paral {ep} {larg} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write(f"*proprad {emis} 0\n")
        f.write("*seulrayon gauche\n")
        f.write(f"*branchethermette {larg * H}\n")
        f.write("*fin\n")

        f.write("\n")
        f.write(f"**objet B_S{i}_R{j}_Sepa\n")
        f.write(f"*geom {gap} 0 0 0 0 0 paral {ep} {delta} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write("*proprad 0 0\n")
        f.write("*seulrayon gauche\n")
        f.write("*fin\n")

        f.write("\n")
        f.write(f"**objet B_S{i}_R{j - 1}_Sepa2\n")  # bug VBA reproduit : j-1
        f.write(f"*geom {gap} {larg + delta} 0 0 0 0 paral {ep} {delta} {H}\n")
        f.write("*maill  1 5 1 5 1 5 1\n")
        f.write("*proprad 0 0\n")
        f.write("*seulrayon gauche\n")
        f.write("*fin\n")


def _tube_obj(f, name: str, geom: str, maill: str,
              no_base: bool = False, no_top: bool = False) -> None:
    """Écrit un objet tube générique (cylindre ou coude)."""
    f.write(f"**objet {name}\n")
    f.write(f"*geom {geom}\n")
    f.write(f"*maill {maill}\n")
    f.write("*proprad 0.85 0\n")
    if no_base:
        f.write("*norayon base\n")
    if no_top:
        f.write("*norayon sommet\n")
    f.write("*fin\n")
    f.write("\n")


def _write_tubes(f, i: int, j: int, L: float,
                 n_tube: int, tube_type: str,
                 prem_abs: float, dist_bt: float, pas: float,
                 a: float, b: float, c: float, d: float) -> None:
    """Remplace MdRes_tube VBA."""
    MAILL_CYL  = "1 8 1 1 1 3 1"
    MAILL_COUD = "1 8 1 1 1 3 1"

    for k in range(1, n_tube + 1):
        pfx = f"_S{i}_R{j}_N{k}"
        even = (k % 2 == 0)

        if tube_type == "U":
            z0m = prem_abs - c + (k - 1) * pas   # z base
            z0p = prem_abs + c + (k - 1) * pas   # z sommet

            if even:
                y_cyl = 0.01;    rot_a = "-90 0 0";   rot_b = "-90 90 0"
                y_coud = a + 0.01
            else:
                y_cyl = L - 0.01; rot_a = "90 0 0";   rot_b = "90 90 0"
                y_coud = L - 0.01 - a

            _tube_obj(f, f"Tubea{pfx}",
                      f"{dist_bt} {y_cyl} {z0m} {rot_a} cylin {b} {a} 360",
                      MAILL_CYL, no_top=True)
            _tube_obj(f, f"Tubeb{pfx}",
                      f"{dist_bt} {y_coud} {z0m} {rot_b} coude {b} {c} 180 360",
                      MAILL_COUD, no_base=True, no_top=True)
            _tube_obj(f, f"Tubec{pfx}",
                      f"{dist_bt} {y_cyl} {z0p} {rot_a} cylin {b} {a} 360",
                      MAILL_CYL, no_top=True)

        elif tube_type == "W":
            z3m = prem_abs - 3 * c + (k - 1) * pas
            z1m = prem_abs - c     + (k - 1) * pas
            z1p = prem_abs + c     + (k - 1) * pas
            z3p = prem_abs + 3 * c + (k - 1) * pas

            if even:
                y0   = 0.01;              rot_str = "-90"
                ya   = a + 0.01;          ya_d = a + 0.01 - d
                rot_coude = f"{rot_str} 90 0"
                rot_coud2 = f"{rot_str} -270 180"
                rot_cyl   = f"{rot_str} 0 0"
            else:
                y0   = L - 0.01;          rot_str = "90"
                ya   = L - 0.01 - a;      ya_d = L - 0.01 - a + d
                rot_coude = f"{rot_str} 90 0"
                rot_coud2 = f"{rot_str} -270 180"
                rot_cyl   = f"{rot_str} 0 0"

            _tube_obj(f, f"Tubea{pfx}",
                      f"{dist_bt} {y0} {z3m} {rot_cyl} cylin {b} {a} 360",
                      MAILL_CYL, no_top=True)
            _tube_obj(f, f"Tubeb{pfx}",
                      f"{dist_bt} {ya} {z3m} {rot_coude} coude {b} {c} 180 360",
                      MAILL_COUD, no_base=True, no_top=True)
            _tube_obj(f, f"Tubec{pfx}",
                      f"{dist_bt} {ya_d} {z1m} {rot_cyl} cylin {b} {d} 360",
                      MAILL_CYL, no_base=True, no_top=True)
            _tube_obj(f, f"Tubed{pfx}",
                      f"{dist_bt} {ya_d} {z1m} {rot_coud2} coude {b} {c} 180 360",
                      MAILL_COUD, no_base=True, no_top=True)
            _tube_obj(f, f"Tubee{pfx}",
                      f"{dist_bt} {ya_d} {z1p} {rot_cyl} cylin {b} {d} 360",
                      MAILL_CYL, no_base=True, no_top=True)
            _tube_obj(f, f"Tubef{pfx}",
                      f"{dist_bt} {ya} {z1p} {rot_coude} coude {b} {c} 180 360",
                      MAILL_COUD, no_base=True, no_top=True)
            _tube_obj(f, f"Tubeg{pfx}",
                      f"{dist_bt} {y0} {z3p} {rot_cyl} cylin {b} {a} 360",
                      MAILL_CYL, no_top=True)

        elif tube_type == "I":
            z0 = prem_abs - c + (k - 1) * pas if k > 1 else prem_abs - c
            _tube_obj(f, f"Tubea{pfx}",
                      f"0.01 {dist_bt} {z0} 0 90 0 cylin {b} {a - 0.01} 360",
                      MAILL_CYL, no_top=True)


def _write_groups(f, i: int, j: int, n_tube: int, tube_type: str) -> None:
    """Remplace MdRes_groupe VBA."""
    if n_tube <= 0:
        return

    grp = f"Tube_S{i}_R{j}"

    if tube_type == "W":
        parts_per = ["Tubea", "Tubeb", "Tubec", "Tubed", "Tubee", "Tubef", "Tubeg"]
        n_parts = 7
    else:   # U (défaut)
        parts_per = ["Tubea", "Tubeb", "Tubec"]
        n_parts = 3

    names = []
    for k in range(1, n_tube + 1):
        for p in parts_per:
            names.append(f"{p}_S{i}_R{j}_N{k}")

    f.write(f"**Groupes 1\n")
    f.write("\n")
    f.write(f"{grp} {n_parts * n_tube} {' '.join(names)}\n")
    f.write("\n")


# ---------------------------------------------------------------------------
# API publique
# ---------------------------------------------------------------------------

def write_modray_section(
    output_file: str | Path,
    cfg: dict,
    zone: int,
    row: int,
    top_zones: frozenset | None = None,
) -> None:
    """
    Écrit le fichier d'entrée Modray1 pour une section (zone, rangée).
    Remplace les appels MdRes_* dans ModrayVrtf VBA.

    cfg doit contenir (même structure que write_thermette_files) :
      n_zones, rows_per_zone,
      zone_height_m, zone_width_m, gap_band_m,
      strip_thickness_m, strip_width_m,
      tubes : [{zone, row, type, a, b, c, d, n_tubes,
                prem_abs_m, dist_band_tube_m, pitch_m, emissivity}]
    """
    if top_zones is None:
        top_zones = _TOP_ZONES_DEFAULT

    n_rows = cfg["rows_per_zone"][zone - 1]
    H      = cfg["zone_height_m"]
    L      = cfg["zone_width_m"]
    gap    = cfg["gap_band_m"][zone - 1]
    ep     = cfg["strip_thickness_m"]
    larg   = cfg["strip_width_m"]

    tube_map = {(t["zone"], t["row"]): t for t in cfg.get("tubes", [])}
    t = tube_map.get((zone, row), {})

    n_tube    = t.get("n_tubes",    0)
    tube_type = t.get("type",       "U").upper()
    prem_abs  = t.get("prem_abs_m", 0.0)
    dist_bt   = t.get("dist_band_tube_m", gap / 2.0)
    pas       = t.get("pitch_m",    0.0)
    a         = t.get("a",          0.0)
    b         = t.get("b",          0.0)
    c         = t.get("c",          0.0)
    d         = t.get("d",          0.0)
    emis      = t.get("emissivity", 0.85)

    with open(output_file, "w") as f:
        _write_header(f, n_rows, row, n_tube, tube_type)
        _write_walls(f, n_rows, zone, row, H, L, gap, top_zones)
        _write_strips(f, n_rows, zone, row, H, L, gap, ep, larg, emis)
        _write_tubes(f, zone, row, L, n_tube, tube_type,
                     prem_abs, dist_bt, pas, a, b, c, d)
        _write_groups(f, zone, row, n_tube, tube_type)


def run_modray_section(
    section_file: str | Path,
    modray1_exe: str | Path,
    modray2_exe: str | Path,
    timeout: int = 300,
) -> None:
    """
    Lance Modray1.exe puis Modray2.exe sur un fichier section.
    Remplace les appels ShellWait de ModrayVrtf VBA.

    section_file : chemin du fichier VRTF_S (sans extension)
    Les executables sont cherchés dans modray_dir.
    """
    sec = Path(section_file)
    m1 = Path(modray1_exe)
    m2 = Path(modray2_exe)

    subprocess.run(
        [str(m1), str(sec), str(sec) + "_m1"],
        cwd=str(m1.parent), check=True, timeout=timeout,
    )
    subprocess.run(
        [str(m2), str(sec) + "_m1"],
        cwd=str(m2.parent), check=True, timeout=timeout,
    )


def read_radex_lines(radex_file: str | Path) -> list[str]:
    """
    Lit un fichier Res_VRTF_R (sortie Modray2) et retourne ses lignes.
    Remplace la lecture de fichier dans ModrayVrtf VBA.
    """
    return Path(radex_file).read_text(encoding="latin-1").splitlines()


def generate_radex_lines(
    output_dir: str | Path,
    cfg: dict,
    modray1_exe: str | Path,
    modray2_exe: str | Path,
    top_zones: frozenset | None = None,
    timeout: int = 300,
) -> list[str]:
    """
    Orchestre la génération des échanges radiatifs pour toutes les sections :
    1. Écrit VRTF_S pour chaque (zone, rangée)
    2. Lance Modray1/Modray2
    3. Collecte les Res_VRTF_R
    Retourne la liste complète de lignes à passer à write_reseau (radex_lines).

    Remplace la boucle principale de ModrayVrtf VBA.
    """
    if top_zones is None:
        top_zones = _TOP_ZONES_DEFAULT

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    all_lines: list[str] = []
    n_zones       = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]

    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1] + 1):
            sec_file = out / "VRTF_S"
            write_modray_section(sec_file, cfg, i, j, top_zones)
            run_modray_section(sec_file, modray1_exe, modray2_exe, timeout)
            radex = out / "Res_VRTF_R"
            if radex.exists():
                all_lines.extend(read_radex_lines(radex))

    # Thermette attend en première ligne le nombre total de lignes d'échanges
    n = len(all_lines)
    return [str(n)] + all_lines
