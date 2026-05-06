"""
Génération des fichiers d'entrée pour le solveur Thermette.exe.
Remplace Module_thermette.bas.
"""

import math
from pathlib import Path
from . import database as db

PI = math.pi

# Zones avec mur Haut (valeur par défaut calquée sur le VBA)
_TOP_ZONES_DEFAULT = frozenset({2, 3, 6, 7, 10})


# ---------------------------------------------------------------------------
# Fichiers de propriétés
# ---------------------------------------------------------------------------

def write_steel_tables(output_dir: str | Path, grade: str) -> None:
    """
    Écrit AcierCp, AcierCo, AcierRo dans output_dir.
    Remplace ThTablesPropAcier VBA.
    """
    out = Path(output_dir)
    T_cp, cp = db.bisra_cp_table(grade)
    T_co, co = db.bisra_co_table(grade)
    rho = db.bisra_density(grade)

    with open(out / "AcierCp", "w") as f:
        f.write(f"{len(T_cp)}\n")
        for T, v in zip(T_cp, cp):
            f.write(f" {T}  {v}\n")

    with open(out / "AcierCo", "w") as f:
        f.write(f"{len(T_co)}\n")
        for T, v in zip(T_co, co):
            f.write(f" {T}  {v}\n")

    # Densité scalaire → table 2 points constante sur la plage Cp
    T_min, T_max = T_cp[0], T_cp[-1]
    with open(out / "AcierRo", "w") as f:
        f.write("2\n")
        f.write(f" {T_min}  {rho}\n")
        f.write(f" {T_max}  {rho}\n")


def write_fluides_prp(output_dir: str | Path) -> None:
    """
    Écrit fluides.prp dans output_dir.
    Remplace la section fluides.prp de Thnetwork VBA.
    """
    out = Path(output_dir)
    with open(out / "fluides.prp", "w") as f:
        f.write("#BISRA 0 3000\n")
        f.write("table(T,AcierCo)\n")
        f.write("table(T,AcierRo)\n")
        f.write("table(T,AcierCp)\n")
        f.write("\n")


# ---------------------------------------------------------------------------
# Sections internes du fichier réseau
# ---------------------------------------------------------------------------

def _write_solicitations(f, cfg: dict) -> None:
    """Remplace ThRes_Sollicitations VBA."""
    n_zones = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]
    prod = cfg["strip_flowrate_kgs"]
    T_in = cfg["T_strip_in_K"]
    tube_temps = {(t["zone"], t["row"]): t["T_tube_K"] for t in cfg.get("tubes", [])}

    f.write("**sollicit Text 293\n")
    f.write("**sollicit Tref 273\n")
    f.write(f"**sollicit prod {prod}\n")
    f.write(f"**sollicit inv_prod {-prod}\n")
    f.write(f"**sollicit T_entreebande {T_in}\n")

    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1] + 1):
            T_tube = tube_temps.get((i, j), 1273.0)
            f.write(f"**sollicit T_tube_Z{i}_R{j} {T_tube}\n")


def _write_walls(f, cfg: dict, top_zones: frozenset) -> None:
    """Remplace ThRes_Parois VBA."""
    n_zones = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]
    H = cfg["zone_height_m"]
    L = cfg["zone_width_m"]
    ep = cfg["wall_thickness_m"]
    cond = cfg["wall_conductivity"]
    gap = cfg["gap_band_m"]
    ext_conv = {1, 2, n_zones - 1, n_zones}

    for i in range(1, n_zones + 1):
        ihaut = i in top_zones
        nr = rows_per_zone[i - 1]
        for j in range(1, nr + 1):

            if j == 1:
                f.write("\n")
                f.write(f"**branche Bmur_S{i}_R{j}Gauche mur_S{i}_R{j}Gauche_0 mur_S{i}_R{j}Gauche_1\n")
                f.write(f"*geom 0 {ep} rect {H} {L}\n")
                f.write(f"*prop brique {cond} 2000 400 273 \n")
                f.write("*maill 1 5 1\n")
                if i in ext_conv:
                    f.write("*sortie conv Text 10+sfrad(T,Text)\n")
                f.write("*fin\n")

            if j == nr:
                f.write("\n")
                f.write(f"**branche Bmur_S{i}_R{j}Droit mur_S{i}_R{j}Droit_0 mur_S{i}_R{j}Droit_1\n")
                f.write(f"*geom 0 {ep} rect {H} {L}\n")
                f.write(f"*prop brique {cond} 2000 400 273 \n")
                f.write("*maill 1 5 1\n")
                if i in ext_conv:
                    f.write("*sortie conv Text 10+sfrad(T,Text)\n")
                f.write("*fin\n")

            hb = "Haut" if ihaut else "Bas"
            f.write("\n")
            f.write(f"**branche Bmur_S{i}_R{j}{hb} mur_S{i}_R{j}{hb}_0 mur_S{i}_R{j}{hb}_1\n")
            f.write(f"*geom 0 {ep} rect {L} {gap[i - 1]}\n")
            f.write(f"*prop brique {cond} 2000 400 273 \n")
            f.write("*maill 1 5 1\n")
            f.write("*sortie conv Text 10+sfrad(T,Text)\n")
            f.write("*fin\n")

            f.write("\n")
            f.write(f"**branche Bmur_S{i}_R{j}Lat mur_S{i}_R{j}Lat_0 mur_S{i}_R{j}Lat_1\n")
            f.write(f"*geom 0 {ep} rect {H} {2 * gap[i - 1]}\n")
            f.write(f"*prop brique {cond} 2000 400 273 \n")
            f.write("*maill 1 5 1\n")
            f.write("*sortie conv Text 10+sfrad(T,Text)\n")
            f.write("*fin\n")


def _write_strips(f, cfg: dict) -> None:
    """Remplace ThRes_Bandes VBA."""
    n_zones = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]
    larg = cfg["strip_width_m"]
    ep = cfg["strip_thickness_m"]
    H = cfg["zone_height_m"]

    k = 0
    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1]):   # NbRangee(i)-1 branches
            f.write(" \n")
            f.write(f"**branche B_S{i}_R{j} NS{k} NS{k + 1}\n")
            f.write(f"*geom 0 {H} rect {larg} {ep}\n")
            f.write("*prop BISRA table(T,AcierCo) table(T,AcierRo) table(T,AcierCp) 293\n")
            f.write("*maill 1 6 1\n")
            f.write("*debmas prod\n")
            if i == 1 and j == 1:
                f.write("*entree temp T_entreebande\n")
            f.write("*fin\n")
            k += 1


def _tube_geom(t: dict) -> tuple[float, float]:
    """Volume [m³] et surface [m²] d'un tube selon la formule VBA (U ou W)."""
    a = t["a"]
    b = t["b"]      # rayon [m]
    c = t["c"]
    d = t.get("d", 0.0)
    n = t.get("n_tubes", 1)
    ttype = t.get("type", "U").upper()

    if ttype == "W":
        vol  = ((a * PI * b**2) * 2 + (PI * c * PI * b**2) * 3 + (d * PI * b**2) * 2) * n
        surf = ((PI * b * a * 4) + (2 * PI * b * PI * c * 3) + (PI * b * d * 4)) * n
    else:  # U — NB : premier terme NON multiplié par n (fidèle au VBA)
        vol  = (a * PI * b**2) * 2 + (PI * c * PI * b**2) * n
        surf = ((PI * b * a * 4) + (2 * PI * b * PI * c)) * n

    return vol, surf


def _write_tubes(f, cfg: dict) -> None:
    """Remplace ThRes_Tubes VBA."""
    n_zones = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]
    tube_map = {(t["zone"], t["row"]): t for t in cfg.get("tubes", [])}

    for i in range(1, n_zones + 1):
        for j in range(1, rows_per_zone[i - 1] + 1):
            f.write("\n")
            f.write(f"**volume Tube_S{i}_R{j}\n")

            t = tube_map.get((i, j))
            if t is None:
                f.write("*geom 0 0\n")
            else:
                vol, surf = _tube_geom(t)
                f.write(f"*geom {vol} {surf}\n")

            f.write("*prop Materiau_tube 51.6 10 100 273\n")
            f.write(f"*soll temp T_tube_Z{i}_R{j}\n")
            f.write("*fin\n")

    f.write("\n")


# ---------------------------------------------------------------------------
# Fichiers publics
# ---------------------------------------------------------------------------

def write_reseau(
    output_dir: str | Path,
    cfg: dict,
    radex_lines: list | None = None,
    top_zones: frozenset | None = None,
) -> None:
    """
    Écrit VRTF_reseau dans output_dir.
    Remplace les appels ThRes_* de Thnetwork VBA.
    """
    if top_zones is None:
        top_zones = _TOP_ZONES_DEFAULT
    if radex_lines is None:
        radex_lines = []

    out = Path(output_dir)
    with open(out / "VRTF_reseau", "w") as f:
        _write_solicitations(f, cfg)
        _write_walls(f, cfg, top_zones)
        _write_strips(f, cfg)
        _write_tubes(f, cfg)
        f.write("\n")
        for line in radex_lines:
            f.write(line + "\n")


def write_calc(
    output_dir: str | Path,
    cfg: dict,
    top_zones: frozenset | None = None,
) -> None:
    """
    Écrit VRTF_calc dans output_dir.
    Remplace ThCalc VBA.
    """
    if top_zones is None:
        top_zones = _TOP_ZONES_DEFAULT

    n_zones = cfg["n_zones"]
    rows_per_zone = cfg["rows_per_zone"]

    Ns = 0
    for i in range(1, n_zones + 1):
        Ns += 2 * (rows_per_zone[i - 1] - 1)   # SStrip + SFluxStrip par branche bande
    Ns += 2                                      # S_Prod_IN + S_Prod_OUT
    for i in range(1, n_zones + 1):
        Ns += 2 * rows_per_zone[i - 1]          # SFluxMur + SFluxTube par zone/rangée

    out = Path(output_dir)
    with open(out / "VRTF_calc", "w") as f:
        f.write("*calcul statique\n")
        f.write(f"*sorties {Ns}\n")

        f.write("S_Prod_IN EntM(BISRA,273,Tn(NS0))*prod 1 0\n")
        k = 1
        for i in range(1, n_zones + 1):
            for j in range(1, rows_per_zone[i - 1]):
                f.write(f"SStrip_S{i}_R{j} Tn(NS{k}) 1 0 \n")
                f.write(f"SFluxStrip_S{i}_R{j} EntM(BISRA,Tn(NS{k - 1}),Tn(NS{k}))*prod 1 0 \n")
                k += 1
        f.write(f"S_Prod_OUT EntM(BISRA,273,Tn(NS{k - 1}))*prod 1 0\n")

        for i in range(1, n_zones + 1):
            ihaut = i in top_zones
            nr = rows_per_zone[i - 1]
            for j in range(1, nr + 1):
                expr = f"SFluxMur_S{i}_R{j} "
                if j == 1:
                    expr += f"Fcond(Bmur_S{i}_R{j}Gauche,0)+"
                if j == nr:
                    expr += f"Fcond(Bmur_S{i}_R{j}Droit,0)+"
                hb = "Haut" if ihaut else "Bas"
                expr += f"Fcond(Bmur_S{i}_R{j}{hb},0)+"
                expr += f"Fcond(Bmur_S{i}_R{j}Lat,0) 1 0"
                f.write(expr + "\n")

        for i in range(1, n_zones + 1):
            for j in range(1, rows_per_zone[i - 1] + 1):
                f.write(f"SFluxTube_S{i}_R{j} Fvnet(Tube_S{i}_R{j}) 1 0\n")


# ---------------------------------------------------------------------------
# Point d'entrée principal
# ---------------------------------------------------------------------------

def write_thermette_files(
    output_dir: str | Path,
    cfg: dict,
    radex_lines: list | None = None,
    top_zones: frozenset | None = None,
) -> None:
    """
    Génère tous les fichiers d'entrée Thermette dans output_dir.
    Remplace Thnetwork VBA (sans le lancement de Thermette.exe).

    cfg doit contenir :
      steel_grade        str          grade BISRA (ex. "BISRA 01")
      strip_thickness_m  float        épaisseur bande [m]
      strip_width_m      float        largeur bande [m]
      strip_flowrate_kgs float        débit massique bande [kg/s]
      T_strip_in_K       float        température entrée bande [K]
      n_zones            int          nombre de zones
      rows_per_zone      list[int]    nb de rangées par zone (longueur n_zones)
      zone_height_m      float        hauteur interne zones [m]
      zone_width_m       float        largeur zones [m]
      wall_thickness_m   float        épaisseur mur [m]
      wall_conductivity  float        conductivité mur [W/m.K]
      gap_band_m         list[float]  écartement bande par zone (longueur n_zones)
      tubes              list[dict]   géométrie des tubes par zone/rangée :
                           zone int, row int, type str ("U"/"W"),
                           a float, b float (rayon [m]), c float, d float,
                           n_tubes int, T_tube_K float

    radex_lines : lignes d'échanges radiatifs (feuille SMesh du classeur)
    top_zones   : zones avec mur Haut (défaut : {2, 3, 6, 7, 10})
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    write_steel_tables(out, cfg["steel_grade"])
    write_fluides_prp(out)
    write_reseau(out, cfg, radex_lines, top_zones)
    write_calc(out, cfg, top_zones)
