"""
Calcul VRTF complet : combustion → radiatif (Modray) → thermique (Thermette) → post-traitement.

Usage
-----
    py run_vrtf.py projet.plf [options]

    --modray1   chemin vers Modray1.exe
    --modray2   chemin vers Modray2.exe
    --thermette chemin vers Thermette.exe
    --workdir   répertoire de travail (défaut : dossier du projet)
    --out       fichier JSON de sortie (défaut : <projet>_resultats.json)
    --dry-run   génère les fichiers sans lancer les solveurs

Si un exécutable est absent ou non fourni, l'étape correspondante est sautée.

Structure attendue du fichier .plf (JSON)
------------------------------------------
{
  "steel_grade":         "BISRA 01",
  "strip_thickness_m":   0.005,
  "strip_width_m":       1.2,
  "strip_flowrate_kgs":  0.5,
  "T_strip_in_K":        573.0,
  "n_zones":             2,
  "rows_per_zone":       [2, 2],
  "zone_height_m":       1.5,
  "zone_width_m":        2.0,
  "wall_thickness_m":    0.2,
  "wall_conductivity":   1.2,
  "gap_band_m":          [0.05, 0.05],
  "tubes": [
    {"zone":1,"row":1,"type":"U","a":1.0,"b":0.05,"c":0.5,"n_tubes":4,"T_tube_K":1200.0,
     "prem_abs_m":0.3,"dist_band_tube_m":0.04,"pitch_m":0.25,"emissivity":0.85,
     "tube_eff_pct":80.0},
    ...
  ],
  "zones_combustion": [
    {"fuel":"Gaz naturel","air":"Air_sec","excess_air_pct":10.0,
     "power_W":100000.0,"T_fuel_K":293.0,"T_air_K":293.0,
     "regen_air_eff_pct":0.0,"regen_air_aspi_pct":0.0},
    ...
  ]
}
"""

import argparse
import io
import json
import sys
from pathlib import Path

# Force UTF-8 sur la sortie console (Windows cp1252 ne supporte pas les accents/symboles)
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import vrtf


# ---------------------------------------------------------------------------
# Affichage
# ---------------------------------------------------------------------------

def _sep(titre=""):
    w = 60
    if titre:
        pad = (w - len(titre) - 2) // 2
        print(f"\n{'-'*pad} {titre} {'-'*(w - pad - len(titre) - 2)}")
    else:
        print("-" * w)


def _fmt_zone(i, res):
    print(f"  Zone {i} : {res['fuel']!r}  excès {res['excess_air_pct']:.0f}%"
          f"  P={res['power_W']/1e3:.1f} kW")
    print(f"    Débit gaz  : {res['fuel_Nm3h']:.3f} Nm³/h")
    print(f"    Débit air  : {res['air_Nm3h']:.2f} Nm³/h")
    print(f"    Débit fum  : {res['flue_kgs']*3600:.3f} kg/h")
    compo = res["flue_compo"]
    co2 = compo.get("CO2", 0)
    h2o = compo.get("H2O", 0)
    o2  = compo.get("O2", 0)
    print(f"    Fumées     : CO2={co2:.1f}%  H2O={h2o:.1f}%  O2={o2:.2f}%")


def _fmt_postprocess(pp):
    _sep("Résultats post-traitement")
    print(f"  Puissance bande entrée  : {pp['E_strip_in_W']/1e3:8.2f} kW")
    print(f"  Puissance bande sortie  : {pp['E_strip_out_W']/1e3:8.2f} kW")
    print(f"  Flux tubes  (four)      : {pp['flux_tube_four_W']/1e3:8.2f} kW")
    print(f"  Flux murs   (four)      : {pp['flux_mur_four_W']/1e3:8.2f} kW")
    print(f"  Flux fumées (four)      : {pp['flux_fum_four_W']/1e3:8.2f} kW")

    zones = sorted(pp["flux_tube_zone_W"])
    if zones:
        print()
        print("  Par zone :")
        for z in zones:
            qt = pp["flux_tube_zone_W"].get(z, 0.0)
            qm = pp["flux_mur_zone_W"].get(z, 0.0)
            qf = pp["flux_fum_zone_W"].get(z, 0.0)
            print(f"    Zone {z} : tubes={qt/1e3:7.2f} kW  murs={qm/1e3:7.2f} kW"
                  f"  fumées={qf/1e3:7.2f} kW")

    rows = sorted(pp["strip_temp_K"])
    if rows:
        print()
        print("  Températures bande par section (K) :")
        for ij in rows:
            print(f"    S{ij[0]}_R{ij[1]} : {pp['strip_temp_K'][ij]:.1f} K")


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def run(plf_path: Path, args) -> dict:
    # 1. Chargement projet
    _sep("Chargement projet")
    cfg = vrtf.open_project(plf_path)
    print(f"  Fichier : {plf_path}")
    print(f"  Acier   : {cfg['steel_grade']}")
    print(f"  Zones   : {cfg['n_zones']}  rangées : {cfg['rows_per_zone']}")

    workdir = Path(args.workdir) if args.workdir else plf_path.parent / "calcul"
    workdir.mkdir(parents=True, exist_ok=True)
    print(f"  Calculs : {workdir}")

    # 2. Résumé combustible / comburant (premier brûleur)
    zones_comb = cfg.get("zones_combustion", [])
    if not zones_comb:
        print("  AVERTISSEMENT : aucune zone de combustion définie dans le projet.")

    # 3. Combustion
    _sep("Combustion")
    comb_results = vrtf.solve_combustion(zones_comb)
    for i, res in enumerate(comb_results, 1):
        _fmt_zone(i, res)

    # 4. Radex : priorité 1 = .plf, 2 = Modray, 3 = aucun
    radex_lines = cfg.get("radex_lines") or None

    _sep("Radiatif Modray")
    if radex_lines:
        print(f"  Radex pre-calcules charges depuis le .plf ({len(radex_lines)} lignes)")
    else:
        m1 = Path(args.modray1) if args.modray1 else None
        m2 = Path(args.modray2) if args.modray2 else None
        if m1 and m2 and m1.exists() and m2.exists():
            if args.dry_run:
                print("  [dry-run] generation des fichiers section uniquement")
                files = vrtf.generate_furnace_section_files(workdir, cfg)
                print(f"  {len(files)} fichiers section ecrits")
            else:
                print(f"  Modray1 : {m1}")
                print(f"  Modray2 : {m2}")
                radex_lines = vrtf.generate_radex_lines(
                    workdir, cfg, m1, m2, timeout=600
                )
                print(f"  {len(radex_lines)} lignes radex collectees")
        else:
            print("  [ignore] pas de radex disponibles (ni .plf, ni Modray)")

    # 5. Génération fichiers Thermette
    _sep("Generation fichiers Thermette")
    vrtf.write_thermette_files(workdir, cfg, radex_lines=radex_lines)
    print(f"  AcierCp, AcierCo, AcierRo, fluides.prp")
    print(f"  VRTF_reseau, VRTF_calc  -> {workdir}")

    # 6. Solveur Thermette
    thermette_exe = Path(args.thermette) if args.thermette else None
    results_file  = workdir / "VRTF_resu"
    pp = None

    if thermette_exe and thermette_exe.exists():
        _sep("Solveur Thermette")
        if args.dry_run:
            print("  [dry-run] lancement Thermette ignore")
        else:
            import shutil
            ther_dir = thermette_exe.parent
            print(f"  Exe : {thermette_exe}")

            # Thermette lit ses fichiers dans son propre repertoire
            for fname in ("VRTF_reseau", "VRTF_calc",
                          "AcierCp", "AcierCo", "AcierRo", "fluides.prp"):
                src = workdir / fname
                if src.exists():
                    shutil.copy2(src, ther_dir / fname)
            print(f"  Fichiers copies vers {ther_dir}")

            # Thermette batch mode : 3 arguments = reseau calc resu
            ret = vrtf.shell_wait(
                [str(thermette_exe), "VRTF_reseau", "VRTF_calc", "VRTF_resu"],
                cwd=ther_dir,
                timeout=120,
            )
            print(f"  Code retour : {ret}")

            # Recuperer le resultat
            resu_src = ther_dir / "VRTF_resu"
            if resu_src.exists():
                shutil.copy2(resu_src, results_file)
                print(f"  Resultat copie -> {results_file}")
            else:
                print("  ERREUR : VRTF_resu absent apres calcul.")
    else:
        _sep("Solveur Thermette")
        print("  [ignore] executable Thermette non fourni ou introuvable")

    # 7. Post-traitement
    if results_file.exists() and not args.dry_run:
        _sep("Post-traitement")
        fuel_name = zones_comb[0]["fuel"] if zones_comb else "Gaz naturel"
        air_name  = zones_comb[0].get("air", "Air_sec") if zones_comb else "Air_sec"
        excess    = zones_comb[0].get("excess_air_pct", 0.0) if zones_comb else 0.0
        try:
            pp = vrtf.postprocess(results_file, cfg, fuel_name, air_name, excess)
            _fmt_postprocess(pp)
        except RuntimeError as e:
            print(f"  ERREUR post-traitement : {e}")
    else:
        if not args.dry_run:
            _sep("Post-traitement")
            print("  [ignoré] fichier VRTF_resu absent")

    # 8. Sauvegarde résultats JSON
    out_path = Path(args.out) if args.out else plf_path.parent / (plf_path.stem + "_resultats.json")
    _sep("Sauvegarde")
    output = {
        "projet": str(plf_path),
        "workdir": str(workdir),
        "combustion": [
            {k: v for k, v in r.items()
             if not k.startswith("poly_")}
            for r in comb_results
        ],
    }
    if pp:
        output["postprocess"] = {
            "E_strip_in_W":    pp["E_strip_in_W"],
            "E_strip_out_W":   pp["E_strip_out_W"],
            "flux_tube_four_W": pp["flux_tube_four_W"],
            "flux_mur_four_W":  pp["flux_mur_four_W"],
            "flux_fum_four_W":  pp["flux_fum_four_W"],
            "flux_tube_zone_W": {str(k): v for k, v in pp["flux_tube_zone_W"].items()},
            "flux_mur_zone_W":  {str(k): v for k, v in pp["flux_mur_zone_W"].items()},
            "flux_fum_zone_W":  {str(k): v for k, v in pp["flux_fum_zone_W"].items()},
            "strip_temp_K":    {f"S{i}_R{j}": v for (i,j), v in pp["strip_temp_K"].items()},
            "T_fum_K":         {f"S{i}_R{j}": v for (i,j), v in pp["T_fum_K"].items()},
            "deb_gaz_Nm3h":    {f"S{i}_R{j}": v for (i,j), v in pp["deb_gaz_Nm3h"].items()},
        }

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(output, fh, ensure_ascii=False, indent=2, default=float)
    print(f"  Résultats → {out_path}")

    return output


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Calcul VRTF complet (combustion → Modray → Thermette → post-traitement)"
    )
    parser.add_argument("projet", help="Fichier projet .plf (JSON)")
    parser.add_argument("--modray1",   default=None, help="Chemin Modray1.exe")
    parser.add_argument("--modray2",   default=None, help="Chemin Modray2.exe")
    parser.add_argument("--thermette", default=None, help="Chemin Thermette.exe")
    parser.add_argument("--workdir",   default=None, help="Répertoire de travail")
    parser.add_argument("--out",       default=None, help="Fichier JSON de sortie")
    parser.add_argument("--dry-run",   action="store_true",
                        help="Génère les fichiers sans lancer les solveurs")
    args = parser.parse_args()

    plf = Path(args.projet)
    if not plf.exists():
        print(f"ERREUR : fichier projet introuvable : {plf}", file=sys.stderr)
        sys.exit(1)

    run(plf, args)
    _sep()
    print("Terminé.")


if __name__ == "__main__":
    main()
