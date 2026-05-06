# VRTF — Vertical Radiant Tube Furnace

Calcul thermique d'un four à tubes radiants verticaux (VRTF) pour le traitement thermique de bandes d'acier.

Traduction complète des macros VBA du classeur `BLD VRTF 1.1.xlsm` en Python.

## Workflow

```
py calculer.py
```

1. Lit les paramètres depuis `BLD VRTF 1.1_modifiable.xlsx` (feuilles *Furnace design*, *Combustion*, *Mesh*)
2. Calcule la combustion et génère les fichiers réseau thermique
3. Lance le solveur **Thermette** (réseau thermique nodal)
4. Écrit les résultats dans `BLD VRTF 1.1_modifiable_résultats.xlsx`

## Structure

```
calculer.py          Script principal (lecture Excel → calcul → résultats Excel)
lire_excel.py        Extraction des données depuis le classeur
run_vrtf.py          Pipeline alternatif via fichier projet .plf (JSON)
vrtf/                Package Python — moteur de calcul
  combustion.py      Calcul stœchiométrique et enthalpique
  posttreatment.py   Post-traitement des résultats Thermette
  fichiers.py        Génération des fichiers réseau Thermette
  thermo.py          Propriétés thermophysiques des gaz
  database.py        Base de données combustibles et aciers
  hottel.py          Facteurs d'échange radiatif (méthode Hottel)
  solver.py          Solveur itératif
  modray.py          Interface avec Modray (échanges radiatifs)
  ...
data/                Base de données (CSV) : aciers, combustibles, gaz
exemple.plf          Fichier projet exemple (2 zones × 2 rangées)
vrtf_reel.plf        Fichier projet réel extrait de l'Excel
```

## Résultats écrits dans Excel

| Feuille | Cellules | Contenu |
|---------|----------|---------|
| Results | F14:F29 | Flux tubes par rangée [kW] |
| Results | G14:G29 | Débit gaz par rangée [Nm³/h] |
| Results | L25:L33 | Synthèse : production, températures, puissances, consommation spécifique |
| Results | K39:Q54 | Tableau détaillé par section |
| ResultsHEAT | E26–E30 | Efficacité et consommation spécifique |
| ResultsHEAT | E36–J38 | Bilan thermique global |

## Dépendances

```
pip install openpyxl
```

Python 3.10+ requis. Le solveur **Thermette.exe** (ARMINES/CES) doit être installé séparément dans `C:\thermette\`.

## Options

```
py calculer.py --excel   "chemin/vers/classeur.xlsx"
               --thermette "C:\thermette\thermette.exe"
               --out    "chemin/résultats.xlsx"
```

```
py run_vrtf.py projet.plf --thermette "C:\thermette\thermette.exe"
```
