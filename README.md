# VRTF — Vertical Radiant Tube Furnace

Calcul thermique d'un four à tubes radiants verticaux (VRTF) pour le traitement thermique de bandes d'acier.

Traduction complète des macros VBA du classeur `BLD VRTF 1.1.xlsm` en Python.

## Workflow

```
py calculer.py
```

1. Lit les paramètres depuis `BLD VRTF 1.1_modifiable.xlsx` (feuilles *Furnace design*, *Combustion*, *Mesh*)
2. *(optionnel)* Recalcule les facteurs de forme via **Modray** et met à jour la feuille *Mesh*
3. Calcule la combustion et génère les fichiers réseau thermique
4. Lance le solveur **Thermette** (réseau thermique nodal)
5. Écrit les résultats dans `BLD VRTF 1.1_modifiable_résultats.xlsx`

## Structure

```
calculer.py          Script principal (lecture Excel → calcul → résultats Excel)
lire_excel.py        Extraction des données depuis le classeur
run_vrtf.py          Pipeline alternatif via fichier projet .plf (JSON)
pyproject.toml       Packaging pip (kappa-combustion)

combustion/          Package autonome — calculs de combustion
  core.py            Stœchiométrie, fumées, T adiabatique, débits
  thermo.py          Propriétés thermophysiques des mélanges gazeux
  database.py        Base de données combustibles, comburants, gaz
  math_utils.py      Bisection, interpolation, polynômes
  data/              CSV : basdo_gaz, Fuels, Comburants

vrtf/                Package Python — moteur de calcul VRTF
  combustion.py      Délègue à combustion/
  thermo.py          Gaz → combustion/ ; acier BISRA local
  database.py        Gaz → combustion/ ; BISRA local
  posttreatment.py   Post-traitement des résultats Thermette
  thermette.py       Génération des fichiers réseau Thermette
  hottel.py          Facteurs d'échange radiatif (méthode Hottel)
  solver.py          Solveur itératif
  modray.py          Interface avec Modray (échanges radiatifs)
  ...

data/                CSV aciers BISRA (Cp, conductivité, émissivité, enthalpie)
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

## Package combustion

Les calculs de combustion sont disponibles comme bibliothèque autonome,
utilisable dans tout projet Python indépendamment de VRTF.

### Installation

```
pip install -e "C:\...\developpement"
```

### Exemple d'utilisation

```python
import combustion

# --- Combustibles et comburants disponibles ---
print(combustion.fuel_names())       # ['Gaz_naturel', 'BFG_Sidmar', ...]
print(combustion.comburant_names())  # ['Air_sec', 'Air_humide', ...]

# --- Stœchiométrie ---
fuel = "Gaz_naturel"
va  = combustion.stoich_air_vol(fuel)              # Nm³_air / Nm³_fuel
print(f"Va = {va:.2f} Nm³_air/Nm³_fuel")          # Va = 9.96 Nm³_air/Nm³_fuel

# --- PCI ---
pci = combustion.lhv_vol(fuel)
print(f"PCI = {pci/1e6:.2f} MJ/Nm³")              # PCI = 37.18 MJ/Nm³

# --- Débits à partir de la puissance ---
flows = combustion.flow_fuel(power_W=10e6, fuel=fuel, air_fuel_ratio=1.1)
print(f"Gaz  = {flows['fuel_Nm3h']:.1f} Nm³/h")
print(f"Air  = {flows['air_Nm3h']:.0f} Nm³/h")
print(f"Fumées = {flows['flue_kgs']:.2f} kg/s")

# --- Composition des fumées ---
compo = combustion.waste_gas_composition(fuel, air_fuel_ratio=1.1)
print(f"CO2={compo['CO2']:.1f}%  H2O={compo['H2O']:.1f}%  O2={compo['O2']:.1f}%")

# --- Température adiabatique ---
Tad = combustion.adiabatic_temperature(fuel, air_fuel_ratio=1.1)
print(f"T adiabatique = {Tad - 273:.0f} °C")      # T adiabatique = 1952 °C

# --- Enthalpie des fumées ---
flue_c = combustion.flue_gas_composition(fuel, air_fuel_ratio=1.1)
h = combustion.mixture_enthalpy_vol(flue_c, T_K=1200 + 273)
print(f"H fumées à 1200°C = {h/1e6:.3f} MJ/Nm³")

# --- Indice de Wobbe ---
w = combustion.wobbe_index(fuel)
print(f"Wobbe = {w/1e6:.2f} MJ/Nm³")
```

### Modules

| Module | Contenu |
|--------|---------|
| `combustion.core` | Stœchiométrie, volume fumées, composition fumées, débits, T adiabatique |
| `combustion.thermo` | Cp, enthalpie, densité des mélanges gazeux, PCI, échangeur |
| `combustion.database` | Combustibles, comburants, propriétés des gaz élémentaires |
| `combustion.math_utils` | Bisection, interpolation, polynômes Cp |

---

## Dépendances

```
pip install openpyxl
```

Python 3.10+ requis. Le solveur **Thermette.exe** (ARMINES/CES) doit être installé séparément dans `C:\thermette\`.

## Options

```
py calculer.py [--excel      "chemin/vers/classeur.xlsx"]
               [--thermette  "C:\thermette\thermette.exe"]
               [--modray1    "C:\modray\Modray1.exe"]
               [--modray2    "C:\modray\Modray2.exe"]
               [--out        "chemin/résultats.xlsx"]
```

| Option | Défaut | Description |
|--------|--------|-------------|
| `--excel` | `BLD VRTF 1.1_modifiable.xlsx` | Classeur source |
| `--thermette` | `C:\thermette\thermette.exe` | Solveur réseau thermique |
| `--modray1` | *(non fourni)* | Modray1.exe — calcul des facteurs de forme |
| `--modray2` | *(non fourni)* | Modray2.exe — calcul des facteurs de forme |
| `--out` | `<nom>_résultats.xlsx` | Classeur de sortie |

Si `--modray1` et `--modray2` sont fournis, les facteurs de forme sont recalculés pour toutes les sections et la feuille **Mesh** du classeur source est mise à jour automatiquement.

```
py run_vrtf.py projet.plf --thermette "C:\thermette\thermette.exe"
               [--modray1 "C:\modray\Modray1.exe"] [--modray2 "C:\modray\Modray2.exe"]
```
