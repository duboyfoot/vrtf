"""Tests de validation des lookups Python vs valeurs Excel attendues."""

import sys
sys.path.insert(0, r"C:\Users\patrick pro\Documents\kappa-dubois\developpement")
import vrtf

PASS = []
FAIL = []

def check(label, got, expected, tol=0.01):
    err = abs(got - expected) / (abs(expected) or 1)
    if err <= tol:
        PASS.append(f"  OK  {label}: {got:.4g}")
    else:
        FAIL.append(f"  FAIL {label}: {got:.4g} (attendu {expected:.4g}, ecart {err:.1%})")

def section(title):
    print(f"\n=== {title} ===")

# ------------------------------------------------------------------
section("Grades BISRA disponibles")
grades = vrtf.bisra_grades()
print(f"  {len(grades)} grades : {grades[:5]} ...")

# ------------------------------------------------------------------
section("Masse volumique BISRA (scalaire)")
check("BISRA 01 rho", vrtf.steel_density("BISRA 01"), 7876)
check("BISRA 06 rho", vrtf.steel_density("BISRA 06"), 7855)

# ------------------------------------------------------------------
section("Cp acier (interpolation)")
check("BISRA 01 Cp @ 273K", vrtf.steel_cp("BISRA 01", 273), 485.6)
check("BISRA 01 Cp @ 373K", vrtf.steel_cp("BISRA 01", 373), 494.0)
# Interpolation entre 273 et 373
cp_interp = vrtf.steel_cp("BISRA 01", 323)
check("BISRA 01 Cp @ 323K (interpolé)", cp_interp, (485.6 + 494.0) / 2, tol=0.02)

# ------------------------------------------------------------------
section("Conductivité thermique acier (interpolation)")
check("BISRA 01 Co @ 253K", vrtf.steel_conductivity("BISRA 01", 253), 65.302)
check("BISRA 01 Co @ 323K", vrtf.steel_conductivity("BISRA 01", 323), 62.791)

# ------------------------------------------------------------------
section("Emissivite acier (interpolation)")
check("BISRA 01 eps @ 253K", vrtf.steel_emissivity("BISRA 01", 253), 0.4)
check("BISRA 02 eps @ 253K", vrtf.steel_emissivity("BISRA 02", 253), 0.8)

# ------------------------------------------------------------------
section("Proprietes gaz elementaires")
ch4 = vrtf.gas_props("CH4")
print(f"  CH4 molar_mass = {ch4['molar_mass']:.3f} g/mol  (attendu 16.043)")
print(f"  CH4 C={ch4['C']}, H={ch4['H']}, O={ch4['O']}")
print(f"  CH4 PCI = {ch4['pci']:.3e} J/kmol  (attendu ~802e6)")
check("CH4 masse molaire",    ch4["molar_mass"], 16.043)
check("CO2 masse molaire",    vrtf.gas_props("CO2")["molar_mass"], 44.01)
check("H2O masse molaire",    vrtf.gas_props("H2O")["molar_mass"], 18.015)
check("H2  masse molaire",    vrtf.gas_props("H2")["molar_mass"],   2.016)

# ------------------------------------------------------------------
section("Cp gaz par polynome")
# Cp O2 @ 273K : a=876.317, +b*T = 876.317 + 0.122828*273 ≈ 909.8
cp_o2 = vrtf.gas_cp("O2", 273)
print(f"  O2 Cp @ 273K = {cp_o2:.2f} J/kg.K")
check("O2 Cp @273K", cp_o2, vrtf.gas_props("O2")["cp_273"], tol=0.05)

# ------------------------------------------------------------------
section("Enthalpie gaz (integrale polynome)")
# H @ T_ref=273 doit être ~ 0
h0 = vrtf.gas_enthalpy("O2", 273, 273)
check("O2 h(273,273)=0", h0, 0.0, tol=1e-6)
# Valeur croissante avec T
h1 = vrtf.gas_enthalpy("CH4", 1000, 273)
print(f"  CH4 h(1000K-273K) = {h1/1000:.1f} kJ/kg")

# ------------------------------------------------------------------
section("Combustibles disponibles")
fuels = vrtf.fuel_names()
print(f"  {len(fuels)} combustibles : {fuels[:5]} ...")

section("Composition combustible")
gn = vrtf.fuel_composition("Gaz naturel")
print(f"  Gaz naturel : {gn}")
check("GN CH4%", gn.get("CH4", 0), 94.0, tol=0.02)

# ------------------------------------------------------------------
section("Comburants disponibles")
print(f"  {vrtf.comburant_names()}")
air = vrtf.comburant_composition("Air_sec")
check("Air O2%", air.get("O2", 0), 20.8)
check("Air N2%", air.get("N2", 0), 79.2)

# ------------------------------------------------------------------
section("PCI combustible")
pci_gn = vrtf.lhv_vol("Gaz naturel")
print(f"  PCI Gaz naturel = {pci_gn/1e6:.2f} MJ/Nm3  (attendu ~36-38)")
check("PCI GN ordre de grandeur", pci_gn, 37e6, tol=0.05)

# ------------------------------------------------------------------
section("Stoechimetrie")
va = vrtf.stoich_air_vol("Gaz naturel")
print(f"  Va Gaz naturel = {va:.3f} Nm3_air/Nm3_fuel  (attendu ~9.5-10.0)")
check("Va GN", va, 9.7, tol=0.05)

# ------------------------------------------------------------------
section("Composition fumees (exces d'air)")
fumees = vrtf.flue_gas_composition("Gaz naturel", air_fuel_ratio=1.1)
print("  Fumees (10% exces air) :")
for k, v in fumees.items():
    if v: print(f"    {k}: {v:.2f}%")
check("CO2 fumees GN",  fumees["CO2"], 8.9, tol=0.1)

# ------------------------------------------------------------------
section("Temperature adiabatique")
T_ad = vrtf.adiabatic_temperature("Gaz naturel", air_fuel_ratio=1.05)
print(f"  T_adiab GN, 5% exces = {T_ad:.0f} K  (attendu ~2100-2250 K)")
check("T_adiab GN ordre", T_ad, 2150, tol=0.1)

# ------------------------------------------------------------------
section("Masse volumique melange")
rho = vrtf.mixture_density({"CH4": 100}, 273)
print(f"  rho CH4 @ 273K = {rho:.4f} kg/m3  (attendu 0.7167)")
check("rho CH4 @ 273K", rho, 16.043 / 22.4136, tol=0.01)

# ------------------------------------------------------------------
section("Emissivite Hottel - CO2 seul")
# pCO2=0.1 atm, pH2O=0, T=1000°C, L=1m
# Calcul manuel : R_CO2=0.80m, pCO2*R=0.08 atm.m → eps_CO2 ≈ 0.094
eps_co2_only = vrtf.gas_emissivity(pCO2=0.1, pH2O=0.0, T_celsius=1000, beam_length=1.0)
print(f"  eps(CO2=0.1, H2O=0, T=1000C, L=1m) = {eps_co2_only:.4f}  (attendu ~0.094)")
check("eps CO2-seul 1000C", eps_co2_only, 0.094, tol=0.05)

section("Emissivite Hottel - H2O seul")
# pCO2=0, pH2O=0.2 atm, T=600°C, L=1m → eps_H2O ~ 0.20-0.25
eps_h2o_only = vrtf.gas_emissivity(pCO2=0.0, pH2O=0.2, T_celsius=600, beam_length=1.0)
print(f"  eps(CO2=0, H2O=0.2, T=600C, L=1m)  = {eps_h2o_only:.4f}  (attendu ~0.20-0.25)")
check("eps H2O-seul 600C", eps_h2o_only, 0.22, tol=0.15)

section("Emissivite Hottel - CO2+H2O (fumees typiques)")
# pCO2=0.09, pH2O=0.17, T=900°C, L=2m → eps_tot attendu ~0.25-0.40
eps_mix = vrtf.gas_emissivity(pCO2=0.09, pH2O=0.17, T_celsius=900, beam_length=2.0)
print(f"  eps(CO2=0.09, H2O=0.17, T=900C, L=2m) = {eps_mix:.4f}  (attendu 0.25-0.40)")
check("eps mix 900C plausible", eps_mix, 0.32, tol=0.30)

section("Emissivite Hottel - croissance avec L")
eps_L1 = vrtf.gas_emissivity(pCO2=0.1, pH2O=0.1, T_celsius=800, beam_length=0.5)
eps_L2 = vrtf.gas_emissivity(pCO2=0.1, pH2O=0.1, T_celsius=800, beam_length=2.0)
print(f"  eps(L=0.5m) = {eps_L1:.4f}   eps(L=2.0m) = {eps_L2:.4f}   (attendu L2 > L1)")
if eps_L2 > eps_L1:
    PASS.append(f"  OK  eps croit avec L : {eps_L1:.4f} < {eps_L2:.4f}")
else:
    FAIL.append(f"  FAIL eps ne croit pas avec L : {eps_L1:.4f} >= {eps_L2:.4f}")

section("Coefficient d'absorption Hottel")
kappa = vrtf.absorption_coefficient(pCO2=0.1, pH2O=0.15, T_celsius=800, beam_length=1.0)
print(f"  kappa(CO2=0.1, H2O=0.15, T=800C, L=1m) = {kappa:.4f} 1/(atm.m)  (attendu > 0)")
check("kappa positif", kappa, 0.35, tol=0.5)

# ------------------------------------------------------------------
section("polint3 - polynome cubique")
# f(x) = 1 + 2x + 3x^2 + 4x^3  => f(0)=1, f(1)=10, f(2)=49, f(3)=142
poly_test = vrtf.polint3([0.0, 1.0, 2.0, 3.0], [1.0, 10.0, 49.0, 142.0])
check("polint3 a", poly_test[0], 1.0,  tol=1e-6)
check("polint3 b", poly_test[1], 2.0,  tol=1e-6)
check("polint3 c", poly_test[2], 3.0,  tol=1e-6)
check("polint3 d", poly_test[3], 4.0,  tol=1e-6)
# Reconstruction aux 4 points
for xi, yi in zip([0, 1, 2, 3], [1, 10, 49, 142]):
    p_val = poly_test[0] + poly_test[1]*xi + poly_test[2]*xi**2 + poly_test[3]*xi**3
    check(f"polint3 f({xi})", p_val, yi, tol=1e-6)

# ------------------------------------------------------------------
section("heat_exchanger - echangeur parfait fluides identiques")
# Meme fluide (air), meme debit, eff=100%, rend=100%
# => T_hot_out = T_cold_in, T_cold_out = T_hot_in
air_compo = vrtf.comburant_composition("Air_sec")
T_hx_hot_out, T_hx_cold_out = vrtf.heat_exchanger(
    air_compo, 1.0, 1000.0,
    air_compo, 1.0,  300.0,
    100.0, 100.0,
)
print(f"  Air-Air eff=100%: T_hot_out={T_hx_hot_out:.1f} K (attendu 300), T_cold_out={T_hx_cold_out:.1f} K (attendu 1000)")
check("heat_exchanger T_hot_out",  T_hx_hot_out,  300.0, tol=0.01)
check("heat_exchanger T_cold_out", T_hx_cold_out, 1000.0, tol=0.01)

section("heat_exchanger - eff=50%")
# eff=50%, meme fluide, meme debit  => T_hot_out = milieu entre hot_in et cold_in
T_hx_hot_out50, T_hx_cold_out50 = vrtf.heat_exchanger(
    air_compo, 1.0, 1000.0,
    air_compo, 1.0,  300.0,
    50.0, 100.0,
)
print(f"  Air-Air eff=50%: T_hot_out={T_hx_hot_out50:.1f} K, T_cold_out={T_hx_cold_out50:.1f} K")
if T_hx_cold_out50 > 300.0 and T_hx_cold_out50 < 1000.0:
    PASS.append(f"  OK  heat_exchanger eff50 T_cold_out in range: {T_hx_cold_out50:.1f}")
else:
    FAIL.append(f"  FAIL heat_exchanger eff50 T_cold_out hors plage: {T_hx_cold_out50:.1f}")

# ------------------------------------------------------------------
section("solve_zone - zone simple sans regenerateur")
zone_gn = {
    "fuel": "Gaz naturel",
    "air": "Air_sec",
    "excess_air_pct": 10.0,
    "power_W": 100e3,      # 100 kW
    "T_fuel_K": 293.0,
    "T_air_K": 293.0,
    "regen_air_eff_pct": 0.0,
    "regen_air_aspi_pct": 0.0,
}
res = vrtf.solve_zone(zone_gn)
print(f"  Debit fuel = {res['fuel_Nm3h']:.3f} Nm3/h  (attendu ~ 10 Nm3/h pour 100kW GN)")
print(f"  Debit air  = {res['air_Nm3h']:.2f} Nm3/h")
print(f"  Debit fum  = {res['flue_kgs']*3600:.3f} kg/h")
print(f"  CO2 fumees = {res['flue_compo']['CO2']:.2f}%  (attendu ~8-9%)")

# PCI GN ~ 36 MJ/Nm3 => 100kW => fuel_Nm3h ~ 100e3/36e6*3600 ~ 10 Nm3/h
check("solve_zone fuel_Nm3h", res["fuel_Nm3h"], 9.9, tol=0.05)

# Le polynome enthalpie air doit valoir 0 a T=273 K
h_air_273 = (res["poly_enthalp_air"][0]
           + res["poly_enthalp_air"][1] * 273
           + res["poly_enthalp_air"][2] * 273**2
           + res["poly_enthalp_air"][3] * 273**3)
check("solve_zone poly_enthalp_air(273K)=0", h_air_273, 0.0, tol=0.01)

# Cp fumees a 273 K doit etre positif et raisonnable (~1000-1200 J/kg.K)
cp_fum_273 = (res["poly_cp_flue"][0]
            + res["poly_cp_flue"][1] * 273
            + res["poly_cp_flue"][2] * 273**2
            + res["poly_cp_flue"][3] * 273**3)
print(f"  Cp fumees @ 273K = {cp_fum_273:.1f} J/kg.K  (attendu 1000-1200)")
check("solve_zone cp_fum_273K", cp_fum_273, 1100.0, tol=0.1)

# T_eq_air doit etre proche de T_air_K (gaz froid a l'entree)
T_eq_air_val = res["poly_T_eq_air"][0]
print(f"  T_eq_air = {T_eq_air_val:.1f} K  (attendu ~ T_air = 293 K)")
check("solve_zone T_eq_air ~ T_air", T_eq_air_val, 293.0, tol=0.02)

section("solve_combustion - liste de zones")
zones_list = [zone_gn, {**zone_gn, "power_W": 200e3}]
results = vrtf.solve_combustion(zones_list)
check("solve_combustion nb_zones", len(results), 2.0, tol=0.0)
check("solve_combustion zone2 debit", results[1]["fuel_Nm3h"], 2 * results[0]["fuel_Nm3h"], tol=0.01)

# ------------------------------------------------------------------
section("Thermette - geometrie tubes U et W")
import math as _math
from vrtf.thermette import _tube_geom
_PI = _math.pi

_tu = {"type": "U", "a": 1.0, "b": 0.05, "c": 0.5, "n_tubes": 2}
_vol_u, _surf_u = _tube_geom(_tu)
_vol_u_exp  = (1.0 * _PI * 0.05**2) * 2 + (_PI * 0.5 * _PI * 0.05**2) * 2
_surf_u_exp = ((_PI * 0.05 * 1.0 * 4) + (2 * _PI * 0.05 * _PI * 0.5)) * 2
check("tube_U vol",  _vol_u,  _vol_u_exp,  tol=1e-6)
check("tube_U surf", _surf_u, _surf_u_exp, tol=1e-6)

_tw = {"type": "W", "a": 1.0, "b": 0.05, "c": 0.5, "d": 0.3, "n_tubes": 2}
_vol_w, _surf_w = _tube_geom(_tw)
_vol_w_exp  = ((1.0*_PI*0.05**2)*2 + (_PI*0.5*_PI*0.05**2)*3 + (0.3*_PI*0.05**2)*2) * 2
_surf_w_exp = ((_PI*0.05*1.0*4) + (2*_PI*0.05*_PI*0.5*3) + (_PI*0.05*0.3*4)) * 2
check("tube_W vol",  _vol_w,  _vol_w_exp,  tol=1e-6)
check("tube_W surf", _surf_w, _surf_w_exp, tol=1e-6)

# U premier terme NON multiplié par n_tubes (fidèle au bug VBA)
_tu1 = {"type": "U", "a": 1.0, "b": 0.05, "c": 0.5, "n_tubes": 1}
_tu4 = {"type": "U", "a": 1.0, "b": 0.05, "c": 0.5, "n_tubes": 4}
_vol_u1, _ = _tube_geom(_tu1)
_vol_u4, _ = _tube_geom(_tu4)
_first_term = (1.0 * _PI * 0.05**2) * 2
_second_n1  = (_PI * 0.5 * _PI * 0.05**2) * 1
_second_n4  = (_PI * 0.5 * _PI * 0.05**2) * 4
check("tube_U vol n=1", _vol_u1, _first_term + _second_n1, tol=1e-6)
check("tube_U vol n=4", _vol_u4, _first_term + _second_n4, tol=1e-6)

# ------------------------------------------------------------------
section("Thermette - ecriture fichiers")
import tempfile, os as _os
from vrtf.thermette import write_fluides_prp, write_calc, write_reseau, write_thermette_files

_tmpdir = tempfile.mkdtemp()

# fluides.prp
write_fluides_prp(_tmpdir)
_lines_prp = open(_os.path.join(_tmpdir, "fluides.prp")).read().splitlines()
assert _lines_prp[0] == "#BISRA 0 3000", f"fluides.prp[0] = {_lines_prp[0]!r}"
assert "table(T,AcierCo)" in _lines_prp, "table AcierCo manquante"
assert "table(T,AcierCp)" in _lines_prp, "table AcierCp manquante"
PASS.append("  OK  fluides.prp contenu")

_cfg = {
    "steel_grade": "BISRA 01",
    "strip_thickness_m": 0.005,
    "strip_width_m": 1.2,
    "strip_flowrate_kgs": 0.5,
    "T_strip_in_K": 573.0,
    "n_zones": 2,
    "rows_per_zone": [2, 2],
    "zone_height_m": 1.5,
    "zone_width_m": 2.0,
    "wall_thickness_m": 0.2,
    "wall_conductivity": 1.2,
    "gap_band_m": [0.05, 0.05],
    "tubes": [
        {"zone": 1, "row": 1, "type": "U", "a": 1.0, "b": 0.05, "c": 0.5, "n_tubes": 4, "T_tube_K": 1200.0},
        {"zone": 1, "row": 2, "type": "U", "a": 1.0, "b": 0.05, "c": 0.5, "n_tubes": 4, "T_tube_K": 1200.0},
        {"zone": 2, "row": 1, "type": "W", "a": 1.0, "b": 0.05, "c": 0.5, "d": 0.3, "n_tubes": 4, "T_tube_K": 1100.0},
        {"zone": 2, "row": 2, "type": "W", "a": 1.0, "b": 0.05, "c": 0.5, "d": 0.3, "n_tubes": 4, "T_tube_K": 1100.0},
    ],
}

# VRTF_reseau
write_reseau(_tmpdir, _cfg)
_res_text = open(_os.path.join(_tmpdir, "VRTF_reseau")).read()
assert "**sollicit prod 0.5" in _res_text, "sollicit prod manquant"
assert "**sollicit T_entreebande 573.0" in _res_text, "T_entreebande manquant"
assert "**sollicit T_tube_Z1_R1 1200.0" in _res_text, "T_tube Z1R1 manquant"
assert "**branche B_S1_R1 NS0 NS1" in _res_text, "branche bande manquante"
assert "**branche B_S2_R1 NS1 NS2" in _res_text, "branche bande zone2 manquante"
assert "**volume Tube_S1_R1" in _res_text, "volume tube manquant"
assert "*entree temp T_entreebande" in _res_text, "*entree manquant"
# Zone 1 → Bas (pas dans top_zones par défaut), zone 2 → Haut
assert "Bmur_S1_R1Bas" in _res_text, "mur Bas zone1 manquant"
assert "Bmur_S2_R1Haut" in _res_text, "mur Haut zone2 manquant"
# Zones 1,2 ont convection externe sur Gauche/Droit
assert "*sortie conv Text 10+sfrad(T,Text)" in _res_text, "sortie conv manquante"
PASS.append("  OK  VRTF_reseau contenu")

# VRTF_calc : Ns = 2*(2-1+2-1)+2 + 2*(2+2) = 2*2+2 + 8 = 14
write_calc(_tmpdir, _cfg)
_calc_text = open(_os.path.join(_tmpdir, "VRTF_calc")).read()
assert "*calcul statique" in _calc_text, "*calcul manquant"
assert "*sorties 14" in _calc_text, f"sorties attendu 14, calcul:\n{_calc_text[:300]}"
assert "S_Prod_IN " in _calc_text, "S_Prod_IN manquant"
assert "S_Prod_OUT " in _calc_text, "S_Prod_OUT manquant"
assert "SFluxMur_S1_R1 " in _calc_text, "SFluxMur manquant"
assert "SFluxTube_S2_R2 " in _calc_text, "SFluxTube manquant"
# Zone 1 mur = Bas, zone 2 = Haut
assert "Fcond(Bmur_S1_R1Bas,0)" in _calc_text, "Bas dans calc manquant"
assert "Fcond(Bmur_S2_R2Haut,0)" in _calc_text, "Haut dans calc manquant"
PASS.append("  OK  VRTF_calc contenu")

# write_thermette_files crée tous les fichiers d'un coup
_tmpdir2 = tempfile.mkdtemp()
write_thermette_files(_tmpdir2, _cfg)
for _fname in ["AcierCp", "AcierCo", "AcierRo", "fluides.prp", "VRTF_reseau", "VRTF_calc"]:
    assert _os.path.exists(_os.path.join(_tmpdir2, _fname)), f"{_fname} non créé"
PASS.append("  OK  write_thermette_files - tous les fichiers créés")

# ------------------------------------------------------------------
section("PostTraitement - parse_results")
import tempfile as _tempfile
from vrtf.posttreatment import parse_results, postprocess

# Fichier résultat factice
_RESU_OK = """\
CALCUL TERMINE NORMALEMENT

Valeurs des sorties
SStrip_S1_R1 : 1123
SFluxStrip_S1_R1 : 45000.0
SFluxMur_S1_R1 : -1200.5
SFluxTube_S1_R1 : -80000.0
SFluxTube_S2_R1 : -60000.0
S_Prod_IN : 5000.0
S_Prod_OUT : 50000.0
"""

_RESU_NOK = "Calcul interrompu\nPas de convergence\n"

_tmp_ok  = _tempfile.NamedTemporaryFile(mode="w", suffix=".resu", delete=False)
_tmp_ok.write(_RESU_OK); _tmp_ok.close()
_tmp_nok = _tempfile.NamedTemporaryFile(mode="w", suffix=".resu", delete=False)
_tmp_nok.write(_RESU_NOK); _tmp_nok.close()

_r = parse_results(_tmp_ok.name)
assert _r["SStrip_S1_R1"] == 1123.0,    f"SStrip attendu 1123, got {_r['SStrip_S1_R1']}"
assert _r["SFluxTube_S1_R1"] == -80000, f"SFluxTube attendu -80000, got {_r['SFluxTube_S1_R1']}"
assert _r["S_Prod_IN"] == 5000.0,       f"S_Prod_IN attendu 5000, got {_r['S_Prod_IN']}"
PASS.append("  OK  parse_results valeurs correctes")

_ok_raised = False
try:
    parse_results(_tmp_nok.name)
except RuntimeError:
    _ok_raised = True
assert _ok_raised, "RuntimeError attendu si non convergé"
PASS.append("  OK  parse_results leve RuntimeError si non converge")

section("PostTraitement - postprocess")
_cfg_pp = {
    "n_zones": 2,
    "rows_per_zone": [1, 1],
    "tubes": [
        {"zone": 1, "row": 1, "type": "U", "a": 1.0, "b": 0.05, "c": 0.5,
         "n_tubes": 4, "T_tube_K": 1200.0, "real_zone": 1, "tube_eff_pct": 80.0},
        {"zone": 2, "row": 1, "type": "U", "a": 1.0, "b": 0.05, "c": 0.5,
         "n_tubes": 4, "T_tube_K": 1100.0, "real_zone": 2, "tube_eff_pct": 80.0},
    ],
    # autres clés non utilisées par postprocess
    "steel_grade": "BISRA 01", "strip_thickness_m": 0.005, "strip_width_m": 1.2,
    "strip_flowrate_kgs": 0.5, "T_strip_in_K": 573.0,
    "zone_height_m": 1.5, "zone_width_m": 2.0,
    "wall_thickness_m": 0.2, "wall_conductivity": 1.2, "gap_band_m": [0.05, 0.05],
}

_res = postprocess(_tmp_ok.name, _cfg_pp, "Gaz naturel", "Air_sec", excess_air_pct=10.0)

# flux_tube_W : signe inversé vs fichier (-80000 → +80000)
assert _res["flux_tube_W"].get((1, 1), 0) == 80000.0, \
    f"flux_tube_W Z1R1 attendu 80000, got {_res['flux_tube_W'].get((1,1))}"
assert _res["flux_tube_W"].get((2, 1), 0) == 60000.0, \
    f"flux_tube_W Z2R1 attendu 60000, got {_res['flux_tube_W'].get((2,1))}"

# flux_mur_W
assert _res["flux_mur_W"].get((1, 1), 0) == -1200.5, \
    f"flux_mur_W Z1R1 attendu -1200.5, got {_res['flux_mur_W'].get((1,1))}"

# flux_fum_W : (1-0.8)/0.8 * 80000 = 20000
_fum_11 = _res["flux_fum_W"].get((1, 1), 0)
check("flux_fum Z1R1", _fum_11, 20000.0, tol=1e-6)

# deb_gaz : Q_tube / PCI / eff  [Nm3/h]
# PCI GN ~ 37.4e6 J/Nm3, eff=0.8
# deb = 80000 / 37.4e6 / 0.8 * 3600 ~ 9.6e-3 * 3600 ~ 0.026... Nm3/h
_deb = _res["deb_gaz_Nm3h"].get((1, 1), 0)
print(f"  deb_gaz Z1R1 = {_deb:.4f} Nm3/h")
assert _deb > 0, "deb_gaz doit etre positif"
PASS.append("  OK  postprocess deb_gaz positif")

# T_fum > T_REF si flux non nul
_T_fum = _res["T_fum_K"].get((1, 1), 0)
print(f"  T_fum Z1R1 = {_T_fum:.1f} K")
assert _T_fum > 273.0, f"T_fum attendu > 273K, got {_T_fum}"
PASS.append("  OK  postprocess T_fum > T_ref")

# Agrégats par zone
assert _res["flux_tube_zone_W"][1] == 80000.0
assert _res["flux_tube_zone_W"][2] == 60000.0
check("flux_tube_four_W", _res["flux_tube_four_W"], 140000.0, tol=1e-6)
check("flux_mur_four_W", _res["flux_mur_four_W"], -1200.5, tol=1e-4)
PASS.append("  OK  postprocess agregats zones et four")

# E_strip
assert _res["E_strip_in_W"]  == 5000.0
assert _res["E_strip_out_W"] == 50000.0
PASS.append("  OK  postprocess E_strip_in / E_strip_out")

import os as _os2
_os2.unlink(_tmp_ok.name)
_os2.unlink(_tmp_nok.name)

# ------------------------------------------------------------------
section("Modray - write_modray_section")
import tempfile as _tempfile2, os as _os3
from vrtf.modray import write_modray_section, _n_tube_parts, _n_wall_strip_objects

# cfg minimaliste pour 2 zones, 2 rangées chacune
_cfg_mr = {
    "n_zones": 3,
    "rows_per_zone": [2, 2, 2],
    "zone_height_m":   1.5,
    "zone_width_m":    2.0,
    "strip_thickness_m": 0.005,
    "strip_width_m":   1.2,
    "gap_band_m":      [0.08, 0.08, 0.08],
    "tubes": [
        {"zone": 1, "row": 1, "type": "U", "a": 1.5, "b": 0.05, "c": 0.5,
         "d": 0.0, "n_tubes": 4, "T_tube_K": 1200.0,
         "prem_abs_m": 0.3, "dist_band_tube_m": 0.04, "pitch_m": 0.25, "emissivity": 0.85},
        {"zone": 1, "row": 2, "type": "W", "a": 1.5, "b": 0.05, "c": 0.5,
         "d": 0.3, "n_tubes": 2, "T_tube_K": 1200.0,
         "prem_abs_m": 0.3, "dist_band_tube_m": 0.04, "pitch_m": 0.25, "emissivity": 0.85},
    ],
}

_tmpd = _tempfile2.mkdtemp()

# --- Zone 1, rangée 1 (bord gauche, tubes U) ---
_sec_file = _os3.path.join(_tmpd, "VRTF_S1_R1")
write_modray_section(_sec_file, _cfg_mr, zone=1, row=1)
_txt = open(_sec_file).read()

# En-tête
assert "**Nbandes 1" in _txt, "Nbandes manquant"
assert "1 diffus Res_VRTF_R" in _txt, "Calculs groupe manquant"
# Nobj : 3*4(U) + 7(bord) = 12 + 7 = 19
assert "**nobjets 19" in _txt, f"nobjets attendu 19, got: {[l for l in _txt.splitlines() if 'nobjets' in l]}"
# Mur gauche (j=1)
assert "**objet mur_S1_R1Gauche_0" in _txt, "mur Gauche manquant"
# Pas de mur droit (j != n_rows=2)
assert "mur_S1_R1Droit_0" not in _txt, "mur Droit ne doit pas être là pour j=1 de 2"
# Mur Lat, Haut, Bas
assert "mur_S1_R1Lat_0" in _txt, "mur Lat manquant"
assert "mur_S1_R1Haut_0" in _txt, "mur Haut manquant"
assert "mur_S1_R1Bas_0"  in _txt, "mur Bas manquant"
# Zone 1 pas dans top_zones → proprad 0 sur Haut, 0.85 sur Bas
_haut_idx = _txt.index("mur_S1_R1Haut_0")
_haut_block = _txt[_haut_idx:_haut_idx+200]
assert "*proprad 0 0" in _haut_block, "Haut doit avoir proprad 0 pour zone 1"
# Strip gauche (j < n_rows) : B_S1_R1
assert "**objet B_S1_R1\n" in _txt, "strip B_S1_R1 manquant"
# Pas de strip principal droit (j=1 donc j>1 faux)
# NB: B_S1_R0_Sepa2 peut exister (bug VBA reproduit fidèlement)
assert "**objet B_S1_R1\n" in _txt,  "strip gauche B_S1_R1 attendu"
assert "**objet B_S1_R0\n" not in _txt, "strip principal B_S1_R0 ne doit pas exister"
# Tubes U — groupes
assert "**Groupes 1" in _txt, "Groupes manquant"
assert "Tube_S1_R1 12" in _txt, "groupe Tube_S1_R1 12 parties attendu (3*4)"
assert "Tubea_S1_R1_N1" in _txt, "objet Tubea manquant"
PASS.append("  OK  modray VRTF_S zone1 rangee1 (bord gauche, U)")

# --- Zone 1, rangée 2 (bord droit, tubes W) ---
_sec_file2 = _os3.path.join(_tmpd, "VRTF_S1_R2")
write_modray_section(_sec_file2, _cfg_mr, zone=1, row=2)
_txt2 = open(_sec_file2).read()
# Nobj : 7*2(W) + 7(bord) = 14 + 7 = 21
assert "**nobjets 21" in _txt2, f"nobjets attendu 21, got {[l for l in _txt2.splitlines() if 'nobjets' in l]}"
# Mur droit (j=2=n_rows)
assert "mur_S1_R2Droit_0" in _txt2, "mur Droit manquant pour j=2"
assert "mur_S1_R2Gauche_0" not in _txt2, "mur Gauche ne doit pas être là pour j=2=n_rows"
# Strip droit (j > 1) : B_S1_R1
assert "**objet B_S1_R1\n" in _txt2, "strip B_S1_R1 doit être écrit pour j=2"
# Tubes W
assert "Tubea_S1_R2_N1" in _txt2, "Tubea W manquant"
assert "Tubeg_S1_R2_N1" in _txt2, "Tubeg W (7ème partie) manquant"
assert "Tube_S1_R2 14" in _txt2, "groupe W 14 parties attendu (7*2)"
PASS.append("  OK  modray VRTF_S zone1 rangee2 (bord droit, W)")

# --- Zone 2, rangée 1 (bord, dans top_zones=={2,3,6,7,10}) ---
_sec_file3 = _os3.path.join(_tmpd, "VRTF_S2_R1")
write_modray_section(_sec_file3, _cfg_mr, zone=2, row=1)
_txt3 = open(_sec_file3).read()
# Zone 2 dans top_zones → Haut proprad 0.85, Bas proprad 0
_haut3_idx = _txt3.index("mur_S2_R1Haut_0")
_haut3_blk = _txt3[_haut3_idx:_haut3_idx+200]
assert "*proprad 0.85 0" in _haut3_blk, "Zone 2 (top_zones) : Haut doit avoir proprad 0.85"
PASS.append("  OK  modray top_zones : proprad Haut correct")

# --- Nombre d'objets : formule ---
check("n_tube_parts U 4", float(_n_tube_parts("U", 4)), 12.0, tol=0)
check("n_tube_parts W 3", float(_n_tube_parts("W", 3)), 21.0, tol=0)
check("n_tube_parts I 2", float(_n_tube_parts("I", 2)),  2.0, tol=0)
check("n_wall_strip bord",   float(_n_wall_strip_objects(3, 1)), 7.0, tol=0)
check("n_wall_strip milieu", float(_n_wall_strip_objects(3, 2)), 9.0, tol=0)

# --- generate_furnace_section_files ---
from vrtf.visualisation import generate_furnace_section_files
_tmpd2 = _tempfile2.mkdtemp()
_files = generate_furnace_section_files(_tmpd2, _cfg_mr)
# 3 zones × 2 rangées = 6 fichiers
assert len(_files) == 6, f"Attendu 6 fichiers section, got {len(_files)}"
for _fp in _files:
    assert _os3.path.exists(_fp), f"Fichier absent : {_fp}"
PASS.append("  OK  generate_furnace_section_files - 6 fichiers créés")

# Nettoyage
import shutil as _shutil
_shutil.rmtree(_tmpd, ignore_errors=True)
_shutil.rmtree(_tmpd2, ignore_errors=True)

# ------------------------------------------------------------------
section("Constants - valeurs physiques et utilitaires")
import math as _math2
from vrtf.constants import (
    SIGMA, PI, R, T_REF, VERSION,
    SOLVER_SIM_TIME, SOLVER_GAIN, SOLVER_INTEGRAL, SOLVER_DERIVATIVE,
    file_exists, shell_wait,
)

# Constantes physiques
check("SIGMA",  SIGMA, 5.67e-8,       tol=1e-3)
check("PI",     PI,    _math2.pi,     tol=1e-9)
check("R",      R,     8314.0,        tol=0)
check("T_REF",  T_REF, 273.0,         tol=0)

# Paramètres solveur
check("SOLVER_SIM_TIME",   SOLVER_SIM_TIME,   240.0, tol=0)
check("SOLVER_GAIN",       SOLVER_GAIN,        10.0, tol=0)
check("SOLVER_INTEGRAL",   SOLVER_INTEGRAL,    20.0, tol=0)
check("SOLVER_DERIVATIVE", SOLVER_DERIVATIVE,   0.0, tol=0)

# VERSION est une chaîne non vide
assert isinstance(VERSION, str) and VERSION, "VERSION doit être une chaîne non vide"
PASS.append("  OK  VERSION est une chaîne non vide")

# file_exists
import tempfile as _tf3, os as _os4
_f_exist = _tf3.NamedTemporaryFile(delete=False)
_f_exist.close()
assert file_exists(_f_exist.name),  "file_exists doit retourner True pour un fichier existant"
assert not file_exists(_f_exist.name + "_absent"), "file_exists doit retourner False si absent"
assert not file_exists(_os4.path.dirname(_f_exist.name)), "file_exists doit retourner False pour un répertoire"
_os4.unlink(_f_exist.name)
PASS.append("  OK  file_exists")

# shell_wait
import sys as _sys
_ret = shell_wait([_sys.executable, "-c", "import sys; sys.exit(0)"])
assert _ret == 0, f"shell_wait exit 0 attendu, got {_ret}"
_ret2 = shell_wait([_sys.executable, "-c", "import sys; sys.exit(42)"])
assert _ret2 == 42, f"shell_wait exit 42 attendu, got {_ret2}"
PASS.append("  OK  shell_wait code de retour")

# ------------------------------------------------------------------
section("Update - fuel_air_summary")
from vrtf.update import fuel_air_summary

_s = fuel_air_summary("Gaz naturel", "Air_sec", excess_air_pct=10.0)

# Champs présents
for _k in ("fuel_name","air_name","fuel_compo","air_compo",
           "density_kgNm3","lhv_J_Nm3","lhv_kJ_Nm3",
           "wobbe_J_Nm3","wobbe_MJ_Nm3","va_Nm3_Nm3","vf_Nm3_Nm3"):
    assert _k in _s, f"Clé manquante : {_k}"
PASS.append("  OK  fuel_air_summary - toutes les clés présentes")

# Cohérence des valeurs
assert _s["fuel_name"] == "Gaz naturel"
assert _s["air_name"]  == "Air_sec"
assert isinstance(_s["fuel_compo"], dict) and _s["fuel_compo"]
assert isinstance(_s["air_compo"],  dict) and _s["air_compo"]

# Densité GN : CH4 94%+C2H6 3%+C3H8 2%+N2 1% → M≈17.1 g/mol → ~0.763 kg/Nm³
check("density GN",    _s["density_kgNm3"], 0.763, tol=0.02)
# LHV GN ≈ 37 MJ/Nm³
check("lhv GN MJ/Nm3", _s["lhv_J_Nm3"]/1e6, 37.4, tol=0.05)
# kJ = J / 1000
check("lhv_kJ = lhv_J/1000", _s["lhv_kJ_Nm3"], _s["lhv_J_Nm3"]/1e3, tol=1e-9)
# Wobbe > LHV (densité relative < 1 pour GN)
assert _s["wobbe_J_Nm3"] > _s["lhv_J_Nm3"], "Wobbe doit être > LHV (GN plus léger que l'air)"
check("wobbe_MJ = wobbe_J/1e6", _s["wobbe_MJ_Nm3"], _s["wobbe_J_Nm3"]/1e6, tol=1e-9)
# Va GN ≈ 10 Nm³/Nm³
check("va GN",  _s["va_Nm3_Nm3"], 10.0, tol=0.05)
# Vf avec 10% excès > Va
assert _s["vf_Nm3_Nm3"] > _s["va_Nm3_Nm3"], "Vf doit être > Va"

# Sans excès d'air, Vf doit être plus petit
_s0 = fuel_air_summary("Gaz naturel", "Air_sec", excess_air_pct=0.0)
assert _s["vf_Nm3_Nm3"] > _s0["vf_Nm3_Nm3"], "Vf doit croître avec l'excès d'air"
PASS.append("  OK  fuel_air_summary - valeurs cohérentes (GN/Air)")

# ------------------------------------------------------------------
section("Fichiers - save_project / open_project")
import tempfile, os
from vrtf.fichiers import save_project, open_project

_proj = {
    "n_zones": 3,
    "rows_per_zone": [2, 2, 2],
    "H": 1.5,
    "tube_type": "U",
    "fuel_name": "Gaz naturel",
    "nested": {"key": "valeur", "num": 3.14},
}

with tempfile.TemporaryDirectory() as _tmpdir:
    _plf = os.path.join(_tmpdir, "test_projet.plf")

    # Sauvegarde
    save_project(_plf, _proj)
    assert os.path.isfile(_plf), "Le fichier .plf doit exister après save_project"
    PASS.append("  OK  save_project - fichier créé")

    # Rechargement exact
    _loaded = open_project(_plf)
    assert _loaded == _proj, f"Données rechargées différentes : {_loaded}"
    PASS.append("  OK  open_project - données identiques après round-trip")

    # Types conservés
    assert isinstance(_loaded["n_zones"], int), "n_zones doit rester int"
    assert isinstance(_loaded["H"], float), "H doit rester float"
    assert isinstance(_loaded["nested"], dict), "nested doit rester dict"
    PASS.append("  OK  open_project - types numériques et dict conservés")

    # Contenu JSON lisible
    import json as _json
    with open(_plf, encoding="utf-8") as _fh:
        _raw = _json.load(_fh)
    assert _raw == _proj
    PASS.append("  OK  save_project - format JSON lisible directement")

# Fichier inexistant
_raised = False
try:
    open_project("/chemin/inexistant/projet.plf")
except FileNotFoundError:
    _raised = True
assert _raised, "open_project doit lever FileNotFoundError pour un fichier absent"
PASS.append("  OK  open_project - FileNotFoundError si fichier absent")

# Contenu invalide (liste JSON au lieu d'objet)
with tempfile.NamedTemporaryFile(mode="w", suffix=".plf", delete=False, encoding="utf-8") as _tf:
    _tf.write("[1, 2, 3]")
    _bad_path = _tf.name
_raised2 = False
try:
    open_project(_bad_path)
except ValueError:
    _raised2 = True
finally:
    os.unlink(_bad_path)
assert _raised2, "open_project doit lever ValueError pour un JSON non-dict"
PASS.append("  OK  open_project - ValueError si contenu non-dict")

# ------------------------------------------------------------------
print("\n" + "="*50)
print(f"RESULTATS : {len(PASS)} OK, {len(FAIL)} ECHEC")
for line in FAIL:
    print(line)
if not FAIL:
    print("  Tous les lookups sont valides.")
