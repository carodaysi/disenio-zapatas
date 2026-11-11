# Contraste del programa con el Excel desarrollado en clase (hoja "Empotrado Art").
# Cargas, dimensiones y resultados de la hoja para las zapatas P3 a P7.
# Uso: python comparar_excel_clase.py  (escribe Datos_Excel_Clase.xlsx en esta carpeta)
import os
import sys
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(AQUI))
import Diseno_Zapatas_Aisladas as dz

SALIDA = AQUI

# Point: (B, H del Excel, [(combo, FZ, MX, MY)])
CARGAS = {
    3: (2.40, 0.45, [("DCON1", 59.1, 0.0, 0.0), ("DCON2", 61.8, 0.0, 0.0),
                     ("DCON3", 25.2, -0.1, -1.2), ("DCON4", 106.9, 0.0, 1.2),
                     ("DCON5", 13.6, 1.7, -0.3), ("DCON6", 118.5, -1.8, 0.3)]),
    4: (2.30, 0.50, [("DCON1", 98.7, 0.0, 0.0), ("DCON2", 108.5, 0.0, 0.0),
                     ("DCON3", 87.4, -0.1, -1.3), ("DCON4", 139.9, 0.0, 1.3),
                     ("DCON5", 121.7, 1.9, 0.0), ("DCON6", 105.6, -2.0, 0.0)]),
    5: (2.10, 0.45, [("DCON1", 66.5, 0.1, 0.0), ("DCON2", 70.0, 0.1, 0.0),
                     ("DCON3", 48.9, 0.0, -1.2), ("DCON4", 100.4, 0.1, 1.2),
                     ("DCON5", 117.3, 1.8, 0.3), ("DCON6", 32.0, -1.7, -0.3)]),
    6: (2.60, 0.60, [("DCON1", 83.3, 0.0, 0.0), ("DCON2", 89.3, 0.0, 0.0),
                     ("DCON3", 16.3, 0.0, -1.4), ("DCON4", 172.7, 0.0, 1.3),
                     ("DCON5", 31.8, 2.1, -0.3), ("DCON6", 157.2, -2.2, 0.3)]),
    7: (3.00, 0.65, [("DCON1", 134.5, 0.0, 0.0), ("DCON2", 152.5, 0.0, 0.0),
                     ("DCON3", 78.0, 0.0, -1.3), ("DCON4", 237.5, 0.0, 1.2),
                     ("DCON5", 175.5, 2.3, 0.0), ("DCON6", 140.1, -2.3, 0.0)]),
}
# Resultados de la hoja de clase
EXCEL = {
    3: dict(Mu_x=20.10, Mu_y=22.91, n=17, db=14, As_min=16.20, As=26.17, d=37.5,
            phiMn=35.83, Vu_x=26.05, Vu_y=29.22, phiVc=29.60, Vu_p=100.98, bo=370,
            phiVp=165.88),
    4: dict(Mu_x=22.37, Mu_y=24.64, n=16, db=14, As_min=17.60, As=24.63, d=42.5,
            phiMn=38.40, Vu_x=26.35, Vu_y=28.89, phiVc=30.65, Vu_p=114.78, bo=390,
            phiVp=None),
    5: dict(Mu_x=15.81, Mu_y=18.47, n=15, db=14, As_min=14.18, As=23.09, d=37.5,
            phiMn=31.60, Vu_x=21.09, Vu_y=24.58, phiVc=25.97, Vu_p=94.62, bo=370,
            phiVp=165.88),
    6: dict(Mu_x=33.65, Mu_y=36.64, n=18, db=14, As_min=24.57, As=27.71, d=52.5,
            phiMn=53.67, Vu_x=32.01, Vu_y=34.90, phiVc=39.82, Vu_p=143.27, bo=430,
            phiVp=None),
    7: dict(Mu_x=57.44, Mu_y=61.86, n=21, db=16, As_min=31.05, As=42.22, d=57.5,
            phiMn=89.13, Vu_x=49.90, Vu_y=53.45, phiVc=53.57, Vu_p=204.18, bo=450,
            phiVp=None),
}

filas = []
for p, (B, H, cs) in CARGAS.items():
    for nombre, fz, mx, my in cs:
        filas.append(dict(Story="BASE", Point=p, Load=nombre, FX=0.0, FY=0.0, FZ=fz,
                          MX=mx, MY=my, MZ=0.0, **{"Bx[m]": B, "By[m]": B,
                                                   "Bcx[m]": 0.60, "Bcy[m]": 0.50}))
ruta = SALIDA + r"\Datos_Excel_Clase.xlsx"
pd.DataFrame(filas).to_excel(ruta, sheet_name="Clase", index=False)

# El Excel de clase no suma el peso propio: q = P/A ± M·c/I
dz.INCLUIR_PESO_PROPIO = False
df = dz.leer_datos(ruta, "Clase")

out = []
for p in CARGAS:
    datos = df[df["Point"] == p].reset_index(drop=True)
    H = CARGAS[p][1]
    ex = EXCEL[p]
    # (a) misma h y mismo diámetro que el Excel
    dz.H_MIN_PUNTO = {p: H}
    r0 = dz.disenar_zapata(p, datos)          # para geo, combos, campos
    rv = dz.verificar(r0["geo"], r0["combos"], r0["campos"], H, ex["db"])
    # (b) diseño libre del programa (h desde H_MIN)
    dz.H_MIN_PUNTO = {}
    rl = dz.disenar_zapata(p, datos)
    out.append((p, H, ex, rv, rl))

def f(x, n=2):
    return "—" if x is None else f"{x:.{n}f}"

for p, H, ex, rv, rl in out:
    print(f"\n=== P{p}  B = {CARGAS[p][0]:.2f} m, columna 0.60x0.50, h = {H*100:.0f} cm, Ø{ex['db']} ===")
    print(f"{'Magnitud':28s}{'Excel':>10s}{'Programa':>10s}{'Dif %':>8s}")
    pares = [
        ("Mu,x [t·m]", ex["Mu_x"], rv["Mu_x"]),
        ("Mu,y [t·m]", ex["Mu_y"], rv["Mu_y"]),
        ("d x [cm]", ex["d"], rv["dx"] * 100), ("d y [cm]", ex["d"], rv["dy"] * 100),
        ("As,min [cm²]", ex["As_min"], rv["As_min_x"]),
        ("As req x [cm²]", None, rv["As_req_x"]), ("As req y [cm²]", None, rv["As_req_y"]),
        ("As prov x [cm²]", ex["As"], rv["bar_x"]["As"]),
        ("As prov y [cm²]", ex["As"], rv["bar_y"]["As"]),
        ("φMn x [t·m]", ex["phiMn"], rv["phiMn_x"]),
        ("φMn y [t·m]", ex["phiMn"], rv["phiMn_y"]),
        ("Vu 1D x [t]", ex["Vu_x"], rv["Vu1_x"]), ("Vu 1D y [t]", ex["Vu_y"], rv["Vu1_y"]),
        ("φVc 1D x [t]", ex["phiVc"], rv["phiVc_x"]),
        ("φVc 1D y [t]", ex["phiVc"], rv["phiVc_y"]),
        ("Vu punz [t]", ex["Vu_p"], rv["Vu_p"]), ("bo [cm]", ex["bo"], rv["detalle"][0]["bo"] * 100),
        ("φVc punz [t]", ex["phiVp"], rv["phi_vc_p"] * rv["detalle"][0]["bo"] * rv["d"]),
    ]
    for nom, a, b in pares:
        dif = "" if a in (None, 0) or b is None else f"{(b - a) / a * 100:+.1f}"
        print(f"{nom:28s}{f(a):>10s}{f(b):>10s}{dif:>8s}")
    print(f"  barras programa (Ø{ex['db']}): X {dz.texto_armado(rv['bar_x'])} | Y {dz.texto_armado(rv['bar_y'])}"
          f" | Excel: {ex['n']} Ø{ex['db']} @ 15 cm")
    print(f"  D/C programa: flex {rv['DC_flex_x']:.3f}/{rv['DC_flex_y']:.3f}  v1D {rv['DC_v1_x']:.3f}/{rv['DC_v1_y']:.3f}"
          f"  punz {rv['DC_punz']:.3f} (vu = {rv['vu_p']:.1f} t/m², φvc = {rv['phi_vc_p']:.1f}); gobierna {rv['gobierna']}")
    print(f"  combos programa: flex {rv['combo_flex_x']}/{rv['combo_flex_y']}  v1D {rv['DC_v1_x_combo']}/{rv['DC_v1_y_combo']}  punz {rv['DC_punz_combo']}")
    print(f"  Diseño libre del programa: h = {rl['h']*100:.0f} cm, {dz.texto_armado(rl['bar_x'])} / {dz.texto_armado(rl['bar_y'])}, "
          f"estado {rl['estado']}, gobierna {rl['gobierna'][0]} D/C {rl['gobierna'][1]:.3f}; {rl['control_h']}")

# -----------------------------------------------------------------------------
# Criterio CLASE: diseño completo del programa (h y armado iterados desde H_MIN)
# con las hipótesis del Excel. P7: Ø16 elegido a mano en el Excel.
# -----------------------------------------------------------------------------
dz.CRITERIO = "CLASE"
dz.DIAMETRO_PUNTO = {7: 16}
dz.H_MIN_PUNTO = {}
dz.validar_parametros()
print("\n\n=== CRITERIO = \"CLASE\" (diseño iterado por el programa, DIAMETRO_PUNTO = {7: 16}) ===")
cab = ["h [cm]", "Ø/n", "As [cm²]", "φMn [t·m]", "Vu1D y [t]", "φVc [t]", "Vu punz [t]", "φVc punz [t]"]
print(f"{'':6s}" + "".join(f"{c:>23s}" for c in cab))
for p in CARGAS:
    ex = EXCEL[p]
    r = dz.disenar_zapata(p, df[df["Point"] == p].reset_index(drop=True))
    bo = r["detalle"][0]["bo"]
    prog = [f"{r['h'] * 100:.0f}", f"Ø{r['db']}/{r['bar_x']['n']}", f"{r['bar_x']['As']:.2f}",
            f"{r['phiMn_x']:.2f}", f"{r['Vu1_y']:.2f}", f"{r['phiVc_x']:.2f}", f"{r['Vu_p']:.2f}",
            f"{r['phi_vc_p'] * bo * r['d']:.2f}"]
    exc = [f"{CARGAS[p][1] * 100:.0f}", f"Ø{ex['db']}/{ex['n']}", f"{ex['As']:.2f}", f"{ex['phiMn']:.2f}",
           f"{ex['Vu_y']:.2f}", f"{ex['phiVc']:.2f}", f"{ex['Vu_p']:.2f}", f(ex["phiVp"])]
    print(f"P{p:<5d}" + "".join(f"{a + ' | ' + b:>23s}" for a, b in zip(exc, prog)))
print("(cada celda: Excel | Programa)")
