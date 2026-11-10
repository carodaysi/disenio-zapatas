# Genera Velecela_Daysi_Cimentaciones_Tarea01.pdf
import math
import os
import sys
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, Image, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle,
                                XPreformatted)

PROY = r"D:\Maestría UPS\Diseño de cimentaciones\Trabajos"
sys.path.insert(0, PROY)
import Diseno_Zapatas_Aisladas as dz  # noqa: E402

SALIDA = os.path.join(PROY, "Velecela_Daysi_Cimentaciones_Tarea01.pdf")
RES = os.path.join(PROY, "Resultados")

# ---------------------------------------------------------------- fuentes
F = r"C:\Windows\Fonts"
pdfmetrics.registerFont(TTFont("Arial", os.path.join(F, "arial.ttf")))
pdfmetrics.registerFont(TTFont("Arial-B", os.path.join(F, "arialbd.ttf")))
pdfmetrics.registerFont(TTFont("Arial-I", os.path.join(F, "ariali.ttf")))
pdfmetrics.registerFont(TTFont("Arial-BI", os.path.join(F, "arialbi.ttf")))
pdfmetrics.registerFont(TTFont("Consolas", os.path.join(F, "consola.ttf")))
pdfmetrics.registerFont(TTFont("Simbolos", os.path.join(F, "seguisym.ttf")))
from reportlab.pdfbase.pdfmetrics import registerFontFamily  # noqa: E402
registerFontFamily("Arial", normal="Arial", bold="Arial-B", italic="Arial-I", boldItalic="Arial-BI")

AZUL = colors.HexColor("#1F3B5C")
GRIS = colors.HexColor("#F1F3F5")
LINEA = colors.HexColor("#9AA5B1")
VERDE = colors.HexColor("#E7F4EA")
AMBAR = colors.HexColor("#FFF4D6")

st = {}
st["body"] = ParagraphStyle("body", fontName="Arial", fontSize=9.2, leading=12.4,
                            alignment=TA_JUSTIFY, spaceAfter=3.5)
st["bul"] = ParagraphStyle("bul", parent=st["body"], leftIndent=11, bulletIndent=2,
                           spaceAfter=2)
st["h1"] = ParagraphStyle("h1", fontName="Arial-B", fontSize=10.8, leading=13, textColor=AZUL,
                          spaceBefore=6, spaceAfter=3)
st["tit"] = ParagraphStyle("tit", fontName="Arial-B", fontSize=14, leading=17,
                           alignment=TA_CENTER, textColor=AZUL)
st["sub"] = ParagraphStyle("sub", fontName="Arial", fontSize=9, leading=11.5,
                           alignment=TA_CENTER)
st["cap"] = ParagraphStyle("cap", fontName="Arial-I", fontSize=8, leading=10,
                           alignment=TA_LEFT, spaceBefore=2, spaceAfter=4)
st["tc"] = ParagraphStyle("tc", fontName="Arial", fontSize=7.6, leading=10.4)
st["tcc"] = ParagraphStyle("tcc", parent=st["tc"], alignment=TA_CENTER)
st["th"] = ParagraphStyle("th", fontName="Arial-B", fontSize=7.4, leading=8.8,
                          alignment=TA_CENTER, textColor=colors.white)
st["tsm"] = ParagraphStyle("tsm", fontName="Arial", fontSize=6.6, leading=7.8)
st["thsm"] = ParagraphStyle("thsm", parent=st["th"], fontSize=6.5, leading=7.6)
st["code"] = ParagraphStyle("code", fontName="Consolas", fontSize=6.3, leading=7.45)


def P(t, s="body"):
    return Paragraph(t, st[s])


def tabla(data, widths, head_rows=1, zebra=True, extra=None, font=7.6, align="CENTER"):
    t = Table(data, colWidths=widths, repeatRows=head_rows, hAlign="CENTER")
    estilo = [
        ("FONT", (0, 0), (-1, -1), "Arial", font),
        ("BACKGROUND", (0, 0), (-1, head_rows - 1), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, head_rows - 1), colors.white),
        ("FONT", (0, 0), (-1, head_rows - 1), "Arial-B", font),
        ("ALIGN", (0, 0), (-1, -1), align),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, LINEA),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
    ]
    if zebra:
        for i in range(head_rows, len(data)):
            if (i - head_rows) % 2 == 1:
                estilo.append(("BACKGROUND", (0, i), (-1, i), GRIS))
    if extra:
        estilo += extra
    t.setStyle(TableStyle(estilo))
    return t


def n(x, d=2):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "—"
    if isinstance(x, str):
        return x
    return f"{x:.{d}f}"


def pct(a, b):
    if a in (None, 0) or b is None:
        return "—"
    v = (b - a) / a * 100
    return "0.0" if abs(v) < 0.05 else f"{v:+.1f}"


# ---------------------------------------------------------------- datos
datos = pd.read_excel(os.path.join(PROY, "InputData01.xlsx"), sheet_name="Tarea1")
res = pd.read_excel(os.path.join(RES, "Resultados_Zapatas.xlsx"), sheet_name="Resumen")
n_cumple = int((res["Estado"] == "CUMPLE").sum())
n_revisar = int(res["Estado"].astype(str).str.startswith("REVISAR").sum())

# ------------------------------------------- contraste con el Excel de clase
# Valores leídos de la hoja "Empotrado Art" (clase del 23/10/2026, 3:29:59 y 4:00:00)
EXCEL = {
    4: dict(h=0.50, Mu_x=22.37, Mu_y=24.64, n=16, As_min=17.60, As=24.63, d=42.5, phiMn=38.40,
            Vu_x=26.35, Vu_y=28.89, phiVc=30.65, Vu_p=114.78, bo=390),
    5: dict(h=0.45, Mu_x=15.81, Mu_y=18.47, n=15, As_min=14.18, As=23.09, d=37.5, phiMn=31.60,
            Vu_x=21.09, Vu_y=24.58, phiVc=25.97, Vu_p=94.62, bo=370, phiVp=165.88),
    6: dict(h=0.60, Mu_x=33.65, Mu_y=36.64, n=18, As_min=24.57, As=27.71, d=52.5, phiMn=53.67,
            Vu_x=32.01, Vu_y=34.90, phiVc=39.82, Vu_p=143.27, bo=430),
}
for p, e in EXCEL.items():   # φVc punz. del Excel: 0.75·vc1·bo·d con vc1 = 1.1√f'c = 15.94
    e.setdefault("phiVp", 0.75 * 15.94 * e["bo"] * e["d"] / 1000.0)

df = dz.leer_datos(os.path.join(PROY, "InputData01.xlsx"), "Tarea1")
CONTRASTE = {}
for p, e in EXCEL.items():
    sub = df[df["Point"] == p].reset_index(drop=True)
    dz.H_MIN_PUNTO = {p: e["h"]}
    r0 = dz.disenar_zapata(p, sub)
    rv = dz.verificar(r0["geo"], r0["combos"], r0["campos"], e["h"], 14)
    dz.H_MIN_PUNTO = {}
    rf = dz.disenar_zapata(p, sub)
    # mismas fórmulas del programa con el As y d del Excel
    B = r0["geo"]["Bx"]
    phiMn_e = dz.momento_resistente(e["As"], B, e["d"] / 100)[0]
    phiVc_e = dz.resistencia_cortante_1D(e["As"], B, e["d"] / 100)[0]
    CONTRASTE[p] = dict(rv=rv, rf=rf, B=B, phiMn_e=phiMn_e, phiVc_e=phiVc_e)

# Criterio CLASE: el programa itera h y armado con las hipótesis del Excel, usando
# los datos transcritos de la hoja de clase (P3–P7, Validacion_Excel_Clase)
EXCEL_CLASE = {3: (2.40, 45, 14, 17, 26.17, 35.83, 29.60, 100.98),
               4: (2.30, 50, 14, 16, 24.63, 38.40, 30.65, 114.78),
               5: (2.10, 45, 14, 15, 23.09, 31.60, 25.97, 94.62),
               6: (2.60, 60, 14, 18, 27.71, 53.67, 39.82, 143.27),
               7: (3.00, 65, 16, 21, 42.22, 89.13, 53.57, 204.18)}
dz.CRITERIO, dz.DIAMETRO_PUNTO = "CLASE", {7: 16}
dz.validar_parametros()
dfc = dz.leer_datos(os.path.join(PROY, "Validacion_Excel_Clase", "Datos_Excel_Clase.xlsx"), "Clase")
CLASE = {p: dz.disenar_zapata(p, dfc[dfc["Point"] == p].reset_index(drop=True)) for p in EXCEL_CLASE}
dz.CRITERIO, dz.DIAMETRO_PUNTO = "ACI", {}

# ---------------------------------------------------------------- documento
MARG = 1.6 * cm


def pie(canvas, doc):
    canvas.saveState()
    w, h = canvas._pagesize
    canvas.setFont("Arial", 7.5)
    canvas.setFillColor(colors.HexColor("#5A6573"))
    canvas.drawString(MARG, 0.95 * cm, "Velecela Daysi · Diseño de Cimentaciones · Tarea 01 · "
                                        "Zapatas aisladas ACI 318-19")
    canvas.drawRightString(w - MARG, 0.95 * cm, f"Página {doc.page}")
    canvas.setStrokeColor(LINEA)
    canvas.setLineWidth(0.4)
    canvas.line(MARG, 1.3 * cm, w - MARG, 1.3 * cm)
    canvas.restoreState()


class Doc(BaseDocTemplate):
    def __init__(self, ruta):
        super().__init__(ruta, pagesize=A4, leftMargin=MARG, rightMargin=MARG,
                         topMargin=1.4 * cm, bottomMargin=1.6 * cm,
                         title="Diseño de zapatas aisladas ACI 318-19 - Tarea 01",
                         author="Daysi Velecela", subject="Diseño de Cimentaciones")
        pw, ph = A4
        fp = Frame(MARG, 1.6 * cm, pw - 2 * MARG, ph - 3.0 * cm, id="fp",
                   leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        lw, lh = landscape(A4)
        fl = Frame(MARG, 1.6 * cm, lw - 2 * MARG, lh - 3.0 * cm, id="fl",
                   leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        self.addPageTemplates([
            PageTemplate("port", frames=[fp], onPage=pie, pagesize=A4),
            PageTemplate("land", frames=[fl], onPage=pie, pagesize=landscape(A4)),
        ])
        self.marcas = {}

    def afterFlowable(self, fl):
        m = getattr(fl, "_marca", None)
        if m:
            self.marcas[m] = self.page


def marca(fl, nombre):
    fl._marca = nombre
    return fl


S = []
ANCHO = A4[0] - 2 * MARG
ANCHO_L = landscape(A4)[0] - 2 * MARG

# ================================================================ INFORME
S.append(marca(P("UNIVERSIDAD POLITÉCNICA SALESIANA · Maestría en Construcción de Estructuras "
                 "Sismorresistentes", "sub"), "inicio"))
S.append(Spacer(1, 3))
S.append(P("Diseño automatizado de zapatas aisladas según ACI 318-19", "tit"))
S.append(Spacer(1, 2))
S.append(P("<b>Diseño de Cimentaciones · Tarea 01</b> &nbsp;·&nbsp; Estudiante: Daysi Velecela "
           "&nbsp;·&nbsp; Noviembre 2026", "sub"))
S.append(Spacer(1, 6))

S.append(P("1. Objetivo y alcance", "h1"))
S.append(P(
    "El script <font face='Consolas'>Diseno_Zapatas_Aisladas.py</font> (Python 3, numpy, pandas, "
    "matplotlib) lee las reacciones en la base exportadas de ETABS (<i>InputData01.xlsx</i>, hoja "
    "<i>Tarea1</i>: 13 apoyos × 6 combinaciones mayoradas DCON1–DCON6) y diseña cada zapata aislada "
    "con columna centrada de 60 × 50 cm. Las dimensiones en planta B<sub>x</sub>, B<sub>y</sub> son "
    "dato (provienen del dimensionamiento con cargas de servicio y presión admisible); el programa "
    "determina el peralte h y el armado, y verifica flexión, cortante unidireccional, punzonamiento, "
    "aplastamiento y anclaje. Los cortantes F<sub>X</sub>, F<sub>Y</sub> los toman las cadenas de "
    "amarre (criterio de clase) y no se trasladan como momento V·h. Se acompaña de una interfaz de "
    "escritorio y una versión web; el código está en <i>github.com/carodaysi/disenio-zapatas</i>."))

S.append(P("2. Algoritmo", "h1"))
pasos = [
    "<b>Lectura y validación.</b> Se comprueban la hoja, las columnas obligatorias (Point, Load, FZ, "
    "MX, MY, Bx, By, Bcx, Bcy), las unidades de los encabezados (Bx[m]; un Bx[cm] se rechaza), celdas "
    "vacías o con texto (se indica la fila de Excel), dimensiones que parecen estar en cm, geometría "
    "que cambia entre combinaciones, columna mayor que la zapata y combinaciones repetidas. También "
    "se validan los parámetros de la sección 0. Cualquier error detiene el cálculo con un mensaje claro.",
    "<b>Presión de contacto por combinación.</b> Zapata rígida sobre una malla de 200 × 200 celdas. "
    "Con contacto total, q = P/A ± M<sub>y</sub>·x/I<sub>y</sub> ± M<sub>x</sub>·y/I<sub>x</sub> "
    "(e<sub>x</sub> = MY/P, e<sub>y</sub> = MX/P). Si aparece tracción, el plano de presiones se "
    "resuelve por Newton-Raphson sin tracción en el suelo (equilibrio de P, M<sub>x</sub>, "
    "M<sub>y</sub>). Cada combinación queda con su estado: <i>OK</i>, <i>contacto parcial</i> "
    "(zapata en REVISAR) o <i>resultante fuera de la zapata / no converge</i> (NO CUMPLE).",
    "<b>Peso propio.</b> W = 0.9·(γ<sub>c</sub>·h + γ<sub>s</sub>·(D<sub>f</sub> − h))·B<sub>x</sub>"
    "·B<sub>y</sub> se suma a P para el contacto y se recalcula con cada h; flexión y cortante usan "
    "la presión neta q<sub>u</sub> − w, de modo que W no se cuenta dos veces.",
    "<b>Iteración del peralte.</b> h = 40, 45, 50 … 150 cm. Para cada h se prueban los diámetros "
    "Ø12, 14, 16 … y se toma el primero con separación ≥ 10 cm y anclaje dentro del voladizo.",
    "<b>Verificaciones para el h de prueba</b> (todas las combinaciones resueltas): M<sub>u</sub> en "
    "la cara, A<sub>s</sub> requerido y mínimo, barras y φM<sub>n</sub>; V<sub>u</sub> a d de la "
    "cara y φV<sub>c</sub>; V<sub>u</sub> en el perímetro a d/2 con transferencia de momento y "
    "φv<sub>c</sub>. Se calcula D/C de cada verificación y combinación; el h final es el primero con "
    "todos los D/C ≤ 1, y se registra qué verificación obligó a subirlo (<i>Controla h</i>). Si "
    "d ≥ voladizo, la zapata trabaja como hormigón en masa y el cortante no aplica.",
    "<b>Complementos y salidas.</b> Aplastamiento columna-zapata, anclaje (recto o con gancho), "
    "distribución en zapatas rectangulares y acero superior si hay levantamiento. Se generan la "
    "memoria de cálculo, un Excel (Resumen, Detalle_Combos, Datos_Entrada, Parametros) y figuras.",
]
for i, t in enumerate(pasos, 1):
    S.append(Paragraph(t, st["bul"], bulletText=f"{i}."))

S.append(P("3. Criterios ACI 318-19", "h1"))
S.append(P("f'c = 210 kg/cm², fy = 4200 kg/cm², λ = λ<sub>s</sub> = 1 (13.2.6.2), recubrimiento "
           "7.5 cm (20.5.1.3.1), γ<sub>c</sub> = 2.4 t/m³, γ<sub>s</sub> = 1.8 t/m³, D<sub>f</sub> = "
           "1.5 m. Unidades: t, m (esfuerzos en kg/cm²). Criterio por defecto "
           "<font face='Consolas'>CRITERIO = \"ACI\"</font>:"))
crit = [
    [P("<b>Verificación</b>", "th"), P("<b>Expresión / criterio</b>", "th"), P("<b>ACI 318-19</b>", "th")],
    [P("Peralte efectivo", "tc"),
     P("d = h − rec − d<sub>b</sub>/2 (capa inferior, dirección de mayor M<sub>u</sub>); "
       "d = h − rec − 1.5·d<sub>b</sub> (capa superior); d ≥ 15 cm", "tc"), P("13.3.1.2", "tcc")],
    [P("Flexión", "tc"),
     P("M<sub>u</sub> en la cara de la columna; φM<sub>n</sub> = φ·A<sub>s</sub>·f<sub>y</sub>·"
       "(d − a/2), φ = 0.90 comprobando ε<sub>t</sub>; A<sub>s,mín</sub> = 0.0018·b·h; "
       "s ≤ mín(2h, 45 cm); s libre ≥ 2.5 cm", "tc"),
     P("13.2.7.1, 21.2.2, 8.6.1.1, 8.7.2.2, 25.2.1", "tcc")],
    [P("Cortante 1D", "tc"),
     P("Sección a d de la cara; V<sub>c</sub> = 2.12·λ<sub>s</sub>·λ·ρ<sub>w</sub><super>1/3</super>"
       "·√f'c·b<sub>w</sub>·d ≤ 1.33·λ·√f'c·b<sub>w</sub>·d; φ = 0.75", "tc"),
     P("13.2.7.2, 22.5.5.1", "tcc")],
    [P("Punzonamiento", "tc"),
     P("Perímetro b<sub>o</sub> a d/2; v<sub>c</sub> = mín[1.06; 0.53(1 + 2/β); 0.27(2 + α<sub>s</sub>"
       "d/b<sub>o</sub>)]·λ<sub>s</sub>·λ·√f'c, α<sub>s</sub> = 40; v<sub>u</sub> = V<sub>u</sub>/"
       "(b<sub>o</sub>d) + γ<sub>v</sub>·M<sub>sc</sub>·c/J<sub>c</sub>", "tc"),
     P("22.6.4.1, Tabla 22.6.5.2, 8.4.4.2", "tcc")],
    [P("Aplastamiento", "tc"),
     P("φB<sub>n</sub> = 0.65·0.85·f'c·A<sub>1</sub>·√(A<sub>2</sub>/A<sub>1</sub>), "
       "√(A<sub>2</sub>/A<sub>1</sub>) ≤ 2", "tc"), P("22.8.3.2", "tcc")],
    [P("Anclaje y zapata rectangular", "tc"),
     P("ℓ<sub>d</sub> recto o ℓ<sub>dh</sub> con gancho dentro del voladizo; banda central "
       "γ<sub>s</sub> = 2/(β + 1)", "tc"), P("25.4.2.3, 25.4.3.1, 13.3.3.3", "tcc")],
]
S.append(tabla(crit, [3.0 * cm, 11.4 * cm, ANCHO - 14.4 * cm], zebra=True,
               extra=[("ALIGN", (0, 1), (1, -1), "LEFT")]))
S.append(P("Criterio de aceptación: CUMPLE si todos los D/C ≤ 1 y todas las combinaciones mantienen "
           "contacto total; REVISAR si alguna combinación pierde contacto parcialmente (la resistencia "
           "se verifica con la presión resuelta, pero B<sub>x</sub>, B<sub>y</sub> deben revisarse); "
           "NO CUMPLE si falta resistencia hasta h = 150 cm o alguna resultante sale de la zapata.",
           "cap"))

# ------------------------------------------------------- 4. contraste
S.append(P("4. Contraste con el Excel de clase", "h1"))
S.append(P(
    "Se usó la hoja <i>Empotrado Art</i> del archivo <i>Dimensionamiento.xlsx</i> desarrollado en "
    "clase (23/10/2026). Parte de las mismas reacciones DCON del modelo ETABS; las zapatas P4, P5 y P6 "
    "tienen el mismo B que en la tarea (2.30, 2.10 y 2.60 m), por lo que el contraste es directo. "
    "El programa se evaluó con el mismo h y la misma barra Ø14 del Excel; las magnitudes por "
    "combinación corresponden a la combinación que gobierna (DCON5 en P5, DCON4 en P6)."))


def fila(nombre, a5, b5, a6, b6, dec=2, dif=True):
    return [nombre, n(a5, dec), n(b5, dec), pct(a5, b5) if dif else "", n(a6, dec), n(b6, dec),
            pct(a6, b6) if dif else ""]


c5, c6 = CONTRASTE[5], CONTRASTE[6]
r5, r6 = c5["rv"], c6["rv"]
e5, e6 = EXCEL[5], EXCEL[6]
bo5, bo6 = r5["detalle"][0]["bo"], r6["detalle"][0]["bo"]
cont = [["Magnitud", "P5 Excel", "P5 Prog.", "Δ %", "P6 Excel", "P6 Prog.", "Δ %"],
        ["B × B [m] · h [cm]", "2.10 · 45", "2.10 · 45", "", "2.60 · 60", "2.60 · 60", ""],
        fila("Mu,x / Mu,y [t·m]", e5["Mu_x"], r5["Mu_x"], e6["Mu_x"], r6["Mu_x"]),
        fila("", e5["Mu_y"], r5["Mu_y"], e6["Mu_y"], r6["Mu_y"]),
        fila("d x / d y [cm]", e5["d"], r5["dx"] * 100, e6["d"], r6["dx"] * 100, 1),
        fila("", e5["d"], r5["dy"] * 100, e6["d"], r6["dy"] * 100, 1),
        fila("As,mín [cm²]", e5["As_min"], r5["As_min_x"], e6["As_min"], r6["As_min_x"]),
        ["As provisto [cm²]", f"{e5['As']:.2f} ({e5['n']}Ø14)",
         f"{r5['bar_x']['As']:.2f} ({r5['bar_x']['n']}Ø14)", pct(e5["As"], r5["bar_x"]["As"]),
         f"{e6['As']:.2f} ({e6['n']}Ø14)", f"{r6['bar_x']['As']:.2f} ({r6['bar_x']['n']}Ø14)",
         pct(e6["As"], r6["bar_x"]["As"])],
        fila("φMn,y [t·m]", e5["phiMn"], r5["phiMn_y"], e6["phiMn"], r6["phiMn_y"]),
        fila("Vu 1D x / y [t]", e5["Vu_x"], r5["Vu1_x"], e6["Vu_x"], r6["Vu1_x"]),
        fila("", e5["Vu_y"], r5["Vu1_y"], e6["Vu_y"], r6["Vu1_y"]),
        fila("φVc 1D y [t]", e5["phiVc"], r5["phiVc_y"], e6["phiVc"], r6["phiVc_y"]),
        fila("D/C cortante 1D y", e5["Vu_y"] / e5["phiVc"], r5["DC_v1_y"],
             e6["Vu_y"] / e6["phiVc"], r6["DC_v1_y"], 3),
        fila("Vu punzonamiento [t]", e5["Vu_p"], r5["Vu_p"], e6["Vu_p"], r6["Vu_p"]),
        fila("bo [cm]", e5["bo"], bo5 * 100, e6["bo"], bo6 * 100, 1),
        fila("φVc punzonamiento [t]", e5["phiVp"], r5["phi_vc_p"] * bo5 * r5["d"],
             e6["phiVp"], r6["phi_vc_p"] * bo6 * r6["d"]),
        fila("D/C punzonamiento", e5["Vu_p"] / e5["phiVp"], r5["DC_punz"],
             e6["Vu_p"] / e6["phiVp"], r6["DC_punz"], 3),
        ["h final de diseño [cm]", "45", f"{c5['rf']['h'] * 100:.0f}", "",
         "60", f"{c6['rf']['h'] * 100:.0f}", ""]]
t_cont = tabla(cont, [3.6 * cm, 2.55 * cm, 2.55 * cm, 1.35 * cm, 2.55 * cm, 2.55 * cm, 1.35 * cm],
               font=7.4, extra=[("ALIGN", (0, 1), (0, -1), "LEFT"),
                                ("BACKGROUND", (0, -1), (-1, -1), VERDE),
                                ("FONT", (0, -1), (-1, -1), "Arial-B", 7.4)])
S.append(KeepTogether([t_cont, P("φVc de punzonamiento del Excel: 0.75·v<sub>c1</sub>·b<sub>o</sub>·d con v<sub>c1</sub> = "
           "1.1√f'c = 15.94 kg/cm² (para P5 coincide con el valor de la hoja, 165.88 t). En P4 "
           f"(h = 50 cm) el resultado es análogo: Mu 22.37/24.64 t·m en ambos y el programa sube a "
           f"h = {CONTRASTE[4]['rf']['h'] * 100:.0f} cm por cortante 1D.", "cap")]))

S.append(P("<b>Justificación de las discrepancias.</b>"))
just = [
    "<b>Solicitaciones (M<sub>u</sub>, V<sub>u</sub> de punzonamiento): coinciden (Δ ≤ 0.7 %).</b> "
    "Ambos usan la presión lineal P/A ± M·c/I (contacto total en estas combinaciones); el Excel la "
    "evalúa sobre los ejes y el programa la integra en toda la planta, lo que con momentos pequeños "
    "da el mismo resultado.",
    "<b>Peralte efectivo.</b> El Excel toma d = H − 7.5 cm en las dos direcciones; el programa descuenta "
    "medio diámetro y, en la capa superior, uno y medio (d<sub>x</sub> 2.1 cm y d<sub>y</sub> 0.7 cm "
    "menores). Con d menor la sección crítica se acerca a la columna y V<sub>u</sub> 1D crece entre "
    "1 % y 6 %; "
    "φV<sub>c</sub> baja.",
    "<b>Acero mínimo y provisto.</b> El Excel usa A<sub>s,mín</sub> = 0.0018·B·d; el programa 0.0018·B·h, "
    "como pide 8.6.1.1 (sobre A<sub>g</sub>). El Excel coloca barras a 15 cm fijos, con más acero del "
    "requerido (P5: 23.09 cm² frente a A<sub>s,req</sub> ≈ 13.6 cm²); el programa coloca "
    "máx(A<sub>s,req</sub>, A<sub>s,mín</sub>). Como V<sub>c</sub> crece con ρ<sub>w</sub>"
    "<super>1/3</super>, el Excel obtiene mayor φV<sub>c</sub>. Con el A<sub>s</sub> y d del Excel, las funciones del "
    f"programa dan exactamente φMn = {c5['phiMn_e']:.2f} t·m y φVc = {c5['phiVc_e']:.2f} t en P5 "
    f"(Excel: 31.60 y 25.97): las fórmulas son las mismas.",
    "<b>Punzonamiento.</b> El Excel usa v<sub>c</sub> = 1.1√f'c; el ACI 318-19 en kg/cm² da "
    "0.33√f'c [MPa] = 1.054√f'c ≈ 1.06√f'c, valor del programa. El programa suma además el término "
    "γ<sub>v</sub>·M<sub>sc</sub>·c/J<sub>c</sub>, por lo que su D/C es algo mayor (más conservador).",
    "<b>Consecuencia.</b> En P6 ambos llegan a h = 60 cm. En P5 el programa, con su armado mínimo, "
    "obtiene D/C de cortante 1D = "
    f"{r5['DC_v1_y']:.3f} con h = 45 cm y sube a {c5['rf']['h'] * 100:.0f} cm; el Excel cumple con "
    "45 cm porque el acero adicional eleva V<sub>c</sub>. Es una decisión de diseño (más peralte o "
    "más acero), no un error de cálculo.",
]
for t in just:
    S.append(Paragraph(t, st["bul"], bulletText="•"))

S.append(P(
    "<b>Reproducción exacta del ejemplo de clase.</b> Para validar el algoritmo, el parámetro "
    "<font face='Consolas'>CRITERIO = \"CLASE\"</font> adopta las hipótesis del Excel (d = h − 7.5 cm, "
    "A<sub>s,mín</sub> = 0.0018·b·d, barras a 15 cm, v<sub>c1</sub> = 1.1√f'c, sin transferencia de "
    "momento). Con él, el programa itera h y armado desde 40 cm y obtiene los valores de la hoja en "
    "las cinco zapatas (datos de la hoja de clase; en P7 el Excel usa Ø16, elegido a mano, que se "
    "impone con <font face='Consolas'>DIAMETRO_PUNTO = {7: 16}</font>). El diseño entregado (Anexo B) "
    "usa el criterio ACI."))
filasC = [["Zapata", "h [cm]\nExcel | Prog.", "Barras\nExcel | Prog.", "As [cm²]\nExcel | Prog.",
           "φMn [t·m]\nExcel | Prog.", "φVc 1D [t]\nExcel | Prog.", "Vu punz. [t]\nExcel | Prog."]]
for p, (B, h, db, nb, As, pMn, pVc, Vup) in EXCEL_CLASE.items():
    r = CLASE[p]
    filasC.append([f"P{p} ({B:.2f} m)", f"{h} | {r['h'] * 100:.0f}",
                   f"{nb}Ø{db} | {r['bar_x']['n']}Ø{r['db']}", f"{As:.2f} | {r['bar_x']['As']:.2f}",
                   f"{pMn:.2f} | {r['phiMn_x']:.2f}", f"{pVc:.2f} | {r['phiVc_x']:.2f}",
                   f"{Vup:.2f} | {r['Vu_p']:.2f}"])
S.append(tabla(filasC, [2.6 * cm, 2.3 * cm, 2.75 * cm, 2.6 * cm, 2.6 * cm, 2.6 * cm, 2.65 * cm],
               font=7.4))

S.append(P("5. Resultados y control de combinaciones", "h1"))
hmin, hmax = res.loc[2:, "h [cm]"].min(), res.loc[2:, "h [cm]"].max()
dc1d = pd.to_numeric(res["D/C cortante 1D Y"], errors="coerce")
dcp = pd.to_numeric(res["D/C punzonamiento"], errors="coerce")
dcf = res[["D/C flexión X", "D/C flexión Y"]].max(axis=1)
S.append(P(
    f"De las 13 zapatas, <b>{n_cumple} cumplen</b> y <b>{n_revisar} quedan en REVISAR</b>. P3–P13 "
    f"requieren h = {hmin:.0f}–{hmax:.0f} cm; en todas gobierna el cortante 1D en Y (D/C "
    f"{dc1d.min():.2f}–{dc1d.max():.3f}), con punzonamiento D/C ≤ {dcp.max():.2f} y flexión D/C ≤ "
    f"{dcf.max():.2f}. Las combinaciones sísmicas DCON3–DCON6 controlan todas las verificaciones. "
    "P1 y P2 (1.10 × 1.10 m) quedan en h = 40 cm con d ≥ voladizo (hormigón en masa, sin cortante "
    "ni punzonamiento) pero pierden contacto parcial con el suelo en DCON5 (99.0 %) y DCON6 (76.6 %): "
    "su resistencia está verificada con la presión sin tracción, y se recomienda revisar B por "
    "estabilidad y presión admisible. Ninguna combinación saca la resultante de la zapata. El detalle "
    "por zapata está en el Anexo B; los datos de entrada en el Anexo A y el script en el Anexo C."))
S.append(marca(Spacer(1, 0.1), "fin_informe"))

# ================================================================ ANEXO A
S.append(PageBreak())
S.append(marca(P("Anexo A. Datos de entrada (InputData01.xlsx, hoja Tarea1)", "h1"), "anexoA"))
S.append(P("Reacciones en la base para las combinaciones de diseño mayoradas (t, t·m) y geometría "
           "(m). Story = BASE en todas las filas. MZ no interviene en el diseño de la zapata.", "cap"))
cab = ["Point", "Load", "FX [t]", "FY [t]", "FZ [t]", "MX [t·m]", "MY [t·m]", "MZ [t·m]",
       "Bx [m]", "By [m]", "Bcx [m]", "Bcy [m]"]
filasA = [cab]
for _, f in datos.iterrows():
    filasA.append([int(f["Point"]), f["Load"], f"{f['FX']:.2f}", f"{f['FY']:.2f}", f"{f['FZ']:.2f}",
                   f"{f['MX']:.3f}", f"{f['MY']:.3f}", f"{f['MZ']:.3f}", f"{f['Bx[m]']:.2f}",
                   f"{f['By[m]']:.2f}", f"{f['Bcx[m]']:.2f}", f"{f['Bcy[m]']:.2f}"])
extraA = []
for i in range(1, len(filasA)):
    p = filasA[i][0]
    if p % 2 == 0:
        extraA.append(("BACKGROUND", (0, i), (-1, i), GRIS))
    if (i - 1) % 6 == 0 and i > 1:
        extraA.append(("LINEABOVE", (0, i), (-1, i), 1.0, AZUL))
anchoA = [1.2 * cm, 1.45 * cm] + [1.55 * cm] * 6 + [1.35 * cm] * 4
tA = Table(filasA, colWidths=anchoA, repeatRows=1, hAlign="CENTER")
tA.setStyle(TableStyle([
    ("FONT", (0, 0), (-1, -1), "Arial", 7.2), ("FONT", (0, 0), (-1, 0), "Arial-B", 7.2),
    ("BACKGROUND", (0, 0), (-1, 0), AZUL), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("ALIGN", (2, 1), (7, -1), "RIGHT"),
    ("GRID", (0, 0), (-1, -1), 0.3, LINEA), ("TOPPADDING", (0, 0), (-1, -1), 1.1),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 1.1)] + extraA))
S.append(tA)

# ================================================================ ANEXO B
S.append(NextPageTemplate("land"))
S.append(PageBreak())
S.append(marca(P("Anexo B. Resultados por zapata", "h1"), "anexoB"))
S.append(P(f"f'c = 210 kg/cm² y fy = 4200 kg/cm² en todas las zapatas; columna 60 × 50 cm centrada. "
           "As en cm², Mu y φMn en t·m. X: barras paralelas a X (voladizo en X); Y: barras paralelas "
           "a Y. Armado: n Ø (mm) @ separación (cm), igual en malla inferior.", "cap"))


def arm(t):
    return str(t).replace(" mm @ ", "@").replace(" cm", "").replace(" Ø", "Ø")


hb1 = ["Punto", "Bx × By\n[m]", "f'c\n[kg/cm²]", "fy\n[kg/cm²]", "h\n[cm]", "dx\n[cm]", "dy\n[cm]",
       "Mu x", "As req\nX", "As mín\nX", "As prov\nX", "Armado X", "φMn x", "D/C\nflex X",
       "Mu y", "As req\nY", "As mín\nY", "As prov\nY", "Armado Y", "φMn y", "D/C\nflex Y"]
B1 = [hb1]
for _, f in res.iterrows():
    B1.append([int(f["Punto"]), f"{f['Bx [m]']:.2f} × {f['By [m]']:.2f}", f"{f[chr(102)+chr(39)+'c [kg/cm²]']:.0f}",
               f"{f['fy [kg/cm²]']:.0f}", f"{f['h [cm]']:.0f}", f"{f['dx [cm]']:.1f}",
               f"{f['dy [cm]']:.1f}", f"{f['Mu x [t·m]']:.2f}", f"{f['As req x [cm²]']:.2f}",
               f"{f['As min x [cm²]']:.2f}", f"{f['As prov x [cm²]']:.2f}", arm(f["Armado X"]),
               f"{f['φMn x [t·m]']:.2f}", f"{f['D/C flexión X']:.3f}", f"{f['Mu y [t·m]']:.2f}",
               f"{f['As req y [cm²]']:.2f}", f"{f['As min y [cm²]']:.2f}",
               f"{f['As prov y [cm²]']:.2f}", arm(f["Armado Y"]), f"{f['φMn y [t·m]']:.2f}",
               f"{f['D/C flexión Y']:.3f}"])
wB1 = [1.0, 1.75, 1.25, 1.25, 0.9, 0.95, 0.95, 1.05, 1.15, 1.15, 1.2, 2.05, 1.1, 1.05,
       1.05, 1.15, 1.15, 1.2, 2.05, 1.1, 1.05]
esc = ANCHO_L / (sum(wB1) * cm)
S.append(P("<b>B.1 Geometría, materiales y flexión</b>", "body"))
S.append(tabla(B1, [w * cm * esc for w in wB1], font=7.0))
S.append(Spacer(1, 8))


def comb(c):
    return "—" if str(c) == "No aplica" else str(c)


def dc(x):
    try:
        return f"{float(x):.3f}"
    except (TypeError, ValueError):
        return "N/A"


def num(x, d=2):
    try:
        return f"{float(x):.{d}f}"
    except (TypeError, ValueError):
        return "N/A"


hb2 = ["Punto", "Vu 1D X\n[t]", "φVc X\n[t]", "D/C\n1D X", "Comb.", "Vu 1D Y\n[t]", "φVc Y\n[t]",
       "D/C\n1D Y", "Comb.", "Vu punz.\n[t]", "φVc punz.\n[t]", "D/C\npunz.", "Comb.",
       "Verificación gobernante", "Comb.\ncontrol.", "Estado"]
B2 = [hb2]
ext2 = []
for i, (_, f) in enumerate(res.iterrows(), 1):
    estado = "CUMPLE" if f["Estado"] == "CUMPLE" else str(f["Estado"]).split(":")[0] + " (1)"
    B2.append([int(f["Punto"]), num(f["Vu 1D x [t]"]), num(f["φVc x [t]"]), dc(f["D/C cortante 1D X"]),
               comb(f["Combo cortante 1D X"]), num(f["Vu 1D y [t]"]), num(f["φVc y [t]"]),
               dc(f["D/C cortante 1D Y"]), comb(f["Combo cortante 1D Y"]), num(f["Vu punz [t]"]),
               num(f["φVc punz [t]"]), dc(f["D/C punzonamiento"]), comb(f["Combo punzonamiento"]),
               str(f["Verificación gobernante"]), str(f["Combinación controlante"]), estado])
    ext2.append(("BACKGROUND", (-1, i), (-1, i), VERDE if estado == "CUMPLE" else AMBAR))
wB2 = [1.0, 1.4, 1.4, 1.15, 1.25, 1.4, 1.4, 1.15, 1.25, 1.5, 1.5, 1.15, 1.25, 4.6, 1.4, 2.2]
esc = ANCHO_L / (sum(wB2) * cm)
S.append(P("<b>B.2 Cortante unidireccional, punzonamiento, combinación controlante y estado</b>", "body"))
S.append(tabla(B2, [w * cm * esc for w in wB2], font=7.0, extra=ext2))
S.append(P("N/A: no aplica, d ≥ voladizo (hormigón en masa; la sección crítica queda fuera de la "
           "zapata). (1) P1 y P2: pérdida de contacto parcial en DCON5 (99.0 % en contacto) y DCON6 "
           "(76.6 %); resistencia verificada con la presión sin tracción, revisar B por estabilidad "
           "y presión admisible.", "cap"))

S.append(PageBreak())
S.append(P("<b>B.3 Control del peralte, presiones y verificaciones complementarias</b>", "body"))
hb3 = ["Punto", "qu máx\n[t/m²]", "Contacto\nmín [%]", "W zapata +\nrelleno [t]", "Controla h",
       "Aplast.", "Anclaje", "Acero superior"]
B3 = [[P(t.replace("\n", "<br/>"), "thsm") for t in hb3]]
for _, f in res.iterrows():
    B3.append([P(str(int(f["Punto"])), "tsm"), P(f"{f['qu max [t/m²]']:.2f}", "tsm"),
               P(f"{f['Contacto mín [%]']:.1f}", "tsm"), P(f"{f['W zapata+relleno [t]']:.2f}", "tsm"),
               P(escape(str(f["Controla h"])), "tsm"), P(str(f["Aplastamiento"]), "tsm"),
               P(str(f["Anclaje"]), "tsm"), P(escape(str(f["Acero superior"])), "tsm")])
wB3 = [1.1, 1.5, 1.6, 1.8, 10.5, 1.7, 2.6, 6.0]
esc = ANCHO_L / (sum(wB3) * cm)
S.append(tabla(B3, [w * cm * esc for w in wB3], extra=[("ALIGN", (0, 1), (-1, -1), "LEFT")]))
S.append(Spacer(1, 8))
img = Image(os.path.join(RES, "Resumen_Demanda_Capacidad.png"))
iw = ANCHO_L * 0.70
img.drawWidth, img.drawHeight = iw, iw * 750 / 1950
S.append(img)
S.append(P("Figura B.1. Relación demanda/capacidad de cada verificación por zapata (salida del "
           "programa).", "cap"))

S.append(PageBreak())
img = Image(os.path.join(RES, "Zapata_P6.png"))
iw = ANCHO_L * 0.92
img.drawWidth, img.drawHeight = iw, iw * 900 / 1950
S.append(img)
S.append(P("Figura B.2. Ejemplo de salida gráfica: zapata P6 (planta con armado, presiones de la "
           "combinación de punzonamiento y secciones críticas).", "cap"))

# ================================================================ ANEXO C
S.append(NextPageTemplate("port"))
S.append(PageBreak())
S.append(marca(P("Anexo C. Script: Diseno_Zapatas_Aisladas.py", "h1"), "anexoC"))
S.append(P("Código completo del programa (Python 3.10; numpy, pandas, matplotlib, openpyxl). "
           "Ejecución: <font face='Consolas'>python Diseno_Zapatas_Aisladas.py</font>, con "
           "InputData01.xlsx en la misma carpeta. Las salidas se escriben en la carpeta Resultados.",
           "cap"))
codigo = open(os.path.join(PROY, "Diseno_Zapatas_Aisladas.py"), encoding="utf-8").read().splitlines()
BLOQUE = 60
for k in range(0, len(codigo), BLOQUE):
    lineas = []
    for j, l in enumerate(codigo[k:k + BLOQUE], k + 1):
        t = escape(l).replace("✓", "<font name='Simbolos'>✓</font>")
        lineas.append(f"<font color='#8A94A0'>{j:4d}</font>  {t}")
    S.append(XPreformatted("\n".join(lineas), st["code"]))

doc = Doc(SALIDA)
doc.build(S)
print("PDF:", SALIDA)
print("páginas:", doc.page, "marcas:", doc.marcas)
