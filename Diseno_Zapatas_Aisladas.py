# -*- coding: utf-8 -*-
"""
===============================================================================
 DISEÑO ESTRUCTURAL DE ZAPATAS AISLADAS DE HORMIGÓN ARMADO  —  ACI 318-19
===============================================================================
 Maestría en Construcción de Estructuras Sismorresistentes — UPS
 Asignatura : Diseño de Cimentaciones
 Tema       : Diseño de plintos aislados: flexión, cortante unidireccional y
              cortante bidireccional (punzonamiento).

 Entrada  : InputData01.xlsx, hoja "Tarea1". Reacciones en la base de las
            columnas (ETABS) para las combinaciones de diseño DCON1..DCON6,
            dimensiones de la zapata (Bx, By) y de la columna (Bcx, Bcy).
 Unidades : fuerzas [tonf], longitudes [m], presiones [tonf/m²],
            materiales [kg/cm²] (ecuaciones mks del ACI 318-19).
            1 kg/cm² = 10 tonf/m².

 Procedimiento para cada zapata:
   1. Presión de contacto última qu(x,y) para cada combinación, sin admitir
      tracciones en el suelo (contacto parcial resuelto numéricamente).
   2. Se propone un peralte h (desde H_MIN con incrementos DH) y un diámetro
      de barra.
   3. Flexión en la cara de la columna (ACI 13.2.7.1, 8.4.2.1) -> As, As,min.
   4. Cortante unidireccional a "d" de la cara (ACI 8.4.3.2, 22.5.5.1 c).
   5. Cortante bidireccional a "d/2" de la cara (ACI 22.6.4, 22.6.5.2, 8.4.4.2).
   6. Si alguna verificación no cumple se aumenta h y se repite.
   7. Verificaciones complementarias: anclaje (25.4), aplastamiento (22.8),
      zapatas rectangulares (13.3.3.3) y acero superior (18.13.2.5).

 Salidas  : carpeta "Resultados" con
            - Resultados_Zapatas.xlsx (resumen, detalle por combinación, parámetros)
            - Memoria_Calculo.txt     (memoria paso a paso de cada zapata)
            - Zapata_P#.png           (planta con presiones y secciones críticas)
            - Resumen_Demanda_Capacidad.png
===============================================================================
"""

import os
import math
import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.patches import Rectangle

# =============================================================================
# 0. DATOS DE ENTRADA Y PARÁMETROS DE DISEÑO (modificables por el usuario)
# =============================================================================
CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(CARPETA, "InputData01.xlsx")
HOJA_EXCEL = "Tarea1"
CARPETA_SALIDA = os.path.join(CARPETA, "Resultados")

# --- Materiales -------------------------------------------------------------
FC = 210.0          # f'c del hormigón [kg/cm²]
FY = 4200.0         # fy del acero de refuerzo [kg/cm²]
LAMBDA = 1.0        # hormigón de peso normal (ACI 19.2.4)
LAMBDA_S = 1.0      # factor de tamaño; se permite despreciar en zapatas (ACI 13.2.6.2)
ALPHA_S = 40        # columna interior para punzonamiento (ACI 22.6.5.3)

# --- Factores de reducción de resistencia (ACI 21.2.1) ----------------------
PHI_FLEXION = 0.90  # sección controlada por tracción
PHI_CORTE = 0.75
PHI_APLAST = 0.65

# --- Geometría y detallado --------------------------------------------------
REC = 0.075         # recubrimiento, hormigón contra el suelo [m] (ACI 20.5.1.3.1)
D_MIN = 0.15        # peralte efectivo mínimo [m] (ACI 13.3.1.2: 150 mm)
H_MIN = 0.30        # peralte inicial de prueba [m] (criterio práctico)
H_MAX = 1.50        # peralte máximo a probar [m]
DH = 0.05           # incremento del peralte [m]
# Peralte inicial adoptado para zapatas específicas {Point: h [m]}. P1 y P2: con
# h = 40 cm, d ≥ voladizo y trabajan como hormigón en masa (sin cortante 1D).
H_MIN_PUNTO = {1: 0.40, 2: 0.40}
RHO_MIN = 0.0018    # cuantía mínima As,min = 0.0018·b·h (ACI 8.6.1.1, fy = 4200)
S_MIN = 0.10        # separación mínima práctica entre barras [m]
S_LIBRE_MIN = 0.025 # espaciamiento libre mínimo entre barras [m] (ACI 25.2.1)
# Diámetros en orden de preferencia [mm]. El Ø10 queda al final solo como
# recurso para anclar barras en voladizos muy cortos.
DIAMETROS = [12, 14, 16, 18, 20, 22, 25, 10]

# --- Suelo y relleno (solo para el acero superior por levantamiento) --------
GAMMA_C = 2.4       # peso específico del hormigón [tonf/m³]
GAMMA_S = 1.8       # peso específico del relleno [tonf/m³]
DF = 1.50           # profundidad de desplante [m]

# --- Opciones de cálculo ----------------------------------------------------
# False: los cortantes FX, FY de la base los toman las cadenas / vigas de
#        amarre (criterio visto en clase), no generan momento en la zapata.
# True : se trasladan a la base de la zapata como momento adicional V·h.
TRASLADAR_CORTANTE = False
N_MALLA = 200       # divisiones por lado de la malla de integración de presiones


# =============================================================================
# 1. FUNCIONES AUXILIARES
# =============================================================================
MEMORIA = []


def log(texto=""):
    """Imprime en consola y guarda la línea para la memoria de cálculo."""
    print(texto)
    MEMORIA.append(texto)


def area_barra(db_mm):
    """Área de una barra [cm²] a partir de su diámetro [mm]."""
    return math.pi * (db_mm / 10.0) ** 2 / 4.0


def beta_1(fc):
    """Factor β1 del bloque de compresiones (ACI 22.2.2.4.3), f'c en kg/cm²."""
    if fc <= 280.0:
        return 0.85
    return max(0.65, 0.85 - 0.05 * (fc - 280.0) / 70.0)


def leer_datos(ruta, hoja):
    """Lee el Excel de entrada y normaliza los nombres de las columnas."""
    df = pd.read_excel(ruta, sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={"Bx[m]": "Bx", "By[m]": "By",
                            "Bcx[m]": "Bcx", "Bcy[m]": "Bcy"})
    return df


# =============================================================================
# 2. PRESIÓN DE CONTACTO ÚLTIMA (sin tracciones en el suelo)
# =============================================================================
def presion_contacto(P, Mx, My, Bx, By, n=N_MALLA):
    """
    Distribución de presión qu(x,y) [tonf/m²] bajo la zapata.

    Se supone zapata rígida: la presión es un plano q = q0 + kx·x + ky·y.
    Si el plano elástico (P/A ± Mx·y/Ix ± My·x/Iy) tiene zonas negativas, se
    eliminan las tracciones y se busca por Newton-Raphson el plano que, con
    q = max(plano, 0), cumple el equilibrio:
        ΣF:  ∫q dA = P      ΣMy: ∫q·x dA = My      ΣMx: ∫q·y dA = Mx
    La zapata es simétrica, por lo que el signo de los momentos solo define
    hacia qué borde se cargan las presiones; las verificaciones se hacen a
    ambos lados de la columna.
    """
    x = (np.arange(n) + 0.5) * Bx / n - Bx / 2.0
    y = (np.arange(n) + 0.5) * By / n - By / 2.0
    X, Y = np.meshgrid(x, y)
    dA = (Bx / n) * (By / n)
    campo = dict(X=X, Y=Y, dA=dA, du=Bx / n, dv=By / n, P=P, Mx=Mx, My=My,
                 q=np.zeros_like(X), estado="OK", contacto=0.0, q_max=0.0,
                 q_min_elastico=np.nan, ex=np.nan, ey=np.nan)
    if P <= 0:
        campo["estado"] = "TRACCION"
        return campo

    ex, ey = My / P, Mx / P
    campo["ex"], campo["ey"] = ex, ey
    A = Bx * By
    Iy = By * Bx ** 3 / 12.0
    Ix = Bx * By ** 3 / 12.0
    campo["q_min_elastico"] = P / A - abs(Mx) * (By / 2) / Ix - abs(My) * (Bx / 2) / Iy
    if abs(ex) >= Bx / 2.0 or abs(ey) >= By / 2.0:
        campo["estado"] = "VOLCAMIENTO"      # resultante fuera de la zapata
        return campo

    objetivo = np.array([P, P * ex, P * ey])

    def residuo(p):
        plano = p[0] + p[1] * X + p[2] * Y
        q = np.maximum(plano, 0.0)
        R = np.array([q.sum(), (q * X).sum(), (q * Y).sum()]) * dA - objetivo
        return q, plano > 0.0, R

    p = np.array([P / A, P * ex / Iy, P * ey / Ix])   # solución elástica inicial
    q, activo, R = residuo(p)
    for _ in range(100):
        if np.linalg.norm(R) <= 1e-9 * P:
            break
        Xa, Ya = X[activo], Y[activo]
        J = np.array([[Xa.size, Xa.sum(), Ya.sum()],
                      [Xa.sum(), (Xa * Xa).sum(), (Xa * Ya).sum()],
                      [Ya.sum(), (Xa * Ya).sum(), (Ya * Ya).sum()]]) * dA
        dp = np.linalg.lstsq(J, -R, rcond=None)[0]
        t = 1.0
        while True:                                   # búsqueda lineal
            q_n, act_n, R_n = residuo(p + t * dp)
            if np.linalg.norm(R_n) < np.linalg.norm(R) or t < 1e-4:
                break
            t *= 0.5
        p = p + t * dp
        q, activo, R = q_n, act_n, R_n

    if np.linalg.norm(R) > 1e-6 * P:
        campo["estado"] = "NO CONVERGE"
    campo["q"] = q
    campo["plano"] = p
    campo["contacto"] = 100.0 * activo.mean()
    # presión máxima exacta: en la esquina más cargada
    campo["q_max"] = max(p[0] + abs(p[1]) * Bx / 2 + abs(p[2]) * By / 2, 0.0)
    return campo


# =============================================================================
# 3. SOLICITACIONES EN LAS SECCIONES CRÍTICAS
# =============================================================================
def _fraccion_mas_alla(U, a, du, lado):
    """Fracción de cada celda (coordenada U, ancho du) situada más allá de
    la sección U·lado = a, y brazo de palanca de esa parte respecto a la sección."""
    w = np.clip((lado * U + du / 2.0 - a) / du, 0.0, 1.0)
    brazo = np.where(w >= 1.0, lado * U - a, w * du / 2.0)
    return w, brazo


def momentos_cara(campo, cx, cy):
    """Momento último en la cara de la columna (ACI 13.2.7.1, Tabla 13.2.7.1).
    Mu_x: momento que resisten las barras paralelas a X (voladizo en X, ancho By).
    Mu_y: momento que resisten las barras paralelas a Y (voladizo en Y, ancho Bx)."""
    X, Y, q, dA = campo["X"], campo["Y"], campo["q"], campo["dA"]
    Mu_x = Mu_y = 0.0
    for lado in (+1, -1):
        w, b = _fraccion_mas_alla(X, cx / 2.0, campo["du"], lado)
        Mu_x = max(Mu_x, float((q * w * b).sum() * dA))
        w, b = _fraccion_mas_alla(Y, cy / 2.0, campo["dv"], lado)
        Mu_y = max(Mu_y, float((q * w * b).sum() * dA))
    return Mu_x, Mu_y


def cortante_1D(campo, cx, cy, dx, dy):
    """Cortante unidireccional a una distancia d de la cara (ACI 8.4.3.2):
    resultante de la presión del suelo fuera de la sección crítica."""
    X, Y, q, dA = campo["X"], campo["Y"], campo["q"], campo["dA"]
    Vu_x = Vu_y = 0.0
    for lado in (+1, -1):
        w, _ = _fraccion_mas_alla(X, cx / 2.0 + dx, campo["du"], lado)
        Vu_x = max(Vu_x, float((q * w).sum() * dA))
        w, _ = _fraccion_mas_alla(Y, cy / 2.0 + dy, campo["dv"], lado)
        Vu_y = max(Vu_y, float((q * w).sum() * dA))
    return Vu_x, Vu_y


def punzonamiento(campo, cx, cy, d):
    """Cortante bidireccional en el perímetro crítico a d/2 de las caras
    (ACI 22.6.4.1) con transferencia de momento por excentricidad del
    cortante (ACI 8.4.4.2). Devuelve el esfuerzo máximo vu [tonf/m²]."""
    X, Y, q, dA = campo["X"], campo["Y"], campo["q"], campo["dA"]
    b1, b2 = cx + d, cy + d                  # lados del perímetro crítico
    bo = 2.0 * (b1 + b2)
    wx = np.clip((b1 / 2.0 + campo["du"] / 2.0 - np.abs(X)) / campo["du"], 0, 1)
    wy = np.clip((b2 / 2.0 + campo["dv"] / 2.0 - np.abs(Y)) / campo["dv"], 0, 1)
    dentro = wx * wy
    R_dentro = float((q * dentro).sum() * dA)
    Vu = campo["P"] - R_dentro               # reacción del suelo fuera del perímetro
    # Momento desbalanceado que debe transferir la sección crítica
    Msc_y = abs(float((q * X * (1 - dentro)).sum() * dA))   # alrededor de Y
    Msc_x = abs(float((q * Y * (1 - dentro)).sum() * dA))   # alrededor de X
    # Fracciones de momento transferidas por excentricidad del cortante
    gv_y = 1.0 - 1.0 / (1.0 + (2.0 / 3.0) * math.sqrt(b1 / b2))
    gv_x = 1.0 - 1.0 / (1.0 + (2.0 / 3.0) * math.sqrt(b2 / b1))
    # Propiedad Jc de la sección crítica (columna interior, R8.4.4.2.3)
    Jc_y = d * b1 ** 3 / 6.0 + b1 * d ** 3 / 6.0 + d * b2 * b1 ** 2 / 2.0
    Jc_x = d * b2 ** 3 / 6.0 + b2 * d ** 3 / 6.0 + d * b1 * b2 ** 2 / 2.0
    vu = (Vu / (bo * d)
          + gv_y * Msc_y * (b1 / 2.0) / Jc_y
          + gv_x * Msc_x * (b2 / 2.0) / Jc_x)
    return dict(Vu=Vu, vu=vu, bo=bo, b1=b1, b2=b2, Msc_x=Msc_x, Msc_y=Msc_y,
                gv_x=gv_x, gv_y=gv_y)


# =============================================================================
# 4. RESISTENCIAS DE DISEÑO (ACI 318-19, unidades mks)
# =============================================================================
def acero_flexion(Mu, b, d):
    """As requerido [cm²] para Mu [tonf·m] en una sección b x d [m].
    Mu = φ·As·fy·(d − a/2), a = As·fy/(0.85·f'c·b). None si no alcanza."""
    Mu_kgcm = Mu * 1e5
    b_cm, d_cm = b * 100.0, d * 100.0
    Rn = Mu_kgcm / (PHI_FLEXION * b_cm * d_cm ** 2)
    disc = 1.0 - 2.0 * Rn / (0.85 * FC)
    if disc < 0:
        return None
    rho = 0.85 * FC / FY * (1.0 - math.sqrt(disc))
    return rho * b_cm * d_cm


def momento_resistente(As, b, d):
    """φMn [tonf·m] de una sección rectangular con As [cm²]. φ según εt (21.2.2)."""
    b_cm, d_cm = b * 100.0, d * 100.0
    a = As * FY / (0.85 * FC * b_cm)
    c = a / beta_1(FC)
    eps_t = 0.003 * (d_cm - c) / c
    eps_y = FY / 2.0e6
    if eps_t >= eps_y + 0.003:
        phi = 0.90
    elif eps_t <= eps_y:
        phi = 0.65
    else:
        phi = 0.65 + 0.25 * (eps_t - eps_y) / 0.003
    return phi * As * FY * (d_cm - a / 2.0) / 1e5, a, eps_t, phi


def resistencia_cortante_1D(As, b, d):
    """φVc [tonf] para cortante unidireccional sin refuerzo transversal,
    ACI 22.5.5.1(c): Vc = 2.12·λs·λ·(ρw)^(1/3)·√f'c·bw·d (Nu = 0),
    con el límite Vc ≤ 1.33·λ·√f'c·bw·d (22.5.5.1.1)."""
    rho_w = As / (b * 100.0 * d * 100.0)
    vc = 2.12 * LAMBDA_S * LAMBDA * rho_w ** (1.0 / 3.0) * math.sqrt(FC)
    vc = min(vc, 1.33 * LAMBDA * math.sqrt(FC))           # kg/cm²
    return PHI_CORTE * vc * 10.0 * b * d, vc, rho_w


def resistencia_punzonamiento(cx, cy, d, bo):
    """φvc [tonf/m²] para cortante bidireccional, ACI Tabla 22.6.5.2 (mks)."""
    beta = max(cx, cy) / min(cx, cy)
    vc = min(1.06,
             0.53 * (1.0 + 2.0 / beta),
             0.27 * (2.0 + ALPHA_S * d / bo)) * LAMBDA_S * LAMBDA * math.sqrt(FC)
    return PHI_CORTE * vc * 10.0, vc


def longitud_desarrollo(db_mm):
    """ld (barra recta, ACI 25.4.2.3) y ldh (gancho estándar, ACI 25.4.3.1) [m].
    ψt = ψe = ψg = 1 (barra inferior, sin recubrimiento epóxico, grado 420)."""
    db = db_mm / 10.0                                        # cm
    k = 6.6 if db_mm <= 19 else 5.3
    ld = max(FY / (k * LAMBDA * math.sqrt(FC)) * db, 30.0)
    fc_mpa, fy_mpa = FC * 0.0980665, FY * 0.0980665
    psi_c = fc_mpa / 105.0 + 0.6 if fc_mpa < 40 else 1.0
    psi_r = 1.6                                              # sin confinamiento
    ldh = fy_mpa * psi_r * psi_c / (23.0 * LAMBDA * math.sqrt(fc_mpa)) * db_mm ** 1.5
    ldh = max(ldh / 10.0, 8.0 * db, 15.0)
    return ld / 100.0, ldh / 100.0


def aplastamiento(Pu, Bx, By, cx, cy, h):
    """Resistencia al aplastamiento en el contacto columna-zapata (ACI 22.8.3.2)."""
    A1 = cx * cy
    # √(A2/A1): A2 geométricamente semejante a A1, limitada por la pirámide 1:2
    k = min(Bx / cx, By / cy, (cx + 4 * h) / cx, (cy + 4 * h) / cy, 2.0)
    phiBn_zap = PHI_APLAST * 0.85 * FC * 10.0 * A1 * k
    phiBn_col = PHI_APLAST * 0.85 * FC * 10.0 * A1
    return min(phiBn_zap, phiBn_col)


def distribuir_barras(As, ancho, db_mm, h):
    """Número y separación de barras para cubrir As [cm²] en un ancho [m].
    Separación máxima min(2h, 45 cm) en secciones críticas (ACI 8.7.2.2)."""
    Ab = area_barra(db_mm)
    s_max = min(2.0 * h, 0.45)
    L = ancho - 2.0 * REC - db_mm / 1000.0
    n = max(math.ceil(As / Ab), math.ceil(L / s_max) + 1, 2)
    s = L / (n - 1)
    return dict(n=n, db=db_mm, s=s, As=n * Ab, s_max=s_max)


def texto_armado(barras):
    return f"{barras['n']} Ø{barras['db']} mm @ {math.floor(barras['s'] * 100)} cm"


def distribucion_rectangular(As, Bx, By, db_mm, h):
    """ACI 13.3.3.3: en zapatas rectangulares el acero de la dirección corta se
    concentra γs·As en una banda central de ancho igual al lado corto."""
    if abs(Bx - By) < 1e-6:
        return "No aplica (zapata cuadrada: distribución uniforme, ACI 13.3.3.2)"
    L, B = max(Bx, By), min(Bx, By)
    gs = 2.0 / (L / B + 1.0)
    banda = distribuir_barras(gs * As, B + 2 * REC, db_mm, h)
    fuera = distribuir_barras((1 - gs) * As / 2.0, (L - B) / 2.0 + 2 * REC, db_mm, h)
    return (f"γs = {gs:.3f}: banda central {B:.2f} m -> {texto_armado(banda)}; "
            f"a cada lado -> {texto_armado(fuera)}")


# =============================================================================
# 5. DISEÑO DE UNA ZAPATA
# =============================================================================
def verificar(geo, combos, campos, h, db_mm):
    """Verifica la zapata para un peralte h y un diámetro de barra dados."""
    Bx, By, cx, cy = geo["Bx"], geo["By"], geo["cx"], geo["cy"]
    dbm = db_mm / 1000.0
    validos = [i for i, c in enumerate(campos) if c["estado"] == "OK"]

    # --- Flexión: momentos en la cara (independientes de d) -----------------
    Mx_list = [momentos_cara(campos[i], cx, cy) for i in validos]
    Mu_x = max(m[0] for m in Mx_list)
    Mu_y = max(m[1] for m in Mx_list)

    # Peralte efectivo = h − recubrimiento − medio diámetro de la barra de la
    # capa analizada. Las barras de una dirección se apoyan sobre las de la otra,
    # así que la capa superior pierde un diámetro: d = h − rec − 1.5·db.
    # La dirección con mayor momento va en la capa inferior (mayor d) y cada
    # dirección se revisa con su propio d, que en la capa superior es el más
    # desfavorable.
    if Mu_x >= Mu_y:
        dx, dy = h - REC - dbm / 2.0, h - REC - 1.5 * dbm
    else:
        dx, dy = h - REC - 1.5 * dbm, h - REC - dbm / 2.0
    dm = (dx + dy) / 2.0                         # d promedio (ACI 22.6.2.1)
    r = dict(h=h, db=db_mm, dx=dx, dy=dy, d=dm, Mu_x=Mu_x, Mu_y=Mu_y,
             cumple=False, motivo="")
    if min(dx, dy) < D_MIN:
        r["motivo"] = "d < 15 cm (ACI 13.3.1.2)"
        return r
    # Hormigón en masa: si d ≥ voladizo, la sección crítica de cortante (a d de
    # la cara) queda fuera de la zapata en esa dirección.
    masa_x = dx >= (Bx - cx) / 2.0
    masa_y = dy >= (By - cy) / 2.0
    bloque = masa_x and masa_y
    r.update(masa_x=masa_x, masa_y=masa_y, bloque=bloque)

    As_min_x = RHO_MIN * By * 100 * h * 100
    As_min_y = RHO_MIN * Bx * 100 * h * 100
    As_req_x = acero_flexion(Mu_x, By, dx)
    As_req_y = acero_flexion(Mu_y, Bx, dy)
    if As_req_x is None or As_req_y is None:
        r["motivo"] = "Sección insuficiente a flexión"
        return r
    bar_x = distribuir_barras(max(As_req_x, As_min_x), By, db_mm, h)
    bar_y = distribuir_barras(max(As_req_y, As_min_y), Bx, db_mm, h)
    phiMn_x, a_x, et_x, phi_x = momento_resistente(bar_x["As"], By, dx)
    phiMn_y, a_y, et_y, phi_y = momento_resistente(bar_y["As"], Bx, dy)
    r.update(As_req_x=As_req_x, As_req_y=As_req_y, As_min_x=As_min_x,
             As_min_y=As_min_y, bar_x=bar_x, bar_y=bar_y,
             phiMn_x=phiMn_x, phiMn_y=phiMn_y, eps_t_x=et_x, eps_t_y=et_y,
             phi_x=phi_x, phi_y=phi_y,
             DC_flex_x=Mu_x / phiMn_x, DC_flex_y=Mu_y / phiMn_y)

    # --- Resistencias a cortante ---------------------------------------------
    phiVc_x, vc_x, rho_x = resistencia_cortante_1D(bar_x["As"], By, dx)
    phiVc_y, vc_y, rho_y = resistencia_cortante_1D(bar_y["As"], Bx, dy)
    b1, b2 = cx + dm, cy + dm
    phi_vc_p, vc_p = resistencia_punzonamiento(cx, cy, dm, 2 * (b1 + b2))
    r.update(phiVc_x=phiVc_x, phiVc_y=phiVc_y, vc_x=vc_x, vc_y=vc_y,
             rho_w_x=rho_x, rho_w_y=rho_y, phi_vc_p=phi_vc_p, vc_p=vc_p)

    # --- Revisión combinación por combinación --------------------------------
    detalle = []
    for k, i in enumerate(validos):
        c, campo = combos[i], campos[i]
        Vu_x, Vu_y = cortante_1D(campo, cx, cy, dx, dy)
        pz = punzonamiento(campo, cx, cy, dm)
        detalle.append(dict(
            Combo=c["nombre"], Pu=c["P"], Mux_base=campo["Mx"], Muy_base=campo["My"],
            ex=campo["ex"], ey=campo["ey"], qu_max=campo["q_max"],
            qu_min_elastico=campo["q_min_elastico"], contacto=campo["contacto"],
            Mu_x=Mx_list[k][0], Mu_y=Mx_list[k][1],
            Vu1_x=Vu_x, Vu1_y=Vu_y, Vu_p=pz["Vu"], Msc_x=pz["Msc_x"],
            Msc_y=pz["Msc_y"], vu_p=pz["vu"],
            DC_flex_x=Mx_list[k][0] / phiMn_x, DC_flex_y=Mx_list[k][1] / phiMn_y,
            DC_v1_x=Vu_x / phiVc_x, DC_v1_y=Vu_y / phiVc_y,
            DC_punz=pz["vu"] / phi_vc_p, idx=i, bo=pz["bo"]))
    r["detalle"] = detalle

    for clave in ("DC_v1_x", "DC_v1_y", "DC_punz"):
        gob = max(detalle, key=lambda e: e[clave])
        r[clave] = gob[clave]
        r[clave + "_combo"] = gob["Combo"]
    r["combo_flex_x"] = max(detalle, key=lambda e: e["Mu_x"])["Combo"]
    r["combo_flex_y"] = max(detalle, key=lambda e: e["Mu_y"])["Combo"]
    r["idx_punz"] = max(detalle, key=lambda e: e["DC_punz"])["idx"]
    # Si la sección a d de la cara queda fuera del voladizo no hay cortante en
    # esa dirección: el voladizo trabaja como hormigón en masa.
    claves = []
    if masa_x:
        claves.append("DC_v1_x")
    if masa_y:
        claves.append("DC_v1_y")
    if bloque:
        # En ambas direcciones: la falla a 45° sale por los bordes, tampoco hay punzonamiento
        claves.append("DC_punz")
        r["idx_punz"] = max(detalle, key=lambda e: e["qu_max"])["idx"]
    for clave in claves:
        r[clave] = float("nan")
        r[clave + "_combo"] = "No aplica"

    # --- Criterio de aceptación por resistencia ------------------------------
    fallas = []
    if r["DC_flex_x"] > 1.0 or r["DC_flex_y"] > 1.0:
        fallas.append("flexión")
    if r["DC_v1_x"] > 1.0 or r["DC_v1_y"] > 1.0:
        fallas.append("cortante 1D")
    if r["DC_punz"] > 1.0:
        fallas.append("punzonamiento")
    if min(bar_x["s"], bar_y["s"]) - dbm < S_LIBRE_MIN:
        fallas.append("espaciamiento libre < 2.5 cm (ACI 25.2.1)")
    r["cumple"] = not fallas
    r["motivo"] = ", ".join(fallas)
    return r


def disenar_zapata(punto, datos):
    """Itera peralte h y diámetro de barra hasta cumplir todas las verificaciones."""
    f0 = datos.iloc[0]
    geo = dict(Bx=float(f0["Bx"]), By=float(f0["By"]),
               cx=float(f0["Bcx"]), cy=float(f0["Bcy"]))
    combos = [dict(nombre=f["Load"], P=float(f["FZ"]), Mx=float(f["MX"]),
                   My=float(f["MY"]), Vx=float(f["FX"]), Vy=float(f["FY"]))
              for _, f in datos.iterrows()]

    campos = None
    ultimo = None
    for h in np.arange(H_MIN_PUNTO.get(punto, H_MIN), H_MAX + 1e-9, DH):
        h = round(float(h), 3)
        # Presiones con los momentos en la base de la zapata. Sin traslado del
        # cortante no dependen de h y se calculan una sola vez.
        if campos is None or TRASLADAR_CORTANTE:
            campos = []
            for c in combos:
                Mx_b, My_b = c["Mx"], c["My"]
                if TRASLADAR_CORTANTE:
                    Mx_b, My_b = c["Mx"] - c["Vy"] * h, c["My"] + c["Vx"] * h
                campos.append(presion_contacto(c["P"], Mx_b, My_b, geo["Bx"], geo["By"]))
        if not any(cp["estado"] == "OK" for cp in campos):
            break
        # Diámetro: el primero de la lista con separación ≥ S_MIN que además se
        # ancla dentro del voladizo. Si ninguno se ancla así, se acepta el primero
        # que se ancla con el espaciamiento libre mínimo del ACI (25.2.1); y si
        # tampoco, el primero con s ≥ S_MIN (el anclaje queda reportado).
        elegido = respaldo = ajustado = None
        for db_mm in DIAMETROS:
            r = verificar(geo, combos, campos, h, db_mm)
            if "bar_x" not in r:
                continue
            s = min(r["bar_x"]["s"], r["bar_y"]["s"])
            r["anclaje"] = revisar_anclaje(geo, r)
            if s >= S_MIN:
                respaldo = respaldo or r
                if r["anclaje"]["cumple"]:
                    elegido = r
                    break
            elif ajustado is None and r["anclaje"]["cumple"]:
                ajustado = r
        r = elegido or ajustado or respaldo or r
        ultimo = (r, campos)
        # Con el armado elegido, si no cumple resistencia se aumenta el peralte
        if r["cumple"]:
            break

    r, campos = ultimo
    if "anclaje" not in r and "bar_x" in r:
        r["anclaje"] = revisar_anclaje(geo, r)
    r["geo"], r["combos"], r["campos"], r["punto"] = geo, combos, campos, punto
    r["estado_combos"] = {c["nombre"]: cp["estado"] for c, cp in zip(combos, campos)}

    # --- Aplastamiento columna-zapata ---------------------------------------
    Pu_max = max(c["P"] for c in combos)
    r["Pu_max"] = Pu_max
    r["phiBn"] = aplastamiento(Pu_max, geo["Bx"], geo["By"], geo["cx"], geo["cy"], r["h"])

    # --- Acero superior si hay levantamiento (ACI 18.13.2.5) ----------------
    levanta = any(cp["estado"] != "OK" or cp["contacto"] < 99.99 for cp in campos)
    r["levantamiento"] = levanta
    r["acero_superior"] = "No requerido (contacto total en todas las combinaciones)"
    if levanta and "bar_x" in r:
        h = r["h"]
        w = 1.2 * (GAMMA_C * h + GAMMA_S * max(DF - h, 0.0))   # peso propio + relleno
        Lx, Ly = (geo["Bx"] - geo["cx"]) / 2.0, (geo["By"] - geo["cy"]) / 2.0
        d_sup = h - REC - r["db"] / 1000.0
        As_tx = max(acero_flexion(w * geo["By"] * Lx ** 2 / 2.0, geo["By"], d_sup) or 0.0,
                    RHO_MIN * geo["By"] * 100 * h * 100)
        As_ty = max(acero_flexion(w * geo["Bx"] * Ly ** 2 / 2.0, geo["Bx"], d_sup) or 0.0,
                    RHO_MIN * geo["Bx"] * 100 * h * 100)
        r["acero_superior"] = (
            f"Paralelo a X: {texto_armado(distribuir_barras(As_tx, geo['By'], r['db'], h))}; "
            f"paralelo a Y: {texto_armado(distribuir_barras(As_ty, geo['Bx'], r['db'], h))}")
    if "bar_x" in r:
        As_ref = max(r["As_req_x"], r["As_min_x"])
        r["rectangular"] = distribucion_rectangular(As_ref, geo["Bx"], geo["By"], r["db"], r["h"])
    return r


def revisar_anclaje(geo, r):
    """Longitud disponible desde la cara de la columna hasta el extremo de la
    barra vs. ld recta o ldh con gancho estándar de 90° (ACI 13.2.8, 25.4)."""
    ld, ldh = longitud_desarrollo(r["db"])
    disp_x = (geo["Bx"] - geo["cx"]) / 2.0 - REC
    disp_y = (geo["By"] - geo["cy"]) / 2.0 - REC
    disp = min(disp_x, disp_y)
    if ld <= disp:
        tipo, ok = "Barra recta", True
    elif ldh <= disp:
        tipo, ok = "Gancho estándar 90°", True
    else:
        tipo, ok = "NO CUMPLE (ampliar zapata o reducir Ø)", False
    return dict(ld=ld, ldh=ldh, disp=disp, tipo=tipo, cumple=ok)


# =============================================================================
# 6. MEMORIA DE CÁLCULO
# =============================================================================
def imprimir_zapata(r):
    g = r["geo"]
    log("=" * 79)
    log(f" ZAPATA PUNTO {r['punto']}   Bx = {g['Bx']:.2f} m   By = {g['By']:.2f} m   "
        f"Columna {g['cx']*100:.0f} x {g['cy']*100:.0f} cm")
    log("=" * 79)
    log(f" Peralte adoptado h = {r['h']*100:.0f} cm,  Ø = {r['db']} mm")
    db = r["db"] / 10.0
    capa_x = "inferior" if r["dx"] >= r["dy"] else "superior"
    capa_y = "superior" if capa_x == "inferior" else "inferior"
    log(f" dx = h − rec − {'0.5' if capa_x == 'inferior' else '1.5'}·db = {r['h']*100:.0f} − "
        f"{REC*100:.1f} − {(0.5 if capa_x == 'inferior' else 1.5)*db:.1f} = {r['dx']*100:.1f} cm "
        f"(barras paralelas a X, capa {capa_x})")
    log(f" dy = h − rec − {'0.5' if capa_y == 'inferior' else '1.5'}·db = {r['h']*100:.0f} − "
        f"{REC*100:.1f} − {(0.5 if capa_y == 'inferior' else 1.5)*db:.1f} = {r['dy']*100:.1f} cm "
        f"(barras paralelas a Y, capa {capa_y})")
    log(f" d promedio = {r['d']*100:.1f} cm   (d ≥ 15 cm ✓)")
    log("")
    log(" 1) PRESIONES DE CONTACTO ÚLTIMAS")
    log(f"   {'Combo':<7}{'Pu[t]':>8}{'Mux[t·m]':>10}{'Muy[t·m]':>10}"
        f"{'qmax[t/m²]':>12}{'Contacto%':>11}  Estado")
    for c, cp in zip(r["combos"], r["campos"]):
        log(f"   {c['nombre']:<7}{c['P']:8.2f}{cp['Mx']:10.3f}{cp['My']:10.3f}"
            f"{cp['q_max']:12.2f}{cp['contacto']:11.1f}  {cp['estado']}")
    if "bar_x" not in r:
        log(f"   *** No fue posible diseñar: {r['motivo']}")
        return
    log("")
    log(" 2) FLEXIÓN EN LA CARA DE LA COLUMNA (ACI 13.2.7.1)")
    for e, b, Mu, Asr, Asm, bar, pM, et, cmb in (
            ("X", g["By"], r["Mu_x"], r["As_req_x"], r["As_min_x"], r["bar_x"],
             r["phiMn_x"], r["eps_t_x"], r["combo_flex_x"]),
            ("Y", g["Bx"], r["Mu_y"], r["As_req_y"], r["As_min_y"], r["bar_y"],
             r["phiMn_y"], r["eps_t_y"], r["combo_flex_y"])):
        log(f"   Barras paralelas a {e}: Mu = {Mu:.2f} t·m ({cmb}), ancho = {b:.2f} m")
        log(f"     As,req = {Asr:.2f} cm²   As,min = 0.0018·b·h = {Asm:.2f} cm²")
        log(f"     Armado: {texto_armado(bar)}  -> As = {bar['As']:.2f} cm²  "
            f"(s,max = {bar['s_max']*100:.0f} cm)")
        log(f"     φMn = {pM:.2f} t·m ≥ Mu  ->  D/C = {Mu/pM:.3f}   εt = {et:.4f} "
            f"{'(controlada por tracción)' if et >= 0.005 else ''}")
    log("")
    imprimir_cortantes(r)
    log("")
    log(" 5) VERIFICACIONES COMPLEMENTARIAS")
    a = r["anclaje"]
    log(f"   Anclaje: disponible = {a['disp']*100:.1f} cm, ld = {a['ld']*100:.1f} cm, "
        f"ldh = {a['ldh']*100:.1f} cm -> {a['tipo']}")
    log(f"   Aplastamiento: Pu,max = {r['Pu_max']:.2f} t ≤ φBn = {r['phiBn']:.2f} t -> "
        f"{'CUMPLE' if r['Pu_max'] <= r['phiBn'] else 'NO CUMPLE (colocar dowels)'}")
    log(f"   Zapata rectangular: {r['rectangular']}")
    log(f"   Acero superior: {r['acero_superior']}")
    log("")
    log(f" RESULTADO: {'CUMPLE' if r['cumple'] else 'NO CUMPLE: ' + r['motivo']}")
    log("")


def imprimir_cortantes(r):
    g = r["geo"]
    log(" 3) CORTANTE UNIDIRECCIONAL a d de la cara (ACI 8.4.3.2, 22.5.5.1c)")
    for e, B, c, d, masa, rho, vc, phiVc, dc, cmb in (
            ("X", g["Bx"], g["cx"], r["dx"], r["masa_x"], r["rho_w_x"], r["vc_x"],
             r["phiVc_x"], r["DC_v1_x"], r["DC_v1_x_combo"]),
            ("Y", g["By"], g["cy"], r["dy"], r["masa_y"], r["rho_w_y"], r["vc_y"],
             r["phiVc_y"], r["DC_v1_y"], r["DC_v1_y_combo"])):
        vol = (B - c) / 2.0
        log(f"   Dir. {e}: voladizo = ({B:.2f} − {c:.2f})/2 = {vol*100:.1f} cm,  d = {d*100:.1f} cm")
        if masa:
            log(f"           d ≥ voladizo: la sección a d de la cara queda fuera de la zapata.")
            log(f"           No hay cortante en una dirección -> hormigón en masa.")
        else:
            log(f"           La sección a d queda {(vol - d)*100:.1f} cm dentro del voladizo.")
            log(f"           ρw = {rho:.5f}, vc = 2.12·λs·ρw^(1/3)·√f'c = {vc:.2f} kg/cm², "
                f"φVc = {phiVc:.2f} t")
            log(f"           D/C máx = {dc:.3f} ({cmb})")
    log("")
    log(" 4) CORTANTE BIDIRECCIONAL - PUNZONAMIENTO a d/2 (ACI 22.6)")
    if r["bloque"]:
        log("   No aplica: d ≥ voladizo en ambas direcciones; la superficie de falla a 45°")
        log("   desde la cara de la columna sale por los bordes (hormigón en masa).")
        return
    gob = max(r["detalle"], key=lambda e: e["DC_punz"])
    log(f"   bo = {gob['bo']:.3f} m,  vc = {r['vc_p']:.2f} kg/cm²,  φvc = {r['phi_vc_p']:.1f} t/m²")
    log(f"   {'Combo':<7}{'Vu[t]':>8}{'Msc,x':>8}{'Msc,y':>8}{'vu[t/m²]':>10}{'D/C':>7}")
    for e in r["detalle"]:
        log(f"   {e['Combo']:<7}{e['Vu_p']:8.2f}{e['Msc_x']:8.3f}{e['Msc_y']:8.3f}"
            f"{e['vu_p']:10.2f}{e['DC_punz']:7.3f}")


# =============================================================================
# 7. GRÁFICOS
# =============================================================================
def campo_nombre(r):
    return r["combos"][r["idx_punz"]]["nombre"]


def figura_zapata(r):
    """Planta con presiones y secciones críticas, y corte con el armado.
    Usa Figure (sin pyplot) para poder dibujarse también dentro de la interfaz."""
    g = r["geo"]
    campo = r["campos"][r["idx_punz"]]
    Bx, By, cx, cy = g["Bx"], g["By"], g["cx"], g["cy"]
    fig = Figure(figsize=(13, 6))
    ax1, ax2 = fig.subplots(1, 2, gridspec_kw={"width_ratios": [1.15, 1]})
    # --- Planta ---------------------------------------------------------------
    im = ax1.pcolormesh(campo["X"], campo["Y"], campo["q"], shading="auto", cmap="YlOrRd")
    fig.colorbar(im, ax=ax1, label="qu [tonf/m²]")
    ax1.add_patch(Rectangle((-Bx / 2, -By / 2), Bx, By, fill=False, lw=2, ec="k"))
    ax1.add_patch(Rectangle((-cx / 2, -cy / 2), cx, cy, fc="0.55", ec="k", lw=1.5))
    ax1.plot([cx / 2] * 2, [-By / 2, By / 2], "k-", lw=1.2, label="Flexión (cara)")
    ax1.plot([-cx / 2] * 2, [-By / 2, By / 2], "k-", lw=1.2)
    ax1.plot([-Bx / 2, Bx / 2], [cy / 2] * 2, "k-", lw=1.2)
    ax1.plot([-Bx / 2, Bx / 2], [-cy / 2] * 2, "k-", lw=1.2)
    if r.get("bloque"):
        titulo = f"Zapata P{r['punto']} - Hormigón en masa, qu ({campo_nombre(r)})"
    else:
        for s in (+1, -1):
            ax1.plot([s * (cx / 2 + r["dx"])] * 2, [-By / 2, By / 2], "b--", lw=1.2,
                     label="Cortante 1D (d)" if s == 1 else None)
            ax1.plot([-Bx / 2, Bx / 2], [s * (cy / 2 + r["dy"])] * 2, "b--", lw=1.2)
        b1, b2 = cx + r["d"], cy + r["d"]
        ax1.add_patch(Rectangle((-b1 / 2, -b2 / 2), b1, b2, fill=False, ec="g", lw=2,
                                ls="-.", label="Punzonamiento (d/2)"))
        titulo = f"Zapata P{r['punto']} - Planta y presión qu ({r['DC_punz_combo']})"
    ax1.set_aspect("equal")
    ax1.set_xlim(-Bx / 2 * 1.05, Bx / 2 * 1.05)
    ax1.set_ylim(-By / 2 * 1.05, By / 2 * 1.05)
    ax1.set_xlabel("X [m]")
    ax1.set_ylabel("Y [m]")
    ax1.set_title(titulo)
    ax1.legend(loc="upper right", fontsize=8)

    # --- Corte A-A (paralelo a X) --------------------------------------------
    h = r["h"]
    ax2.add_patch(Rectangle((-Bx / 2, 0), Bx, h, fc="0.88", ec="k", lw=1.5))
    ax2.add_patch(Rectangle((-cx / 2, h), cx, 0.5, fc="0.6", ec="k", lw=1.5))
    if r.get("bloque"):
        for s in (+1, -1):
            ax2.plot([s * cx / 2, s * (cx / 2 + h)], [h, 0], "g--", lw=1.3,
                     label="Falla a 45° (sale de la zapata)" if s == 1 else None)
    db = r["db"] / 1000.0
    z_inf_x = REC + db / 2 if r["dx"] >= r["dy"] else REC + 1.5 * db
    z_inf_y = REC + 1.5 * db if r["dx"] >= r["dy"] else REC + db / 2
    ax2.plot([-Bx / 2 + REC, Bx / 2 - REC], [z_inf_x] * 2, "r-", lw=2.5,
             label=f"As paralelo a X: {texto_armado(r['bar_x'])}")
    n_y = r["bar_y"]["n"]
    xs = np.linspace(-Bx / 2 + REC + db / 2, Bx / 2 - REC - db / 2, n_y)
    ax2.plot(xs, [z_inf_y] * n_y, "o", color="navy", ms=5,
             label=f"As paralelo a Y: {texto_armado(r['bar_y'])}")
    if r["levantamiento"]:
        ax2.plot([-Bx / 2 + REC, Bx / 2 - REC], [h - REC] * 2, "m-", lw=2,
                 label="Acero superior (levantamiento)")
    ax2.annotate("", xy=(Bx / 2 + 0.08, 0), xytext=(Bx / 2 + 0.08, h),
                 arrowprops=dict(arrowstyle="<->"))
    ax2.text(Bx / 2 + 0.11, h / 2, f"h = {h*100:.0f} cm", va="center")
    ax2.annotate("", xy=(-Bx / 2, -0.08), xytext=(Bx / 2, -0.08),
                 arrowprops=dict(arrowstyle="<->"))
    ax2.text(0, -0.15, f"Bx = {Bx:.2f} m", ha="center")
    ax2.set_aspect("equal")
    ax2.set_xlim(-Bx / 2 - 0.2, Bx / 2 + 0.55)
    ax2.set_ylim(-0.3, h + 0.8)
    ax2.axis("off")
    ax2.set_title("Corte A-A (paralelo a X)")
    ax2.legend(loc="upper center", bbox_to_anchor=(0.5, -0.02), fontsize=8, ncol=1)
    fig.tight_layout()
    return fig


def graficar_zapata(r, ruta):
    figura_zapata(r).savefig(ruta, dpi=150)


def figura_resumen(resultados):
    """Gráfico de barras de la relación demanda/capacidad de todas las zapatas."""
    puntos = [f"P{r['punto']}" for r in resultados]
    series = [("DC_flex_x", "Flexión X"), ("DC_flex_y", "Flexión Y"),
              ("DC_v1_x", "Cortante 1D X"), ("DC_v1_y", "Cortante 1D Y"),
              ("DC_punz", "Punzonamiento")]
    x = np.arange(len(puntos))
    ancho = 0.16
    fig = Figure(figsize=(13, 5))
    ax = fig.subplots()
    for k, (clave, nombre) in enumerate(series):
        ax.bar(x + (k - 2) * ancho, [r.get(clave, np.nan) for r in resultados],
               ancho, label=nombre)
    ax.axhline(1.0, color="r", ls="--", lw=1.5, label="Límite D/C = 1.0")
    ax.set_xticks(x)
    ax.set_xticklabels(puntos)
    ax.set_ylabel("Demanda / Capacidad")
    ax.set_title("Relación demanda/capacidad de las zapatas diseñadas (ACI 318-19)")
    ax.legend(fontsize=8, loc="upper left", bbox_to_anchor=(1.01, 1.0))
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig


def graficar_resumen(resultados, ruta):
    figura_resumen(resultados).savefig(ruta, dpi=150)


# =============================================================================
# 8. EXPORTACIÓN A EXCEL
# =============================================================================
def dc_texto(valor):
    return "No aplica" if math.isnan(valor) else round(valor, 3)


def exportar_excel(resultados, ruta):
    resumen, detalle = [], []
    for r in resultados:
        g = r["geo"]
        fila = {"Punto": r["punto"], "Bx [m]": g["Bx"], "By [m]": g["By"],
                "Columna [cm]": f"{g['cx']*100:.0f}x{g['cy']*100:.0f}",
                "h [cm]": r["h"] * 100, "dx [cm]": round(r["dx"] * 100, 1),
                "dy [cm]": round(r["dy"] * 100, 1), "Pu max [t]": r["Pu_max"],
                "qu max [t/m²]": round(max(c["q_max"] for c in r["campos"]), 2),
                "Contacto mín [%]": round(min(c["contacto"] for c in r["campos"]), 1)}
        if "bar_x" in r:
            fila.update({
                "Mu x [t·m]": round(r["Mu_x"], 2), "As req x [cm²]": round(r["As_req_x"], 2),
                "As min x [cm²]": round(r["As_min_x"], 2), "Armado X": texto_armado(r["bar_x"]),
                "Mu y [t·m]": round(r["Mu_y"], 2), "As req y [cm²]": round(r["As_req_y"], 2),
                "As min y [cm²]": round(r["As_min_y"], 2), "Armado Y": texto_armado(r["bar_y"]),
                "D/C flexión X": round(r["DC_flex_x"], 3), "D/C flexión Y": round(r["DC_flex_y"], 3),
                "D/C cortante 1D X": dc_texto(r["DC_v1_x"]), "D/C cortante 1D Y": dc_texto(r["DC_v1_y"]),
                "D/C punzonamiento": dc_texto(r["DC_punz"]),
                "Combo punzonamiento": r["DC_punz_combo"],
                "Cortante 1D": ("No aplica en X ni Y (hormigón en masa)" if r["bloque"] else
                                "No aplica en X (hormigón en masa)" if r["masa_x"] else
                                "No aplica en Y (hormigón en masa)" if r["masa_y"] else
                                "Se verifica en X e Y"),
                "Anclaje": r["anclaje"]["tipo"],
                "Aplastamiento": "CUMPLE" if r["Pu_max"] <= r["phiBn"] else "NO CUMPLE",
                "Acero superior": r["acero_superior"],
                "Estado": "CUMPLE" if r["cumple"] else "NO CUMPLE: " + r["motivo"]})
            for e in r["detalle"]:
                d = {"Punto": r["punto"]}
                d.update({k: (round(v, 4) if isinstance(v, float) else v)
                          for k, v in e.items() if k not in ("idx",)})
                if r.get("bloque"):
                    for k in ("Vu1_x", "Vu1_y", "Vu_p", "Msc_x", "Msc_y", "vu_p",
                              "DC_v1_x", "DC_v1_y", "DC_punz", "bo"):
                        d[k] = "No aplica"
                detalle.append(d)
        resumen.append(fila)
    parametros = pd.DataFrame({
        "Parámetro": ["f'c [kg/cm²]", "fy [kg/cm²]", "Recubrimiento [m]", "φ flexión",
                      "φ cortante", "λ", "λs (ACI 13.2.6.2)", "αs", "ρ mínima",
                      "h mínimo [m]", "Δh [m]", "h inicial por zapata [m]",
                      "Traslado de cortante V·h", "Malla"],
        "Valor": [FC, FY, REC, PHI_FLEXION, PHI_CORTE, LAMBDA, LAMBDA_S, ALPHA_S,
                  RHO_MIN, H_MIN, DH, str(H_MIN_PUNTO) if H_MIN_PUNTO else "—",
                  TRASLADAR_CORTANTE, f"{N_MALLA}x{N_MALLA}"]})
    with pd.ExcelWriter(ruta, engine="openpyxl") as wr:
        pd.DataFrame(resumen).to_excel(wr, sheet_name="Resumen", index=False)
        pd.DataFrame(detalle).to_excel(wr, sheet_name="Detalle_Combos", index=False)
        parametros.to_excel(wr, sheet_name="Parametros", index=False)
        for hoja in wr.sheets.values():
            for col in hoja.columns:
                largo = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                hoja.column_dimensions[col[0].column_letter].width = min(largo + 2, 60)


# =============================================================================
# 9. PROGRAMA PRINCIPAL
# =============================================================================
def main(progreso=None):
    """Diseña todas las zapatas del Excel y guarda los resultados.
    progreso(k, n, punto): función opcional que se llama al terminar cada zapata
    (la usa la interfaz gráfica)."""
    MEMORIA.clear()
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    df = leer_datos(ARCHIVO_EXCEL, HOJA_EXCEL)
    log("DISEÑO DE ZAPATAS AISLADAS - ACI 318-19")
    log(f"Archivo: {os.path.basename(ARCHIVO_EXCEL)}  |  {len(df)} filas, "
        f"{df['Point'].nunique()} zapatas, combinaciones: {', '.join(df['Load'].unique())}")
    log(f"f'c = {FC:.0f} kg/cm²   fy = {FY:.0f} kg/cm²   rec = {REC*100:.1f} cm   "
        f"traslado V·h = {TRASLADAR_CORTANTE}")
    log("")

    resultados = []
    grupos = list(df.groupby("Point", sort=True))
    for k, (punto, datos) in enumerate(grupos, start=1):
        r = disenar_zapata(punto, datos)
        inicio = len(MEMORIA)
        imprimir_zapata(r)
        r["memoria"] = "\n".join(MEMORIA[inicio:])
        if "bar_x" in r:
            graficar_zapata(r, os.path.join(CARPETA_SALIDA, f"Zapata_P{punto}.png"))
        resultados.append(r)
        if progreso:
            progreso(k, len(grupos), punto)

    log("=" * 79)
    log(" RESUMEN GENERAL")
    log("=" * 79)
    log(f" {'Pto':<4}{'BxBy[m]':>10}{'h[cm]':>7}  {'Armado X':<22}{'Armado Y':<22}"
        f"{'D/C máx':>8}  Estado")
    for r in resultados:
        if "bar_x" not in r:
            log(f" {r['punto']:<4} NO DISEÑADA: {r['motivo']}")
            continue
        dc = np.nanmax([r["DC_flex_x"], r["DC_flex_y"], r["DC_v1_x"], r["DC_v1_y"], r["DC_punz"]])
        log(f" {r['punto']:<4}{r['geo']['Bx']:>5.2f}x{r['geo']['By']:<4.2f}{r['h']*100:>7.0f}  "
            f"{texto_armado(r['bar_x']):<22}{texto_armado(r['bar_y']):<22}{dc:8.3f}  "
            f"{'CUMPLE' if r['cumple'] else 'NO CUMPLE'}"
            f"{'  (hormigón en masa)' if r['bloque'] else ''}")

    graficar_resumen([r for r in resultados if "bar_x" in r],
                     os.path.join(CARPETA_SALIDA, "Resumen_Demanda_Capacidad.png"))
    exportar_excel(resultados, os.path.join(CARPETA_SALIDA, "Resultados_Zapatas.xlsx"))
    with open(os.path.join(CARPETA_SALIDA, "Memoria_Calculo.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(MEMORIA))
    print(f"\nResultados guardados en: {CARPETA_SALIDA}")
    return resultados


if __name__ == "__main__":
    resultados = main()
