# -*- coding: utf-8 -*-
"""
===============================================================================
 VERSIÓN WEB — DISEÑO DE ZAPATAS AISLADAS (ACI 318-19)
===============================================================================
 Interfaz web con Streamlit para el programa Diseno_Zapatas_Aisladas.py:
   - subir el Excel de entrada (o usar los datos de ejemplo),
   - modificar los parámetros de diseño,
   - ejecutar el diseño y revisar el resumen, el detalle de cada zapata, el
     gráfico de demanda/capacidad y la memoria de cálculo,
   - descargar el Excel de resultados, la memoria y las figuras.

 Ejecutar en la PC:   streamlit run app.py      (se abre en el navegador)
 Publicar:            Streamlit Community Cloud, con este repositorio de GitHub.
===============================================================================
"""

import io
import os
import math
import zipfile
import tempfile
import threading
import traceback

import openpyxl
import pandas as pd
import streamlit as st

import Diseno_Zapatas_Aisladas as dz

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EJEMPLO = os.path.join(CARPETA, "InputData01.xlsx")
COLUMNAS_REQUERIDAS = ["Point", "Load", "FX", "FY", "FZ", "MX", "MY", "Bx", "By", "Bcx", "Bcy"]

# Parámetros editables: (variable del programa, etiqueta, unidad, factor, paso, formato).
# valor en el programa = valor en pantalla × factor
PARAMETROS = [
    ("FC", "f'c del hormigón", "kg/cm²", 1.0, 10.0, "%.0f"),
    ("FY", "fy del acero", "kg/cm²", 1.0, 100.0, "%.0f"),
    ("REC", "Recubrimiento", "cm", 0.01, 0.5, "%.1f"),
    ("H_MIN", "Peralte inicial (h mín)", "cm", 0.01, 5.0, "%.0f"),
    ("H_MAX", "Peralte máximo (h máx)", "cm", 0.01, 5.0, "%.0f"),
    ("DH", "Incremento de h (Δh)", "cm", 0.01, 1.0, "%.0f"),
    ("S_MIN", "Separación mínima de barras", "cm", 0.01, 1.0, "%.0f"),
    ("DF", "Profundidad de desplante Df", "m", 1.0, 0.1, "%.2f"),
    ("GAMMA_S", "Peso específico del relleno", "t/m³", 1.0, 0.1, "%.2f"),
]

COLOR_OK = "#E8F5E9"
COLOR_MASA = "#E3F2FD"
COLOR_FALLA = "#FDECEA"

st.set_page_config(page_title="Diseño de zapatas aisladas — ACI 318-19",
                   page_icon="🏗️", layout="wide")

AUTORA = "Daysi Velecela"
PROGRAMA = "Maestría en Construcción de Estructuras Sismorresistentes — UPS"

# Encabezado fijo en la parte superior de toda la página (sobre el panel lateral y
# el contenido) y ocultamiento del botón "Deploy" y del menú de tres puntos
# (respaldo de toolbarMode = "minimal" en .streamlit/config.toml).
st.markdown(f"""
<style>
:root {{ --alto-encabezado: 3.2rem; }}
[data-testid="stAppDeployButton"], [data-testid="stMainMenu"], footer {{ display: none !important; }}
/* Toda la aplicación (panel lateral, barra de Streamlit y contenido) empieza debajo del encabezado */
[data-testid="stAppViewContainer"] {{
    top: var(--alto-encabezado) !important;
    height: calc(100% - var(--alto-encabezado)) !important;
}}
.block-container {{ padding-top: 3.5rem; }}
.encabezado {{
    position: fixed; top: 0; left: 0; right: 0; height: var(--alto-encabezado);
    z-index: 1000000; display: flex; justify-content: space-between; align-items: center;
    gap: 1rem; padding: 0 1.5rem; background: #1F3A5F; color: #FFFFFF;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.25);
}}
.encabezado .autora {{ font-size: 1.15rem; font-weight: 700; white-space: nowrap; flex-shrink: 0; }}
.encabezado .programa {{
    font-size: 0.95rem; opacity: 0.92; min-width: 0;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}}
@media (max-width: 640px) {{
    :root {{ --alto-encabezado: 3.9rem; }}
    .encabezado {{ padding: 0 0.9rem; }}
    .encabezado .autora {{ font-size: 1rem; }}
    .encabezado .programa {{
        font-size: 0.75rem; line-height: 1.25; text-align: right;
        white-space: normal; overflow: visible;
    }}
}}
</style>
<div class="encabezado">
  <span class="autora">{AUTORA}</span>
  <span class="programa">{PROGRAMA}</span>
</div>
""", unsafe_allow_html=True)


# =============================================================================
# Recursos compartidos entre usuarios
# =============================================================================
@st.cache_resource
def candado():
    """El módulo de cálculo guarda los parámetros en variables globales, así que
    se diseña de a un usuario a la vez para que nadie cambie los datos de otro."""
    return threading.Lock()


@st.cache_resource
def valores_por_defecto():
    """Valores de la sección 0 de Diseno_Zapatas_Aisladas.py (antes de cualquier diseño)."""
    valores = {attr: getattr(dz, attr) for attr, *_ in PARAMETROS}
    valores["H_MIN_PUNTO"] = dict(dz.H_MIN_PUNTO)
    valores["TRASLADAR_CORTANTE"] = dz.TRASLADAR_CORTANTE
    return valores


@st.cache_data
def bytes_ejemplo():
    with open(ARCHIVO_EJEMPLO, "rb") as f:
        return f.read()


@st.cache_data
def hojas_excel(contenido):
    libro = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True)
    try:
        return libro.sheetnames
    finally:
        libro.close()


@st.cache_data
def leer_hoja(contenido, hoja):
    return dz.leer_datos(io.BytesIO(contenido), hoja)


# =============================================================================
# Utilidades
# =============================================================================
def formato_dc(valor):
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return "No aplica"
    return f"{valor:.3f}"


def maximo_valido(*valores):
    validos = [v for v in valores if not math.isnan(v)]
    return max(validos) if validos else float("nan")


def texto_h_puntos(diccionario):
    return ", ".join(f"{k}:{v * 100:g}" for k, v in diccionario.items())


def leer_h_puntos(texto):
    """«1:40, 2:40» -> {1: 0.40, 2: 0.40}. Lanza ValueError si el formato no es válido."""
    resultado = {}
    for parte in texto.replace(";", ",").split(","):
        parte = parte.strip()
        if not parte:
            continue
        try:
            punto, h = parte.split(":")
            resultado[int(punto)] = float(h) / 100.0
        except ValueError:
            raise ValueError("«Peralte inicial por zapata» debe tener el formato zapata:h[cm], "
                             "por ejemplo 1:40, 2:40.")
    return resultado


def restaurar_valores():
    defecto = valores_por_defecto()
    for attr, _, _, factor, _, _ in PARAMETROS:
        st.session_state[attr] = float(defecto[attr] / factor)
    st.session_state["h_puntos"] = texto_h_puntos(defecto["H_MIN_PUNTO"])
    st.session_state["traslado"] = defecto["TRASLADAR_CORTANTE"]


def tabla_resumen(resultados):
    filas = []
    for r in resultados:
        if "bar_x" not in r:
            filas.append({"Zapata": f"P{r['punto']}", "Estado": f"NO DISEÑADA: {r['motivo']}"})
            continue
        g = r["geo"]
        filas.append({
            "Zapata": f"P{r['punto']}",
            "B [m]": f"{g['Bx']:.2f} × {g['By']:.2f}",
            "h [cm]": f"{r['h'] * 100:.0f}",
            "Armado X": dz.texto_armado(r["bar_x"]),
            "Armado Y": dz.texto_armado(r["bar_y"]),
            "D/C flexión": formato_dc(max(r["DC_flex_x"], r["DC_flex_y"])),
            "D/C cortante 1D": formato_dc(maximo_valido(r["DC_v1_x"], r["DC_v1_y"])),
            "D/C punzonamiento": formato_dc(r["DC_punz"]),
            "Tipo": "Hormigón en masa" if r["bloque"] else "Zapata flexible",
            "Estado": "CUMPLE" if r["cumple"] else "NO CUMPLE",
        })
    return pd.DataFrame(filas)


def colorear_filas(fila):
    if fila.get("Estado") != "CUMPLE":
        color = COLOR_FALLA
    elif fila.get("Tipo") == "Hormigón en masa":
        color = COLOR_MASA
    else:
        color = COLOR_OK
    return [f"background-color: {color}; color: #1a1a1a"] * len(fila)


# =============================================================================
# Diseño
# =============================================================================
def disenar(contenido, nombre_archivo, hoja, valores, barra):
    """Ejecuta el diseño en una carpeta temporal y devuelve todo lo que se muestra
    o se descarga, ya leído en memoria (la carpeta se borra al terminar)."""
    with candado(), tempfile.TemporaryDirectory() as tmp:
        ruta_excel = os.path.join(tmp, os.path.basename(nombre_archivo) or "datos.xlsx")
        with open(ruta_excel, "wb") as f:
            f.write(contenido)
        salida = os.path.join(tmp, "Resultados")
        for attr, valor in valores.items():
            setattr(dz, attr, valor)
        dz.ARCHIVO_EXCEL, dz.HOJA_EXCEL, dz.CARPETA_SALIDA = ruta_excel, hoja, salida

        def progreso(k, n, punto):
            barra.progress(k / n, text=f"Zapata P{punto} diseñada ({k} de {n})")

        resultados = dz.main(progreso=progreso)

        def leer(nombre):
            with open(os.path.join(salida, nombre), "rb") as f:
                return f.read()

        zapatas = {}
        for r in resultados:
            if "bar_x" not in r:
                continue
            g = r["geo"]
            zapatas[f"P{r['punto']}"] = {
                "figura": leer(f"Zapata_P{r['punto']}.png"),
                "memoria": r.get("memoria", ""),
                "info": (f"{g['Bx']:.2f} × {g['By']:.2f} m · columna {g['cx'] * 100:.0f} × "
                         f"{g['cy'] * 100:.0f} cm · h = {r['h'] * 100:.0f} cm · "
                         f"{'hormigón en masa' if r['bloque'] else 'zapata flexible'} · "
                         f"{'CUMPLE' if r['cumple'] else 'NO CUMPLE'}"),
            }
        archivo_zip = io.BytesIO()
        with zipfile.ZipFile(archivo_zip, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre in sorted(os.listdir(salida)):
                z.write(os.path.join(salida, nombre), arcname=f"Resultados/{nombre}")
        return {
            "resumen": tabla_resumen(resultados),
            "total": len(resultados),
            "cumplen": sum(1 for r in resultados if r.get("cumple")),
            "masa": sum(1 for r in resultados if r.get("bloque")),
            "zapatas": zapatas,
            "dc": leer("Resumen_Demanda_Capacidad.png"),
            "memoria": leer("Memoria_Calculo.txt").decode("utf-8"),
            "excel": leer("Resultados_Zapatas.xlsx"),
            "zip": archivo_zip.getvalue(),
        }


# =============================================================================
# Página
# =============================================================================
defecto = valores_por_defecto()
if "FC" not in st.session_state:
    restaurar_valores()

st.title("Diseño de zapatas aisladas de hormigón armado")
st.caption("Flexión, cortante unidireccional y punzonamiento según ACI 318-19")

# --- Panel lateral: datos y parámetros ----------------------------------------
with st.sidebar:
    st.header("1. Datos de entrada")
    subido = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])
    if subido is None:
        contenido, nombre_archivo = bytes_ejemplo(), "InputData01.xlsx"
        st.info("Sin archivo: se usan los datos de ejemplo (InputData01.xlsx).")
    else:
        contenido, nombre_archivo = subido.getvalue(), subido.name
    st.download_button("Descargar datos de ejemplo / plantilla", data=bytes_ejemplo(),
                       file_name="InputData01.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    try:
        hojas = hojas_excel(contenido)
    except Exception as error:
        st.error(f"No se pudo abrir el Excel: {error}")
        st.stop()
    hoja = st.selectbox("Hoja", hojas, index=hojas.index("Tarea1") if "Tarea1" in hojas else 0)

    st.header("2. Parámetros de diseño")
    for attr, etiqueta, unidad, factor, paso, formato in PARAMETROS:
        st.number_input(f"{etiqueta} [{unidad}]", min_value=0.01, step=paso, format=formato, key=attr)

    st.header("3. Opciones")
    st.text_input("Peralte inicial por zapata", key="h_puntos",
                  help="Formato zapata:h[cm], por ejemplo 1:40, 2:40. Déjelo vacío para usar "
                       "el peralte inicial general en todas las zapatas.")
    st.checkbox("Trasladar cortantes FX, FY a la base (V·h)", key="traslado",
                help="Desactivado: las cadenas toman el cortante (criterio de clase).")
    st.button("Restaurar valores por defecto", on_click=restaurar_valores)
    disenar_ahora = st.button("▶ Diseñar zapatas", type="primary")

# --- Lectura y validación -----------------------------------------------------
errores = []
try:
    datos = leer_hoja(contenido, hoja)
except Exception as error:
    datos = None
    errores.append(f"No se pudieron leer los datos de la hoja «{hoja}»: {error}")
if datos is not None:
    faltan = [c for c in COLUMNAS_REQUERIDAS if c not in datos.columns]
    if faltan:
        errores.append("Al Excel le faltan las columnas: " + ", ".join(faltan) +
                       ". Revise la plantilla de ejemplo.")
    elif datos.empty:
        errores.append("La hoja no tiene filas de datos.")

valores = {attr: st.session_state[attr] * factor for attr, _, _, factor, _, _ in PARAMETROS}
if valores["H_MIN"] >= valores["H_MAX"]:
    errores.append("El peralte inicial debe ser menor que el peralte máximo.")
try:
    valores["H_MIN_PUNTO"] = leer_h_puntos(st.session_state["h_puntos"])
except ValueError as error:
    errores.append(str(error))
valores["TRASLADAR_CORTANTE"] = st.session_state["traslado"]
firma = (nombre_archivo, len(contenido), hoja, tuple(sorted((k, str(v)) for k, v in valores.items())))

if disenar_ahora:
    if errores:
        st.error("No se puede diseñar:\n\n- " + "\n- ".join(errores))
    else:
        barra = st.progress(0.0, text="Iniciando el diseño…")
        try:
            st.session_state["resultado"] = disenar(contenido, nombre_archivo, hoja, valores, barra)
            st.session_state["firma"] = firma
            barra.empty()
        except Exception:
            barra.empty()
            st.error("El diseño se detuvo por un error.")
            with st.expander("Detalle del error"):
                st.code(traceback.format_exc(), language=None)
elif errores:
    for e in errores:
        st.warning(e)

res = st.session_state.get("resultado")

# --- Descargas y avisos -------------------------------------------------------
if res:
    if st.session_state.get("firma") != firma:
        st.warning("Cambió los datos o los parámetros. Los resultados mostrados corresponden al "
                   "diseño anterior: presione «Diseñar zapatas» para actualizarlos.")
    else:
        st.success(f"Diseño terminado: {res['total']} zapatas · {res['cumplen']} cumplen · "
                   f"{res['masa']} trabajan como hormigón en masa.")
    c1, c2, c3 = st.columns(3)
    c1.download_button("Descargar Excel de resultados", res["excel"], "Resultados_Zapatas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("Descargar memoria de cálculo", res["memoria"].encode("utf-8"),
                       "Memoria_Calculo.txt", mime="text/plain")
    c3.download_button("Descargar todo (.zip)", res["zip"], "Resultados_Zapatas.zip",
                       mime="application/zip")

pestanas = st.tabs(["Datos de entrada", "Resumen", "Detalle por zapata",
                    "Demanda / Capacidad", "Memoria completa"])

with pestanas[0]:
    if datos is not None:
        info = f"{len(datos)} filas"
        if {"Point", "Load"} <= set(datos.columns):
            info += (f" · {datos['Point'].nunique()} zapatas · combinaciones: "
                     f"{', '.join(map(str, datos['Load'].unique()))}")
        st.caption(info + " · Unidades: fuerzas [tonf], momentos [tonf·m], dimensiones [m]")
        st.dataframe(datos, hide_index=True, width="stretch", height=480)

sin_resultados = "Aún no se ha ejecutado el diseño. Revise los datos y presione «Diseñar zapatas»."

with pestanas[1]:
    if res:
        st.dataframe(res["resumen"].style.apply(colorear_filas, axis=1), hide_index=True,
                     width="stretch", height=36 * (len(res["resumen"]) + 1) + 4)
        st.caption("Verde: cumple · Azul: cumple como hormigón en masa (d ≥ voladizo, sin "
                   "cortante) · Rojo: no cumple. El detalle de cada zapata está en la pestaña "
                   "siguiente.")
    else:
        st.info(sin_resultados)

with pestanas[2]:
    if res and res["zapatas"]:
        nombre = st.selectbox("Zapata", list(res["zapatas"]))
        z = res["zapatas"][nombre]
        st.caption(z["info"])
        st.image(z["figura"], width="stretch")
        st.download_button(f"Descargar figura de {nombre}", z["figura"], f"Zapata_{nombre}.png",
                           mime="image/png")
        st.subheader("Memoria de cálculo")
        st.code(z["memoria"], language=None, height=520)
    else:
        st.info(sin_resultados)

with pestanas[3]:
    if res:
        st.image(res["dc"], width="stretch")
    else:
        st.info(sin_resultados)

with pestanas[4]:
    if res:
        st.code(res["memoria"], language=None, height=700)
    else:
        st.info(sin_resultados)
