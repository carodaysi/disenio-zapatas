# -*- coding: utf-8 -*-
"""
===============================================================================
 INTERFAZ GRÁFICA — DISEÑO DE ZAPATAS AISLADAS (ACI 318-19)
===============================================================================
 Ventana para usar el programa Diseno_Zapatas_Aisladas.py sin editar código:
   - elegir el Excel de entrada y la hoja, y ver sus datos,
   - modificar los parámetros de diseño,
   - ejecutar el diseño (con barra de progreso),
   - revisar el resumen, la figura y la memoria de cada zapata, el gráfico de
     demanda/capacidad y la memoria completa.

 Ejecución: python Interfaz_Zapatas.py   (o F5 en Spyder, o doble clic en
            Ejecutar_Interfaz.bat). Diseno_Zapatas_Aisladas.py debe estar en
            la misma carpeta.
===============================================================================
"""

import os
import sys
import math
import queue
import threading
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

import openpyxl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

CARPETA = os.path.dirname(os.path.abspath(__file__))
if CARPETA not in sys.path:
    sys.path.insert(0, CARPETA)
import Diseno_Zapatas_Aisladas as dz  # noqa: E402

# Parámetros editables: (variable del programa, etiqueta, unidad, factor).
# valor en el programa = valor en la interfaz × factor
PARAMETROS = [
    ("FC", "f'c del hormigón", "kg/cm²", 1.0),
    ("FY", "fy del acero", "kg/cm²", 1.0),
    ("REC", "Recubrimiento", "cm", 0.01),
    ("H_MIN", "Peralte inicial (h mín)", "cm", 0.01),
    ("H_MAX", "Peralte máximo (h máx)", "cm", 0.01),
    ("DH", "Incremento de h (Δh)", "cm", 0.01),
    ("S_MIN", "Separación mínima de barras", "cm", 0.01),
    ("DF", "Profundidad de desplante Df", "m", 1.0),
    ("GAMMA_S", "Peso específico del relleno", "t/m³", 1.0),
]

AZUL = "#1F3A5F"
COLOR_OK = "#E8F5E9"
COLOR_MASA = "#E3F2FD"
COLOR_FALLA = "#FDECEA"


def formato_dc(valor):
    """Relación D/C con 3 decimales o «No aplica» si no corresponde (NaN)."""
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return "No aplica"
    return f"{valor:.3f}"


def maximo_valido(*valores):
    validos = [v for v in valores if not math.isnan(v)]
    return max(validos) if validos else float("nan")


class PanelDesplazable(ttk.Frame):
    """Marco con barra de desplazamiento vertical (para el panel de datos)."""

    def __init__(self, padre, ancho=340):
        super().__init__(padre)
        self.lienzo = tk.Canvas(self, width=ancho, highlightthickness=0)
        barra = ttk.Scrollbar(self, orient="vertical", command=self.lienzo.yview)
        self.interior = ttk.Frame(self.lienzo, padding=(10, 6))
        self.interior.bind("<Configure>", lambda e: self.lienzo.configure(
            scrollregion=self.lienzo.bbox("all")))
        ventana = self.lienzo.create_window((0, 0), window=self.interior, anchor="nw")
        self.lienzo.bind("<Configure>", lambda e: self.lienzo.itemconfigure(ventana, width=e.width))
        self.lienzo.configure(yscrollcommand=barra.set)
        self.lienzo.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        # Rueda del ratón solo mientras el cursor está sobre el panel
        self.interior.bind("<Enter>", lambda e: self.lienzo.bind_all("<MouseWheel>", self._rueda))
        self.interior.bind("<Leave>", lambda e: self.lienzo.unbind_all("<MouseWheel>"))

    def _rueda(self, evento):
        self.lienzo.yview_scroll(int(-evento.delta / 120), "units")


class AplicacionZapatas:
    def __init__(self, root):
        self.root = root
        root.title("Diseño de Zapatas Aisladas — ACI 318-19")
        root.minsize(1000, 620)
        try:
            root.state("zoomed")
        except tk.TclError:
            root.geometry("1400x850")

        # Factor de escala de la pantalla (1.0 = 96 ppp). Las medidas en píxeles
        # se multiplican por él para que la ventana se vea igual con escalado de Windows.
        self.escala = root.winfo_fpixels("1i") / 96.0
        self.cola = queue.Queue()
        self.resultados = []
        self.ejecutando = False
        self.canvas_zapata = self.toolbar_zapata = self.canvas_resumen = None
        self.valores_iniciales = {p[0]: getattr(dz, p[0]) for p in PARAMETROS}
        self.valores_iniciales["H_MIN_PUNTO"] = dict(dz.H_MIN_PUNTO)
        self.valores_iniciales["TRASLADAR_CORTANTE"] = dz.TRASLADAR_CORTANTE

        self._estilos()
        self._construir()
        self._restaurar_parametros()
        self._cargar_hojas(dz.ARCHIVO_EXCEL)

    # ------------------------------------------------------------------ estilo
    def px(self, valor):
        """Píxeles a 96 ppp -> píxeles reales de la pantalla."""
        return int(round(valor * self.escala))

    def _estilos(self):
        estilo = ttk.Style()
        if "vista" in estilo.theme_names():
            estilo.theme_use("vista")
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 15, "bold"), foreground=AZUL)
        estilo.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#555555")
        estilo.configure("Nota.TLabel", font=("Segoe UI", 8), foreground="#666666")
        estilo.configure("Seccion.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=AZUL)
        estilo.configure("Accion.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        alto_fila = int(tkfont.Font(family="Segoe UI", size=9).metrics("linespace") * 1.45)
        estilo.configure("Treeview", rowheight=alto_fila, font=("Segoe UI", 9))
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        estilo.configure("Estado.TLabel", font=("Segoe UI", 9), foreground="#333333")

    # --------------------------------------------------------------- ventana
    def _construir(self):
        encabezado = ttk.Frame(self.root, padding=(12, 8, 12, 4))
        encabezado.pack(fill="x")
        ttk.Label(encabezado, text="Diseño de zapatas aisladas de hormigón armado",
                  style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(encabezado, text="Flexión, cortante unidireccional y punzonamiento según "
                  "ACI 318-19  ·  Maestría en Construcción de Estructuras Sismorresistentes — UPS",
                  style="Sub.TLabel").pack(anchor="w")
        ttk.Separator(self.root).pack(fill="x", padx=8)

        cuerpo = ttk.Frame(self.root)
        cuerpo.pack(fill="both", expand=True)
        panel = PanelDesplazable(cuerpo, ancho=self.px(330))
        panel.pack(side="left", fill="y")
        self._panel_izquierdo(panel.interior)
        ttk.Separator(cuerpo, orient="vertical").pack(side="left", fill="y")

        self.notebook = ttk.Notebook(cuerpo)
        self.notebook.pack(side="left", fill="both", expand=True, padx=(6, 8), pady=6)
        self._pestana_datos()
        self._pestana_resumen()
        self._pestana_zapata()
        self._pestana_dc()
        self._pestana_memoria()

        barra = ttk.Frame(self.root, padding=(10, 3))
        barra.pack(fill="x", side="bottom")
        self.var_estado = tk.StringVar(value="Listo. Revise los datos y presione «Diseñar zapatas».")
        ttk.Label(barra, textvariable=self.var_estado, style="Estado.TLabel").pack(side="left")

    def _panel_izquierdo(self, p):
        # --- Datos de entrada ---------------------------------------------
        f = ttk.LabelFrame(p, text="1. Datos de entrada", style="Seccion.TLabelframe", padding=8)
        f.pack(fill="x", pady=(0, 8))
        ttk.Label(f, text="Archivo Excel").grid(row=0, column=0, sticky="w")
        self.var_excel = tk.StringVar(value=dz.ARCHIVO_EXCEL)
        ttk.Entry(f, textvariable=self.var_excel).grid(row=1, column=0, sticky="ew")
        ttk.Button(f, text="…", width=3, command=self._elegir_excel).grid(row=1, column=1, padx=(4, 0))
        ttk.Label(f, text="Hoja").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.var_hoja = tk.StringVar(value=dz.HOJA_EXCEL)
        self.combo_hoja = ttk.Combobox(f, textvariable=self.var_hoja, state="readonly")
        self.combo_hoja.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.combo_hoja.bind("<<ComboboxSelected>>", lambda e: self._cargar_datos())
        ttk.Label(f, text="Carpeta de resultados").grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.var_salida = tk.StringVar(value=dz.CARPETA_SALIDA)
        ttk.Entry(f, textvariable=self.var_salida).grid(row=5, column=0, sticky="ew")
        ttk.Button(f, text="…", width=3, command=self._elegir_salida).grid(row=5, column=1, padx=(4, 0))
        f.columnconfigure(0, weight=1)

        # --- Parámetros -----------------------------------------------------
        f = ttk.LabelFrame(p, text="2. Parámetros de diseño", style="Seccion.TLabelframe", padding=8)
        f.pack(fill="x", pady=(0, 8))
        self.vars_param = {}
        for i, (attr, etiqueta, unidad, _) in enumerate(PARAMETROS):
            ttk.Label(f, text=etiqueta).grid(row=i, column=0, sticky="w", pady=2)
            var = tk.StringVar()
            ttk.Entry(f, textvariable=var, width=9, justify="right").grid(row=i, column=1, padx=4)
            ttk.Label(f, text=unidad, style="Sub.TLabel").grid(row=i, column=2, sticky="w")
            self.vars_param[attr] = var
        f.columnconfigure(0, weight=1)

        # --- Opciones -------------------------------------------------------
        f = ttk.LabelFrame(p, text="3. Opciones", style="Seccion.TLabelframe", padding=8)
        f.pack(fill="x", pady=(0, 8))
        ttk.Label(f, text="Peralte inicial por zapata").grid(row=0, column=0, sticky="w")
        self.var_h_puntos = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_h_puntos).grid(row=1, column=0, sticky="ew")
        ttk.Label(f, text="Formato zapata:h[cm], ej. 1:40, 2:40", style="Nota.TLabel").grid(
            row=2, column=0, sticky="w")
        self.var_traslado = tk.BooleanVar()
        ttk.Checkbutton(f, text="Trasladar cortantes FX, FY a la base (V·h)",
                        variable=self.var_traslado).grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Label(f, text="Desactivado: las cadenas toman el cortante (criterio de clase)",
                  style="Nota.TLabel", wraplength=self.px(290)).grid(row=4, column=0, sticky="w")
        f.columnconfigure(0, weight=1)

        # --- Acciones -------------------------------------------------------
        self.btn_disenar = ttk.Button(p, text="▶  Diseñar zapatas", style="Accion.TButton",
                                      command=self.ejecutar)
        self.btn_disenar.pack(fill="x", pady=(4, 6))
        self.progreso = ttk.Progressbar(p, mode="determinate")
        self.progreso.pack(fill="x")
        self.var_progreso = tk.StringVar(value="")
        ttk.Label(p, textvariable=self.var_progreso, style="Sub.TLabel").pack(anchor="w", pady=(2, 8))
        fila = ttk.Frame(p)
        fila.pack(fill="x")
        ttk.Button(fila, text="Abrir carpeta", command=self._abrir_carpeta).pack(
            side="left", fill="x", expand=True)
        ttk.Button(fila, text="Abrir Excel", command=self._abrir_excel_resultados).pack(
            side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Button(p, text="Restaurar valores por defecto",
                   command=self._restaurar_parametros).pack(fill="x", pady=(4, 0))

    # ------------------------------------------------------------- pestañas
    def _tabla(self, padre, columnas, anchos):
        marco = ttk.Frame(padre)
        arbol = ttk.Treeview(marco, columns=columnas, show="headings", selectmode="browse")
        for col, ancho in zip(columnas, anchos):
            arbol.heading(col, text=col)
            arbol.column(col, width=self.px(ancho), anchor="center", stretch=True)
        sy = ttk.Scrollbar(marco, orient="vertical", command=arbol.yview)
        sx = ttk.Scrollbar(marco, orient="horizontal", command=arbol.xview)
        arbol.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        arbol.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        return marco, arbol

    def _texto(self, padre):
        marco = ttk.Frame(padre)
        texto = tk.Text(marco, wrap="none", font=("Consolas", 10), relief="flat",
                        background="#FAFAFA", padx=8, pady=6, height=8)
        sy = ttk.Scrollbar(marco, orient="vertical", command=texto.yview)
        sx = ttk.Scrollbar(marco, orient="horizontal", command=texto.xview)
        texto.configure(yscrollcommand=sy.set, xscrollcommand=sx.set, state="disabled")
        texto.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        return marco, texto

    @staticmethod
    def _escribir(texto, contenido):
        texto.configure(state="normal")
        texto.delete("1.0", "end")
        texto.insert("1.0", contenido)
        texto.configure(state="disabled")

    def _pestana_datos(self):
        tab = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab, text="  Datos de entrada  ")
        self.var_info_datos = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.var_info_datos, style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        marco, self.tabla_datos = self._tabla(tab, ("",), (100,))
        marco.pack(fill="both", expand=True)

    def _pestana_resumen(self):
        tab = ttk.Frame(self.notebook, padding=6)
        self.tab_resumen = tab
        self.notebook.add(tab, text="  Resumen  ")
        self.var_info_resumen = tk.StringVar(value="Aún no se ha ejecutado el diseño.")
        ttk.Label(tab, textvariable=self.var_info_resumen, style="Sub.TLabel").pack(anchor="w", pady=(0, 4))
        columnas = ("Zapata", "B [m]", "h [cm]", "Armado X", "Armado Y", "D/C flexión",
                    "D/C cortante 1D", "D/C punzonamiento", "Tipo", "Estado")
        anchos = (60, 90, 60, 150, 150, 90, 110, 120, 130, 90)
        marco, self.tabla_resumen = self._tabla(tab, columnas, anchos)
        marco.pack(fill="both", expand=True)
        self.tabla_resumen.tag_configure("ok", background=COLOR_OK)
        self.tabla_resumen.tag_configure("masa", background=COLOR_MASA)
        self.tabla_resumen.tag_configure("falla", background=COLOR_FALLA)
        self.tabla_resumen.bind("<Double-1>", self._abrir_desde_resumen)
        leyenda = ttk.Frame(tab)
        leyenda.pack(anchor="w", pady=(6, 0))
        for color, texto in ((COLOR_OK, "Cumple"), (COLOR_MASA, "Cumple — hormigón en masa"),
                             (COLOR_FALLA, "No cumple")):
            tk.Label(leyenda, width=2, background=color, relief="solid", borderwidth=1).pack(side="left")
            ttk.Label(leyenda, text=f" {texto}    ", style="Sub.TLabel").pack(side="left")
        ttk.Label(leyenda, text="Doble clic en una fila para ver el detalle de la zapata.",
                  style="Nota.TLabel").pack(side="left", padx=(12, 0))

    def _pestana_zapata(self):
        tab = ttk.Frame(self.notebook, padding=6)
        self.tab_zapata = tab
        self.notebook.add(tab, text="  Detalle por zapata  ")
        barra = ttk.Frame(tab)
        barra.pack(fill="x", pady=(0, 4))
        ttk.Label(barra, text="Zapata:").pack(side="left")
        self.var_zapata = tk.StringVar()
        self.combo_zapata = ttk.Combobox(barra, textvariable=self.var_zapata, state="readonly", width=8)
        self.combo_zapata.pack(side="left", padx=6)
        self.combo_zapata.bind("<<ComboboxSelected>>", lambda e: self._mostrar_zapata())
        self.var_info_zapata = tk.StringVar(value="")
        ttk.Label(barra, textvariable=self.var_info_zapata, style="Sub.TLabel").pack(side="left", padx=10)
        division = ttk.PanedWindow(tab, orient="vertical")
        division.pack(fill="both", expand=True)
        self.marco_figura = ttk.Frame(division)
        marco_texto, self.texto_zapata = self._texto(division)
        division.add(self.marco_figura, weight=3)
        division.add(marco_texto, weight=1)

    def _pestana_dc(self):
        tab = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab, text="  Demanda / Capacidad  ")
        self.marco_dc = tab

    def _pestana_memoria(self):
        tab = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab, text="  Memoria completa  ")
        marco, self.texto_memoria = self._texto(tab)
        marco.pack(fill="both", expand=True)

    # ------------------------------------------------------ datos de entrada
    def _elegir_excel(self):
        ruta = filedialog.askopenfilename(title="Seleccionar archivo de datos",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
                                          initialdir=os.path.dirname(self.var_excel.get()) or CARPETA)
        if ruta:
            ruta = os.path.normpath(ruta)
            self.var_excel.set(ruta)
            self.var_salida.set(os.path.join(os.path.dirname(ruta), "Resultados"))
            self._cargar_hojas(ruta)

    def _elegir_salida(self):
        ruta = filedialog.askdirectory(title="Carpeta de resultados",
                                       initialdir=self.var_salida.get() or CARPETA)
        if ruta:
            self.var_salida.set(os.path.normpath(ruta))

    def _cargar_hojas(self, ruta):
        if not os.path.isfile(ruta):
            self.var_info_datos.set("No se encontró el archivo Excel. Selecciónelo con el botón «…».")
            return
        try:
            libro = openpyxl.load_workbook(ruta, read_only=True)
            hojas = libro.sheetnames
            libro.close()
        except Exception as error:
            messagebox.showerror("Excel", f"No se pudo abrir el archivo:\n{error}")
            return
        self.combo_hoja["values"] = hojas
        if self.var_hoja.get() not in hojas:
            self.var_hoja.set(hojas[0])
        self._cargar_datos()

    def _cargar_datos(self):
        try:
            df = dz.leer_datos(self.var_excel.get(), self.var_hoja.get())
        except Exception as error:
            self.var_info_datos.set(f"No se pudieron leer los datos: {error}")
            return
        arbol = self.tabla_datos
        arbol.delete(*arbol.get_children())
        columnas = list(df.columns)
        arbol["columns"] = columnas
        for col in columnas:
            arbol.heading(col, text=col)
            arbol.column(col, width=self.px(80), anchor="center", stretch=True)
        for _, fila in df.iterrows():
            arbol.insert("", "end", values=[f"{v:g}" if isinstance(v, float) else v for v in fila])
        info = f"{len(df)} filas"
        if {"Point", "Load"} <= set(columnas):
            info += (f"  ·  {df['Point'].nunique()} zapatas  ·  combinaciones: "
                     f"{', '.join(map(str, df['Load'].unique()))}")
        info += "  ·  Unidades: fuerzas [tonf], momentos [tonf·m], dimensiones [m]"
        self.var_info_datos.set(info)

    # ------------------------------------------------------------ parámetros
    def _restaurar_parametros(self):
        for attr, _, _, factor in PARAMETROS:
            self.vars_param[attr].set(f"{self.valores_iniciales[attr] / factor:g}")
        self.var_h_puntos.set(", ".join(f"{k}:{v * 100:g}" for k, v in
                                        self.valores_iniciales["H_MIN_PUNTO"].items()))
        self.var_traslado.set(self.valores_iniciales["TRASLADAR_CORTANTE"])

    def _leer_parametros(self):
        """Lee y valida la ventana. Devuelve {variable del programa: valor}."""
        valores = {}
        for attr, etiqueta, unidad, factor in PARAMETROS:
            texto = self.vars_param[attr].get().strip().replace(",", ".")
            try:
                valor = float(texto)
            except ValueError:
                raise ValueError(f"«{etiqueta}» debe ser un número (en {unidad}).")
            if valor <= 0:
                raise ValueError(f"«{etiqueta}» debe ser mayor que cero.")
            valores[attr] = valor * factor
        if valores["H_MIN"] >= valores["H_MAX"]:
            raise ValueError("El peralte inicial debe ser menor que el peralte máximo.")
        h_puntos = {}
        for parte in self.var_h_puntos.get().replace(";", ",").split(","):
            parte = parte.strip()
            if not parte:
                continue
            try:
                punto, h = parte.split(":")
                h_puntos[int(punto)] = float(h) / 100.0
            except ValueError:
                raise ValueError("«Peralte inicial por zapata» debe tener el formato "
                                 "zapata:h[cm], por ejemplo 1:40, 2:40.")
        valores["H_MIN_PUNTO"] = h_puntos
        valores["TRASLADAR_CORTANTE"] = self.var_traslado.get()
        excel = self.var_excel.get()
        if not os.path.isfile(excel):
            raise ValueError("No se encontró el archivo Excel de entrada.")
        valores["ARCHIVO_EXCEL"] = excel
        valores["HOJA_EXCEL"] = self.var_hoja.get()
        valores["CARPETA_SALIDA"] = self.var_salida.get() or os.path.join(
            os.path.dirname(excel), "Resultados")
        return valores

    # -------------------------------------------------------------- ejecución
    def ejecutar(self):
        if self.ejecutando:
            return
        try:
            valores = self._leer_parametros()
        except ValueError as error:
            messagebox.showerror("Datos no válidos", str(error))
            return
        for attr, valor in valores.items():
            setattr(dz, attr, valor)
        self.ejecutando = True
        self.btn_disenar.state(["disabled"])
        self.progreso.configure(value=0, maximum=1)
        self.var_progreso.set("Iniciando el diseño…")
        self.var_estado.set("Diseñando zapatas…")
        threading.Thread(target=self._trabajo, daemon=True).start()
        self.root.after(100, self._revisar_cola)

    def _trabajo(self):
        """Corre en segundo plano para que la ventana no se congele."""
        try:
            res = dz.main(progreso=lambda k, n, p: self.cola.put(("progreso", k, n, p)))
            self.cola.put(("fin", res))
        except PermissionError as error:
            self.cola.put(("error", "No se pudieron guardar los resultados. Si el archivo "
                                    "Resultados_Zapatas.xlsx está abierto en Excel, ciérrelo y "
                                    f"vuelva a intentarlo.\n\n{error}"))
        except Exception:
            self.cola.put(("error", traceback.format_exc()))

    def _revisar_cola(self):
        try:
            while True:
                mensaje = self.cola.get_nowait()
                if mensaje[0] == "progreso":
                    _, k, n, punto = mensaje
                    self.progreso.configure(value=k, maximum=n)
                    self.var_progreso.set(f"Zapata P{punto} diseñada ({k} de {n})")
                elif mensaje[0] == "fin":
                    self._terminar()
                    self._mostrar_resultados(mensaje[1])
                else:
                    self._terminar()
                    self.var_estado.set("El diseño se detuvo por un error.")
                    messagebox.showerror("Error durante el diseño", mensaje[1])
        except queue.Empty:
            pass
        if self.ejecutando:
            self.root.after(100, self._revisar_cola)

    def _terminar(self):
        self.ejecutando = False
        self.btn_disenar.state(["!disabled"])

    # ------------------------------------------------------------ resultados
    def _mostrar_resultados(self, resultados):
        self.resultados = resultados
        arbol = self.tabla_resumen
        arbol.delete(*arbol.get_children())
        n_ok = n_masa = 0
        for r in resultados:
            if "bar_x" not in r:
                arbol.insert("", "end", iid=str(r["punto"]), tags=("falla",),
                             values=(f"P{r['punto']}", "", "", "", "", "", "", "", "",
                                     f"NO DISEÑADA: {r['motivo']}"))
                continue
            g = r["geo"]
            masa = r["bloque"]
            n_ok += r["cumple"]
            n_masa += masa
            etiqueta = "falla" if not r["cumple"] else ("masa" if masa else "ok")
            arbol.insert("", "end", iid=str(r["punto"]), tags=(etiqueta,), values=(
                f"P{r['punto']}", f"{g['Bx']:.2f} × {g['By']:.2f}", f"{r['h'] * 100:.0f}",
                dz.texto_armado(r["bar_x"]), dz.texto_armado(r["bar_y"]),
                formato_dc(max(r["DC_flex_x"], r["DC_flex_y"])),
                formato_dc(maximo_valido(r["DC_v1_x"], r["DC_v1_y"])),
                formato_dc(r["DC_punz"]),
                "Hormigón en masa" if masa else "Zapata flexible",
                "CUMPLE" if r["cumple"] else "NO CUMPLE"))
        total = len(resultados)
        self.var_info_resumen.set(
            f"{total} zapatas diseñadas  ·  {n_ok} cumplen  ·  {n_masa} trabajan como hormigón en "
            f"masa  ·  f'c = {dz.FC:g} kg/cm², fy = {dz.FY:g} kg/cm², rec = {dz.REC * 100:g} cm")

        disenadas = [r for r in resultados if "bar_x" in r]
        self.combo_zapata["values"] = [f"P{r['punto']}" for r in disenadas]
        if disenadas:
            self.var_zapata.set(f"P{disenadas[0]['punto']}")
            self._mostrar_zapata()
            self._mostrar_dc(disenadas)
        self._escribir(self.texto_memoria, "\n".join(dz.MEMORIA))
        self.notebook.select(self.tab_resumen)
        self.var_progreso.set(f"Diseño terminado: {total} zapatas")
        self.var_estado.set(f"Diseño terminado. Resultados guardados en: {dz.CARPETA_SALIDA}")

    def _incrustar(self, figura, marco, con_barra=True):
        """Dibuja una figura de matplotlib dentro de un marco de la ventana."""
        for hijo in marco.winfo_children():
            hijo.destroy()
        figura.set_layout_engine("tight")
        # Tamaño inicial pequeño: el lienzo se expande al espacio disponible. Con
        # el tamaño original (13 × 6 in) Tk reservaría más espacio del que hay.
        figura.set_size_inches(4, 2)
        lienzo = FigureCanvasTkAgg(figura, master=marco)
        # Aplica ya el escalado de pantalla (Windows al 125 %, 150 %…). Si se deja
        # para cuando el lienzo aparece, la figura queda más grande que su espacio.
        if hasattr(lienzo, "_update_device_pixel_ratio"):
            lienzo._update_device_pixel_ratio()
        if con_barra:
            barra = NavigationToolbar2Tk(lienzo, marco, pack_toolbar=False)
            barra.update()
            barra.pack(side="bottom", fill="x")
        lienzo.get_tk_widget().pack(fill="both", expand=True)
        lienzo.draw()
        return lienzo

    def _mostrar_zapata(self):
        nombre = self.var_zapata.get()
        r = next((r for r in self.resultados if f"P{r['punto']}" == nombre), None)
        if r is None:
            return
        self.canvas_zapata = self._incrustar(dz.figura_zapata(r), self.marco_figura)
        self._escribir(self.texto_zapata, r.get("memoria", ""))
        g = r["geo"]
        self.var_info_zapata.set(
            f"{g['Bx']:.2f} × {g['By']:.2f} m  ·  columna {g['cx'] * 100:.0f} × {g['cy'] * 100:.0f} cm  ·  "
            f"h = {r['h'] * 100:.0f} cm  ·  {'hormigón en masa' if r['bloque'] else 'zapata flexible'}"
            f"  ·  {'CUMPLE' if r['cumple'] else 'NO CUMPLE'}")

    def _mostrar_dc(self, disenadas):
        self.canvas_resumen = self._incrustar(dz.figura_resumen(disenadas), self.marco_dc)

    def _abrir_desde_resumen(self, _evento=None):
        seleccion = self.tabla_resumen.selection()
        if not seleccion:
            return
        nombre = f"P{seleccion[0]}"
        if nombre in self.combo_zapata["values"]:
            self.var_zapata.set(nombre)
            self._mostrar_zapata()
            self.notebook.select(self.tab_zapata)

    # --------------------------------------------------------------- archivos
    def _abrir_carpeta(self):
        carpeta = self.var_salida.get()
        if os.path.isdir(carpeta):
            os.startfile(carpeta)
        else:
            messagebox.showinfo("Resultados", "La carpeta de resultados todavía no existe. "
                                              "Ejecute primero el diseño.")

    def _abrir_excel_resultados(self):
        ruta = os.path.join(self.var_salida.get(), "Resultados_Zapatas.xlsx")
        if os.path.isfile(ruta):
            os.startfile(ruta)
        else:
            messagebox.showinfo("Resultados", "Todavía no hay un Excel de resultados. "
                                              "Ejecute primero el diseño.")


def iniciar():
    # Ventana nítida en pantallas con escalado (Windows)
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    app = AplicacionZapatas(root)
    root.mainloop()
    return app


if __name__ == "__main__":
    iniciar()
