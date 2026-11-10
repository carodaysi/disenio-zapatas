# Diseño de zapatas aisladas de hormigón armado — ACI 318-19

Procedimiento computacional reproducible, escrito en Python, para el diseño estructural de
zapatas aisladas de hormigón armado. Verifica **flexión**, **cortante unidireccional** y
**cortante bidireccional (punzonamiento)** según el ACI 318-19 y los criterios desarrollados en
clase. Puede usarse con una **interfaz gráfica** o desde la consola.

> Maestría en Construcción de Estructuras Sismorresistentes — Universidad Politécnica Salesiana  
> Asignatura: Diseño de Cimentaciones

![Interfaz gráfica: resumen del diseño](docs/img/interfaz_resumen.png)

---

## Contenido

- [Características](#características)
- [Estructura del repositorio](#estructura-del-repositorio)
- [Instalación](#instalación)
- [Uso](#uso)
- [Datos de entrada](#datos-de-entrada)
- [Parámetros de diseño](#parámetros-de-diseño)
- [Resultados que genera](#resultados-que-genera)
- [Metodología](#metodología)
- [Resultados del proyecto](#resultados-del-proyecto)
- [Validación](#validación)
- [Hipótesis y limitaciones](#hipótesis-y-limitaciones)

---

## Características

- Lee las reacciones de ETABS desde Excel (13 zapatas × 6 combinaciones mayoradas).
- Calcula la presión de contacto **sin admitir tracciones en el suelo**, incluso con contacto
  parcial biaxial.
- Busca automáticamente el peralte `h` y el diámetro de barra que cumplen todas las verificaciones.
- Detecta cuándo la sección de cortante queda fuera del voladizo (`d ≥ voladizo`): la zapata
  trabaja como **hormigón en masa** y no se verifica el cortante.
- Revisa además el anclaje del refuerzo, el aplastamiento columna-zapata, el acero superior por
  pérdida de contacto y la distribución del acero en zapatas rectangulares.
- Genera una memoria de cálculo paso a paso, un Excel de resultados y las figuras de cada zapata.
- Incluye una interfaz gráfica (Tkinter) que valida los datos y permite revisar cada zapata.

## Estructura del repositorio

```
├── Diseno_Zapatas_Aisladas.py         # Motor de cálculo (también se ejecuta solo, por consola)
├── Interfaz_Zapatas.py                # Interfaz gráfica (Tkinter)
├── Ejecutar_Interfaz.bat              # Acceso directo para abrir la interfaz en Windows
├── InputData01.xlsx                   # Datos de entrada (hoja "Tarea1")
├── Explicacion_Programa_Zapatas.docx  # Documento que explica el programa paso a paso
├── requirements.txt                   # Dependencias (pip)
├── environment.yml                    # Entorno conda
├── docs/img/                          # Capturas usadas en este README
└── Resultados/                        # Salidas generadas por el programa
```

## Instalación

Se necesita **Python 3.10** (probado con 3.10.22) y las librerías `numpy`, `pandas`, `matplotlib`
y `openpyxl`. La interfaz usa `tkinter`, que viene incluido con Python.

**Opción A — conda (recomendada, igual que en el curso):**

```bash
git clone https://github.com/carodaysi/disenio-zapatas.git
cd disenio-zapatas
conda env create -f environment.yml
conda activate zapatas
```

**Opción B — pip:**

```bash
git clone https://github.com/carodaysi/disenio-zapatas.git
cd disenio-zapatas
pip install -r requirements.txt
```

Si ya tienes el entorno del curso (`spyder-env`), solo falta `openpyxl`, que pandas necesita para
leer archivos `.xlsx`:

```bash
conda activate spyder-env
pip install openpyxl
```

## Uso

### Interfaz gráfica

```bash
python Interfaz_Zapatas.py
```

En Windows también se abre con doble clic en `Ejecutar_Interfaz.bat`. Ese archivo apunta a
`D:\Programas_inst\miniconda\envs\spyder-env\pythonw.exe`; en otra computadora hay que cambiar
esa ruta por la del `pythonw.exe` de su entorno.

1. **Datos de entrada:** elige el Excel, la hoja y la carpeta de resultados.
2. **Parámetros de diseño:** revisa f'c, fy, recubrimiento, peraltes, separación de barras, etc.
3. **Opciones:** peralte inicial por zapata (por ejemplo `1:40, 2:40`) y traslado del cortante.
4. Presiona **Diseñar zapatas**. Los resultados aparecen en las pestañas *Resumen*, *Detalle por
   zapata*, *Demanda / Capacidad* y *Memoria completa*. Doble clic en una fila del resumen abre
   el detalle de esa zapata.

![Interfaz gráfica: detalle de la zapata P2](docs/img/interfaz_detalle_P2.png)

### Consola o Spyder

```bash
python Diseno_Zapatas_Aisladas.py
```

En Spyder: abre `Diseno_Zapatas_Aisladas.py` y presiona **F5**. Los parámetros se cambian en la
**sección 0** del archivo. El programa imprime la memoria de cálculo y guarda los resultados en
`Resultados/`.

> Si `Resultados_Zapatas.xlsx` está abierto en Excel, ciérralo antes de ejecutar; de lo contrario
> no se puede sobrescribir.

## Datos de entrada

Hoja `Tarea1` de `InputData01.xlsx`: una fila por zapata y combinación.

| Columna | Significado | Unidad |
|---|---|---|
| `Story` | Nivel (BASE) | — |
| `Point` | Número de la zapata / columna | — |
| `Load` | Combinación mayorada (DCON1 … DCON6) | — |
| `FX`, `FY` | Cortantes en la base de la columna | tonf |
| `FZ` | Carga axial (positiva = compresión) | tonf |
| `MX`, `MY` | Momentos alrededor de X e Y | tonf·m |
| `MZ` | Torsión (no interviene en el diseño) | tonf·m |
| `Bx[m]`, `By[m]` | Dimensiones en planta de la zapata | m |
| `Bcx[m]`, `Bcy[m]` | Dimensiones de la columna | m |

Las dimensiones en planta vienen del dimensionamiento previo por presión admisible. Este programa
hace el diseño **por resistencia** con cargas mayoradas.

## Parámetros de diseño

Están al inicio de `Diseno_Zapatas_Aisladas.py` (sección 0) y se pueden editar en la interfaz.

| Parámetro | Variable | Valor | Referencia |
|---|---|---|---|
| Resistencia del hormigón f'c | `FC` | 210 kg/cm² | — |
| Fluencia del acero fy | `FY` | 4200 kg/cm² | — |
| Recubrimiento (contra el suelo) | `REC` | 7.5 cm | ACI 20.5.1.3.1 |
| φ flexión / cortante / aplastamiento | `PHI_*` | 0.90 / 0.75 / 0.65 | ACI 21.2.1 |
| Factor de tamaño λs | `LAMBDA_S` | 1.0 | ACI 13.2.6.2 |
| αs (columna interior) | `ALPHA_S` | 40 | ACI 22.6.5.3 |
| Cuantía mínima | `RHO_MIN` | 0.0018 | ACI 8.6.1.1 |
| Peralte efectivo mínimo | `D_MIN` | 15 cm | ACI 13.3.1.2 |
| Peralte inicial / máximo / incremento | `H_MIN` / `H_MAX` / `DH` | 30 / 150 / 5 cm | Criterio práctico |
| Peralte inicial por zapata | `H_MIN_PUNTO` | P1 y P2: 40 cm | Hormigón en masa |
| Separación mínima entre barras | `S_MIN` | 10 cm | Constructibilidad |
| Espaciamiento libre mínimo | `S_LIBRE_MIN` | 2.5 cm | ACI 25.2.1 |
| Diámetros, en orden de preferencia | `DIAMETROS` | 12 … 25 mm, 10 mm al final | — |
| Traslado del cortante a la base (V·h) | `TRASLADAR_CORTANTE` | `False` | Las cadenas toman el cortante |

Unidades internas: fuerzas en tonf, longitudes en m y presiones en tonf/m². Las resistencias usan
las expresiones mks del ACI 318-19 (1 kg/cm² = 10 tonf/m²).

## Resultados que genera

En la carpeta `Resultados/`:

| Archivo | Contenido |
|---|---|
| `Resultados_Zapatas.xlsx` | Hojas *Resumen* (una fila por zapata), *Detalle_Combos* (solicitaciones y D/C por combinación) y *Parametros* |
| `Memoria_Calculo.txt` | Memoria de cálculo paso a paso de cada zapata |
| `Zapata_P1.png` … `Zapata_P13.png` | Planta con presiones y secciones críticas, y corte con el armado |
| `Resumen_Demanda_Capacidad.png` | Relación demanda/capacidad de todas las verificaciones |

## Metodología

Para cada zapata y cada combinación:

1. **Presión de contacto.** Se supone una zapata rígida:
   `qu = P/A + My·x/Iyy + Mx·y/Ixx`. Si aparecen tracciones, se resuelve el contacto parcial por
   Newton-Raphson sobre una malla de 200 × 200 celdas, con `q = máx(plano, 0)` y las tres
   ecuaciones de equilibrio (ΣF, ΣMx, ΣMy).
2. **Peralte y peralte efectivo.** Se prueba `h` desde 30 cm en pasos de 5 cm. La dirección con
   mayor momento va en la capa inferior:
   `d = h − rec − db/2` (capa inferior) y `d = h − rec − 1.5·db` (capa superior, el caso más
   desfavorable). Cada dirección se revisa con su propio `d`.
3. **Hormigón en masa.** Si `d ≥ (B − c)/2` en una dirección, la sección crítica de cortante queda
   fuera de la zapata y no hay cortante en esa dirección. Si ocurre en ambas, tampoco hay
   punzonamiento.
4. **Flexión** en la cara de la columna (ACI 13.2.7.1): `Mu = ∫ qu·(x − c/2) dA`,
   `As = máx(As,req; 0.0018·b·h)`, con separación ≤ mín(2h; 45 cm) (ACI 8.7.2.2) y control de εt.
5. **Cortante unidireccional** a `d` de la cara (ACI 8.4.3.2), sin estribos (ACI 22.5.5.1c):
   `vc = 2.12·λs·λ·ρw^(1/3)·√f'c ≤ 1.33·λ·√f'c`, con `φVc ≥ Vu`.
6. **Punzonamiento** en el perímetro a `d/2` (ACI 22.6.4) con transferencia de momento
   (ACI 8.4.4.2):
   `vu = Vu/(bo·d) + γv·Msc·c/Jc ≤ φ·mín{1.06; 0.53(1 + 2/β); 0.27(2 + αs·d/bo)}·√f'c`.
7. **Verificaciones complementarias:** anclaje con barra recta o gancho de 90° (ACI 25.4),
   aplastamiento (ACI 22.8), zapatas rectangulares (ACI 13.3.3.3) y acero superior por pérdida
   de contacto (ACI 18.13.2.5).
8. Si alguna verificación no cumple, se aumenta `h` y se repite.

La explicación detallada, con fórmulas, diagrama de flujo y referencias, está en
[`Explicacion_Programa_Zapatas.docx`](Explicacion_Programa_Zapatas.docx).

## Resultados del proyecto

Las 13 zapatas cumplen. Columna de 60 × 50 cm, f'c = 210 kg/cm², fy = 4200 kg/cm².

| Zapata | B [m] | h [cm] | Armado (X e Y) | D/C flexión | D/C cortante 1D | D/C punzonamiento | Tipo |
|---|---|---|---|---|---|---|---|
| P1 | 1.10 | 40 | 11 Ø10 mm @ 9 cm | 0.110 | No aplica | No aplica | Hormigón en masa |
| P2 | 1.10 | 40 | 11 Ø10 mm @ 9 cm | 0.121 | No aplica | No aplica | Hormigón en masa |
| P3 | 2.60 | 50 | 21 Ø12 mm @ 12 cm | 0.703 | 0.907 | 0.572 | Flexible |
| P4 | 2.30 | 55 | 21 Ø12 mm @ 10 cm | 0.601 | 0.811 | 0.529 | Flexible |
| P5 | 2.10 | 50 | 17 Ø12 mm @ 12 cm | 0.623 | 0.838 | 0.519 | Flexible |
| P6 | 2.60 | 60 | 25 Ø12 mm @ 10 cm | 0.677 | 0.887 | 0.578 | Flexible |
| P7 | 2.85 | 70 | 24 Ø14 mm @ 11 cm | 0.684 | 0.894 | 0.610 | Flexible |
| P8 | 2.50 | 60 | 24 Ø12 mm @ 10 cm | 0.673 | 0.882 | 0.568 | Flexible |
| P9 | 2.50 | 55 | 22 Ø12 mm @ 11 cm | 0.757 | 0.997 | 0.637 | Flexible |
| P10 | 2.85 | 65 | 22 Ø14 mm @ 12 cm | 0.711 | 0.930 | 0.619 | Flexible |
| P11 | 2.20 | 50 | 18 Ø12 mm @ 11 cm | 0.715 | 0.971 | 0.598 | Flexible |
| P12 | 2.50 | 60 | 24 Ø12 mm @ 10 cm | 0.648 | 0.847 | 0.554 | Flexible |
| P13 | 2.20 | 50 | 18 Ø12 mm @ 11 cm | 0.723 | 0.982 | 0.605 | Flexible |

Observaciones:

- El **cortante unidireccional controla el peralte**: con cuantías bajas (ρw ≈ 0.002),
  `vc ≈ 3.9 kg/cm²`, casi la mitad de `0.53·√f'c`.
- En todas las zapatas el acero de flexión lo define la **cuantía mínima** 0.0018·b·h.
- **P1 y P2** tienen voladizos de 25 y 30 cm. Con h = 40 cm, `d ≥ voladizo` y trabajan como
  hormigón en masa. Usan barras Ø10 con gancho de 90° para anclarse y llevan acero superior
  porque pierden contacto parcial en las combinaciones sísmicas.

## Validación

Se verificó a mano la zapata P7 (2.85 × 2.85 m, combinación DCON4, Pu = 237.52 t):

| Cálculo | A mano | Programa |
|---|---|---|
| qu,máx | 29.57 t/m² | 29.57 t/m² |
| Mu,x en la cara | 53.16 t·m | 53.17 t·m |
| As,req,x | 23.70 cm² | 23.67 cm² |
| φVc,y (cortante 1D) | 51.9 t | 51.95 t |
| Vu de punzonamiento | 198.18 t | 198.18 t |
| vc de punzonamiento | 15.36 kg/cm² | 15.36 kg/cm² |

![Interfaz gráfica: detalle de la zapata P7](docs/img/interfaz_detalle_P7.png)

## Hipótesis y limitaciones

- Zapata rígida, con distribución lineal de presiones y sin tracciones en el suelo.
- Columna centrada (zapata interior). Para zapatas medianeras o esquineras hay que cambiar αs
  (30 o 20) y la geometría del perímetro crítico.
- Los cortantes FX y FY los toman las cadenas (`TRASLADAR_CORTANTE = False`). Con `True` se
  trasladan a la base como un momento adicional V·h.
- La presión admisible del suelo y los asentamientos se verifican en el dimensionamiento con
  cargas de servicio, que queda fuera de este programa.
- Se supone que el hormigón de la columna tiene el mismo f'c que la zapata.
