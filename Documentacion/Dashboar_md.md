# Documentación técnica — Dashboard Marketing KPI

Referencia completa del sistema: dependencias, ejecutable, arquitectura, formato de entrada,
validaciones, salida y mantenimiento.

> Este documento es el **manual técnico**. Para el uso diario (copiar el Excel, doble clic,
> compartir el HTML) el documento corto es [README.md](README.md); aquí está todo lo demás.

**Versión documentada:** `dashboard_kpi.py` de 3.358 líneas (18/08/2026 16:43) · 63 funciones · 3 clases
**Fecha de la documentación:** 21 de agosto de 2026

---

## Índice

| § | Sección |
|---|---|
| 1 | [Qué es y qué hace](#1-qué-es-y-qué-hace) |
| 2 | [Requisitos del sistema](#2-requisitos-del-sistema) |
| 3 | [Librerías y dependencias](#3-librerías-y-dependencias) |
| 4 | [El ejecutable](#4-el-ejecutable) |
| 5 | [Formas de ejecución](#5-formas-de-ejecución) |
| 6 | [Estructura de archivos](#6-estructura-de-archivos) |
| 7 | [Arquitectura: el pipeline](#7-arquitectura-el-pipeline) |
| 8 | [Entrada: el Excel esperado](#8-entrada-el-excel-esperado) |
| 9 | [Configuración](#9-configuración) |
| 10 | [Sistema de validación](#10-sistema-de-validación) |
| 11 | [Salida: anatomía del dashboard](#11-salida-anatomía-del-dashboard) |
| 12 | [Referencia del código](#12-referencia-del-código) |
| 13 | [Capa de diseño (CSS/JS)](#13-capa-de-diseño-cssjs) |
| 14 | [Distribución y despliegue](#14-distribución-y-despliegue) |
| 15 | [Solución de problemas](#15-solución-de-problemas) |
| 16 | [Rarezas conocidas de los datos](#16-rarezas-conocidas-de-los-datos) |
| 17 | [Notas de mantenimiento](#17-notas-de-mantenimiento) |

---

## 1. Qué es y qué hace

Un script de Python que lee el Excel mensual de KPIs del departamento de ventas y genera
`dashboard.html`: un tablero **autocontenido** que se abre con doble clic, sin internet, sin
Python y sin librerías instaladas.

**Entrada** → `Entrada\*.xlsx` (el más reciente por fecha de modificación)
**Salida** → `dashboard.html` (~4,8 MB, un solo archivo)
**Duración** → 1–2 segundos

Tres garantías que definen el diseño del sistema:

1. **El Excel de entrada nunca se modifica.** Se abre en modo solo lectura (`read_only=True`).
2. **Si los datos no cuadran, no se genera nada.** El script aborta y explica qué falla; el
   `dashboard.html` anterior queda intacto y sigue siendo válido.
3. **El HTML resultante no depende de nada externo.** Plotly va incrustado en línea, el CSS y el
   JS son propios, y no se usan fuentes web.

> La filosofía está escrita en el propio código: *un dashboard con números equivocados y buena
> pinta es peor que uno que no se genera.*

---

## 2. Requisitos del sistema

| Requisito | Detalle |
|---|---|
| **Sistema operativo** | Windows (por el `.bat` y por `os.startfile`). El núcleo Python es multiplataforma; solo la apertura automática del navegador es específica de Windows y está protegida con `hasattr(os, "startfile")`. |
| **Python** | Probado en **3.13.14**. Mínimo real **3.8+** (lo exige openpyxl 3.1.5). El código no usa sintaxis posterior a 3.6 —sin operador morsa ni `match`—, así que la restricción viene de la librería, no del script. |
| **PATH** | Durante la instalación de Python hay que marcar **«Add Python to PATH»**, o el `.bat` no lo encontrará. |
| **Permisos** | Escritura en `CARPETA_SALIDA` (por defecto, la carpeta del propio script). |
| **Internet** | Solo la primera vez, para instalar las librerías. Después, nunca. |
| **Navegador** | Cualquiera moderno. El panel de Configuración usa `showSaveFilePicker` si existe y cae a descarga normal si no. |

**Para el destinatario del `dashboard.html`: ningún requisito.** Ni Python, ni librerías, ni
internet, ni VS Code. Solo un navegador.

---

## 3. Librerías y dependencias

### 3.1 Dependencias externas

| Librería | Versión instalada | ¿Se usa? | Para qué |
|---|---|---|---|
| **openpyxl** | 3.1.5 | Sí, siempre | Único lector del Excel. `load_workbook(ruta, data_only=True, read_only=True)`. Importada arriba del todo ([dashboard_kpi.py:22](dashboard_kpi.py#L22)). |
| **plotly** | 6.9.0 | Sí, al renderizar | Los tres gráficos interactivos. **Import perezoso** dentro de `construir_html` ([dashboard_kpi.py:2561-2562](dashboard_kpi.py#L2561-L2562)): `plotly.graph_objects` y `plotly.offline.plot`. |
| **pandas** | 3.0.5 | **No** | Cero referencias en el código. Ver la nota de abajo. |

**Instalación:**

```
pip install pandas openpyxl plotly
```

> **Nota sobre pandas.** El script **no lo importa ni lo usa** en ninguna línea, pero
> `Generar_dashboard.bat` lo exige en su comprobación (`python -c "import pandas, openpyxl,
> plotly"`) y lo instala. Consecuencia práctica: si pandas falta, el `.bat` cree que faltan todas
> las librerías y relanza el `pip install` completo. No rompe nada —el script funcionaría igual
> sin pandas—, pero es una dependencia declarada de más. Ver [§17](#17-notas-de-mantenimiento).

### 3.2 Librería estándar

Sin instalación, viene con Python:

| Módulo | Uso en el script |
|---|---|
| `base64` | Incrustar el logo en el HTML como data URI |
| `glob` | Buscar los `.xlsx` de `Entrada` |
| `json` | Leer y validar `configuracion.json` |
| `math` | Escalas de los gráficos y regresión de la tendencia |
| `os` | Rutas, creación de carpetas, `os.startfile` |
| `re` | Normalización de texto |
| `sys` | Salida de error y código de retorno |
| `unicodedata` | Quitar acentos al comparar nombres de cola |
| `collections.defaultdict` | Agregación de llamadas por (fecha, cola) |
| `datetime` (`date`, `datetime`, `timedelta`) | Fechas de corte, serie diaria, sello de generación |

### 3.3 Dependencias del HTML generado

**Ninguna en tiempo de ejecución.** Todo viaja dentro del archivo:

- **plotly.js** incrustado con `include_plotlyjs="inline"` en el primer gráfico; los otros dos usan
  `include_plotlyjs=False` para no repetirlo. Es el grueso de los ~4,8 MB.
- **CSS propio** (`CSS_DASHBOARD`), sin frameworks ni hojas externas.
- **JS propio** (`JS_DASHBOARD`), sin librerías.
- **Fuentes del sistema** (Segoe UI en Windows). Nada de Google Fonts: el archivo tiene que
  abrirse sin internet.
- **Iconos y logo** como SVG en línea o base64.

---

## 4. El ejecutable

**No hay `.exe`.** El ejecutable es `Generar_dashboard.bat`, un lanzador de 84 líneas que se abre
con doble clic. No pide escribir nada.

### 4.1 Qué hace, paso a paso

| Paso | Acción | Si falla |
|---|---|---|
| 0 | `cd /d "%~dp0"` — se sitúa en su propia carpeta, funcione desde donde funcione | — |
| 1 | `python --version` — comprueba que Python existe | → `:sin_python`, código **1** |
| 2 | `python -c "import pandas, openpyxl, plotly"` — comprueba las librerías; si faltan, `pip install` | → `:fallo_pip`, código **1** |
| 3 | `python dashboard_kpi.py` — genera el dashboard | → `:fallo_script`, código **1** |
| 4 | Espera 6 s con `ping -n 7 127.0.0.1` y cierra | código **0** |

### 4.2 Detalles deliberados

- **El `.bat` no abre el navegador.** Lo abre el propio script, porque solo él conoce la ruta
  final: si se edita `CARPETA_SALIDA`, el `.bat` no sabría dónde buscar el archivo.
- **`ping` en lugar de `timeout`.** `timeout` falla en entornos con la entrada redirigida; el
  `ping` a loopback es la pausa portable.
- **El `.bat` no contiene lógica.** Cambiar el script no obliga a tocarlo nunca.

### 4.3 Los tres mensajes de error

| Etiqueta | Causa | Mensaje al usuario |
|---|---|---|
| `:sin_python` | Python no está o no está en el PATH | Enlace de descarga + recordatorio de marcar «Add Python to PATH» |
| `:fallo_pip` | `pip install` no pudo completarse | Apunta al proxy o al antivirus corporativo y da el comando manual |
| `:fallo_script` | El script abortó (datos inválidos) | Los tres casos habituales + recuerda que el dashboard anterior sigue siendo válido |

En los tres casos la ventana se queda abierta con `pause` para poder leer el mensaje.

---

## 5. Formas de ejecución

| Vía | Cómo | Cuándo |
|---|---|---|
| **Doble clic** | `Generar_dashboard.bat` | Uso normal. Comprueba e instala dependencias. |
| **VS Code** | Abrir `dashboard_kpi.py` y pulsar ▶ Run | Desarrollo. Hace exactamente lo mismo, sin la comprobación de librerías. |
| **Terminal** | `python dashboard_kpi.py` | Automatización, servidores, tareas programadas. |

No acepta argumentos de línea de comandos ni interacción. Toda la configuración vive en el bloque
`CONFIG` del script y en `configuracion.json`.

**Salida por consola** en una ejecución correcta:

```
OK. Dashboard generado en: C:\dashboard\dashboard.html
Archivo de origen: C:\dashboard\Entrada\KPI Sales Department August.xlsx
Corte: 2026-08-12 | Dias habiles: 10
Configuracion aplicada desde: C:\dashboard\configuracion.json   (si existe)
  Nombres cambiados: 3
  Campanas con colas definidas ahi: 2

Avisos:
  - ...
```

**Códigos de retorno:** `0` correcto · `1` `ErrorDatosExcel` (mensaje explicativo en stderr).

Para uso desatendido, poner `ABRIR_AL_TERMINAR = False` para que no se abra el navegador.

---

## 6. Estructura de archivos

```
c:\dashboard\
├── Entrada\                          <- el Excel del mes va aquí
│   ├── KPI Sales Department June.xlsx
│   ├── KPI Sales Department July.xlsx
│   └── KPI Sales Department August.xlsx
├── dashboard_kpi.py                  <- el script (151 KB, 3.358 líneas)
├── Generar_dashboard.bat             <- el ejecutable (2,4 KB)
├── dashboard.html                    <- el resultado (4,8 MB, se sobrescribe)
├── configuracion.json                <- opcional, lo escribe el panel Configuración
├── README.md                         <- guía de uso diario
├── DOCUMENTACION.md                  <- este documento
├── Generar_dashboard.zip             <- paquete para repartir (ver §17)
├── imagenes de referencia\           <- capturas de diseño, no las usa el script
└── __pycache__\                      <- caché de Python, regenerable
```

### Qué es imprescindible

| Archivo | ¿Necesario? |
|---|---|
| `dashboard_kpi.py` | **Sí.** Es todo el programa. |
| `Entrada\` con un `.xlsx` | **Sí.** Sin Excel no hay nada que leer. |
| `Generar_dashboard.bat` | Conveniencia. Se puede ejecutar el `.py` directamente. |
| `configuracion.json` | Opcional. Si no existe, manda el código. |
| `dashboard.html` | Es la salida, no una entrada. Se regenera siempre. |
| `__pycache__\` | Borrable. Se regenera solo. |
| `imagenes de referencia\` | Solo documentación visual. |

**La carpeta se puede mover entera.** Todas las rutas cuelgan de
`CARPETA_BASE = os.path.dirname(os.path.abspath(__file__))`
([dashboard_kpi.py:31](dashboard_kpi.py#L31)). No hay ni una ruta escrita a mano en el resto del
código. Lo único que debe mantenerse es que `Entrada` esté **al lado** del `.py`.

### Selección del archivo de entrada

`archivo_mas_reciente()` toma el `.xlsx` con la **fecha de modificación más reciente**, ignorando
los temporales de Excel (`~$...`) — por eso se puede tener el Excel abierto mientras se ejecuta.

> **Cuidado a principios de mes:** si nadie ha pegado los datos nuevos, el script tomará el archivo
> del mes pasado sin quejarse. Por eso la cabecera del dashboard muestra siempre **nombre del
> archivo, fecha de corte y hora de generación**.

---

## 7. Arquitectura: el pipeline

Un solo proceso, sin estado intermedio en disco, orquestado por `main()`
([dashboard_kpi.py:3257](dashboard_kpi.py#L3257)).

```
configuracion.json  ─┐
                     ├─→ [0] cargar_configuracion()   fusiona colas y nombres
Entrada\*.xlsx     ──┘        ↓
                        [1] EXTRACCIÓN        una sola apertura, read_only
                              ├─ extraer_print()                 hoja Print
                              ├─ extraer_llamadas_diarias()      RawDataRingCentral (+auxiliares)
                              ├─ extraer_bloques_raw_data_general()  Raw Data General
                              └─ extraer_costo_corte()           Valores$$$
                                    ↓
                        [2] EMPAREJAMIENTO
                              └─ asignar_bloques_a_campanas()    bloque ↔ campaña, 1 a 1
                                    ↓
                        [3] SERIE DIARIA
                              └─ construir_serie_diaria()        mes completo, día a día
                                    ↓
                        [4] VALIDACIÓN        ← aborta aquí si algo no cuadra
                              └─ validar()
                                    ↓
                        [5] RENDER
                              └─ construir_html()                CSS + JS + 3 gráficos Plotly
                                    ↓
                              dashboard.html
```

### Decisiones de rendimiento

**El libro se abre una sola vez y en `read_only`.** Es lo que hace que tarde ~1,5 s en vez de
~16 s: en modo normal openpyxl carga también `RawDataRingCentral HM2-IA`, que declara más de un
millón de filas aunque no se use.

**Las hojas pequeñas se vuelcan a memoria.** `read_only` no permite acceso por coordenadas, así que
`hoja_a_memoria()` copia las hojas chicas a un `HojaEnMemoria` —una clase mínima con la misma
interfaz que openpyxl (`cell`, `iter_rows`, `max_row`, `max_column`)— con tope de 2.000 filas y
corte tras 50 filas vacías seguidas.

**Las hojas grandes se leen en streaming.** `_leer_hoja_llamadas()` itera con `values_only=True` y
corta tras 500 filas vacías consecutivas, agregando sobre la marcha en un `defaultdict`.

### Principio de lectura

> **Las tablas se localizan buscando etiquetas de texto, nunca por coordenadas fijas.**

Vale también para las columnas dentro de cada bloque de `Raw Data General`: en agosto 2026 venía
primero `Count of WO Sale Date` y en julio 2026 primero `Sum of V-linea`. Leerlas por posición
intercambiaba clientes con líneas sin que saltara ningún error.

La comparación de etiquetas pasa siempre por `_norm()`: sin acentos, minúsculas, espacios
colapsados. Así `Flatexco` y `FLATEXCO` son la misma cola.

---

## 8. Entrada: el Excel esperado

### 8.1 Hojas

| Hoja | Papel | ¿Obligatoria? |
|---|---|---|
| `Print` | Métricas del corte: las dos tablas y el total general | **Sí** |
| `RawDataRingCentral` | Llamadas por día y por cola | **Sí** |
| `Raw Data General` | Ventas por día y por campaña (tablas dinámicas) | **Sí** |
| `Valores$$$` | Costo capturado por día | No (si falta, se omite el dato) |
| `RawDataRingCentral *` | Hojas de llamadas auxiliares, detectadas **por prefijo** | No |

**Ignoradas a propósito:** `Print (2)`, `Print (3)`, `Grupos (2)` (contienen `#REF!` y `#VALUE!`),
`Grupos` y los rankings por asesor, ciudad o proveedor (fuera de alcance).

### 8.2 Hoja `Print`

**Cabecera del corte** (se busca en las primeras 10 filas):

| Etiqueta | Qué se lee |
|---|---|
| `Corte` | Celda a la derecha = fecha de corte · dos a la derecha = días hábiles (debe ser numérico) |
| `Mes` | Celda a la derecha = mes. Si falta, se deduce como día 1 de la fecha de corte |

**Tabla 1 — Llamadas.** Se localiza por la fila que contenga **todas** estas etiquetas:

```
Campaña · Monto · Numero de llamadas · #Perdidas · Costo Llamada
```

Las demás columnas de la cabecera se leen igual (se mapean dinámicamente) y se muestran tal cual en
la tabla de detalle. Las que el script usa para calcular, si están presentes:

`% Llamadas Perdidas` · `Ventas (Cliente)` · `Ventas (Linea)` · `Llamadas /Cliente` ·
`Llamadas /Linea` · `Eficiencia /Cliente` · `Eficiencia /Linea` · `Costo /Cliente` · `Costo /Linea`

**Tabla 2 — Leads.** Etiquetas clave: `Campaña` · `Numero de leads` · `Costo Lead`

**Fila de totales:** en ambas tablas, la primera fila con la columna `Campaña` vacía pero con datos
numéricos en el resto. Su ausencia es un error.

**Total general:** una fila etiquetada `Total`, debajo del bloque de leads, con el importe en la
celda inmediatamente a la derecha. Incluye llamadas + leads.

### 8.3 Hoja `RawDataRingCentral`

Cabecera en la primera fila. Columnas **obligatorias**:

| Columna | Contenido |
|---|---|
| `Date` | Fecha. Admite fecha real o texto en `MM/DD/YYYY`, `DD/MM/YYYY` o `YYYY-MM-DD` |
| `Queue Name` | Nombre de la cola |
| `#Calls` | Llamadas recibidas |
| `Missed` | Llamadas perdidas |

Se agrega por `(fecha, cola)`. **No se recalcula lógica de negocio**: `#Calls` y `Missed` ya vienen
calculadas por fórmula en el propio Excel.

**Hojas auxiliares.** Cualquier hoja cuyo nombre empiece por `RawDataRingCentral` y no sea la
principal. El nombre cambia según el mes (`... HM2` en mayo, `... HM2-IA` desde junio), por eso se
detectan por prefijo y no por nombre exacto.

Son hojas tramposas porque cambian de papel:

| Mes | Papel de la hoja auxiliar |
|---|---|
| Agosto 2026 | Sobra: esas campañas estaban en la principal. Sumarla habría inflado las cifras ~50 % |
| Julio y junio 2026 | Única fuente de `Heri Mob 2` e `IA FCO Internet`. Sin ella faltaban 3.681 llamadas en julio |
| Mayo 2026 | Única fuente de `Heri Mob 2` (1.577 llamadas), con la hoja llamada `HM2` |

**La regla que funciona siempre:** la hoja principal manda; en las auxiliares solo se buscan las
campañas que declaran llamadas en `Print` pero cuyas colas **no aparecieron** en la principal.
Cuando ocurre, el dashboard avisa de qué hoja las sacó y la validación final confirma la elección.

### 8.4 Hoja `Raw Data General`

Contiene varios bloques de tabla dinámica. Cada bloque se localiza por la celda `Row Labels`.

Dentro de cada bloque, las columnas se localizan **por su cabecera** (nunca por posición):

| Cabecera | Significado |
|---|---|
| `Count of WO Sale Date` | Ventas por cliente |
| `Sum of V-linea` | Ventas por línea |
| `Costos` | Costo del día (opcional) |

Las filas cuya primera celda es una fecha son la serie diaria; la fila `Grand Total` cierra el
bloque. El título del bloque se busca hasta 3 filas por encima.

**Deduplicación y emparejamiento.** La hoja trae bloques repetidos con los mismos datos; sumarlos
todos duplicaría las ventas. `deduplicar_bloques()` conserva uno de cada grupo idéntico,
prefiriendo el que tiene título. Después `asignar_bloques_a_campanas()` empareja en dos pasos:

1. **Por título del bloque** — coincidencia de prefijo, normalizada.
2. **Por `Grand Total`** — solo se aceptan correspondencias **1 a 1** contra
   `(Ventas (Cliente), Ventas (Linea))` de `Print`.

Si un bloque encaja con dos campañas (mismas ventas por casualidad o por una referencia cruzada del
Excel), se recurre al `Monto` para desempatar: gana la única campaña cuyo importe respalden los
costos del bloque, ya sea porque suman su `Monto` o porque un día suelto lo iguala (tolerancia:
`max(0.02, |monto| × 0.001)`).

> **El `Monto` solo desempata, nunca descarta.** Los costos a veces vienen en un único día y otras
> repartidos, y no siempre suman el total de `Print`. Usarlo para filtrar rechazaba bloques
> correctos.

Si aun así queda ambiguo, **no adivina**: esa campaña se queda fuera de la serie diaria y se
reporta como aviso. Sus totales del corte se siguen mostrando.

### 8.5 Hoja `Valores$$$`

Cabecera en la fila 1, con una columna `Total`. La columna 1 debe contener fechas. Se recogen solo
los días con importe distinto de cero.

Se presenta como **dato puntual del corte, nunca como serie diaria**: la captura es manual y va con
retraso, así que un día en cero significa «aún no cargado», no «sin inversión». Dibujarlo como
serie sugeriría un desplome del gasto que no ocurrió.

Si la hoja no existe o no tiene columna `Total`, el dashboard se genera igual sin ese dato.

---

## 9. Configuración

Hay dos niveles: el bloque `CONFIG` del script (código) y `configuracion.json` (sin tocar código).

### 9.1 Bloque `CONFIG` — [dashboard_kpi.py:24-205](dashboard_kpi.py#L24-L205)

| Constante | Valor por defecto | Para qué |
|---|---|---|
| `CARPETA_BASE` | Carpeta del `.py` | Raíz de todas las rutas. No tocar. |
| `CARPETA_ENTRADA` | `CARPETA_BASE\Entrada` | Dónde buscar el `.xlsx` |
| `CARPETA_SALIDA` | `CARPETA_BASE` | Dónde escribir el HTML. Se crea sola si no existe. |
| `NOMBRE_SALIDA` | `dashboard.html` | Nombre del archivo generado |
| `NOMBRE_MARCA` | `Marketing KPI` | Cabecera, título de la pestaña, subtítulo y nombre de la copia descargada |
| `ARCHIVOS_LOGO` | `logo.svg/.png/.webp/.jpg/.jpeg` | Si existe alguno en la carpeta, se incrusta en base64; si no, se dibuja el emblema vectorial de reserva |
| `ABRIR_AL_TERMINAR` | `True` | `False` para uso desatendido |
| `MAPEO_COLAS` | 14 campañas | Campaña de `Print` → lista de colas de `RawDataRingCentral` |
| `COLAS_IGNORADAS_CONOCIDAS` | 9 colas | Colas del raw que no son campañas de llamadas. Solo documentan; el script ya las ignora por no estar en `MAPEO_COLAS` |
| `CIFRAS_DE_CONTROL_POR_CORTE` | 3 cortes | Cifras exactas de archivos ya revisados a mano |
| `CAMPANAS_SIN_SERIE_DIARIA_CONOCIDAS` | `set()` (vacío) | Silencia el aviso de campañas sin bloque propio |
| `NOMBRE_CONFIG` / `RUTA_CONFIG` | `configuracion.json` | Archivo opcional de configuración |

**Publicar en red** — una sola línea:

```python
CARPETA_SALIDA = CARPETA_BASE                 # por defecto: junto al script
CARPETA_SALIDA = r"\\servidor\carpeta\kpi"    # publicar en red
```

El prefijo `r"..."` es obligatorio: evita que las barras invertidas se interpreten mal. No hay nada
que recompilar y el `.bat` no se toca.

### 9.2 `MAPEO_COLAS`

Los nombres de campaña de `Print` **no coinciden** con los nombres de cola del reporte de llamadas.
La traducción vive aquí:

```python
MAPEO_COLAS = {
    "TAEKNO Mobile A": ["TAEKNO Mobile"],
    "TMETRO":          [],                          # lista vacía = no tiene llamadas
    "ALLI WL1":        ["Allitech WL1"],
    "ALLI WL2":        ["Allitech WL2"],
    "ALLI WL3":        ["Allitech WL3"],
    "ALLI WL4":        ["Allitech WL4"],
    "ROYI FIBRA":      ["ROYI FIBRA"],
    "ITSFIBERNET":     [],
    "R-FRONTIER":      [],
    "Heri Mob 1":      ["Heri Mob 1", "HERIMOBILE"], # el nombre cambió de un mes a otro
    "Heri Mob 2":      ["Heri Mob 2"],
    "NEXTL2":          ["NEXOTL 2", "NEXOTL 1"],     # una campaña puede sumar varias colas
    "FLATEXCO":        ["Flatexco"],
    "IA FCO Internet": ["IA FCO Internet"],
}
```

Tres reglas que ahorran trabajo:

- **Mayúsculas y acentos dan igual.** `Flatexco` y `FLATEXCO` son la misma cola.
- **Se pueden poner varios nombres.** Si una cola se renombra, se deja el viejo y se añade el
  nuevo: los archivos de meses anteriores siguen funcionando.
- **Si una campaña recibe de varias colas**, se ponen todas y se suman.

**Para añadir una campaña nueva:** mirar cómo se llama en la columna `Campaña` de `Print`, cómo se
llama en `Queue Name` de `RawDataRingCentral`, y añadir la línea. Si todavía no tiene llamadas,
lista vacía `[]`.

**No hace falta acordarse de esto.** Si aparece una campaña sin mapear, el script se detiene y
escribe la línea exacta que hay que pegar. Nunca genera un dashboard al que le falten llamadas.

### 9.3 `configuracion.json`

Archivo **opcional** que vive al lado del `.py`. Lo escribe el panel Configuración del propio
dashboard, así que nadie tiene que editar Python para renombrar una campaña o arreglar una cola.

```json
{
  "nombres_campanas": { "ALLI WL1": "Allitech Wireless 1" },
  "mapeo_colas":      { "Heri Mob 1": ["Heri Mob 1", "HERIMOBILE"] },
  "mostrar_notas_en_dashboard": false
}
```

| Clave | Efecto |
|---|---|
| `nombres_campanas` | **Solo cosmético.** La lectura del Excel sigue usando el nombre original de `Print`, así que renombrar no puede mover ninguna cifra. |
| `mapeo_colas` | **Se fusiona** sobre `MAPEO_COLAS` (`dict.update`): manda solo en las campañas que nombra; el resto sigue con lo que diga el código. |
| `mostrar_notas_en_dashboard` | `false` (por defecto) deja las notas dentro de Configuración; `true` las trae arriba del tablero. |

**Se carga lo primero de todo**, antes de leer el Excel, porque puede cambiar las colas de una
campaña.

**Tolerancias de formato:** acepta BOM (`utf-8-sig`); en `mapeo_colas` acepta una cadena con comas
además de una lista; descarta nombres vacíos y los que sean idénticos a la clave.

> **Un JSON roto detiene el proceso.** No se ignora en silencio: si alguien corrigió ahí el nombre
> de una cola y el archivo no se lee, el dashboard saldría con llamadas de menos y sin avisar.

### 9.4 El panel Configuración (dentro del dashboard)

Botón arriba a la derecha. Abre un panel **por encima** del tablero, no comparte pantalla con las
cifras. Contiene cuatro cosas:

**1. Nombres de las campañas** — una fila por campaña, tres columnas:

| Columna | Qué es |
|---|---|
| Campaña en el Excel | El nombre tal cual viene en `Print`. Es la clave de lectura: no se toca. |
| Nombre a mostrar | Lo único que cambia en el tablero. Renombrar es seguro. |
| Colas en RawDataRingCentral | Los nombres de cola que suman sus llamadas, separados por comas. |

La tercera columna es la que evita editar código cuando en el Excel cambia el nombre de una cola.
Se listan también las campañas de `MAPEO_COLAS` que no salen en el Excel de este mes: son las que
están a la espera, y si no aparecieran, guardar el archivo las borraría.

**2. Notas de datos** — las rarezas del corte, con un globo naranja en el botón indicando cuántas
hay. Un interruptor las devuelve al tablero.

**3. Guardar los cambios** — dos niveles, a propósito:

- **Aplicar ahora** → se ven al instante y se recuerdan **en ese navegador**. Quien abra el archivo
  en otro sitio ve los nombres originales. El panel avisa mientras haya cambios así.
- **Guardar `configuracion.json`** → los hace definitivos. Se elige la carpeta del dashboard, al
  lado del `.py`. A partir de la siguiente generación salen de ahí, para todo el mundo.

**Restablecer** borra lo guardado en el navegador y vuelve a los nombres del dashboard recién
generado.

**4. Enviar el tablero a otras personas** — un botón que descarga una copia **sin** la sección
Configuración. La arma el navegador clonando la página: quita el panel y sus botones, deja los
gráficos vacíos para que se redibujen solos al abrir la copia (si no, se guardaría el SVG ya
pintado) y, si se ha renombrado algo, reescribe los ejes de los dos gráficos de campaña con los
nombres nuevos. Una casilla decide si la copia se lleva las notas de datos.

La copia sale **tal como se está viendo el tablero**, con los cambios de nombre aplicados aunque no
se haya guardado el JSON. Pesa lo mismo que el original.

---

## 10. Sistema de validación

`validar()` corre **antes de dibujar nada** ([dashboard_kpi.py:981](dashboard_kpi.py#L981)). Si algo
falla, lanza `ErrorDatosExcel` y **no se escribe ningún archivo**.

### 10.1 Validaciones universales (siempre, sea el mes que sea)

| # | Comprobación | Por qué importa |
|---|---|---|
| 1 | Toda campaña de `Print` está en `MAPEO_COLAS` | Sin esto, una campaña nueva perdería sus llamadas en silencio |
| 2 | La suma de llamadas día a día **cuadra exactamente** con el total de `Print` | Detecta cualquier cola sin asignar |
| 3 | No hay `#REF!`, `#VALUE!`, `#N/A` ni `#DIV/0!` en ninguna celda necesaria | Un error de fórmula propagado a una cifra |
| 4 | Están todas las cabeceras esperadas y las filas de totales | Cambios de layout del Excel |

Además, un **aviso** (no error) si hay colas en el raw que no están en `MAPEO_COLAS` ni en
`COLAS_IGNORADAS_CONOCIDAS`: puede ser una campaña nueva que todavía no llegó a `Print`.

Estas cuatro son las que protegen a los meses futuros.

### 10.2 Cifras de control (solo cortes ya revisados a mano)

Para los cortes presentes en `CIFRAS_DE_CONTROL_POR_CORTE` se comparan además **12 cifras exactas**:

llamadas totales · llamadas perdidas · % perdidas · ventas cliente · ventas línea · monto ·
eficiencia /cliente · eficiencia /línea · llamadas por día · leads de META HERI · total general ·
suma de la serie diaria

Tolerancia: `max(0.01, |esperado| × 0.001)`.

> **La clave de esa tabla es `(mes, fecha de corte)`, no solo el mes.** El mismo Excel se actualiza
> cada día con un corte nuevo, así que unas cifras atadas solo al mes fallarían en cuanto avanzara
> el corte, aunque los datos estuvieran perfectos.

**Rellenarla es opcional.** Un corte que no esté en la tabla se salta estas comprobaciones y se
queda con las universales. No hay que añadir nada cada día.

Cortes registrados actualmente: **1/8/2026 – corte 11/08** · **1/8/2026 – corte 12/08** ·
**1/7/2026 – corte 31/07**.

### 10.3 Avisos (no bloquean)

Se acumulan en la lista `avisos` y salen tanto por consola como en el recuadro **Notas de datos**
del dashboard:

- Campañas rescatadas de una hoja auxiliar (con el nombre de la hoja).
- Colas del raw sin mapear.
- Bloques de `Raw Data General` que encajan con varias campañas y no se pueden desempatar.
- Campañas con ventas en `Print` pero sin bloque propio de detalle.

---

## 11. Salida: anatomía del dashboard

Un HTML de ~4,8 MB con barra de navegación fija, siete secciones y el panel de Configuración.

**Cabecera:** nombre del mes, marca, y chips con corte, días hábiles, **nombre del archivo Excel
usado** y **hora de generación**. Esos dos últimos son los que permiten saber si los datos están
frescos antes de compartir el enlace.

### 11.1 Secciones

| # | Sección | Contenido |
|---|---|---|
| 1 | **Resumen** | 8 tarjetas KPI |
| 2 | **Actividad** | Llamadas por campaña (top 5, bloque de columnas) + Monto del corte con reparto |
| 3 | **Ritmo diario** | Área escalonada de llamadas/día + matrices de puntos de contestadas y perdidas + tarjeta de lectura |
| 4 | **Eficiencia** | Serie diaria de eficiencia /Línea con línea de tendencia y reproductor |
| 5 | **Campañas** | Barras de llamadas vs. perdidas + barras de eficiencia por campaña |
| 6 | **Detalle** | Tabla completa, mismas columnas y nombres que `Print` |
| 7 | **Leads** | Tabla de campañas captadas por formulario de META |

### 11.2 Las 8 tarjetas KPI

| Tarjeta | Cálculo | Subtítulo |
|---|---|---|
| Llamadas totales | `Print` → `Numero de llamadas` | Llamadas por día hábil |
| Llamadas perdidas | `Print` → `#Perdidas` | % del total |
| Ventas (Cliente) | `Print` | 1 cada N llamadas |
| Ventas (Línea) | `Print` | 1 cada N llamadas |
| Monto | `Print` | Total general (llamadas + leads) |
| Eficiencia /Cliente | `Ventas (Cliente) ÷ llamadas` | — |
| Eficiencia /Línea | `Ventas (Linea) ÷ llamadas` | — |
| Llamadas / día | `llamadas ÷ días hábiles` | Nº de días hábiles |

Cada tarjeta grande tiene un botón redondo `···` que abre una nota corta explicando qué mide y de
qué hoja del Excel sale el dato (las variables `ayuda_*` en `construir_html`).

### 11.3 Gráficos

**Tres con Plotly:**

1. **Eficiencia /Línea día a día** (`go.Scatter`, 380 px) — con línea de tendencia por regresión
   lineal e insignia de subida/bajada/estable.
2. **Llamadas y perdidas por campaña** (`go.Bar`, 400 px).
3. **Eficiencia por campaña** (`go.Bar`, 400 px) — sobre cliente y sobre línea.

Configuración común: `displaylogo: False`, `responsive: True`, `displayModeBar: False`, y
`_estilo_plotly()` unifica tipografía, colores y márgenes con el resto del tablero.

**Cuatro construidos a mano en HTML/SVG:** bloque de columnas del top 5, barras de reparto del
monto, área escalonada de llamadas por día y matrices de puntos (una por día, con globo de fecha y
cifra).

### 11.4 Reglas de representación

- **Los días posteriores al corte se dibujan como huecos (`None`), nunca como cero.** Un cero
  dibujaría una caída a plomo que no ocurrió.
- **Las campañas sin llamadas van con guion `—`, no con cero.** Un 0 sugeriría que se midió y dio
  cero.
- **La eficiencia diaria divide ventas y llamadas del mismo conjunto de campañas.** Si una campaña
  no aporta numerador (no tiene bloque de detalle), tampoco aporta denominador; si no, la curva se
  hundiría por una diferencia de cobertura y no por rendimiento real.
- **El gráfico de eficiencia por campaña excluye las eficiencias imposibles (>100 %)** y lo dice
  bajo el título. Con ALLI WL2 dentro, su 3.400 % aplastaba la escala. La campaña sigue completa en
  la tabla de detalle y en las notas.
- **El costo se presenta como dato del corte, no como serie diaria** (ver §8.5).

### 11.5 El reproductor

Botón «Reproducir» + control deslizante que redibuja la serie de eficiencia día a día.

> **No usa la animación de Plotly.** `Plotly.animate` deja la serie sin línea ni relleno al saltar
> de fotograma (solo quedan los puntos sueltos). La reproducción va por `Plotly.restyle`,
> recortando la serie desde el propio tablero. Si algún día se vuelve a tocar ese gráfico, conviene
> **no reintroducir `frames` ni el slider de Plotly**.

---

## 12. Referencia del código

`dashboard_kpi.py` — 3.358 líneas, 63 funciones, 3 clases. Organizado en bloques con separadores
`# ====`.

### Utilidades — [L211](dashboard_kpi.py#L211)

| Función | Qué hace |
|---|---|
| `_norm(txt)` | Normaliza para comparar: sin acentos, minúsculas, espacios colapsados |
| `archivo_mas_reciente(carpeta)` | El `.xlsx` más reciente por `mtime`, ignorando `~$...` |
| `a_fecha(valor)` | Fecha desde `datetime`, `date` o texto (`MM/DD/YYYY`, `DD/MM/YYYY`, `YYYY-MM-DD`) |
| `_Celda` | Celda mínima con la interfaz de openpyxl (`row`, `column`, `value`), con `__slots__` |
| `HojaEnMemoria` | Copia en memoria de una hoja pequeña, con `cell()` e `iter_rows()` |
| `hoja_a_memoria(wb, hoja, limite_filas=2000)` | Vuelca una hoja al objeto anterior |
| `colas_de_campana(nombre)` | Colas de una campaña; acepta lista o cadena suelta |
| `campana_por_cola()` | Índice inverso: cola normalizada → campaña |
| `nombre_visible(clave)` | Nombre a mostrar. **Solo cosmético** |
| `cargar_configuracion()` | Lee y valida `configuracion.json`; aplica sobre `MAPEO_COLAS` y nombres |
| `verificar_no_error(valor, contexto)` | Aborta si la celda trae `#REF!`, `#VALUE!`, `#N/A` o `#DIV/0!` |

### Fase 1 — Extracción: hoja `Print` — [L416](dashboard_kpi.py#L416)

| Función | Qué hace |
|---|---|
| `leer_encabezado_corte(ws)` | Busca `Corte` y `Mes`; devuelve `(mes, fecha_corte, dias_habiles)` |
| `_fila_a_dict(ws, fila, col_map)` | Convierte una fila en dict, verificando errores de celda |
| `leer_tabla(ws, etiquetas_clave, nombre)` | Localiza una tabla por sus etiquetas, lee las filas y para en la de totales |
| `leer_total_general(ws)` | La fila `Total` bajo el bloque de leads |
| `extraer_print(ws)` | Orquesta todo lo anterior y devuelve el dict de datos del corte |

### Fase 1 — Extracción: llamadas — [L556](dashboard_kpi.py#L556)

| Función | Qué hace |
|---|---|
| `_leer_hoja_llamadas(wb, hoja)` | Lectura en streaming; agrega `#Calls` y `Missed` por `(fecha, cola)` |
| `hojas_llamadas_auxiliares(wb)` | Hojas distintas de la principal, detectadas por prefijo |
| `extraer_llamadas_diarias(wb, campanas)` | La principal manda; rescata de las auxiliares solo lo que falte |

### Fase 1 — Extracción: `Raw Data General` — [L671](dashboard_kpi.py#L671)

| Función | Qué hace |
|---|---|
| `_titulo_arriba(ws, fila, col)` | Busca el título hasta 3 filas por encima del bloque |
| `_leer_bloque(ws, fila_cab, col)` | Lee un bloque; columnas **por cabecera**, nunca por posición |
| `extraer_bloques_raw_data_general(ws)` | Localiza todos los bloques por la celda `Row Labels` |
| `deduplicar_bloques(bloques)` | Conserva uno de cada grupo idéntico, prefiriendo el que tiene título |
| `_bloque_compatible(bloque, vc, vl)` | `Grand Total == (ventas cliente, ventas línea)`. El Monto no interviene |
| `_costo_respalda(bloque, monto)` | Solo para desempatar: los costos suman el Monto o un día lo iguala |
| `asignar_bloques_a_campanas(bloques, objetivo)` | Emparejamiento 1 a 1 en dos pasos; devuelve asignaciones, avisos y excluidas |

### Fase 1 — Extracción: costos — [L888](dashboard_kpi.py#L888)

| Función | Qué hace |
|---|---|
| `extraer_costo_corte(ws)` | Días con importe en la columna `Total` de `Valores$$$`. Si falta, devuelve `None` sin bloquear |

### Serie diaria y validación — [L919](dashboard_kpi.py#L919)

| Función | Qué hace |
|---|---|
| `construir_serie_diaria(...)` | Eje de mes completo; llamadas, perdidas, ventas y llamadas para eficiencia, por día |
| `validar(...)` | Las comprobaciones de §10. Lanza `ErrorDatosExcel` o devuelve los avisos |

### Formato — [L1083](dashboard_kpi.py#L1083)

`_a_formato_es` · `fmt_num` · `fmt_pct` · `fmt_usd` · `fmt_fecha_es` · `fmt_mes_es`
— formato español: punto de miles, coma decimal.

### Render — [L1137](dashboard_kpi.py#L1137)

| Elemento | Qué es |
|---|---|
| `C_AZUL`, `C_VERDE`, `C_ROSA`, `C_AMBAR`, `C_VIOLETA`, `C_TINTA`… | Paleta compartida entre Plotly y el CSS |
| `CSS_DASHBOARD` | Hoja de estilos completa (L1164) |
| `JS_DASHBOARD` | Reproductor, menú activo y panel de Configuración (L1663) |
| `_esc`, `_mezclar`, `_escala_bonita` | Ayudantes de escape, mezcla de color y escalas redondas |
| `_svg_marca`, `_svg_marca_reserva`, `_icono_*` | Logo (base64 o vectorial de reserva) e iconos SVG |
| `_cabecera_tarjeta`, `_bloque_columnas`, `_barras_reparto`, `_area_escalonada`, `_matriz_puntos` | Bloques visuales hechos a mano |
| `_fila_config`, `_panel_configuracion` | Panel de Configuración |
| `_estilo_plotly(fig, altura, margenes)` | Unifica los tres gráficos con el resto del tablero |
| `construir_html(contexto)` | Ensambla el documento completo (L2560) |

### Main — [L3257](dashboard_kpi.py#L3257)

`main()` orquesta el pipeline de §7. El bloque `if __name__ == "__main__"` captura
`ErrorDatosExcel`, lo imprime en `stderr` y sale con código **1**.

---

## 13. Capa de diseño (CSS/JS)

Maquetado a mano, sin frameworks ni CSS externo. Sin fuentes web: se usa la del sistema (Segoe UI
en Windows) porque el archivo tiene que abrirse sin internet.

Todo el diseño vive en dos constantes, justo encima de `construir_html`:

- **`CSS_DASHBOARD`** — colores y medidas arriba del todo, en el bloque `:root`. Cambiar ahí un
  color lo cambia en todo el tablero.
- **`JS_DASHBOARD`** — tres piezas: el reproductor del gráfico diario, el menú que se ilumina según
  la sección visible, y el panel de Configuración.

### Paleta (`:root`)

| Token | Valor | Uso |
|---|---|---|
| `--bg` | `#F0F0EE` | Fondo |
| `--tarjeta` | `#FAFAF9` | Tarjetas |
| `--tinta` / `--tinta-2` | `#141412` / `#3E3E39` | Texto |
| `--suave` / `--tenue` | `#8A8A80` / `#B7B7AE` | Texto secundario |
| `--linea` / `--linea-fuerte` | `#E7E7E2` / `#DADAD3` | Bordes |
| `--azul` / `--azul-osc` | `#2F6FED` / `#1E4FD8` | Acento principal |
| `--verde` | `#1FA84A` | Positivo |
| `--rosa` | `#EC2E7B` | Perdidas |
| `--ambar` | `#E8963C` | Avisos y globo de notas |
| `--violeta` | `#7C5CE6` | Serie secundaria |
| `--rojo` | `#DC2626` | Crítico |

Los mismos valores existen como constantes Python (`C_AZUL`, `C_ROSA`…) para que Plotly y el CSS no
se desincronicen.

### Tres detalles del maquetado que no son casuales

> Por si alguien los «simplifica» sin querer.

1. **El desfase del ancla y el resaltado del menú salen de la misma medida.** El JS mide la barra
   fija y la publica en `--h-barra`; el CSS la usa en el `scroll-margin-top` de las secciones y el
   JS la usa para la línea de corte. Antes había un `scroll-padding-top` en `html` **y** un
   `scroll-margin-top` en `section`, que se suman: la sección aterrizaba 192 px más abajo del
   umbral fijo de 150 px con el que se decidía el resaltado, y por eso al pulsar un enlace se
   quedaba iluminada la sección anterior.
2. **La actualización del menú lleva `requestAnimationFrame` y un temporizador.** El rAF solo corre
   cuando el navegador dibuja un fotograma; sin la red de seguridad, el menú puede quedarse clavado.
3. **Las alturas del bloque de columnas son elásticas.** Comparte fila con «Monto del corte», cuya
   lista crece con el número de campañas; con la zona de barras fija, ese alto de más quedaba como
   un hueco vacío bajo el gráfico. `--h-zona-col` es el **mínimo**, no el alto exacto, y la lista de
   montos scrollea a partir de 430 px.

En la matriz de puntos, cada punto es redondo y su diámetro sale del ancho disponible con tope de
9 px. Ocupa el ancho entero de la tarjeta a propósito: encajada entre la cifra y el porcentaje le
quedaban unos 180 px para 31 días y los puntos salían como rayitas de 3 px.

---

## 14. Distribución y despliegue

### 14.1 Compartir solo el resultado

Mandar únicamente `dashboard.html`. Lleva todo dentro. El destinatario no necesita nada.

**Mejor todavía: la copia limpia.** En **Configuración → Enviar el tablero a otras personas** hay un
botón que descarga una copia **sin** la sección de Configuración. La genera el propio navegador:
sale un `.html` en Descargas, con el mes en el nombre, listo para adjuntar a un correo. Quien la
recibe ve el tablero completo —mismas cifras, mismos gráficos, mismo reproductor— pero sin poder
cambiar nombres ni colas sin querer.

### 14.2 Repartir la herramienta

Mandar la carpeta completa **menos** `dashboard.html` y el Excel de `Entrada`:

```
dashboard_kpi.py · Generar_dashboard.bat · README.md · DOCUMENTACION.md · Entrada\ (vacía)
```

`configuracion.json` conviene mandarlo si ya se han cambiado nombres o colas: es lo que hace que
todos vean lo mismo.

### 14.3 Repartir un cambio del script

Cada compañero tiene **su propia copia** del `.py`. Editar el propio no cambia el de los demás.

**Opción A — mandar el archivo.** Se pasa el `dashboard_kpi.py` nuevo y se pega encima del suyo.
Nada más cambia: ni el `.bat`, ni `Entrada`, ni el Excel. Lo más simple si son pocos.

**Opción B — una sola copia en red.** La carpeta del script vive en una unidad compartida y todos
ejecutan el `.bat` desde ahí. Solo existe un `.py` y cualquier cambio llega solo.

> Con la opción B, **dejar `CARPETA_SALIDA` como está**. Si además se apunta a una carpeta fija de
> red, todos escribirían en el mismo `dashboard.html` y el último que ejecute pisaría el de los
> demás. Publicar en red tiene sentido cuando **una sola persona** genera el tablero para todos.

### 14.4 Publicar en red

Editar una línea (§9.1) y guardar. La carpeta destino se crea sola. El nombre del archivo no
cambia nunca, así que **un enlace compartido no deja de funcionar**.

---

## 15. Solución de problemas

Cuando algo va mal, el script **no genera el HTML**: aborta e imprime qué pasó y en qué hoja. El
dashboard anterior no se toca.

| Situación | Mensaje | Solución |
|---|---|---|
| No hay `.xlsx` en `Entrada` | `No hay ningun archivo .xlsx en la carpeta: ...` | Copiar el Excel del mes |
| El Excel está abierto | Error de acceso al archivo | Cerrarlo y reintentar |
| Falta la hoja `Print` | `El Excel no tiene la hoja 'Print'...` | Comprobar que es el archivo correcto |
| Falta una cabecera | `No se encontro la cabecera de Tabla 1 (Llamadas)...` | Revisar que las etiquetas de §8.2 siguen ahí |
| `#REF!` o `#VALUE!` en una celda necesaria | `Celda con error '#REF!' en Print (...)` | Arreglar la fórmula en el Excel de origen |
| Campaña nueva sin mapear | `Hay campañas en la hoja 'Print' que no están en MAPEO_COLAS: '...'` | El propio mensaje escribe la línea a pegar. O añadirla desde Configuración |
| Las llamadas no cuadran | `La suma de llamadas día a día (...) no coincide con el total de 'Print'` | Casi siempre una cola sin asignar; revisar los avisos de colas desconocidas |
| `configuracion.json` roto | `El archivo 'configuracion.json' no es un JSON valido (...)` | Volver a guardarlo desde Configuración, o borrarlo |
| Cifra fuera de control | `Llamadas totales: esperado 9124, se leyo ...` | El corte está en `CIFRAS_DE_CONTROL_POR_CORTE` y no coincide: o cambió el Excel, o se rompió la lectura |
| Python no encontrado | `ERROR: Python no esta instalado en este equipo` | Instalar Python marcando «Add Python to PATH» |
| `pip install` falla | `ERROR: no se pudieron instalar las librerias` | Proxy o antivirus corporativo. Probar el comando a mano en una terminal |

**Nombres cambiados que no se ven en otro PC** → se aplicaron con «Aplicar ahora» (solo en ese
navegador). Hay que guardar `configuracion.json` en la carpeta del dashboard.

**El menú se queda iluminado en la sección anterior** → ver §13, detalle 1. No tocar el
`scroll-margin-top`.

---

## 16. Rarezas conocidas de los datos

No son fallas del script, son cosas del Excel de origen. El dashboard las recoge en **Notas de
datos**, dentro de Configuración.

### Ya corregidas en el archivo con corte 12/08

Se dejan documentadas porque el script las sigue detectando solo si vuelven a aparecer.

- **ALLI WL2 y ALLI WL4 sin desglose día a día.** En el archivo con corte 11/08 sus ventas
  (102/155 y 77/104) eran un calco exacto, día por día, de ALLI WL3 y Heri Mob 1, sin bloque propio
  en `Raw Data General`.
- **ALLI WL2 con eficiencia superior al 100 %** (3 llamadas, 102 ventas): imposible en la operación
  real, y consecuencia del punto anterior.

Si algo así reaparece, esas campañas quedan fuera del gráfico de eficiencia diaria —tanto sus
ventas como sus llamadas, para que numerador y denominador cubran las mismas campañas—. En las
tarjetas y en la tabla siguen completas, porque esas cifras vienen de `Print`.

### Ventas (Cliente) duplicando a Ventas (Línea) — mayo y junio 2026

**Pendiente de revisar con quien arma el Excel.** En esos dos meses, `ALLI WL3` y `ALLI WL4` tienen
en `Print` el mismo número en *Ventas (Cliente)* que en *Ventas (Línea)*, y no coincide con el
detalle de `Raw Data General`:

| Mes | Campaña | `Print` cliente/línea | Bloque de detalle |
|---|---|---|---|
| Mayo | ALLI WL3 | 454 / 454 | **341** / 454 |
| Mayo | ALLI WL4 | 146 / 146 | **114** / 146 |
| Junio | ALLI WL3 | 239 / 239 | **188** / 239 |
| Junio | ALLI WL4 | 293 / 293 | **209** / 293 |

Tiene toda la pinta de una fórmula copiada de la columna de al lado. La hoja `Print` es coherente
consigo misma, así que el error se propaga al total: en mayo la eficiencia /Cliente sale 7,11 %
cuando con las cifras del detalle sería 6,57 %.

**El script no lo corrige**, porque sería inventar datos. Muestra lo que dice `Print` y deja esas
campañas fuera de la serie diaria de ventas, avisándolo. En julio y agosto no ocurre.

> Que *cliente* y *línea* coincidan no es por sí solo un error — pasa cuando cada venta es de una
> sola línea (`FLATEXCO` en agosto, `NEXTL2` en julio, y ambas cuadran con su detalle). Lo
> sospechoso es que además **no cuadre** con el bloque de `Raw Data General`.

### Costos con retraso

`Valores$$$` se captura a mano; en agosto solo el día 1 tiene importes. Por eso el costo se
presenta como dato del corte y no como serie diaria: los días en cero significan «aún no cargado»,
no «sin inversión».

### Historial de cambios entre meses

De julio 2026 a agosto 2026, por si vuelve a pasar:

| Cambio | Cómo se resolvió |
|---|---|
| Campaña `ITSFIBERNET` sustituida por `R-FRONTIER` | Añadida a `MAPEO_COLAS` con lista vacía |
| Cola `Heri Mob 1` renombrada a `HERIMOBILE` | Las dos conviven como alias de la misma campaña |
| Cola `Flatexco` pasó a `FLATEXCO` | La comparación ignora mayúsculas; no hizo falta tocar nada |
| Apareció la cola `NEXOTL 1` junto a `NEXOTL 2` | `NEXTL2` ahora suma las dos |
| `Heri Mob 2` e `IA FCO Internet` salieron de la hoja principal | Se buscan en `RawDataRingCentral HM2-IA` |
| Se invirtió el orden de las columnas en `Raw Data General` | Ahora se leen por el nombre de la cabecera |
| Los costos pasaron de un único día a repartidos por día | El `Monto` ya solo se usa para desempatar |
| El corte de agosto avanzó del 11 al 12 (9 → 10 días hábiles) | Las cifras de control se indexan por (mes, corte) |
| `R-FRONTIER` (julio) volvió a ser `ITSFIBERNET` (agosto) | Las dos conviven en `MAPEO_COLAS`, sin llamadas |
| `ALLI WL2` y `ALLI WL4` ya traen ventas propias | Se vació `CAMPANAS_SIN_SERIE_DIARIA_CONOCIDAS` |
| La hoja auxiliar cambió de `... HM2` a `... HM2-IA` | Se detectan por prefijo, no por nombre exacto |

> **Ninguno de estos cambios se detectó a ojo: todos los destapó la validación al no cuadrar las
> cifras.** Por eso conviene no desactivarla nunca.

---

## 17. Notas de mantenimiento

Estado del repositorio a **21 de agosto de 2026**. Tres cosas que conviene saber:

### 1. `pandas` es una dependencia declarada pero no usada

El script no importa pandas en ninguna línea (verificado: 0 referencias). Sin embargo,
`Generar_dashboard.bat` lo comprueba y lo instala:

```bat
python -c "import pandas, openpyxl, plotly" >nul 2>&1
...
python -m pip install pandas openpyxl plotly
```

**Impacto:** ninguno funcional, pero si pandas falta, el `.bat` cree que faltan todas las librerías
y relanza el `pip install` completo. Quitarlo de las dos líneas del `.bat` (y de la instrucción del
README) aligeraría la puesta en marcha en equipos nuevos. **Es un cambio opcional, no un fallo.**

### 2. `Generar_dashboard.zip` está desactualizado

El paquete de reparto contiene:

| Archivo | En el `.zip` | En la carpeta | Estado |
|---|---|---|---|
| `dashboard_kpi.py` | 144.467 B · 18/08 09:44 | 151.335 B · 18/08 16:43 | **Desfasado ~7 KB** |
| `README.md` | 25.638 B · 18/08 09:53 | 25.638 B · 18/08 09:53 | Al día |
| `Generar_dashboard.bat` | 2.465 B · 13/08 08:57 | 2.465 B · 13/08 08:57 | Al día |

Quien reciba ese `.zip` recibe una versión anterior del script. **Antes de repartirlo hay que
regenerarlo**, e incluir además una carpeta `Entrada` vacía (el `.zip` actual no la trae) y este
`DOCUMENTACION.md`.

### 3. El proyecto no está bajo control de versiones

No hay repositorio git en `c:\dashboard`. Con un script de 3.358 líneas que cambia cada mes al
ritmo del Excel, y con copias repartidas entre varias personas, un `git init` daría historial,
comparación entre versiones y una forma de saber qué copia es la buena. Es la mejora estructural
más rentable si el sistema va a seguir vivo.

### Al cerrar cada mes

1. Copiar el Excel del mes en `Entrada` y ejecutar.
2. Si el script se detiene por una campaña nueva, añadirla desde **Configuración** (se guarda en
   `configuracion.json`, sin tocar código).
3. Revisar los **avisos** de la consola y las **Notas de datos** del tablero.
4. Opcional: cuando el corte esté revisado a mano, añadir sus cifras a
   `CIFRAS_DE_CONTROL_POR_CORTE` con la clave `(mes, fecha_corte)`.
5. Anotar en §16 cualquier cambio de estructura del Excel.

### Antes de tocar el código

- No leer nunca por coordenadas fijas: buscar la etiqueta (§7).
- No usar el `Monto` para descartar bloques: solo desempata (§8.4).
- No reintroducir `frames` ni el slider de Plotly en el gráfico diario (§11.5).
- No sumar un `scroll-padding-top` al `scroll-margin-top` de las secciones (§13).
- No desactivar la validación: es lo que ha destapado todos los cambios del Excel (§16).
