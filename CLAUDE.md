# Proyecto CallLog — notas para Claude

Limpieza del Excel de registro de llamadas (CallLog) de la centralita.
Proyecto independiente: no tiene relación con `C:\Restaurantes`.

## Restricción principal

Los equipos donde se usa son de empresa **con CMD/terminal bloqueado**. Todo tiene
que poder ejecutarse abriendo el archivo en VS Code y pulsando Run / F5. Nada de
comandos, ni de pasos que exijan una consola.

## Entorno

- Python **3.12.10** en `%LOCALAPPDATA%\Programs\Python\Python312\python.exe`
  (el mismo que usa el proyecto Restaurantes, instalado a nivel de usuario).
- **No hay venv**: las librerías van en ese Python del sistema.
- Instalado: `openpyxl 3.1.5` + `et_xmlfile 2.0.0` (dependencia suya).
- VS Code con la extensión `ms-python.python`. Carpeta marcada como confiable.

## Archivos

- `limpiar_calllog.py` — el script. Rutas configurables arriba del todo:
  `CARPETA_ENTRADA` / `ARCHIVO_ENTRADA` y `CARPETA_SALIDA` / `ARCHIVO_SALIDA`.
  Carpeta vacía `""` = la carpeta del propio `.py` (no la del workspace de VS Code).
  **Entrada**: `.xlsx` o `.csv` (detecta el tipo por la extensión; en CSV detecta
  solo la codificación —`utf-8-sig`/`cp1252`/`latin-1`— y el separador).
  **Salida**: siempre `.xlsx`, y siempre un archivo NUEVO — el nombre lleva
  `_aaaammdd-hhmm` (`NOMBRE_CON_FECHA_HORA`). Nunca sobrescribe: si coinciden
  dos ejecuciones en el mismo minuto añade `_2`, `_3`.
- `instalar_librerias.py` — instala openpyxl llamando a pip vía `sys.executable`,
  ejecutable con F5. Es la vía para los equipos donde no hay terminal.
- `prueba_en\` — archivos de entrada. `prueba_sal\` — resultados.

## Reglas de limpieza (definidas por el usuario)

- `Date`: quitar el prefijo del día de la semana (Mon–Sun) y dejar `mm/dd/yyyy`.
  **Nunca reordenar día y mes.** Los valores sin cero delante (`8/5/2026`) se rellenan.
- **NINGUNA fila se borra.** Todo el filtrado de filas es un autofiltro de Excel:
  las filas se conservan en el archivo y quedan ocultas. Decisión explícita del
  usuario — "es un filtro", y un filtro oculta, no elimina. `MODO_FILTRADO = "eliminar"`
  cambia a borrado real si algún día hace falta.
- Se ocultan: las de `Direction` distinto de `Incoming` (incluye `Outgoing` y vacías)
  y las de `To` vacío (`COLUMNAS_SIN_VACIOS = ["To"]`).
- Los criterios del filtro se construyen con los valores **tal como vienen** en el
  archivo, así que variantes como `incoming` o `Incoming ` con espacio siguen
  funcionando si el usuario reaplica el filtro en Excel.
- El archivo de salida lleva siempre las flechas de autofiltro en la cabecera
  (`ACTIVAR_FILTRO = True`).
- Lo único que sí se borra son las 5 **columnas**: una columna no se puede ocultar
  con un autofiltro.
- Eliminar 5 columnas: `Extension`, `Forwarded To`, `Result Description`,
  `Included`, `Purchased`. El resto se conserva tal cual.
- El archivo original **nunca** se modifica; la salida es un `.xlsx` nuevo.

## Decisiones técnicas

- **openpyxl y no pandas**: una sola dependencia ligera (250 KB frente a pandas+numpy),
  menos superficie para que falle el proxy corporativo, y pandas reinterpreta las
  fechas por su cuenta, que es justo lo que hay que evitar aquí.
- La columna `Date` se escribe como **texto** con formato de celda `@`, para que Excel
  no la muestre según la configuración regional (en España saldría `30/07/2026`).
- Comparaciones de `Direction` y de nombres de columna sin distinguir
  mayúsculas y recortando espacios.
- Las rutas se resuelven contra `Path(__file__).parent`, no contra el directorio de
  trabajo: el botón Run de VS Code no siempre arranca desde la carpeta del archivo.

## Estado (3 agosto 2026)

Verificado con los datos reales: 990 filas → **990 en el archivo, 134 visibles**
(856 ocultas: 703 por `Direction` y 153 por `To` vacío), 15 → 10 columnas,
todas las fechas en `mm/dd/yyyy`.

Comprobado que **CSV y XLSX dan resultado idéntico**, con dos CSV de formatos
distintos (utf-8 + coma, y cp1252 + punto y coma): los tres dan 990/134.

Verificado a nivel de XML: `<autoFilter>` presente con criterio en las dos columnas
(`Direction` y `To`, ambas con `blank=False`), y las filas ocultas coinciden una a una
con las que incumplen alguna regla. Y que tres ejecuciones seguidas generan tres
archivos distintos sin pisarse.

Archivos de prueba en `prueba_en\`: `archivo de prueba 1.xlsx` (15 filas, los 7 días de
la semana) y `archivo de prueba 2.xlsx` (12 filas de casos límite: fechas sin cero
delante, `incoming` en minúsculas, espacios sobrantes, fila vacía, valor `Missed`).

## Pendiente

- **Reparto a otros equipos**: como openpyxl es Python puro, la vía sin instalaciones
  es copiar las carpetas `openpyxl\` y `et_xmlfile\` junto al `.py`.
- Con el tiempo, `prueba_sal\` se llenará de archivos (uno por ejecución, por diseño).
  Si molesta, tocaría añadir un borrado de los más antiguos.
