@echo off
REM ============================================================
REM  Genera el dashboard KPI Sales Department.
REM  Doble clic sobre este archivo. No hay que escribir nada.
REM ============================================================
setlocal
cd /d "%~dp0"

echo.
echo ===========================================
echo   Dashboard KPI Sales Department
echo ===========================================
echo.

REM --- 1) Comprobar que Python esta instalado -----------------
python --version >nul 2>&1
if errorlevel 1 goto sin_python

REM --- 2) Instalar las librerias la primera vez ---------------
python -c "import pandas, openpyxl, plotly" >nul 2>&1
if errorlevel 1 (
    echo Primera ejecucion: instalando librerias necesarias.
    echo Esto tarda un par de minutos y solo pasa una vez.
    echo.
    python -m pip install pandas openpyxl plotly
    if errorlevel 1 goto fallo_pip
    echo.
)

REM --- 3) Generar el dashboard --------------------------------
REM El propio script abre el navegador: es el unico que sabe la ruta
REM final, que cambia si se edita CARPETA_SALIDA.
python dashboard_kpi.py
if errorlevel 1 goto fallo_script

echo.
echo Listo. Esta ventana se cierra sola.
REM ping en lugar de timeout: timeout falla en algunos entornos
ping -n 7 127.0.0.1 >nul 2>&1
exit /b 0

REM ============================================================
:sin_python
echo.
echo  ERROR: Python no esta instalado en este equipo.
echo.
echo  Descargalo desde:  https://www.python.org/downloads/
echo.
echo  IMPORTANTE: durante la instalacion marca la casilla
echo  "Add Python to PATH", si no, este archivo no lo encontrara.
echo.
pause
exit /b 1

:fallo_pip
echo.
echo  ERROR: no se pudieron instalar las librerias.
echo.
echo  Suele ser el proxy o el antivirus de la empresa.
echo  Prueba a ejecutar esto en una terminal:
echo      python -m pip install pandas openpyxl plotly
echo.
pause
exit /b 1

:fallo_script
echo.
echo ===========================================
echo   EL DASHBOARD NO SE GENERO
echo ===========================================
echo.
echo  Lee el mensaje de error de aqui arriba: dice que falta
echo  y en que hoja del Excel.
echo.
echo  Lo mas habitual:
echo   - No hay ningun .xlsx en la carpeta Entrada.
echo   - El Excel esta abierto: cierralo y vuelve a intentarlo.
echo   - Falta una campana nueva en MAPEO_COLAS (ver README).
echo.
echo  El dashboard anterior NO se ha tocado, sigue siendo valido.
echo.
pause
exit /b 1
