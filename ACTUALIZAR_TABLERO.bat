@echo off
chcp 65001 > nul
title ACTUALIZACION BIENVENIDA FACILITADORES CUN
color 0A

echo ==============================================================================
echo              ACTUALIZACION BIENVENIDA FACILITADORES CUN
echo ==============================================================================
echo.

set SCRIPT_DIR=%~dp0
set PYTHON_EXE="C:\Program Files\Blender Foundation\Blender 4.0\4.0\python\bin\python.exe"

if not exist %PYTHON_EXE% (
    set PYTHON_EXE=python
)

echo Ejecutando script de consolidacion Excel y actualizacion del Tablero HTML...
echo.

%PYTHON_EXE% "%SCRIPT_DIR%Scripts\procesar_bienvenida.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Ocurrio un problema durante la ejecucion del proceso.
    echo Por favor revise los datos y vuelva a intentarlo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ==============================================================================
echo       PROCESO COMPLETADO EXITOSAMENTE!
echo       - Excel consolidado generado: Consolidado_Bienvenida_Facilitadores.xlsx
echo       - Tablero HTML actualizado: tablero.html y tablero_respaldo.html
echo ==============================================================================
echo.
pause
