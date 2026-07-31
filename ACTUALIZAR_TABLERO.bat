@echo off
chcp 65001 > nul
title ACTUALIZACION BIENVENIDA FACILITADORES CUN
color 0A

echo ==============================================================================
echo              ACTUALIZACION BIENVENIDA FACILITADORES CUN
echo ==============================================================================
echo.

set SCRIPT_DIR=%~dp0

rem Buscar ejecutable de Python disponible
set PYTHON_EXE=
if exist "%USERPROFILE%\scoop\apps\python\3.14.6\python.exe" set PYTHON_EXE="%USERPROFILE%\scoop\apps\python\3.14.6\python.exe"
if "%PYTHON_EXE%"=="" if exist "%USERPROFILE%\scoop\apps\python\current\python.exe" set PYTHON_EXE="%USERPROFILE%\scoop\apps\python\current\python.exe"
if "%PYTHON_EXE%"=="" if exist "%USERPROFILE%\scoop\shims\python.exe" set PYTHON_EXE="%USERPROFILE%\scoop\shims\python.exe"
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Blender Foundation\Blender 4.0\4.0\python\bin\python.exe" set PYTHON_EXE="C:\Program Files\Blender Foundation\Blender 4.0\4.0\python\bin\python.exe"
if "%PYTHON_EXE%"=="" set PYTHON_EXE=python

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
echo       - Excel de faltantes generado: Facilitadores_Pendientes_Asistencia.xlsx
echo       - Tableros HTML actualizados: tablero.html, tablero_respaldo.html e index.html
echo       - Repositorio GitHub actualizado: https://github.com/efi-cun/Cunectate_al_parche
echo ==============================================================================
echo.
pause
