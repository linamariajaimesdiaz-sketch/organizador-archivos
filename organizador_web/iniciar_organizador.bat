@echo off
title Organizador Inteligente de Archivos
chcp 65001 > nul
echo ========================================================
echo    ORGANIZADOR INTELIGENTE DE ARCHIVOS
echo ========================================================
echo.
echo [1/3] Verificando instalacion de Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [AVISO] No se detecto Python en tu sistema.
    echo Por favor descarga e instala Python desde: https://www.python.org/downloads/
    echo IMPORTANTE: Marca la casilla "Add python.exe to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

echo [2/3] Verificando dependencias (Streamlit, Pandas)...
python -c "import streamlit, pandas" >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando componentes necesarios (solo la primera vez)...
    pip install streamlit pandas pdfplumber python-docx --quiet
)

echo [3/3] Abriendo Organizador Inteligente...
echo.
echo ========================================================
echo  La aplicacion abrira automaticamente en tu navegador!
echo  Deja esta ventana abierta mientras uses la app.
echo  Para cerrarla, presiona Ctrl+C o cierra esta ventana.
echo ========================================================
echo.

python -m streamlit run "%~dp0app.py"
pause

