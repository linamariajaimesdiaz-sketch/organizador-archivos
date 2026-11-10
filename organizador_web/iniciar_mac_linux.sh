#!/bin/bash
echo "========================================================"
echo "   ORGANIZADOR INTELIGENTE DE ARCHIVOS"
echo "========================================================"
echo ""

if ! command -v python3 &> /dev/null; then
    echo "[AVISO] No se encontro python3 en este equipo."
    echo "Por favor instala Python 3 desde https://www.python.org/downloads/"
    exit 1
fi

echo "[1/2] Verificando librerías..."
python3 -c "import streamlit, pandas" 2>/dev/null || pip3 install streamlit pandas pdfplumber python-docx --quiet

echo "[2/2] Iniciando aplicación..."
echo "Se abrirá en tu navegador automáticamente."
echo ""
python3 -m streamlit run "$(dirname "$0")/app.py"

