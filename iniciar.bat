@echo off
title Reporte Operativo LiveOps
echo.
echo  =========================================
echo   REPORTE OPERATIVO - LiveOps PedidosYa
echo  =========================================
echo.

:: Verifica si el entorno virtual existe
IF NOT EXIST ".venv\Scripts\activate.bat" (
    echo [SETUP] Creando entorno virtual...
    python -m venv .venv
    echo [SETUP] Instalando dependencias...
    .venv\Scripts\pip install -r requirements.txt --quiet
    echo [OK] Dependencias instaladas.
)

echo [INFO] Iniciando aplicacion...
echo [INFO] Abre tu navegador en: http://localhost:8501
echo.
.venv\Scripts\streamlit run reporte_operativo.py --server.headless false

pause
