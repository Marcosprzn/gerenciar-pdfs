@echo off
title Verificador de Mes em PDFs
echo ================================================
echo   Iniciando Verificador de PDFs com OCR...
echo ================================================
echo.
cd /d "%~dp0"

python verificar_mes_scans.py
if errorlevel 1 pause
