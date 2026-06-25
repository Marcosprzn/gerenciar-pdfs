@echo off
title Verificador de Mes em PDFs
echo ================================================
echo   Iniciando Verificador de PDFs com OCR...
echo ================================================
echo.
cd /d "%~dp0"

:: Se rodado da pasta de fora, entra automaticamente na pasta oficial
if exist "gerenciador_pdfs\verificar_mes_scans.py" (
    cd gerenciador_pdfs
)

python -c "import PIL, fitz, pytesseract" 2>nul
if errorlevel 1 (
    echo Instalando dependencias necessarias pela primeira vez... aguarde.
    pip install Pillow PyMuPDF pytesseract -q
    echo Dependencias instaladas com sucesso!
    echo.
)

python verificar_mes_scans.py
if errorlevel 1 pause
