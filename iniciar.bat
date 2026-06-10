@echo off
title Gerenciador de PDFs
echo ================================================
echo   Gerenciador de PDFs - Iniciando...
echo ================================================
echo.
cd /d "%~dp0"
pip install flask pandas openpyxl xlrd -q
echo.
echo  Acesse no navegador: http://localhost:5000
echo  Para encerrar, feche esta janela.
echo.
python app.py
pause
