@echo off
title Gerenciador de PDFs
echo ================================================
echo   Gerenciador de PDFs - Iniciando...
echo ================================================
echo.
cd /d "%~dp0"

python -c "import flask, watchdog" 2>nul
if errorlevel 1 (
    echo Instalando dependencias necessarias pela primeira vez... aguarde.
    pip install -r requirements.txt -q
    echo Dependencias instaladas com sucesso!
    echo.
)

echo  Acesse no navegador: http://localhost:5000
echo  Para encerrar, feche esta janela.
echo.
python app.py
pause
