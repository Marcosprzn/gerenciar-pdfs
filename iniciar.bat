@echo off
title Gerenciador de PDFs
echo ================================================
echo   Gerenciador de PDFs - Iniciando...
echo ================================================
echo.
cd /d "%~dp0"

:: Se rodado da pasta de fora, entra automaticamente na pasta oficial
if exist "gerenciador_pdfs\app.py" (
    cd gerenciador_pdfs
)
if not exist "app.py" (
    echo ERRO: Arquivo app.py nao encontrado. Certifique-se de estar na pasta correta.
    pause
    exit /b
)

python -c "import flask, watchdog, pandas, openpyxl" 2>nul
if errorlevel 1 (
    echo Instalando dependencias necessarias pela primeira vez... aguarde.
    pip install -r requirements.txt -q
    echo Dependencias instaladas com sucesso!
    echo.
)

echo  Acesse no navegador: http://localhost:5000
echo  Para encerrar o servidor pelo navegador, use o botao "Sair" na interface.
echo.
python app.py
if errorlevel 1 pause
