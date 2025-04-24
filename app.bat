@echo off
echo Iniciando aplicação Flask...

:: (1) Ativa o ambiente virtual (se estiver usando)
call app_qualidade\Scripts\activate

:: (2) Define variáveis Flask (ajuste se seu app for diferente)
set FLASK_APP=main.py
set FLASK_ENV=development

:: (3) Executa o Flask
python main.py

pause
