@echo off
rem Abre o Planejador MRP/PCP no seu computador (precisa do Python instalado: python.org, marcando "Add Python to PATH").
cd /d %~dp0
if not exist .venv (
  echo Preparando pela primeira vez - pode levar alguns minutos...
  py -m venv .venv
)
call .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
pause
