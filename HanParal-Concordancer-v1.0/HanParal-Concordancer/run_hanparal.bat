@echo off
cd /d "%~dp0"

echo Installing HanParal requirements...
python -m pip install -r requirements.txt

echo Starting HanParal...
python -m streamlit run app\streamlit_app.py

pause