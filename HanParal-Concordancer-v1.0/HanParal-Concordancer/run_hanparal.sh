#!/usr/bin/env bash
cd "$(dirname "$0")"

echo "Installing HanParal requirements..."
python3 -m pip install -r requirements.txt

echo "Starting HanParal..."
python3 -m streamlit run app/streamlit_app.py