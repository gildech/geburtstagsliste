#!/bin/bash

# Pfad zu deinem Projektordner (ohne Dateinamen)
PROJECT_DIR="/Users/milan/Library/Mobile Documents/com~apple~CloudDocs/Gilde"

cd "$PROJECT_DIR" || exit 1

# Streamlit-App starten (nutzt das venv-Python)
"/Users/milan/Library/Mobile Documents/com~apple~CloudDocs/Gilde/venv/bin/python3" -m streamlit run app.py
