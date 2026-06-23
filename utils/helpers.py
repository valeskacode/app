# utils/helpers.py
import streamlit as st
import pandas as pd

def load_css(file_path):
    """Carga el archivo de estilos externo."""
    try:
        with open(file_path) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    except FileNotFoundError:
        st.error(f"No se encontró el archivo: {file_path}")

def clean_dni(value):
    """Limpia y formatea el DNI."""
    return str(value).strip() if pd.notna(value) else ""

def safe_float(value, default=0.0):
    """Convierte a float de forma segura."""
    try:
        return float(value)
    except:
        return default
