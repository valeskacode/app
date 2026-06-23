# app.py
import streamlit as st
from utils.helpers import load_css

# 1. Configuración inicial
st.set_page_config(page_title="App Visitas", layout="centered", initial_sidebar_state="collapsed")

# 2. Cargar estilos
load_css("assets/style.css")

# 3. Estado inicial
if "view" not in st.session_state: 
    st.session_state.view = "busqueda"

# 4. Router (Navegación)
if st.session_state.view == "busqueda":
    st.title("Pantalla de Búsqueda")
    # Aquí iría el llamado a tu función de búsqueda
elif st.session_state.view == "ficha":
    st.title("Ficha de Cliente")
    # Aquí iría el llamado a tu función de ficha
