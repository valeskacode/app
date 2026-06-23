import io
import pandas as pd
import streamlit as st
from docx import Document
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="App Visitas", layout="centered", initial_sidebar_state="collapsed")

# --- CSS PARA ESTILO MÓVIL (IDÉNTICO A MOCKUPS) ---
st.markdown("""
<style>
    .stApp { background-color: #F8FAFC; }
    .block-container { max-width: 450px !important; padding: 1rem !important; }
    .card { 
        background: white; padding: 20px; border-radius: 20px; 
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); margin-bottom: 15px; 
    }
    div.stButton > button { 
        width: 100%; border-radius: 12px; height: 50px; 
        background-color: #D31118; color: white; font-weight: bold; border: none;
    }
    .stTextInput input, .stSelectbox select { border-radius: 10px !important; }
    header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# --- ESTADO DE NAVEGACIÓN ---
if "view" not in st.session_state: st.session_state.view = "busqueda"
if "cliente" not in st.session_state: st.session_state.cliente = None

# --- LÓGICA DE VISTAS ---

def vista_busqueda():
    st.markdown('<div class="card"><h2>🏦 Búsqueda y Carga</h2>', unsafe_allow_html=True)
    archivo = st.file_uploader("Subir base de datos", type="xlsx")
    if archivo:
        st.session_state.df = pd.read_excel(archivo)
        st.success("Base de datos cargada")
    
    busqueda = st.text_input("Ingresar DNI o Nombre")
    if busqueda and "df" in st.session_state:
        res = st.session_state.df[st.session_state.df['CLIENTE'].str.contains(busqueda, case=False)]
        if not res.empty:
            if st.button("Cargar Cliente"):
                st.session_state.cliente = res.iloc[0].to_dict()
                st.session_state.view = "ficha"
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def vista_ficha():
    c = st.session_state.cliente
    st.markdown(f'<div class="card"><h2>👤 {c["CLIENTE"]}</h2>', unsafe_allow_html=True)
    st.write(f"**DNI:** {c['PENDOC']}")
    st.write(f"**Saldo:** S/. {c['SALDO_MN']}")
    if st.button("Ir a Evaluación"):
        st.session_state.view = "evaluacion"
        st.rerun()
    if st.button("Atrás", type="secondary"):
        st.session_state.view = "busqueda"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def vista_evaluacion():
    st.markdown('<div class="card"><h2>📋 Evaluación</h2>', unsafe_allow_html=True)
    # Aquí irían tus checkboxes de riesgos
    st.checkbox("Documentos enmiendas")
    st.checkbox("Sustento de ingresos")
    if st.button("Continuar a Ubicación"):
        st.session_state.view = "ubicacion"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def vista_ubicacion():
    st.markdown('<div class="card"><h2>📍 Ubicación</h2>', unsafe_allow_html=True)
    st.camera_input("Foto de fachada")
    if st.button("Finalizar Reporte"):
        st.session_state.view = "reporte"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --- RUTEO PRINCIPAL ---
if st.session_state.view == "busqueda": vista_busqueda()
elif st.session_state.view == "ficha": vista_ficha()
elif st.session_state.view == "evaluacion": vista_evaluacion()
elif st.session_state.view == "ubicacion": vista_ubicacion()
