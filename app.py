import streamlit as st
import pandas as pd
from utils.helpers import load_css

# Configuración
st.set_page_config(page_title="App Visitas", layout="centered", initial_sidebar_state="collapsed")
load_css("assets/style.css")

# Inicialización de estado
if "view" not in st.session_state: st.session_state.view = "busqueda"
if "cliente_actual" not in st.session_state: st.session_state.cliente_actual = None
if "df" not in st.session_state: st.session_state.df = None

# --- Lógica de Pantalla de Búsqueda ---
def pantalla_busqueda():
    st.markdown('<div class="card" style="text-align: center;">', unsafe_allow_html=True)
    st.title("🏦 Visitas")
    
    archivo = st.file_uploader("Cargar Base Excel", type=["xlsx"])
    if archivo:
        df = pd.read_excel(archivo, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        st.session_state.df = df
        st.success("Base cargada exitosamente")

    if st.session_state.df is not None:
        busqueda = st.text_input("Buscar por DNI o Nombre")
        if busqueda:
            res = st.session_state.df[
                st.session_state.df["PENDOC"].str.contains(busqueda, na=False) | 
                st.session_state.df["CLIENTE"].str.contains(busqueda, case=False, na=False)
            ]
            if not res.empty:
                st.write(f"Resultados: {len(res)}")
                if st.button("Cargar Cliente"):
                    st.session_state.cliente_actual = res.iloc[0].to_dict()
                    st.session_state.view = "ficha"
                    st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --- Lógica de Pantalla de Ficha ---
def pantalla_ficha():
    cliente = st.session_state.cliente_actual
    st.markdown(f'<div class="card"><h2>{cliente.get("CLIENTE")}</h2><p>DNI: {cliente.get("PENDOC")}</p></div>', unsafe_allow_html=True)
    
    # Aquí irían tus TABS (Info, Domicilio, etc.)
    st.write("Contenido de la ficha aquí...")
    
    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# --- Router ---
if st.session_state.view == "busqueda":
    pantalla_busqueda()
else:
    pantalla_ficha()
