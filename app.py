# app.py
import streamlit as st
import pandas as pd
from utils.helpers import load_css

# Configuración
st.set_page_config(page_title="App Visitas", layout="centered", initial_sidebar_state="collapsed")
load_css("assets/style.css")

if "view" not in st.session_state: st.session_state.view = "busqueda"
if "cliente_actual" not in st.session_state: st.session_state.cliente_actual = None
if "df" not in st.session_state: st.session_state.df = None

def pantalla_busqueda():
    st.markdown('<div class="card" style="text-align: center;">', unsafe_allow_html=True)
    st.title("🏦 Visitas")
    
    archivo = st.file_uploader("Cargar Base Excel", type=["xlsx"])
    
    if archivo:
        # Leemos el Excel
        df = pd.read_excel(archivo, dtype=str)
        # Limpiamos nombres de columnas (espacios y mayúsculas)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # --- AQUÍ ESTÁ LA CORRECCIÓN ---
        # Verificamos si PENDOC existe
        if "PENDOC" in df.columns and "CLIENTE" in df.columns:
            st.session_state.df = df
            st.success("Base cargada correctamente")
        else:
            st.error(f"Error: El archivo no tiene las columnas requeridas.")
            st.write(f"Columnas detectadas en el Excel: {list(df.columns)}")
            st.info("Asegúrate de que tu Excel tenga los encabezados: PENDOC y CLIENTE")

    if st.session_state.df is not None:
        busqueda = st.text_input("Buscar por DNI o Nombre")
        if busqueda:
            # Filtramos usando las columnas limpias
            df_filtrado = st.session_state.df[
                st.session_state.df["PENDOC"].str.contains(busqueda, na=False) | 
                st.session_state.df["CLIENTE"].str.contains(busqueda, case=False, na=False)
            ]
            
            if not df_filtrado.empty:
                st.write(f"Resultados encontrados: {len(df_filtrado)}")
                if st.button("Cargar Cliente"):
                    st.session_state.cliente_actual = df_filtrado.iloc[0].to_dict()
                    st.session_state.view = "ficha"
                    st.rerun()
            else:
                st.warning("No se encontraron resultados.")
                
    st.markdown('</div>', unsafe_allow_html=True)

def pantalla_ficha():
    cliente = st.session_state.cliente_actual
    if cliente:
        st.markdown(f'<div class="card"><h2>{cliente.get("CLIENTE")}</h2><p>DNI: {cliente.get("PENDOC")}</p></div>', unsafe_allow_html=True)
    
    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# Router
if st.session_state.view == "busqueda":
    pantalla_busqueda()
else:
    pantalla_ficha()
