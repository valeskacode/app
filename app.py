import streamlit as st
import pandas as pd
from utils.helpers import load_css

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="App Visitas", layout="centered", initial_sidebar_state="collapsed")
load_css("assets/style.css")

# --- ESTADO INICIAL ---
if "view" not in st.session_state: st.session_state.view = "busqueda"
if "cliente_actual" not in st.session_state: st.session_state.cliente_actual = None
if "df" not in st.session_state: st.session_state.df = None

# --- LÓGICA DE PANTALLA ---
def pantalla_busqueda():
    st.markdown('<div class="card" style="text-align: center;">', unsafe_allow_html=True)
    st.title("🏦 Visitas")
    
    archivo = st.file_uploader("Cargar Base Excel", type=["xlsx"])
    
    if archivo:
        try:
            df = pd.read_excel(archivo, sheet_name="MUESTRA_FINAL", header=0, dtype=str)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # --- CORRECCIÓN AQUÍ: Usamos DOCPEN ---
            if "DOCPEN" in df.columns and "CLIENTE" in df.columns:
                st.session_state.df = df
                st.success("Base cargada exitosamente")
            else:
                st.error("Error: No se encontraron las columnas DOCPEN y CLIENTE.")
                st.write("Columnas detectadas:", list(df.columns))
                
        except Exception as e:
            st.error(f"Error al procesar: {e}")

    if st.session_state.df is not None:
        busqueda = st.text_input("Buscar por DNI o Nombre")
        if busqueda:
            df = st.session_state.df
            # --- CORRECCIÓN AQUÍ: Filtramos por DOCPEN ---
            res = df[df["DOCPEN"].str.contains(busqueda, na=False) | 
                     df["CLIENTE"].str.contains(busqueda, case=False, na=False)]
            
            if not res.empty:
                st.write(f"Resultados: {len(res)}")
                if st.button("Cargar Cliente"):
                    st.session_state.cliente_actual = res.iloc[0].to_dict()
                    st.session_state.view = "ficha"
                    st.rerun()
            else:
                st.warning("No se encontró coincidencia.")
    st.markdown('</div>', unsafe_allow_html=True)

def pantalla_ficha():
    cliente = st.session_state.cliente_actual
    if cliente:
        st.markdown(f'<div class="card"><h2>{cliente.get("CLIENTE")}</h2><p>DNI: {cliente.get("DOCPEN")}</p></div>', unsafe_allow_html=True)
    
    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# --- NAVEGACIÓN ---
if st.session_state.view == "busqueda":
    pantalla_busqueda()
else:
    pantalla_ficha()
