import streamlit as st
import pandas as pd
from utils.helpers import load_css

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="App Visitas", layout="centered", initial_sidebar_state="collapsed")
load_css("assets/style.css")

# --- ESTADO INICIAL ---
if "view" not in st.session_state: st.session_state.view = "busqueda"
if "cliente_actual" not in st.session_state: st.session_state.cliente_actual = None
if "df" not in st.session_state: st.session_state.df = None

# --- FUNCIONES DE PANTALLA ---

def pantalla_busqueda():
    st.markdown('<div class="card" style="text-align: center;">', unsafe_allow_html=True)
    st.title("🏦 Buscador de Clientes")
    
    archivo = st.file_uploader("Cargar Base Excel (MUESTRA_FINAL)", type=["xlsx"])
    
    if archivo:
        try:
            df = pd.read_excel(archivo, sheet_name="MUESTRA_FINAL", header=0, dtype=str)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            if "DOCPEN" in df.columns and "CLIENTE" in df.columns:
                st.session_state.df = df
                st.success("Base cargada correctamente")
            else:
                st.error("Error: Columnas DOCPEN o CLIENTE no encontradas.")
        except Exception as e:
            st.error(f"Error al leer archivo: {e}")

    if st.session_state.df is not None:
        busqueda = st.text_input("Buscar por DNI o Nombre")
        if busqueda:
            res = st.session_state.df[
                st.session_state.df["DOCPEN"].str.contains(busqueda, na=False) | 
                st.session_state.df["CLIENTE"].str.contains(busqueda, case=False, na=False)
            ]
            if not res.empty:
                st.write(f"Resultados: {len(res)}")
                if st.button("Abrir Ficha de Cliente"):
                    st.session_state.cliente_actual = res.iloc[0].to_dict()
                    st.session_state.view = "ficha"
                    st.rerun()
            else:
                st.warning("No se encontraron coincidencias.")
    st.markdown('</div>', unsafe_allow_html=True)

def pantalla_ficha():
    c = st.session_state.cliente_actual
    
    st.markdown(f'<div class="card"><h1>{c.get("CLIENTE")}</h1><p>DNI: {c.get("DOCPEN")}</p></div>', unsafe_allow_html=True)
    
    tabs = st.tabs(["📋 Evaluar", "🏠 Domicilio", "💼 Negocio", "💰 Financiero", "📍 GPS", "📸 Fotos", "📄 Reporte"])
    
    with tabs[0]: # Pestaña de Evaluación (Mockup de Dictamen)
        st.subheader("Evaluación de Crédito")
        c1, c2 = st.columns(2)
        c1.metric("Categoría", c.get("CATEG_RESULTANTE", "N/A"))
        c2.metric("Atraso", c.get("DIAS_ATRASO", "0"))
        
        with st.form("evaluacion_form"):
            decision = st.selectbox("Decisión sugerida", ["Aprobado", "Observado", "Rechazado"])
            comentarios = st.text_area("Comentarios de la visita")
            if st.form_submit_button("Guardar Evaluación"):
                st.success(f"Dictamen: {decision} guardado.")

    with tabs[1]:
        st.subheader("Domicilio")
        st.write(f"**Dirección:** {c.get('DIRECCION_DOM')}")
        st.write(f"**Distrito:** {c.get('DISTRITO_DOM')}")
        
    with tabs[2]:
        st.subheader("Negocio")
        st.write(f"**Actividad:** {c.get('ACTIVIDAD_ECON')}")
        st.write(f"**Dirección:** {c.get('DIRECCION_NEG')}")
        
    with tabs[3]:
        st.subheader("Situación Financiera")
        st.metric("Saldo Total", f"S/ {c.get('SALDO_MN', '0')}")
        st.bar_chart(pd.DataFrame({"Monto": [float(c.get('SALDO_VIGE',0)), float(c.get('SALDO_VENC',0))]}, index=["Vigente", "Vencido"]))
        
    with tabs[4]:
        st.subheader("Ubicación")
        st.info("Sistema de GPS activo.")
        
    with tabs[5]:
        st.subheader("Evidencias")
        st.file_uploader("Subir foto de visita", accept_multiple_files=True)
        
    with tabs[6]:
        st.subheader("Reporte Word")
        if st.button("Generar Informe"):
            st.info("Generando documento Word...")

    st.divider()
    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# --- ROUTER ---
if st.session_state.view == "busqueda":
    pantalla_busqueda()
else:
    pantalla_ficha()
