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

# --- FUNCIONES DE INTERFAZ ---

def render_panel_riesgos():
    """Panel jerárquico para evaluación de campo"""
    st.subheader("Criterio para visita a clientes")
    
    categorias = {
        "Indicio de dolo o fraude en la evaluación de céditos": ["Documentos con enmendaduras", "Documentos con datos inconsistentes", "Documentos sin datos del cliente","Documentos sin firmas o que no coinciden","Documentos duplicados en más de un cliente"],
        "Evaluaciones deficientes o con sustento insuficiente": ["No se evidencio sustento de actividad económica", "No se evidencio sustento de ingresos", "No se evidenció sustento de activos representativos","Se omitió al cónyugue"],
        "Créditos reprogramados y refinanciados":["Reprogramado","Refinanciado"],
        "Clientes con créditos con calificación diferente a normal a la fecha de revisión":["Indicar la calificación a la fecha de revisión"]
    
    }
    
    for cat, items in categorias.items():
        # Lógica para cambio de color del contenedor
        riesgo_activo = any(st.session_state.get(f"chk_{item.replace(' ', '_')}", False) for item in items)
        
        bg_color = "#FEF2F2" if riesgo_activo else "#F8FAFC"
        border_color = "#DC2626" if riesgo_activo else "#E2E8F0"
        
        with st.expander(f"🔴 {cat}" if riesgo_activo else f"🟢 {cat}", expanded=True):
            st.markdown(f'<div style="background-color:{bg_color}; padding:15px; border-left: 5px solid {border_color}; border-radius: 8px;">', unsafe_allow_html=True)
            for item in items:
                key = f"chk_{item.replace(' ', '_')}"
                st.checkbox(item, key=key)
            st.markdown("</div>", unsafe_allow_html=True)

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
            st.error(f"Error al leer: {e}")

    if st.session_state.df is not None:
        busqueda = st.text_input("Buscar por DNI o Nombre")
        if busqueda:
            res = st.session_state.df[
                st.session_state.df["DOCPEN"].str.contains(busqueda, na=False) | 
                st.session_state.df["CLIENTE"].str.contains(busqueda, case=False, na=False)
            ]
            if not res.empty:
                if st.button("Abrir Ficha"):
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
    
    with tabs[0]:
        render_panel_riesgos()
        if st.button("Guardar Evaluación"): st.success("Datos guardados")
    with tabs[1]:
        st.write(f"**Dirección:** {c.get('DIRECCION_DOM')}")
    with tabs[2]:
        st.write(f"**Actividad:** {c.get('ACTIVIDAD_ECON')}")
    with tabs[3]:
        st.metric("Saldo Total", f"S/ {c.get('SALDO_MN', '0')}")
    with tabs[4]:
        st.write("Funcionalidad GPS...")
    with tabs[5]:
        st.file_uploader("Subir evidencias", accept_multiple_files=True)
    with tabs[6]:
        if st.button("Generar Informe Word"): st.info("Generando...")

    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# --- ROUTER ---
if st.session_state.view == "busqueda": pantalla_busqueda()
else: pantalla_ficha()
