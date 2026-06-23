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
        "Indicio de dolo o fraude en la evaluación de créditos": ["Documentos con enmendaduras", "Documentos con datos inconsistentes", "Documentos sin datos del cliente","Documentos sin firmas o que no coinciden","Documentos duplicados en más de un cliente"],
        "Evaluaciones deficientes o con sustento insuficiente": ["No se evidencio sustento de actividad económica", "No se evidencio sustento de ingresos", "No se evidenció sustento de activos representativos","Se omitió al cónyugue"],
        "Créditos reprogramados y refinanciados":["Reprogramado","Refinanciado"],
        "Clientes con créditos con calificación diferente a normal a la fecha de revisión":["Indicar la calificación a la fecha de revisión"]
    }
    
    for cat, items in categorias.items():
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
    st.title("🏦 Buscador de Clientes")
    archivo = st.file_uploader("Cargar Base Excel", type=["xlsx"])
    
    if archivo is not None:
        try:
            # 1. Creamos un objeto ExcelFile (Esto es muy ligero, no gasta RAM)
            xls = pd.ExcelFile(archivo, engine="openpyxl")
            
            # 2. Elegimos la hoja de forma segura
            hoja_objetivo = "MUESTRA_FINAL"
            if hoja_objetivo not in xls.sheet_names:
                st.error(f"No se encuentra la hoja '{hoja_objetivo}'. Hojas disponibles: {xls.sheet_names}")
                return

            # 3. Leemos solo la hoja necesaria y solo las columnas necesarias (CLAVE PARA MÓVIL)
            # Esto evita el desbordamiento de memoria (Error 111)
            columnas_a_leer = ["DOCPEN", "CLIENTE", "DIAS_ATRASO", "ANALISTA", "BCCTA", "PRODUCTO_CAJA", "AGENCIA", "IMPDESEMB_MN", "SALDO_MN", "DIRECCION_DOM", "ACTIVIDAD_ECON"]
            
            df = pd.read_excel(xls, sheet_name=hoja_objetivo, usecols=columnas_a_leer, engine="openpyxl")
            
            # 4. Limpieza rápida
            df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
            st.session_state.df = df.astype(str)
            st.success(f"Hoja '{hoja_objetivo}' cargada con éxito.")
            
        except Exception as e:
            st.error(f"Error técnico: {e}")

def pantalla_ficha():
    c = st.session_state.cliente_actual
    if not c: return
    
    # Dashboard
    atraso = int(c.get("DIAS_ATRASO", 0))
    color = "#10B981" if atraso <= 30 else ("#F59E0B" if atraso <= 60 else "#EF4444")
    
    st.markdown(f'<div style="background-color:{color}; padding:20px; border-radius:15px; color:white; text-align:center;"><h2>{c.get("CLIENTE")}</h2><p>DNI: {c.get("DOCPEN")}</p></div>', unsafe_allow_html=True)
    
    tabs = st.tabs(["📋 Evaluar", "🏠 Domicilio", "💼 Negocio", "💰 Financiero", "📍 GPS", "📸 Fotos", "📄 Reporte"])
    
    with tabs[0]:
        render_panel_riesgos()
    with tabs[1]:
        st.write(f"**Dirección:** {c.get('DIRECCION_DOM', 'N/A')}")
    with tabs[2]:
        st.write(f"**Negocio:** {c.get('ACTIVIDAD_ECON', 'N/A')}")
    with tabs[6]:
        if st.button("Generar Informe Word"): st.info("Generando...")

    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# --- ROUTER ÚNICO ---
if st.session_state.view == "busqueda": pantalla_busqueda()
else: pantalla_ficha()
