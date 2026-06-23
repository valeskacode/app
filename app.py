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
    # Usamos un contenedor para organizar mejor el buscador
    st.markdown('<div class="card" style="text-align: center;">', unsafe_allow_html=True)
    st.title("🏦 Buscador de Clientes")
    
    # 1. Carga del archivo
    archivo = st.file_uploader("Cargar Base Excel (MUESTRA_FINAL)", type=["xlsx"])
    
    if archivo is not None:
        try:
            # Forzamos lectura con openpyxl
            df = pd.read_excel(archivo, sheet_name="MUESTRA_FINAL", header=0, engine="openpyxl")
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # Guardamos el df en session_state para que no se pierda
            st.session_state.df = df
            st.success("Base cargada correctamente")
        except Exception as e:
            st.error(f"Error al cargar: {e}")

    # 2. Buscador: Solo aparece si ya cargamos datos
    if st.session_state.df is not None:
        st.write("---")
        busqueda = st.text_input("Ingresa el DNI o Nombre para buscar:")
        
        if busqueda:
            df = st.session_state.df
            # Filtro robusto
            res = df[
                df["DOCPEN"].astype(str).str.contains(busqueda, na=False) | 
                df["CLIENTE"].astype(str).str.contains(busqueda, case=False, na=False)
            ]
            
            if not res.empty:
                st.write(f"Resultados encontrados: {len(res)}")
                # Mostramos una tabla breve de los resultados encontrados
                st.dataframe(res[["DOCPEN", "CLIENTE"]], use_container_width=True)
                
                if st.button("Abrir Ficha del Cliente"):
                    # Seleccionamos la primera coincidencia
                    st.session_state.cliente_actual = res.iloc[0].to_dict()
                    st.session_state.view = "ficha"
                    st.rerun()
            else:
                st.warning("No se encontraron coincidencias para esa búsqueda.")
    else:
        st.info("Por favor, carga el archivo Excel para habilitar la búsqueda.")
        
    st.markdown('</div>', unsafe_allow_html=True)

def pantalla_ficha():
    c = st.session_state.cliente_actual
    
    # 1. Dashboard de Estado Superior
    atraso = int(c.get("DIAS_ATRASO", 0))
    # Lógica de color: Verde (<30), Amarillo (31-60), Rojo (>60)
    color_riesgo = "#10B981" if atraso <= 30 else ("#F59E0B" if atraso <= 60 else "#EF4444")
    
    st.markdown(f"""
        <div style="background-color:{color_riesgo}; padding:20px; border-radius:15px; color:white; text-align:center;">
            <h2 style="margin:0;">{c.get("CLIENTE")}</h2>
            <p style="font-size:1.2rem; margin:5px 0;">DNI: {c.get("DOCPEN")}</p>
            <div style="font-weight:bold; font-size:1.1rem;">Días de Atraso: {atraso}</div>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("---") 

    # 2. Expander para detalles técnicos (Cerrado por defecto)
    with st.expander("ℹ️ Ver detalles técnicos del crédito"):
        col1, col2 = st.columns(2)
        col1.write(f"**Analista:** {c.get('ANALISTA')}")
        col2.write(f"**Cuenta:** {c.get('BCCTA')}")
        st.write(f"**Producto:** {c.get('PRODUCTO_CAJA')}")
        st.write(f"**Agencia:** {c.get('AGENCIA')}")

    # 3. Inputs de lectura rápida (Fuentes grandes)
    st.subheader("Control Financiero")
    # Usamos valores convertidos a float para evitar errores
    st.number_input("Importe Desembolsado (S/)", value=float(c.get("IMPDESEMB_MN", 0)), format="%.2f")
    st.number_input("Saldo Actual (S/)", value=float(c.get("SALDO_MN", 0)), format="%.2f")

    # 4. TABS ÚNICOS (Estructura jerárquica)
    tabs = st.tabs(["📋 Evaluar", "🏠 Domicilio", "💼 Negocio", "💰 Financiero", "📍 GPS", "📸 Fotos", "📄 Reporte"])
    
    with tabs[0]: # Evaluar
        render_panel_riesgos()
        if st.button("Guardar Evaluación"): 
            st.success("Evaluación guardada exitosamente")
            
    with tabs[1]: # Domicilio
        st.write(f"**Dirección:** {c.get('DIRECCION_DOM')}")
        
    with tabs[2]: # Negocio
        st.write(f"**Actividad:** {c.get('ACTIVIDAD_ECON')}")
        
    with tabs[3]: # Financiero
        st.metric("Saldo Total", f"S/ {c.get('SALDO_MN', '0')}")
        
    with tabs[4]: # GPS
        st.write("Funcionalidad GPS en desarrollo...")
        
    with tabs[5]: # Fotos
        st.file_uploader("Subir evidencias", accept_multiple_files=True)
        
    with tabs[6]: # Reporte
        if st.button("Generar Informe Word"): 
            st.info("Generando informe automático...")

    # 5. Botón de navegación global (Fuera de los tabs)
    st.divider()
    if st.button("← Volver a Búsqueda"):
        st.session_state.cliente_actual = None
        st.session_state.view = "busqueda"
        st.rerun()

# --- ROUTER (El final del archivo) ---
if st.session_state.view == "busqueda": 
    pantalla_busqueda()
else: 
    pantalla_ficha()

