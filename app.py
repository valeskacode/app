# -*- coding: utf-8 -*-
"""
Formulario de verificación de datos visita - Rediseño Premium Mobile-First
Basado exactamente en las maquetas visuales del sistema de evaluación.
"""

import io
import json
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

try:
    from streamlit_js_eval import get_geolocation
    GEO_OK = True
except Exception:
    GEO_OK = False

# --------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL Y ESTILOS (UI MÓVIL PREMIUM)
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Evaluación de Crédito",
    page_icon="🏦",
    layout="centered",  # Forzar contenedor angosto ideal para móviles/emulación
    initial_sidebar_state="collapsed",
)

EXCEL_COLUMNS = [
    "RECNO", "PEPAIS", "PETDOC", "PENDOC", "CODCLI", "BCEMP", "BCSUC", "BCMDA",
    "BCPAP", "BCCTA", "BCOPER", "BCSBOP", "BCTOP", "BCMOD", "CODCRE", "REGION",
    "ZONA", "AGENCIA", "CLIENTE", "DIRECCION_DOM", "DISTRITO_DOM",
    "PROVINCIA_DOM", "DEPARTAMENTO_DOM", "DIRECCION_NEG", "DISTRITO_NEG",
    "PROVINCIA_NEG", "DEPARTAMENTO_NEG", "ACTIVIDAD_ECON", "ANALISTA",
    "PRODUCTO_CAJA", "SALDO_MN", "SALDO_VIGE", "SALDO_REFI", "SALDO_VENC",
    "SALDO_JUDI", "MORA_CONT", "TIPO_SBS", "FECDES", "IMPDESEMB_MN",
    "COD_MODULO", "MODULO", "COD_TIPO_OPERACION", "TIPO_OPERACION",
    "ANALISTA_EVAL", "USUARIO_APROB", "USUARIO_DESEM", "FECHA_EVAL",
    "DIAS_ATRASO", "ESTADO_CREDITO", "ATRANT_1M", "ATRANT_2M", "ATRANT_3M",
    "ATRANT_4M", "ATRANT_5M", "ATRANT_6M", "TIPO_SOLI", "NUMERO_CUOTAS",
    "CUOTAS_PAGADAS", "TIPO", "SEGMENTACION_MYPE", "CATEG_RESULTANTE",
    "CATEG_RESULTANTE_SINALIN", "CUENTA_AVAL", "FECHA_UTLPAGO", "UAI_IND",
    "ESTRATO", "TIPO_EXPEDIENTE",
]

# Paleta de Colores de las Maquetas Adjuntas
ROJO_INST = "D31118"
AZUL_TEXTO = "0F172A"
GRIS_BG = "F8FAFC"
VERDE_EXITO = "16A34A"

PREMIUM_CSS = f"""
<style>
/* Contenedor Base Móvil */
.block-container {{
    max-width: 460px !important;
    padding-top: 1rem !important;
    padding-bottom: 6rem !important;
    padding-left: 14px !important;
    padding-right: 14px !important;
}}

.stApp {{ background-color: {GRIS_BG}; }}

/* Quitar elementos por defecto de Streamlit innecesarios en App */
#MainMenu, header, footer {{ visibility: hidden; }}

/* Tarjetas Estilo iOS / Android Moderno */
.app-card {{
    background: #FFFFFF;
    padding: 1.2rem;
    border-radius: 16px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
    margin-bottom: 1rem;
}}

/* Perfil Destacado del Cliente */
.profile-banner {{
    background: #F0FDF4;
    border: 1px solid #DCFCE7;
    border-radius: 16px;
    padding: 1.2rem;
    margin-bottom: 1.2rem;
}}

/* Indicadores de Estado Rápidos */
.status-pill {{
    background: #E2E8F0;
    color: #475569;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
    display: inline-block;
}}
.status-pill-active {{ background: #DCFCE7; color: #15803D; }}
.status-pill-risk {{ background: #FEE2E2; color: #991B1B; }}

/* Botón Principal de Acción (Grande, Táctil e Inferior) */
div.stButton > button {{
    background-color: {ROJO_INST} !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 14px 20px !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    width: 100% !important;
    box-shadow: 0 4px 12px rgba(211, 17, 24, 0.2) !important;
    transition: all 0.2s ease;
}}
div.stButton > button:hover {{ transform: translateY(-1px); }}

/* Botón de Retroceso / Secundario */
div.stButton > button[key^="btn_back"] {{
    background-color: #FFFFFF !important;
    color: #475569 !important;
    border: 1px solid #CBD5E1 !important;
    box-shadow: none !important;
    margin-top: 0.5rem;
}}

/* Bloque Desplegable Interno de Criterios de Riesgo */
.risk-row {{
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 10px 14px;
    margin-bottom: 6px;
    display: flex;
    align-items: center;
}}

/* Inputs del Sistema Estilizados */
.stTextInput input, .stNumberInput input, .stSelectbox select {{
    border-radius: 10px !important;
    border: 1px solid #CBD5E1 !important;
    background: #FFFFFF !important;
}}

/* Barra de Navegación Inferior Simulada Fija */
.bottom-nav {{
    position: fixed;
    bottom: 0;
    left: 0;
    right: 0;
    background: #FFFFFF;
    border-top: 1px solid #E2E8F0;
    padding: 10px 0;
    display: flex;
    justify-content: space-around;
    z-index: 999;
}}
</style>
"""
st.markdown(PREMIUM_CSS, unsafe_allow_html=True)

# --------------------------------------------------------------------------
# HELPERS Y ESTADO DE SESIÓN (FLOW SEQUENTIAL DE VISTAS)
# --------------------------------------------------------------------------
def safe_str(v, default=""):
    if v is None or pd.isna(v): return default
    s = str(v).strip()
    return default if s.lower() in ("nan", "none") else s

def safe_float(v, default=0.0):
    try:
        f = float(v)
        return default if pd.isna(f) else f
    except Exception: return default

def fmt_money(v):
    return f"S/. {safe_float(v):,.2f}"

def limpiar_texto_dni(val):
    if pd.isna(val) or val is None: return ""
    txt = str(val).strip()
    if txt.endswith(".0"): txt = txt[:-2]
    if txt.isdigit() and len(txt) < 8 and len(txt) > 0: txt = txt.zfill(8)
    return txt

# Estados de Navegación Móvil
if "current_view" not in st.session_state: st.session_state.current_view = "carga"
if "clientes_df" not in st.session_state: st.session_state.clientes_df = None
if "cliente_actual" not in st.session_state: st.session_state.cliente_actual = {}
if "visitas" not in st.session_state: st.session_state.visitas = {}
if "validaciones_marcadas" not in st.session_state: st.session_state.validaciones_marcadas = {}
if "click_timestamps" not in st.session_state: st.session_state.click_timestamps = {}
if "garantias" not in st.session_state: st.session_state.garantias = []
if "rcc" not in st.session_state: st.session_state.rcc = []

cliente = st.session_state.cliente_actual

# --------------------------------------------------------------------------
# LÓGICA DE VALIDACIONES DE RIESGO AUTOMÁTICAS
# --------------------------------------------------------------------------
def validar_visita():
    validaciones = {
        "documentos_enmiendas": False, "documentos_inconsistentes": False,
        "documentos_sin_datos": False, "documentos_sin_firmas": False,
        "documentos_duplicados": False, "sin_sustento_actividad": False,
        "sin_sustento_ingresos": False, "sin_sustento_activos": False,
        "conyuge_omitido": False, "credito_reprogramado": False,
        "credito_refinanciado": False, "calificacion_diferente": False,
    }
    if not safe_str(cliente.get("CLIENTE")): validaciones["documentos_sin_datos"] = True
    if safe_str(cliente.get("DIAS_ATRASO")) and int(safe_float(cliente.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
    return validaciones


# ==========================================================================
# 📱 VISTA 1: BÚSQUEDA Y CARGA DE BASE DE DATOS
# ==========================================================================
if st.session_state.current_view == "carga":
    st.markdown(f"<h2 style='text-align:center;font-size:1.4rem;margin-bottom:1.5rem;color:#0F172A;'>🏦 Evaluación de Crédito</h2>", unsafe_allow_html=True)
    
    # Tarjeta de Carga de Archivo
    st.markdown('<div class="app-card">', unsafe_allow_html=True)
    st.markdown("<p style='font-weight:600;margin-bottom:4px;color:#1E293B;'>Carga de Base de Datos</p>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.8rem;color:#64748B;margin-bottom:12px;'>Selecciona tu archivo Excel con la cartera de clientes.</p>", unsafe_allow_html=True)
    
    excel_file = st.file_uploader("Elegir archivo (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    
    if excel_file is not None:
        try:
            excel_lector = pd.ExcelFile(excel_file)
            hoja_objetivo = "MUESTRA_FINAL" if "MUESTRA_FINAL" in excel_lector.sheet_names else excel_lector.sheet_names[0]
            df_cargado = pd.read_excel(excel_file, sheet_name=hoja_objetivo, dtype=str)
            df_cargado.columns = [str(c).strip().upper() for c in df_cargado.columns]
            
            if len(df_cargado.columns) >= 4:
                df_cargado = df_cargado.rename(columns={df_cargado.columns[3]: "PENDOC"})
            if "PENDOC" in df_cargado.columns:
                df_cargado["PENDOC"] = df_cargado["PENDOC"].apply(limpiar_texto_dni)
                
            st.session_state.clientes_df = df_cargado
        except Exception as e:
            st.error(f"Error: {e}")
            
    if st.session_state.clientes_df is not None:
        st.markdown(f"""
        <div style='background:#F0FDF4; border:1px solid #BBF7D0; border-radius:10px; padding:10px; text-align:center; margin-top:8px;'>
            <span style='color:#16A34A; font-weight:600; font-size:0.9rem;'>✓ Archivo procesado correctamente</span><br>
            <span style='font-size:1.4rem; font-weight:700; color:#14532D;'>{len(st.session_state.clientes_df):,}</span>
            <p style='font-size:0.75rem; color:#166534; margin:0;'>Registros cargados en el sistema</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Tarjeta de Búsqueda Inteligente
    st.markdown('<div class="app-card">', unsafe_allow_html=True)
    st.markdown("<p style='font-weight:600;margin-bottom:2px;color:#1E293B;'>Búsqueda Inteligente</p>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.8rem;color:#64748B;margin-bottom:12px;'>Encuentra a tu cliente por nombre, DNI o código.</p>", unsafe_allow_html=True)
    
    busq = st.text_input("Buscar por datos...", placeholder="Ej. Perez Garcia o DNI")
    
    df = st.session_state.clientes_df
    if df is not None and busq:
        b = busq.strip().lower()
        mask = (
            df.get("PENDOC", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False) |
            df.get("CODCLI", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False) |
            df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        )
        resultados = df[mask]
        
        if len(resultados) > 0:
            opciones = resultados.apply(lambda r: f"{safe_str(r.get('PENDOC'))} - {safe_str(r.get('CLIENTE'))}", axis=1).tolist()
            sel = st.selectbox("Coincidencias encontradas en tiempo real:", opciones)
            if sel:
                idx_sel = opciones.index(sel)
                st.session_state.temp_client_dict = resultados.iloc[idx_sel].to_dict()
                
                # Botón de confirmación idéntico al diseño de la maqueta
                if st.button("✓ Confirmar este cliente"):
                    st.session_state.cliente_actual = st.session_state.temp_client_dict
                    st.session_state.visitas = {}
                    st.session_state.validaciones_marcadas = {}
                    st.session_state.current_view = "ficha"
                    st.rerun()
        else:
            st.markdown("<p style='font-size:0.85rem;color:#EF4444;text-align:center;'>No encontramos coincidencias exactas.</p>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ==========================================================================
# 📱 VISTA 2: FICHA DE IDENTIDAD (CLIENTE Y CRÉDITO)
# ==========================================================================
elif st.session_state.current_view == "ficha":
    st.markdown("<h3 style='text-align:center;font-size:1.2rem;margin-bottom:1rem;color:#0F172A;'>👤 Cliente y Crédito</h3>", unsafe_allow_html=True)
    
    # Banner Principal de Identidad
    mora = int(safe_float(cliente.get("DIAS_ATRASO")))
    mora_style = "status-pill-risk" if mora > 0 else "status-pill-active"
    mora_txt = "Riesgo Alto" if mora > 0 else "Riesgo Bajo"
    
    st.markdown(f"""
    <div class="profile-banner">
        <span class="{mora_style}" style="float:right; margin-top:2px;">{mora_txt}</span>
        <small style="color:#64748B; font-weight:500;">CLIENTE</small>
        <div style="font-size:1.2rem; font-weight:700; color:#0F172A; margin-bottom:4px;">{safe_str(cliente.get('CLIENTE'))}</div>
        <span style="font-size:0.85rem; color:#475569;"><b>DNI:</b> {safe_str(cliente.get('PENDOC'))}</span> &nbsp;|&nbsp; 
        <span style="font-size:0.85rem; color:#475569;"><b>Atraso:</b> {mora} días</span>
    </div>
    """, unsafe_allow_html=True)

    # Bloque de Información del Crédito Numérico
    st.markdown('<div class="app-card">', unsafe_allow_html=True)
    st.markdown("<p style='font-weight:700; font-size:0.9rem; margin-bottom:10px; color:#1E293B;'>📋 Información del Crédito</p>", unsafe_allow_html=True)
    
    sub_c1, sub_c2 = st.columns(2)
    with sub_c1:
        st.markdown(f"<small style='color:#64748B;'>Importe Original</small><br><span style='font-size:1.25rem; font-weight:700; color:#0F172A;'>{fmt_money(cliente.get('IMPDESEMB_MN'))}</span>", unsafe_allow_html=True)
    with sub_c2:
        st.markdown(f"<small style='color:#64748B;'>Saldo Actual</small><br><span style='font-size:1.25rem; font-weight:700; color:#{VERDE_EXITO};'>{fmt_money(cliente.get('SALDO_MN'))}</span>", unsafe_allow_html=True)
    
    st.markdown("<hr style='margin:12px 0; border:0; border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)
    
    st.text_input("N° de Cuenta", value=safe_str(cliente.get("BCCTA")), disabled=True)
    st.text_input("Tipo de Crédito", value=safe_str(cliente.get("PRODUCTO_CAJA")), disabled=True)
    st.text_input("Fecha de Desembolso", value=safe_str(cliente.get("FECDES")), disabled=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Campos Desplegables de Gestión Administrativa
    with st.expander("⚙️ Información Administrativa y de Gestión", expanded=False):
        st.text_input("Analista Asignado", value=safe_str(cliente.get("ANALISTA")), disabled=True)
        st.text_input("Zona / Sector", value=safe_str(cliente.get("ZONA")), disabled=True)
        st.text_input("Calificación Interna", value=safe_str(cliente.get("CATEG_RESULTANTE")), disabled=True)

    # Botonera de Acción Inferior Fija
    if st.button("Guardar y continuar ➡️"):
        st.session_state.current_view = "criterios"
        st.rerun()
        
    if st.button("⬅️ Cambiar de Cliente", key="btn_back_to_carga"):
        st.session_state.current_view = "carga"
        st.session_state.cliente_actual = {}
        st.rerun()


# ==========================================================================
# 📱 VISTA 3: CRITERIOS PARA VISITA A CLIENTES (RIESGOS)
# ==========================================================================
elif st.session_state.current_view == "criterios":
    st.markdown("<h3 style='text-align:center;font-size:1.2rem;margin-bottom:0.2rem;color:#0F172A;'>📝 Evaluación de Crédito</h3>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.8rem;color:#64748B;text-align:center;margin-bottom:1.2rem;'>Seleccione los criterios identificados durante la visita.</p>", unsafe_allow_html=True)
    
    validaciones_auto = validar_visita()
    
    criterios_maqueta = {
        "🔴 Indicio de dolo o fraude": [
            ("documentos_enmiendas", "Documentos con enmendaduras"),
            ("documentos_inconsistentes", "Documentos con datos inconsistentes"),
            ("documentos_sin_datos", "Documentos sin datos del cliente"),
            ("documentos_sin_firmas", "Documentos sin firmas o que no coinciden"),
            ("documentos_duplicados", "Documentos duplicados en más de un cliente")
        ],
        "⚠️ Evaluaciones deficientes o sustento insuficiente": [
            ("sin_sustento_actividad", "No se evidenció sustento de actividad económica"),
            ("sin_sustento_ingresos", "No se evidenció sustento de ingresos"),
            ("sin_sustento_activos", "No se evidenció sustento de activos representativos"),
            ("conyuge_omitido", "Se omitió al cónyuge")
        ],
        "🔄 Créditos reprogramados y refinanciados": [
            ("credito_reprogramado", "Reprogramado"),
            ("credito_refinanciado", "Refinanciado")
        ]
    }
    
    # Construcción de Bloques Colapsables Idénticos a la Interfaz Gráfica Solicitada
    for bloque_titulo, items in criterios_maqueta.items():
        with st.expander(bloque_titulo, expanded=True):
            for key, label in items:
                # Chequeo dinámico en session_state
                is_active = st.session_state.validaciones_marcadas.get(key, validaciones_auto.get(key, False))
                
                # Checkbox móvil estilizado nativo de Streamlit
                check_val = st.checkbox(label, value=is_active, key=f"chk_{key}")
                if check_val != is_active:
                    st.session_state.validaciones_marcadas[key] = check_val
                    st.rerun()

    # Navegación del flujo móvil
    if st.button("Ir al Registro de Ubicación ➡️"):
        st.session_state.current_view = "ubicacion"
        st.rerun()
        
    if st.button("⬅️ Regresar a Ficha", key="btn_back_to_ficha"):
        st.session_state.current_view = "ficha"
        st.rerun()


# ==========================================================================
# 📱 VISTA 4: NUEVA VISITA (CAPTURA DE FOTO Y UBICACIÓN GPS)
# ==========================================================================
elif st.session_state.current_view == "ubicacion":
    st.markdown("<h3 style='text-align:center;font-size:1.2rem;margin-bottom:0.2rem;color:#0F172A;'>📍 Nueva Visita</h3>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.8rem;color:#64748B;text-align:center;margin-bottom:1.2rem;'>Captura tu ubicación actual y toma una foto de la fachada.</p>", unsafe_allow_html=True)
    
    entorno = st.radio("Entorno a verificar:", ["Domicilio", "Negocio"], horizontal=True)
    clave_v = entorno.lower()
    
    # Card de Captura Fotográfica Obligatoria
    st.markdown('<div class="app-card" style="text-align:center;">', unsafe_allow_html=True)
    st.markdown("<p style='font-weight:600; font-size:0.85rem; text-align:left; margin-bottom:8px; color:#475569;'>FOTO DE VERIFICACIÓN (OBLIGATORIA)</p>", unsafe_allow_html=True)
    
    f_cam = st.camera_input("Capturar instantánea", key=f"camera_{clave_v}", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Card de Geoposicionamiento
    st.markdown('<div class="app-card">', unsafe_allow_html=True)
    st.markdown("<p style='font-weight:600; font-size:0.85rem; margin-bottom:8px; color:#475569;'>UBICACIÓN ACTUAL REAL</p>", unsafe_allow_html=True)
    
    visitas_data = st.session_state.visitas.get(clave_v, {})
    lat, lon = visitas_data.get("lat"), visitas_data.get("lon")
    
    if st.button("📡 Capturar Coordenadas GPS (In Situ)", key=f"btn_gps_{clave_v}"):
        if GEO_OK:
            loc = get_geolocation(key=f"geo_proc_{clave_v}_{datetime.now().timestamp()}")
            if loc and "coords" in loc:
                lat, lon = loc["coords"]["latitude"], loc["coords"]["longitude"]
                st.session_state.visitas[clave_v] = {
                    "lat": lat, "lon": lon,
                    "foto_bytes": f_cam.getvalue() if f_cam is not None else None,
                    "fecha": str(datetime.now().date()), "hora": str(datetime.now().time())[:5]
                }
                st.rerun()
        else:
            st.warning("El módulo de posicionamiento global no se encuentra activo.")
            
    if lat and lon:
        st.markdown(f"""
        <div style="background:#F0FDF4; padding:8px; border-radius:8px; font-size:0.8rem; margin-bottom:8px; border:1px solid #BBF7D0; color:#16A34A; text-align:center;">
            <b>Latitud:</b> {lat:.6f} | <b>Longitud:</b> {lon:.6f}
        </div>
        """, unsafe_allow_html=True)
        st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=15, height=160)
    st.markdown('</div>', unsafe_allow_html=True)

    # Finalización del Reporte
    if st.button("Finalizar y Descargar Informe 🏆"):
        st.session_state.current_view = "reporte_final"
        st.rerun()
        
    if st.button("⬅️ Volver a Criterios", key="btn_back_to_crit"):
        st.session_state.current_view = "criterios"
        st.rerun()


# ==========================================================================
# 📱 VISTA 5: REPORTE Y CIERRE (DESCARGA DEL DOCUMENTO)
# ==========================================================================
elif st.session_state.current_view == "reporte_final":
    st.markdown("<h3 style='text-align:center;font-size:1.2rem;margin-bottom:1.2rem;color:#0F172A;'>🏁 Cierre y Reporte</h3>", unsafe_allow_html=True)
    
    st.markdown('<div class="app-card">', unsafe_allow_html=True)
    st.markdown("<p style='font-weight:700; font-size:0.95rem; margin-bottom:8px;'>Resumen de Calidad</p>", unsafe_allow_html=True)
    
    dom_status = "🟢 Domicilio Verificado" if "domicilio" in st.session_state.visitas else "🟡 Domicilio Pendiente"
    neg_status = "🟢 Negocio Verificado" if "negocio" in st.session_state.visitas else "🟡 Negocio Pendiente"
    
    st.markdown(f"<p style='font-size:0.85rem; margin:2px 0;'>• {dom_status}</p>", unsafe_allow_html=True)
    st.markdown(f"<p style='font-size:0.85rem; margin:2px 0;'>• {neg_status}</p>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Inyección exacta del motor de empaquetado Word (.docx) original del script
    def construir_documento_word():
        doc = Document()
        doc.add_heading("VISITA A CLIENTES DE PEQUEÑA EMPRESA", level=0)
        
        # Bloque I
        p = doc.add_paragraph()
        r = p.add_run("I. Datos del cliente")
        r.bold = True
        
        table = doc.add_table(rows=2, cols=2)
        table.style = "Light Grid Accent 1"
        row1 = table.rows[0].cells
        row1[0].text = "Cliente:"
        row1[1].text = safe_str(cliente.get("CLIENTE"))
        row2 = table.rows[1].cells
        row2[0].text = "DNI / LE:"
        row2[1].text = safe_str(cliente.get("PENDOC"))
        
        # Inserción de riesgos marcados
        doc.add_heading("II. Criterios de Riesgo Encontrados", level=2)
        marcados = [k for k, v in st.session_state.validaciones_marcadas.items() if v]
        if marcados:
            for m in marcados: doc.add_paragraph(f"• {m}", style='List Bullet')
        else:
            doc.add_paragraph("No se registraron riesgos críticos.")
            
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    st.markdown('<div style="text-align:center; margin-top:20px;">', unsafe_allow_html=True)
    try:
        doc_bytes = construir_documento_word()
        nombre_salida = f"Reporte_Visita_{safe_str(cliente.get('PENDOC', 'cliente'))}.docx"
        
        st.download_button(
            label="📥 Descargar Reporte (.docx)",
            data=doc_bytes,
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Error al estructurar informe final: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🔄 Reiniciar e Evaluar Otro Cliente", key="btn_back_clear_all"):
        st.session_state.current_view = "carga"
        st.session_state.cliente_actual = {}
        st.session_state.visitas = {}
        st.session_state.validaciones_marcadas = {}
        st.rerun()
