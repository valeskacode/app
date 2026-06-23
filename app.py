# -*- coding: utf-8 -*-
"""
Formulario de verificación de datos visita - Versión Control de Clic / Doble Clic Móvil
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
# CONFIGURACIÓN GENERAL
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Formulario - Visita de clientes",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
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

NARANJA = "C8102E"   
AZUL = "1B3A5C"
VERDE = "137333"
ROJO = "a50e0e"

CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: #f7f7f9; }}
section[data-testid="stSidebar"] {{ background-color: #1B3A5C; }}
section[data-testid="stSidebar"] * {{ color: #ffffff !important; }}
h1, h2, h3 {{ color: #{AZUL}; }}

/* Botones institucionales generales */
div.stButton > button {{
    background-color: #{NARANJA}; color: white; border: none;
    border-radius: 6px; font-weight: 600;
}}
div.stButton > button:hover {{ background-color: #a30d24; color: white; }}

/* DISEÑO ESPECIAL PARA FILAS DEL PANEL DE VALIDACIÓN (OPTIMIZADO MÓVIL) */
.validation-box div.stButton > button {{
    background-color: #ffffff !important;
    color: #1B3A5C !important;
    border: 1px solid #e0e0e0 !important;
    text-align: left !important;
    display: flex !important;
    justify-content: flex-start !important;
    align-items: center !important;
    padding: 8px 12px !important;
    margin-bottom: 4px !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
}}
.validation-box div.stButton > button:hover {{
    background-color: #f8f9fa !important;
    border-color: #{AZUL} !important;
}}

.card {{
    background: white; padding: 1.2rem 1.4rem; border-radius: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 1rem;
}}
.badge-ok {{ background:#e6f4ea; color:#137333; padding:6px 12px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
.badge-pend {{ background:#fce8e6; color:#a50e0e; padding:6px 12px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
.validation-box {{
    border: 2px solid #{AZUL}; border-radius: 10px; padding: 1.2rem;
    background: white; margin: 0.5rem 0;
}}
.validation-title {{
    font-size: 1.1rem; font-weight: 700; color: #{AZUL}; margin-bottom: 0.6rem;
    border-bottom: 3px solid #{NARANJA}; padding-bottom: 0.3rem;
}}
@media (max-width: 768px) {{
    .card {{ padding: 0.8rem 1rem; }}
    h1 {{ font-size: 1.4rem; }}
}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------
# HELPERS DE LIMPIEZA
# --------------------------------------------------------------------------
def safe_str(v, default=""):
    if v is None: return default
    try:
        if pd.isna(v): return default
    except Exception: pass
    s = str(v).strip()
    return default if s.lower() in ("nan", "none") else s

def safe_float(v, default=0.0):
    try:
        f = float(v)
        if pd.isna(f): return default
        return f
    except Exception: return default

def fmt_money(v):
    return f"S/. {safe_float(v):,.2f}"

def limpiar_texto_dni(val):
    if pd.isna(val) or val is None: return ""
    txt = str(val).strip()
    if txt.endswith(".0"): txt = txt[:-2]
    if txt.isdigit() and len(txt) < 8 and len(txt) > 0: txt = txt.zfill(8)
    return txt

def init_state():
    defaults = {
        "clientes_df": None, "cliente_actual": {}, "visitas": {}, "garantias": [],
        "rcc": [], "validaciones_marcadas": {}, "click_timestamps": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v

init_state()
cliente = st.session_state.cliente_actual

# --------------------------------------------------------------------------
# FUNCIONES DE LÓGICA
# --------------------------------------------------------------------------
def buscar_cliente_por_dni(dni_input, df):
    if not dni_input or df is None or len(df) == 0: return None
    txt_buscar = limpiar_texto_dni(dni_input)
    mask = (df.get("PENDOC", pd.Series("", index=df.index)).astype(str).apply(limpiar_texto_dni) == txt_buscar)
    resultados = df[mask]
    return resultados.iloc[0].to_dict() if len(resultados) > 0 else None

def validar_visita():
    validaciones = {k: False for k in [
        "documentos_enmiendas", "documentos_inconsistentes", "documentos_sin_datos", 
        "documentos_sin_firmas", "documentos_duplicados", "sin_sustento_actividad", 
        "sin_sustento_ingresos", "sin_sustento_activos", "conyuge_omitido", 
        "credito_reprogramado", "credito_refinanciado", "calificacion_diferente"
    ]}
    if not safe_str(cliente.get("CLIENTE")): validaciones["documentos_sin_datos"] = True
    if safe_str(cliente.get("DIAS_ATRASO")) and int(safe_float(cliente.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
    for clave in ["domicilio", "negocio", "aval"]:
        if clave not in st.session_state.visitas or not st.session_state.visitas[clave].get("foto_bytes"):
            validaciones["documentos_sin_firmas"] = True
            break
    return validaciones

def mostrar_panel_validacion():
    st.markdown('<div class="validation-box">', unsafe_allow_html=True)
    st.markdown('<div class="validation-title">🔍 Panel de Validación - Criterios de Riesgo</div>', unsafe_allow_html=True)
    validaciones_auto = validar_visita()
    criterios = {
        "documentos_enmiendas": ("Documentos con enmiendas", "⚠️"),
        "documentos_inconsistentes": ("Datos inconsistentes en documentos", "⚠️"),
        "documentos_sin_datos": ("Documentos sin datos del cliente", "❌"),
        "documentos_sin_firmas": ("Documentos sin firmas o fotos", "❌"),
        "documentos_duplicados": ("Documentos duplicados", "⚠️"),
        "sin_sustento_actividad": ("Sin sustento de actividad económica", "❌"),
        "sin_sustento_ingresos": ("Sin sustento de ingresos", "❌"),
        "sin_sustento_activos": ("Sin sustento de activos representativos", "⚠️"),
        "conyuge_omitido": ("Cónyuge omitido en evaluación", "⚠️"),
        "credito_reprogramado": ("Crédito reprogramado", "ℹ️"),
        "credito_refinanciado": ("Crédito refinanciado", "ℹ️"),
        "calificacion_diferente": ("Calificación diferente a la fecha de revisión", "⚠️"),
    }
    items_por_categoria = {"❌": [], "⚠️": [], "ℹ️": []}
    for key, (label, icon) in criterios.items(): items_por_categoria[icon].append((key, label))
    
    with st.container(height=280, border=True):
        for icon in ["❌", "⚠️", "ℹ️"]:
            for key, label in items_por_categoria[icon]:
                is_checked = st.session_state.validaciones_marcadas.get(key, validaciones_auto.get(key, False))
                marcador_visual = "[ X ]" if is_checked else "[   ]"
                if st.button(f"{icon} {marcador_visual} {label}", key=f"btn_criterio_{key}", use_container_width=True):
                    ahora = datetime.now().timestamp()
                    ultimo_clic = st.session_state.click_timestamps.get(key, 0)
                    st.session_state.click_timestamps[key] = ahora
                    if not is_checked:
                        st.session_state.validaciones_marcadas[key] = True
                    elif (ahora - ultimo_clic) < 0.8:
                        st.session_state.validaciones_marcadas[key] = False
                    st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------------------------------
# SIDEBAR
# --------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📂 Base de clientes")
    excel_file = st.file_uploader("Cargar archivo .xlsx", type=["xlsx"])
    if excel_file:
        df = pd.read_excel(excel_file, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        if "PENDOC" in df.columns: df["PENDOC"] = df["PENDOC"].apply(limpiar_texto_dni)
        st.session_state.clientes_df = df
        st.success("Cargado correctamente")
    
    if st.session_state.clientes_df is not None:
        if st.button("🗑️ Limpiar todo"):
            for k in ["cliente_actual", "visitas", "garantias", "rcc", "validaciones_marcadas"]:
                st.session_state[k] = {} if k != "garantias" else []
            st.rerun()

# --------------------------------------------------------------------------
# INTERFAZ PRINCIPAL
# --------------------------------------------------------------------------
st.title("Visita a Clientes")
tabs = st.tabs(["1️⃣ Cliente", "2️⃣ Historial", "3️⃣ Domicilio", "4️⃣ Negocio", "5️⃣ Ingresos", "6️⃣ Garantías", "7️⃣ Reporte"])

# TABS LÓGICA (Simplificada para espacio)
with tabs[0]:
    st.subheader("Titular")
    dni_input = st.text_input("DNI Titular", value=safe_str(cliente.get("PENDOC")))
    if dni_input and not cliente:
        c = buscar_cliente_por_dni(dni_input, st.session_state.clientes_df)
        if c: st.session_state.cliente_actual = c; st.rerun()
    mostrar_panel_validacion()

# ... (Insertar aquí bloques para tabs 1-6 de tu código previo) ...

# --------------------------------------------------------------------------
# REPORTE WORD
# --------------------------------------------------------------------------
def generar_reporte():
    doc = Document()
    doc.add_heading("VISITA A CLIENTES", level=0)
    # (Aquí iría la lógica completa de generar_reporte que tenías)
    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf

with tabs[6]:
    st.subheader("Generar reporte")
    if st.button("Descargar Word", type="primary"):
        st.download_button("Descargar", data=generar_reporte(), file_name="Reporte.docx")





# TAB 2: Historial
with tabs[1]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("3. Historial crediticio")
    df = st.session_state.clientes_df
    if df is not None and (cliente.get("PENDOC")):
        hist = df[df["PENDOC"] == cliente.get("PENDOC")]
        st.dataframe(hist, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 3: Domicilio
with tabs[2]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("5. Verificación Domicilio")
    bloque_verificacion("domicilio", "Domicilio")
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 4: Negocio
with tabs[3]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("6. Verificación Negocio")
    bloque_verificacion("negocio", "Negocio")
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 5: Ingresos
with tabs[4]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("8-9. Ingresos y Gastos")
    ventas = st.number_input("Ventas mensuales (S/.)", value=0.0)
    gastos = st.number_input("Gastos operativos (S/.)", value=0.0)
    st.metric("Utilidad Bruta", fmt_money(ventas - gastos))
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 6: Garantías
with tabs[5]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("7. Garantías y Aval")
    bloque_verificacion("aval", "Aval")
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 7: Reporte
with tabs[6]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Generar reporte Word")
    if st.button("Descargar Documento", type="primary"):
        buf = generar_reporte() # Asegúrate de que esta función esté definida abajo
        st.download_button("Descargar .docx", data=buf, file_name="Reporte_Visita.docx")
    st.markdown('</div>', unsafe_allow_html=True)
