# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime

# 1. Configuración de página optimizada para móvil
st.set_page_config(
    page_title="App Visitas",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. CSS para el look & feel de Mockup (Centrado, bordes redondos, sombras suaves)
def inject_custom_css():
    st.markdown("""
    <style>
        /* Fondo gris suave como en los mockups */
        .stApp { background-color: #F8FAFC; }
        
        /* Contenedor principal tipo "Pantalla de Celular" */
        .block-container { 
            max-width: 450px !important; 
            padding-top: 2rem !important;
        }
        
        /* Tarjetas blancas con sombra */
        .card { 
            background: white; 
            padding: 25px; 
            border-radius: 25px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            margin-bottom: 20px;
        }
        
        /* Botón de acción principal (Rojo corporativo) */
        div.stButton > button { 
            width: 100%; 
            border-radius: 15px; 
            height: 50px; 
            background-color: #C8102E; 
            color: white; 
            font-weight: bold; 
            border: none;
        }
        
        /* Inputs estilizados */
        .stTextInput input, .stNumberInput input {
            border-radius: 12px !important;
            border: 1px solid #E2E8F0 !important;
        }
    </style>
    """, unsafe_allow_html=True)

inject_custom_css()
# 3. Estado de navegación para que la App sea "de una sola página" (Single Page App)
if "view" not in st.session_state:
    st.session_state.view = "busqueda"

# --- 2. LÓGICA DE PANTALLA DE BÚSQUEDA ---

def pantalla_busqueda():
    st.markdown('<div class="card" style="text-align: center;">', unsafe_allow_html=True)
    st.title("🏦 Visitas")
    st.write("Carga tu base de datos para iniciar")
    
    # Uploader estilizado
    archivo = st.file_uploader("Seleccionar Excel", type=["xlsx"], label_visibility="collapsed")
    
    if archivo:
        # Procesamiento rápido
        df = pd.read_excel(archivo, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        st.session_state.clientes_df = df
        st.success("Base de datos cargada")
        
        # Buscador prominente
        busqueda = st.text_input("Buscar DNI o Nombre", placeholder="Ingrese datos...")
        
        if busqueda:
            # Filtrado rápido
            mask = (df["PENDOC"].str.contains(busqueda, na=False) | 
                    df["CLIENTE"].str.contains(busqueda, case=False, na=False))
            resultados = df[mask]
            
            if not resultados.empty:
                st.write(f"Resultados encontrados: {len(resultados)}")
                if st.button("Cargar Cliente"):
                    st.session_state.cliente_actual = resultados.iloc[0].to_dict()
                    st.session_state.view = "ficha" # Cambiamos de estado
                    st.rerun()
            else:
                st.warning("Cliente no encontrado")
                
    st.markdown('</div>', unsafe_allow_html=True)

# 3. Router Principal
if st.session_state.view == "busqueda":
    pantalla_busqueda()
elif st.session_state.view == "ficha":
    # Aquí irá la Parte 3 (Tabs y detalles)
    st.write("Pantalla de Ficha - En desarrollo")


# --- 3. LÓGICA DE PANTALLA DE FICHA (CONTENIDO DE TABS) ---

def pantalla_ficha():
    cliente = st.session_state.cliente_actual
    
    # Header del cliente (Tipo Mockup)
    st.markdown(f'''
        <div class="card">
            <h2 style="margin:0;">{cliente.get("CLIENTE", "Nombre Cliente")}</h2>
            <p style="color:#64748B;">DNI: {cliente.get("PENDOC", "N/A")}</p>
        </div>
    ''', unsafe_allow_html=True)

    # Definición de los Tabs
    tabs = st.tabs(["📋 Info", "📍 Domicilio", "💼 Negocio", "📊 Reporte"])

    with tabs[0]:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write(f"**Saldo Actual:** S/. {cliente.get('SALDO_MN', '0.00')}")
        st.write(f"**Agencia:** {cliente.get('AGENCIA', 'N/A')}")
        # Aquí puedes agregar más datos del cliente
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[1]:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        # Aquí reutilizamos tu lógica de bloque_verificacion que ya tienes
        # bloque_verificacion("domicilio", "Domicilio") 
        st.write("Configuración de domicilio...")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[2]:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        # bloque_verificacion("negocio", "Negocio")
        st.write("Configuración de negocio...")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[3]:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        if st.button("Finalizar y Generar Word"):
            # Aquí llamarías a tu función generar_reporte()
            st.success("Reporte generado con éxito")
        st.markdown('</div>', unsafe_allow_html=True)

    # Botón de regreso fijo al final
    if st.button("← Volver a Búsqueda"):
        st.session_state.view = "busqueda"
        st.rerun()

# Router final actualizado
if st.session_state.view == "busqueda":
    pantalla_busqueda()
elif st.session_state.view == "ficha":
    pantalla_ficha()










