import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

# CONFIGURACIÓN INICIAL
st.set_page_config(page_title="PCD - Operaciones Internacional", layout="wide")

# ==========================================
# CONFIGURACIÓN DE CREDENCIALES (NUBE)
# ==========================================
# Se obtienen las llaves desde el panel de Secrets de Streamlit
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

def guardar_en_google_sheets_directo(nombre_hoja, df):
    try:
        alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # --- CORRECCIÓN PARA LA NUBE ---
        credenciales = ServiceAccountCredentials.from_json_keyfile_dict(CREDENCIALES_GOOGLE, alcance)
        cliente = gspread.authorize(credenciales)
        
        # Abrir la base de datos (asegúrate de que este nombre coincida con tu archivo en Drive)
        doc = cliente.open("PCD_BaseDatos")
        try:
            hoja = doc.worksheet(nombre_hoja)
        except gspread.exceptions.WorksheetNotFound:
            return False, f"La hoja '{nombre_hoja}' no existe."
            
        df_guardar = df.copy().astype(str)
        
        # Marca de tiempo automática
        if "Fecha Sistema" not in df_guardar.columns:
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        
        hoja.append_rows(df_guardar.values.tolist())
        return True, "Datos guardados exitosamente."
    except Exception as e:
        return False, str(e)

# ==========================================
# DATA MAESTRA Y USUARIOS
# ==========================================
USUARIOS_VALIDOS = {
    "venezuela": "pcd.ve.2024",
    "dominicana": "pcd.do.2024",
    "master": "pcd.master.2025"
}

# ==========================================
# INTERFAZ DE LOGIN
# ==========================================
def login():
    st.markdown("""
        <style>
        .main { background-color: #f0f2f6; }
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #003366; color: white; font-weight: bold; }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        # Intenta cargar el logo si existe
        if os.path.exists("logo.png"):
            st.image("logo.png", width=250)
        
        st.title("🔐 Control de Operaciones")
        st.subheader("PCD Internacional")
        
        usuario_input = st.text_input("Usuario").lower().strip()
        clave_input = st.text_input("Contraseña", type="password")
        
        if st.button("Iniciar Sesión"):
            if usuario_input in USUARIOS_VALIDOS and USUARIOS_VALIDOS[usuario_input] == clave_input:
                st.session_state.autenticado = True
                st.session_state.usuario = usuario_input.capitalize()
                
                # Definir país según usuario
                if usuario_input == "venezuela": st.session_state.pais = "Venezuela"
                elif usuario_input == "dominicana": st.session_state.pais = "Dominicana"
                else: st.session_state.pais = "Global"
                
                st.success(f"Bienvenido al sistema, {st.session_state.usuario}")
                st.rerun()
            else:
                st.error("Credenciales inválidas. Por favor intente de nuevo.")

# ==========================================
# DASHBOARD PRINCIPAL
# ==========================================
def main():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        login()
    else:
        # BARRA LATERAL DE NAVEGACIÓN
        st.sidebar.image("logo.png", width=150) if os.path.exists("logo.png") else st.sidebar.title("PCD")
        st.sidebar.markdown(f"### 📍 Sede: {st.session_state.pais}")
        st.sidebar.markdown(f"**Usuario:** {st.session_state.usuario}")
        
        st.sidebar.divider()
        
        menu = st.sidebar.selectbox(
            "Seleccione un Módulo",
            ["🏠 Inicio", "🚚 Gestión de Flota", "🛰️ Monitoreo GPS", "💰 Gastos y Combustible", "📋 Cierres Diarios"]
        )
        
        if st.sidebar.button("Cerrar Sesión"):
            st.session_state.autenticado = False
            st.rerun()

        # CONTENIDO SEGÚN SELECCIÓN
        if menu == "🏠 Inicio":
            st.title(f"Bienvenido al Panel {st.session_state.pais}")
            st.info("Seleccione una opción en el menú lateral para gestionar las operaciones.")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Estatus Sistema", "En Línea 🟢")
            col2.metric("Sede Activa", st.session_state.pais)
            col3.metric("Base de Datos", "Sincronizada ☁️")
            
        elif menu == "🚚 Gestión de Flota":
            st.title("Gestión de Flota")
            st.write("Módulo de mantenimiento y operatividad vehicular.")
            # Aquí iría tu código de flota o el import correspondiente
            
        elif menu == "🛰️ Monitoreo GPS":
            st.title("Monitoreo Satelital")
            st.write("Seguimiento de unidades en tiempo real.")

        elif menu == "💰 Gastos y Combustible":
            st.title("Control de Gastos")
            st.write("Registro de combustible e insumos.")

        elif menu == "📋 Cierres Diarios":
            st.title("Cierres de Operación")
            st.write("Reportes de cierre de jornada.")

if __name__ == "__main__":
    main()
