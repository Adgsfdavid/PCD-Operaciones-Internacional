import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# CONFIGURACIÓN INICIAL
st.set_page_config(page_title="PCD - Sistema de Gestión Logística", layout="wide")

# ==========================================
# CONFIGURACIÓN DE CREDENCIALES (NUBE)
# ==========================================
# Esta variable lee directamente desde el panel de Secrets de Streamlit
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

def guardar_en_google_sheets_directo(nombre_hoja, df):
    try:
        alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # --- CAMBIO CLAVE: Usamos el diccionario de secretos en lugar del archivo .json ---
        credenciales = ServiceAccountCredentials.from_json_keyfile_dict(CREDENCIALES_GOOGLE, alcance)
        cliente = gspread.authorize(credenciales)
        
        # Abrir la base de datos principal
        doc = cliente.open("PCD_BaseDatos")
        try:
            hoja = doc.worksheet(nombre_hoja)
        except gspread.exceptions.WorksheetNotFound:
            return False, f"La hoja '{nombre_hoja}' no existe en la base de datos."
            
        df_guardar = df.copy().astype(str)
        
        # Agregar marca de tiempo si no existe
        if "Fecha Sistema" not in df_guardar.columns:
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        
        hoja.append_rows(df_guardar.values.tolist())
        return True, "Datos sincronizados correctamente con la nube."
    except Exception as e:
        return False, str(e)

# ==========================================
# LÓGICA DE LOGIN Y NAVEGACIÓN
# ==========================================
def main():
    # Estilo CSS para mejorar la apariencia
    st.markdown("""
        <style>
        .main { background-color: #f5f7f9; }
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #1a237e; color: white; }
        .login-container { max-width: 400px; margin: auto; padding: 2rem; background: white; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        </style>
    """, unsafe_content_type=True)

    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image("logo.png", width=200)
            st.title("🔐 Acceso al Sistema PCD")
            
            usuario = st.text_input("Usuario")
            clave = st.text_input("Contraseña", type="password")
            
            if st.button("Ingresar"):
                # Aquí puedes integrar con seguridad.py para validar usuarios
                if usuario == "admin" and clave == "pcd2024": # Ejemplo simple
                    st.session_state.autenticado = True
                    st.session_state.usuario = usuario
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos")
    else:
        mostrar_dashboard()

def mostrar_dashboard():
    st.sidebar.title(f"Bienvenido, {st.session_state.usuario}")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    st.title("🚀 Panel de Control Principal")
    st.write("Selecciona un módulo en la barra lateral para comenzar a trabajar.")
    
    # Aquí puedes agregar indicadores rápidos (KPIs)
    col1, col2, col3 = st.columns(3)
    col1.metric("Estatus Flota", "Operativa")
    col2.metric("Último Cierre", "Exitoso")
    col3.metric("Conexión Base de Datos", "🟢 Online")

if __name__ == "__main__":
    main()
