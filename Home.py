import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# CONFIGURACIÓN INICIAL DE LA PÁGINA
st.set_page_config(page_title="PCD - Sistema de Gestión Logística", layout="wide")

# ==========================================
# CONFIGURACIÓN DE CREDENCIALES (NUBE)
# ==========================================
# Se obtienen las llaves desde el panel de Secrets de Streamlit
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

def guardar_en_google_sheets_directo(nombre_hoja, df):
    try:
        alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # Conexión usando el diccionario de secretos
        credenciales = ServiceAccountCredentials.from_json_keyfile_dict(CREDENCIALES_GOOGLE, alcance)
        cliente = gspread.authorize(credenciales)
        
        # Nombre de tu base de datos principal en Google Sheets
        doc = cliente.open("PCD_BaseDatos")
        try:
            hoja = doc.worksheet(nombre_hoja)
        except gspread.exceptions.WorksheetNotFound:
            return False, f"La hoja '{nombre_hoja}' no existe."
            
        df_guardar = df.copy().astype(str)
        
        # Insertar marca de tiempo automática
        if "Fecha Sistema" not in df_guardar.columns:
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        
        hoja.append_rows(df_guardar.values.tolist())
        return True, "Datos sincronizados correctamente."
    except Exception as e:
        return False, str(e)

# ==========================================
# LÓGICA DE LOGIN Y NAVEGACIÓN
# ==========================================
def main():
    # Estilo visual corregido (unsafe_allow_html=True)
    st.markdown("""
        <style>
        .main { background-color: #f5f7f9; }
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #1a237e; color: white; font-weight: bold; }
        .login-container { max-width: 400px; margin: auto; padding: 2rem; background: white; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        </style>
    """, unsafe_allow_html=True)

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
                # Validación simple (ajustar según tus necesidades en seguridad.py)
                if usuario == "admin" and clave == "pcd2024":
                    st.session_state.autenticado = True
                    st.session_state.usuario = usuario
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas. Verifica el usuario y la clave.")
    else:
        mostrar_dashboard()

def mostrar_dashboard():
    # Barra lateral de navegación
    st.sidebar.title(f"👤 {st.session_state.usuario}")
    st.sidebar.write("Sistema de Gestión de Operaciones")
    
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    st.title("🚀 Panel de Control Principal")
    st.info("Utiliza los módulos en la barra lateral para gestionar la flota, el monitoreo o los cierres diarios.")
    
    # Vista rápida de indicadores (KPIs)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Estado de Flota", "Operativa")
    with col2:
        st.metric("Conexión Nube", "Estable 🟢")
    with col3:
        st.metric("Base de Datos", "Sincronizada")

    st.divider()
    st.write("### Actividades recientes")
    st.write("No hay reportes nuevos en la última hora.")

if __name__ == "__main__":
    main()
