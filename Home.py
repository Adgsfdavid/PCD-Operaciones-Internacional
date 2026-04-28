import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import traceback
from datetime import datetime
import os

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered")

# ==========================================
# CREDENCIALES Y CONEXIÓN DINÁMICA (MODERNIZADA)
# ==========================================
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

# RECONSTRUCTOR BLINDADO DE LLAVE
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"
CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def guardar_en_google_sheets_directo(nombre_hoja, df):
    try:
        cliente = obtener_cliente_sheets()
        u_data = st.session_state.get("user_data", {})
        # Mantenemos la lógica multitenant por nombre de base de datos
        nombre_bd = u_data.get("sheet_name", "PCD_BaseDatos_VZLA")
        doc = cliente.open(nombre_bd)
        hoja = doc.worksheet(nombre_hoja)
        
        df_guardar = df.copy().astype(str)
        if "Fecha Sistema" not in df_guardar.columns:
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            
        valores = df_guardar.values.tolist()
        
        try:
            hoja.append_rows(valores, value_input_option='USER_ENTERED')
        except Exception as error_interno:
            if "200" in str(error_interno):
                pass
            else:
                raise error_interno
                
        return True, "Datos sincronizados correctamente."
    except Exception as e:
        return False, f"Error: {e}"

# --- CONFIGURACIÓN DE PAÍSES Y ACCESOS ---
CONFIG_PAISES = {
    "admin_vzla": {"pais": "VENEZUELA", "clave": "pcd2026", "sheet_name": "PCD_BaseDatos_VZLA"},
    "admin_rd": {"pais": "REPUBLICA DOMINICANA", "clave": "pcd_rd_2026", "sheet_name": "PCD_BaseDatos_RD"},
    "master_vzla": {"pais": "MASTER_VZLA", "clave": "pcd_master_2026", "sheet_name": "PCD_BaseDatos_VZLA"}
}

# --- LÓGICA DE LOGIN (CORREGIDA CON FORMULARIO) ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

def login():
    st.image("https://pcdinternacional.com/wp-content/uploads/2022/05/logo-pcd.png", width=200)
    st.title("🔐 Acceso PCD Internacional")
    
    # USAMOS UN FORMULARIO PARA CAPTURAR EL AUTOCOMPLETADO AL INSTANTE
    with st.form("login_form"):
        usuario = st.text_input("Usuario", placeholder="Ej: admin_vzla")
        clave = st.text_input("Contraseña", type="password")
        submit_button = st.form_submit_button("🚀 Ingresar", use_container_width=True)
        
        if submit_button:
            if usuario in CONFIG_PAISES and CONFIG_PAISES[usuario]["clave"] == clave:
                st.session_state["logged_in"] = True
                st.session_state["user_data"] = CONFIG_PAISES[usuario]
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")

# --- LÓGICA DE NAVEGACIÓN ---
if not st.session_state["logged_in"]:
    login()
else:
    u_data = st.session_state["user_data"]
    st.sidebar.title(f"📍 {u_data['pais']}")
    if st.sidebar.button("🚪 Cerrar Sesión"):
        st.session_state["logged_in"] = False
        st.rerun()

    # IMPORTANTE: Rutas en minúsculas como detecta tu servidor
    if u_data['pais'] == "VENEZUELA" or u_data['pais'] == "MASTER_VZLA":
        paginas = [
            st.Page("vzla/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("vzla/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("vzla/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("vzla/seguridad.py", title="Prevención y Control", icon="🛡️"),
            st.Page("app.py", title="Trafico y Salidas", icon="📊")
        ]
    elif u_data['pais'] == "REPUBLICA DOMINICANA":
        paginas = [
            st.Page("rd/logistica_rd.py", title="Operaciones RD", icon="🇩🇴")
        ]
    
    pg = st.navigation(paginas)
    pg.run()
