import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered")

# --- CREDENCIALES CLOUD (CONEXIÓN SEGURA) ---
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

def guardar_en_google_sheets_directo(nombre_hoja, df):
    try:
        alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # Uso de secretos en lugar de archivo físico para evitar el FileNotFoundError
        credenciales = ServiceAccountCredentials.from_json_keyfile_dict(CREDENCIALES_GOOGLE, alcance)
        cliente = gspread.authorize(credenciales)

        u_data = st.session_state.get("user_data", {})
        nombre_bd = u_data.get("sheet_name", "PCD_BaseDatos_VZLA")
        doc = cliente.open(nombre_bd)

        try:
            hoja = doc.worksheet(nombre_hoja)
        except gspread.exceptions.WorksheetNotFound:
            return False, f"La hoja '{nombre_hoja}' no existe."

        df_guardar = df.copy().astype(str)
        if "Fecha Sistema" not in df_guardar.columns:
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

        hoja.append_rows(df_guardar.values.tolist())
        return True, "Datos sincronizados correctamente."
    except Exception as e:
        return False, str(e)

# --- BASE DE DATOS DE USUARIOS ---
CONFIG_PAISES = {
    "admin_vzla": {
        "clave": "Vzla2026*", 
        "pais": "VENEZUELA", 
        "sheet_name": "PCD_BaseDatos_VZLA"
    },
    "admin_rd": {
        "clave": "Dom2026*", 
        "pais": "DOMINICANA", 
        "sheet_name": "PCD_BaseDatos_DOM"
    },
    "david_master": {
        "clave": "Master123", 
        "pais": "MASTER_VZLA", 
        "sheet_name": "PCD_BaseDatos_VZLA"
    }
}

# --- ESTADO DE SESIÓN ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# --- PANTALLA DE LOGIN ---
def login():
    try:
        st.image("logo.png", width=200)
    except:
        pass
        
    st.title("Acceso Corporativo PCD")
    
    usuario = st.text_input("Usuario")
    clave = st.text_input("Contraseña", type="password")
    
    if st.button("Ingresar", type="primary", use_container_width=True):
        if usuario in CONFIG_PAISES and CONFIG_PAISES[usuario]["clave"] == clave:
            st.session_state["logged_in"] = True
            st.session_state["user_data"] = CONFIG_PAISES[usuario]
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos")

# --- LÓGICA DE ENRUTAMIENTO (RUTAS) ---
if not st.session_state["logged_in"]:
    login()
else:
    u_data = st.session_state["user_data"]
    
    st.sidebar.title(f"📍 {u_data['pais']}")
    st.sidebar.caption(f"Conectado a: {u_data['sheet_name']}")
    
    if st.sidebar.button("🚪 Cerrar Sesión"):
        st.session_state["logged_in"] = False
        st.rerun()

    # Mapeo de páginas según el país del usuario
    if u_data['pais'] == "VENEZUELA" or u_data['pais'] == "MASTER_VZLA":
        paginas = [
            st.Page("vnzl/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("vnzl/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("vnzl/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("vnzl/seguridad.py", title="Prevención y Control", icon="🛡️")
        ]
        
    elif u_data['pais'] == "DOMINICANA":
        paginas = [
            st.Page("rd/cierre_diario.py", title="Cierre Diario (RD)", icon="📋"),
            st.Page("rd/flota.py", title="Flota y Combustible (RD)", icon="🚛")
        ]
    
    # Ejecutar la navegación automática con las páginas correspondientes
    nav = st.navigation(paginas)
    nav.run()
