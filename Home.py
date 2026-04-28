import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered")

# --- CREDENCIALES CLOUD (SEGURIDAD) ---
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

def guardar_en_google_sheets_directo(nombre_hoja, df):
    try:
        alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credenciales = ServiceAccountCredentials.from_json_keyfile_dict(CREDENCIALES_GOOGLE, alcance)
        cliente = gspread.authorize(credenciales)
        u_data = st.session_state.get("user_data", {})
        nombre_bd = u_data.get("sheet_name", "PCD_BaseDatos_VZLA")
        doc = cliente.open(nombre_bd)
        hoja = doc.worksheet(nombre_hoja)
        df_guardar = df.copy().astype(str)
        if "Fecha Sistema" not in df_guardar.columns:
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        hoja.append_rows(df_guardar.values.tolist())
        return True, "Datos sincronizados correctamente."
    except Exception as e:
        return False, str(e)

# --- BASE DE DATOS DE USUARIOS ---
CONFIG_PAISES = {
    "admin_vzla": {"clave": "Vzla2026*", "pais": "VENEZUELA", "sheet_name": "PCD_BaseDatos_VZLA"},
    "admin_rd": {"clave": "Dom2026*", "pais": "DOMINICANA", "sheet_name": "PCD_BaseDatos_DOM"},
    "david_master": {"clave": "Master123", "pais": "MASTER_VZLA", "sheet_name": "PCD_BaseDatos_VZLA"}
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

def login():
    try: st.image("logo.png", width=200)
    except: pass
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

# --- LÓGICA DE NAVEGACIÓN (RUTAS CORREGIDAS A MINÚSCULAS) ---
if not st.session_state["logged_in"]:
    login()
else:
    u_data = st.session_state["user_data"]
    st.sidebar.title(f"📍 {u_data['pais']}")
    if st.sidebar.button("🚪 Cerrar Sesión"):
        st.session_state["logged_in"] = False
        st.rerun()

    # IMPORTANTE: Aquí usamos "vzla/" y "rd/" en minúsculas como detectó el servidor
    if u_data['pais'] == "VENEZUELA" or u_data['pais'] == "MASTER_VZLA":
        paginas = [
            st.Page("vzla/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("vzla/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("vzla/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("vzla/seguridad.py", title="Prevención y Control", icon="🛡️"),
            st.Page("vzla/app.py", title="Tráfico y Transbordos", icon="🚦") 
        ]
        
    elif u_data['pais'] == "DOMINICANA":
        paginas = [
            st.Page("rd/cierre_diario.py", title="Cierre Diario (RD)", icon="📋"),
            st.Page("rd/flota.py", title="Flota y Combustible (RD)", icon="🚛")
        ]
    
    nav = st.navigation(paginas)
    nav.run()
