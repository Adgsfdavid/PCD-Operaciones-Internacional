import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import traceback
from datetime import datetime
import os
import extra_streamlit_components as stx 

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
        doc = cliente.open_by_key(u_data.get("sheet_id"))
        
        try:
            hoja = doc.worksheet(nombre_hoja)
        except Exception:
            hoja = doc.add_worksheet(title=nombre_hoja, rows=1000, cols=20)
            hoja.append_row(list(df.columns))
            
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

# ==========================================
# CONFIGURACIÓN REAL DE ACCESOS CON IDs DE SHEETS
# ==========================================
CONFIG_PAISES = {
    "admin_vzla": {
        "clave": "Vzla2026*", 
        "pais": "VENEZUELA", 
        "sheet_id": "1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0" 
    },
    "admin_rd": {
        "clave": "Dom2026*", 
        "pais": "DOMINICANA", 
        "sheet_id": "1ourNW6VifjXiJFsyVKamjeBL7iACKEpH0ozdWo8rCMc" 
    },
    "david_master": {
        "clave": "Master123", 
        "pais": "MASTER_VZLA", 
        "sheet_id": "1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0" 
    }
}

# ==========================================
# MANEJO DE SESIÓN Y COOKIES (EVITA EL DESLOGUEO)
# ==========================================
cookie_manager = stx.CookieManager()

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# MAGIA DE LAS COOKIES: Verificamos si el usuario ya se había logueado antes
usuario_guardado = cookie_manager.get(cookie="pcd_usuario_valido")

# Si hay una cookie guardada en el navegador y el usuario no está logueado en la sesión actual, lo dejamos pasar
if usuario_guardado in CONFIG_PAISES and not st.session_state["logged_in"]:
    st.session_state["logged_in"] = True
    st.session_state["user_data"] = CONFIG_PAISES[usuario_guardado]

def login():
    try:
        st.image("logo.png", width=200) 
    except Exception:
        pass
        
    st.title("🔐 Acceso PCD Internacional")
    
    with st.form("login_form"):
        usuario = st.text_input("Usuario", placeholder="Ingresa tu usuario")
        clave = st.text_input("Contraseña", type="password")
        submit_button = st.form_submit_button("🚀 Ingresar", use_container_width=True)
        
        if submit_button:
            if usuario in CONFIG_PAISES and CONFIG_PAISES[usuario]["clave"] == clave:
                # Guardamos la cookie por 30 días
                cookie_manager.set("pcd_usuario_valido", usuario, max_age=2592000)
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
        # SOLUCIÓN AL ERROR DE KEYERROR
        try:
            cookie_manager.delete("pcd_usuario_valido")
        except Exception:
            pass # Si la cookie no existe o no la encuentra, lo ignoramos sin dar error rojo
        
        st.session_state["logged_in"] = False
        st.rerun()

    # Rutas para Venezuela y Master
    if u_data['pais'] == "VENEZUELA" or u_data['pais'] == "MASTER_VZLA":
        paginas = [
            st.Page("vzla/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("vzla/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("vzla/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("vzla/seguridad.py", title="Prevención y Control", icon="🛡️"),
            st.Page("vzla/app.py", title="Trafico y Salidas", icon="📊")
        ]
    # Rutas para República Dominicana
    elif u_data['pais'] == "DOMINICANA":
        paginas = [
            st.Page("rd/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("rd/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("rd/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("rd/seguridad.py", title="Prevención y Control", icon="🛡️")
        ]
    
    pg = st.navigation(paginas)
    pg.run()
