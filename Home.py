import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import traceback
from datetime import datetime, timedelta
import pandas as pd
import os
import extra_streamlit_components as stx 

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered")

# ==========================================
# CREDENCIALES Y CONEXIÓN DINÁMICA
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

# ==========================================
# GESTIÓN DE USUARIOS Y ROLES
# ==========================================
USUARIOS = {
    "admin_vzla": {"pass": "Admin1234*", "rol": "Admin", "pais": "VENEZUELA"},
    "flota_vzla": {"pass": "Flota2026*", "rol": "Coordinador", "pais": "VENEZUELA"},
    "flota_rd": {"pass": "RDFlota2026*", "rol": "Coordinador", "pais": "DOMINICANA"},
    "admin_master": {"pass": "MasterDrotaca*", "rol": "Master", "pais": "MASTER_VZLA"},
    "compras_vzla": {"pass": "Compras2026*", "rol": "Compras", "pais": "COMPRAS_VZLA"}
}

@st.cache_resource
def get_cookie_manager():
    return stx.CookieManager()

cookie_manager = get_cookie_manager()

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
        st.session_state["usuario"] = None

    if st.session_state["logged_in"]:
        return True

    # Revisar cookie para autologin
    cookie_user = cookie_manager.get("pcd_usuario_valido")
    if cookie_user and cookie_user in USUARIOS:
        st.session_state["logged_in"] = True
        st.session_state["usuario"] = cookie_user
        return True

    return False

# ==========================================
# INTERFAZ DE LOGIN
# ==========================================
if not check_login():
    st.markdown("<h1 style='text-align: center; color: #0d47a1;'>PCD Internacional</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #666;'>Control Tower 2026</h3>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        st.write("🔒 Ingrese sus credenciales")
        usuario_input = st.text_input("Usuario")
        password_input = st.text_input("Contraseña", type="password")
        submit_button = st.form_submit_button("Ingresar al Sistema")
        
        if submit_button:
            if usuario_input in USUARIOS and USUARIOS[usuario_input]["pass"] == password_input:
                st.session_state["logged_in"] = True
                st.session_state["usuario"] = usuario_input
                # Guardar sesión por 30 días
                cookie_manager.set("pcd_usuario_valido", usuario_input, expires_at=datetime.now() + timedelta(days=30))
                st.success("✅ Acceso Concedido. Cargando módulos...")
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos")
else:
    # ==========================================
    # ENRUTAMIENTO DINÁMICO (MENÚ LATERAL)
    # ==========================================
    usuario_actual = st.session_state["usuario"]
    u_data = USUARIOS[usuario_actual]
    
    st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=80)
    st.sidebar.markdown(f"**Usuario:** {usuario_actual.upper()}")
    st.sidebar.markdown(f"**Rol:** {u_data['rol']}")
    st.sidebar.markdown(f"**Región:** {u_data['pais']}")
    
    if st.sidebar.button("🚪 Cerrar Sesión"):
        try:
            cookie_manager.delete("pcd_usuario_valido")
        except:
            pass 
        st.session_state["logged_in"] = False
        st.rerun()

    # Lógica de navegación por permisos
    if u_data['pais'] in ["VENEZUELA", "MASTER_VZLA"]:
        paginas = [
            st.Page("vzla/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("vzla/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("vzla/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("vzla/seguridad.py", title="Prevención y Control", icon="🛡️"),
            st.Page("vzla/cierre_semanal.py", title="Reporte Semanal", icon="📊"),
            st.Page("vzla/compras_flota.py", title="Solicitud de Compras", icon="🛒")
        ]
    elif u_data['pais'] == "COMPRAS_VZLA":
        paginas = [
            st.Page("vzla/compras_flota.py", title="Solicitud de Compras", icon="🛒")
        ]
    elif u_data['pais'] == "DOMINICANA":
        paginas = [
            st.Page("rd/cierre_diario.py", title="Cierre Diario (RD)", icon="📋"),
            st.Page("rd/flota.py", title="Flota y Gastos (RD)", icon="🚛")
        ]
    else:
        st.error("Configuración de región no encontrada.")
        st.stop()

    pg = st.navigation(paginas)
    pg.run()
