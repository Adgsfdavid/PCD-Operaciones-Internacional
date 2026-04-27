import streamlit as st

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered")

# --- BASE DE DATOS DE USUARIOS ---
# Aquí controlamos las claves, el país y la Base de Datos a la que apuntan
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
    # Intenta cargar el logo si existe
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
            st.rerun() # Recarga la página
        else:
            st.error("Usuario o contraseña incorrectos")

# --- LÓGICA DE ENRUTAMIENTO (RUTAS) ---
if not st.session_state["logged_in"]:
    login()
else:
    # SI YA INICIÓ SESIÓN:
    u_data = st.session_state["user_data"]
    
    # 1. Menú Lateral (Botón de Cerrar Sesión)
    st.sidebar.title(f"📍 {u_data['pais']}")
    st.sidebar.caption(f"Conectado a: {u_data['sheet_name']}")
    
    if st.sidebar.button("🚪 Cerrar Sesión"):
        st.session_state["logged_in"] = False
        st.rerun()

    # 2. Definir qué páginas ve cada país usando st.Page()
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
    
    # 3. Ejecutar la navegación automática
    nav = st.navigation(paginas)
    nav.run()
