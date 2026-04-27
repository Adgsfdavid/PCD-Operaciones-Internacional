import streamlit as st
import time

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered", initial_sidebar_state="collapsed")

# --- BASE DE DATOS DE USUARIOS ---
CONFIG_PAISES = {
    "admin_vzla": {
        "clave": "Vzla2026*", 
        "pais": "VENEZUELA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "bandera": "🇻🇪"
    },
    "admin_rd": {
        "clave": "Dom2026*", 
        "pais": "DOMINICANA", 
        "sheet_name": "PCD_BaseDatos_DOM",
        "bandera": "🇩🇴"
    },
    "david_master": {
        "clave": "Master123", 
        "pais": "MASTER_VZLA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "bandera": "🇻🇪👑"
    }
}

# --- ESTILOS CSS PERSONALIZADOS ---
st.markdown("""
<style>
    .card-vzla {
        background-color: #f8f9fa;
        border-top: 8px solid #FFD000; /* Amarillo Vzla */
        border-bottom: 8px solid #CE1126; /* Rojo Vzla */
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        text-align: center;
    }
    .card-rd {
        background-color: #f8f9fa;
        border-top: 8px solid #002D62; /* Azul RD */
        border-bottom: 8px solid #CE1126; /* Rojo RD */
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        text-align: center;
    }
    .titulo-card {
        color: #1a237e;
        font-weight: 900;
        font-size: 20px;
        margin-bottom: 5px;
    }
</style>
""", unsafe_allow_html=True)

# --- ESTADO DE SESIÓN ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# --- PANTALLA DE LOGIN DOBLE ---
def login():
    st.markdown("<h1 style='text-align: center; color: #1a237e;'>PCD GLOBAL</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>Selecciona tu región operativa para ingresar</p>", unsafe_allow_html=True)
    st.write("")
    st.write("")
    
    col_v, col_r = st.columns(2)
    
    # --- PANEL VENEZUELA ---
    with col_v:
        st.markdown('<div class="card-vzla">', unsafe_allow_html=True)
        st.markdown("<div style='font-size: 50px;'>🇻🇪</div>", unsafe_allow_html=True)
        st.markdown('<div class="titulo-card">VENEZUELA</div>', unsafe_allow_html=True)
        
        with st.form("form_vzla"):
            usr_v = st.text_input("Usuario")
            pwd_v = st.text_input("Contraseña", type="password")
            btn_v = st.form_submit_button("Ingresar 🇻🇪", use_container_width=True)
            
            if btn_v:
                if usr_v in CONFIG_PAISES and CONFIG_PAISES[usr_v]["clave"] == pwd_v and "VZLA" in CONFIG_PAISES[usr_v]["pais"]:
                    st.session_state["logged_in"] = True
                    st.session_state["user_data"] = CONFIG_PAISES[usr_v]
                    st.success("Acceso concedido...")
                    time.sleep(0.5)
                    st.rerun()
                else:
                    st.error("Credenciales inválidas para Venezuela")
        st.markdown('</div>', unsafe_allow_html=True)

    # --- PANEL DOMINICANA ---
    with col_r:
        st.markdown('<div class="card-rd">', unsafe_allow_html=True)
        st.markdown("<div style='font-size: 50px;'>🇩🇴</div>", unsafe_allow_html=True)
        st.markdown('<div class="titulo-card">DOMINICANA</div>', unsafe_allow_html=True)
        
        with st.form("form_rd"):
            usr_d = st.text_input("Usuario", key="usr_d")
            pwd_d = st.text_input("Contraseña", type="password", key="pwd_d")
            btn_d = st.form_submit_button("Ingresar 🇩🇴", use_container_width=True)
            
            if btn_d:
                if usr_d in CONFIG_PAISES and CONFIG_PAISES[usr_d]["clave"] == pwd_d and "DOMINICANA" in CONFIG_PAISES[usr_d]["pais"]:
                    st.session_state["logged_in"] = True
                    st.session_state["user_data"] = CONFIG_PAISES[usr_d]
                    st.success("Acceso concedido...")
                    time.sleep(0.5)
                    st.rerun()
                else:
                    st.error("Credenciales inválidas para Dominicana")
        st.markdown('</div>', unsafe_allow_html=True)


# --- LÓGICA DE ENRUTAMIENTO (RUTAS) ---
if not st.session_state["logged_in"]:
    login()
else:
    # SI YA INICIÓ SESIÓN:
    u_data = st.session_state["user_data"]
    
    # 1. Menú Lateral (Bandera Gigante y Botón de Cerrar Sesión)
    st.sidebar.markdown(f"<div style='text-align: center; font-size: 100px; margin-bottom: -20px;'>{u_data['bandera']}</div>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<h2 style='text-align: center; color: #1a237e;'>{u_data['pais']}</h2>", unsafe_allow_html=True)
    st.sidebar.caption(f"Base de datos: {u_data['sheet_name']}")
    st.sidebar.markdown("---")
    
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()

    # 2. Definir qué páginas ve cada país
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
    
    # 3. Ejecutar la navegación automática
    nav = st.navigation(paginas)
    nav.run()
