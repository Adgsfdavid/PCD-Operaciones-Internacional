import streamlit as st
import time
from pathlib import Path

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered", initial_sidebar_state="collapsed")

# --- BASE DE DATOS DE USUARIOS Y LOGOS ---
CONFIG_PAISES = {
    "admin_vzla": {
        "clave": "Vzla2026*", 
        "pais": "VENEZUELA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "logo": "logo.png",
        "color": "#FFD000"
    },
    "admin_rd": {
        "clave": "Dom2026*", 
        "pais": "DOMINICANA", 
        "sheet_name": "PCD_BaseDatos_DOM",
        "logo": "logo_rd.png", # Recuerda renombrar el archivo en GitHub
        "color": "#002D62"
    },
    "david_master": {
        "clave": "Master123", 
        "pais": "MASTER_VZLA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "logo": "logo.png",
        "color": "#FFD000"
    }
}

# --- ESTILOS CSS PARA LAS TARJETAS ---
st.markdown("""
<style>
    .card-login {
        background-color: #ffffff;
        padding: 30px 20px;
        border-radius: 15px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        text-align: center;
        border: 1px solid #eee;
        height: 100%;
    }
    .titulo-region {
        color: #1a237e;
        font-weight: 800;
        font-size: 1.2rem;
        margin-top: 15px;
        letter-spacing: 1px;
    }
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
</style>
""", unsafe_allow_html=True)

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# --- PANTALLA DE ACCESO ---
def login():
    st.markdown("<h1 style='text-align: center; color: #1a237e; font-weight: 900; margin-bottom: 0;'>PCD GLOBAL</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666; margin-bottom: 40px;'>Selección de región operativa</p>", unsafe_allow_html=True)
    
    col_v, col_r = st.columns(2)
    
    # PANEL VENEZUELA
    with col_v:
        st.markdown('<div class="card-login" style="border-top: 5px solid #FFD000;">', unsafe_allow_html=True)
        if Path("logo.png").exists():
            st.image("logo.png", width=180)
        st.markdown('<div class="titulo-region">VENEZUELA</div>', unsafe_allow_html=True)
        
        with st.form("form_vzla"):
            usr_v = st.text_input("Usuario")
            pwd_v = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar a Drotaca VZLA", use_container_width=True):
                if usr_v in CONFIG_PAISES and CONFIG_PAISES[usr_v]["clave"] == pwd_v and "VZLA" in CONFIG_PAISES[usr_v]["pais"]:
                    st.session_state["logged_in"] = True
                    st.session_state["user_data"] = CONFIG_PAISES[usr_v]
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
        st.markdown('</div>', unsafe_allow_html=True)

    # PANEL DOMINICANA
    with col_r:
        st.markdown('<div class="card-login" style="border-top: 5px solid #002D62;">', unsafe_allow_html=True)
        # Aquí usamos el nombre nuevo sugerido
        if Path("logo_rd.png").exists():
            st.image("logo_rd.png", width=180)
        else:
            st.warning("Subir logo_rd.png")
            
        st.markdown('<div class="titulo-region">DOMINICANA</div>', unsafe_allow_html=True)
        
        with st.form("form_rd"):
            usr_d = st.text_input("Usuario", key="u_rd")
            pwd_d = st.text_input("Contraseña", type="password", key="p_rd")
            if st.form_submit_button("Entrar a Drotafarma RD", use_container_width=True):
                if usr_d in CONFIG_PAISES and CONFIG_PAISES[usr_d]["clave"] == pwd_d and "DOMINICANA" in CONFIG_PAISES[usr_d]["pais"]:
                    st.session_state["logged_in"] = True
                    st.session_state["user_data"] = CONFIG_PAISES[usr_d]
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
        st.markdown('</div>', unsafe_allow_html=True)

# --- NAVEGACIÓN ---
if not st.session_state["logged_in"]:
    login()
else:
    u_data = st.session_state["user_data"]
    
    # Sidebar con el logo correspondiente
    if Path(u_data["logo"]).exists():
        st.sidebar.image(u_data["logo"], use_container_width=True)
    
    st.sidebar.markdown(f"<h3 style='text-align: center; color: #1a237e; margin-top: 0;'>{u_data['pais']}</h3>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()

    # Enrutamiento de páginas
    if "VZLA" in u_data['pais']:
        paginas = [
            st.Page("vzla/cierre_diario.py", title="Cierre Diario Master", icon="📋"),
            st.Page("vzla/flota.py", title="Flota y Mantenimiento", icon="🚛"),
            st.Page("vzla/monitoreo.py", title="Monitoreo de Despachos", icon="🖥️"),
            st.Page("vzla/seguridad.py", title="Prevención y Control", icon="🛡️"),
            st.Page("vzla/app.py", title="Tráfico y Transbordos", icon="🚦")
        ]
    else:
        paginas = [
            st.Page("rd/cierre_diario.py", title="Cierre Diario (RD)", icon="📋"),
            st.Page("rd/flota.py", title="Flota y Combustible (RD)", icon="🚛")
        ]
    
    nav = st.navigation(paginas)
    nav.run()
