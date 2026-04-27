import streamlit as st
import time
from pathlib import Path

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered", initial_sidebar_state="collapsed")

# --- BASE DE DATOS DE USUARIOS, LOGOS Y BANDERAS ---
CONFIG_PAISES = {
    "admin_vzla": {
        "clave": "Vzla2026*", 
        "pais": "VENEZUELA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "logo": "logo.png",
        "bandera_url": "https://flagcdn.com/w160/ve.png"
    },
    "admin_rd": {
        "clave": "Dom2026*", 
        "pais": "DOMINICANA", 
        "sheet_name": "PCD_BaseDatos_DOM",
        "logo": "logo_rd.png",
        "bandera_url": "https://flagcdn.com/w160/do.png"
    },
    "david_master": {
        "clave": "Master123", 
        "pais": "MASTER_VZLA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "logo": "logo.png",
        "bandera_url": "https://flagcdn.com/w160/ve.png"
    }
}

# --- ESTILOS CSS PARA EL DISEÑO SOLICITADO ---
st.markdown("""
<style>
    .card-login {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        text-align: center;
        border: 1px solid #eee;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .titulo-region {
        color: #1a237e;
        font-weight: 800;
        font-size: 1.1rem;
        margin: 10px 0;
        letter-spacing: 1px;
    }
    .bandera-footer {
        width: 80px;
        border-radius: 4px;
        margin-top: 15px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
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
    st.markdown("<h1 style='text-align: center; color: #1a237e; font-weight: 900; margin-bottom: 40px;'>PCD GLOBAL</h1>", unsafe_allow_html=True)
    
    col_v, col_r = st.columns(2)
    
    # --- PANEL VENEZUELA ---
    with col_v:
        # Se eliminó el estilo de borde de color
        st.markdown('<div class="card-login">', unsafe_allow_html=True)
        # LOGO ARRIBA
        if Path("logo.png").exists():
            st.image("logo.png", width=160)
        
        st.markdown('<div class="titulo-region">VENEZUELA</div>', unsafe_allow_html=True)
        
        with st.form("form_vzla"):
            usr_v = st.text_input("Usuario")
            pwd_v = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar VZLA", use_container_width=True):
                if usr_v in CONFIG_PAISES and CONFIG_PAISES[usr_v]["clave"] == pwd_v and "VZLA" in CONFIG_PAISES[usr_v]["pais"]:
                    st.session_state["logged_in"] = True
                    st.session_state["user_data"] = CONFIG_PAISES[usr_v]
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
        
        # BANDERA ABAJO
        st.markdown(f'<center><img src="{CONFIG_PAISES["admin_vzla"]["bandera_url"]}" class="bandera-footer"></center>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # --- PANEL DOMINICANA ---
    with col_r:
        # Se eliminó el estilo de borde de color
        st.markdown('<div class="card-login">', unsafe_allow_html=True)
        # LOGO ARRIBA
        if Path("logo_rd.png").exists():
            st.image("logo_rd.png", width=160)
        else:
            st.info("Subir logo_rd.png")
            
        st.markdown('<div class="titulo-region">DOMINICANA</div>', unsafe_allow_html=True)
        
        with st.form("form_rd"):
            usr_d = st.text_input("Usuario", key="u_rd")
            pwd_d = st.text_input("Contraseña", type="password", key="p_rd")
            if st.form_submit_button("Entrar RD", use_container_width=True):
                if usr_d in CONFIG_PAISES and CONFIG_PAISES[usr_d]["clave"] == pwd_d and "DOMINICANA" in CONFIG_PAISES[usr_d]["pais"]:
                    st.session_state["logged_in"] = True
                    st.session_state["user_data"] = CONFIG_PAISES[usr_d]
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas")
        
        # BANDERA ABAJO
        st.markdown(f'<center><img src="{CONFIG_PAISES["admin_rd"]["bandera_url"]}" class="bandera-footer"></center>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# --- NAVEGACIÓN ---
if not st.session_state["logged_in"]:
    login()
else:
    u_data = st.session_state["user_data"]
    
    # Sidebar con el LOGO arriba
    if Path(u_data["logo"]).exists():
        st.sidebar.image(u_data["logo"], use_container_width=True)
    
    st.sidebar.markdown(f"<h3 style='text-align: center; color: #1a237e; margin-top: 0;'>{u_data['pais']}</h3>", unsafe_allow_html=True)
    
    # Pequeña bandera discreta debajo del nombre del país en el sidebar
    st.sidebar.markdown(f"""
        <center><img src='{u_data['bandera_url']}' style='width: 40px; border-radius: 3px; border: 1px solid #ddd;'></center>
    """, unsafe_allow_html=True)
    
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
