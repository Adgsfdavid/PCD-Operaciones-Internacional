import streamlit as st
import time

# Configuración de página principal
st.set_page_config(page_title="PCD Internacional - Login", layout="centered", initial_sidebar_state="collapsed")

# --- BASE DE DATOS DE USUARIOS ---
# Añadimos enlaces a banderas en Alta Resolución (HD)
CONFIG_PAISES = {
    "admin_vzla": {
        "clave": "Vzla2026*", 
        "pais": "VENEZUELA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "bandera_url": "https://flagcdn.com/w160/ve.png"
    },
    "admin_rd": {
        "clave": "Dom2026*", 
        "pais": "DOMINICANA", 
        "sheet_name": "PCD_BaseDatos_DOM",
        "bandera_url": "https://flagcdn.com/w160/do.png"
    },
    "david_master": {
        "clave": "Master123", 
        "pais": "MASTER_VZLA", 
        "sheet_name": "PCD_BaseDatos_VZLA",
        "bandera_url": "https://flagcdn.com/w160/ve.png"
    }
}

# --- ESTILOS CSS PERSONALIZADOS ---
st.markdown("""
<style>
    .card-vzla {
        background-color: #f8f9fa;
        border-top: 8px solid #FFD000; /* Amarillo Vzla */
        border-bottom: 8px solid #CE1126; /* Rojo Vzla */
        padding: 25px 20px 20px 20px;
        border-radius: 12px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        text-align: center;
        transition: transform 0.3s;
    }
    .card-rd {
        background-color: #f8f9fa;
        border-top: 8px solid #002D62; /* Azul RD */
        border-bottom: 8px solid #CE1126; /* Rojo RD */
        padding: 25px 20px 20px 20px;
        border-radius: 12px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        text-align: center;
        transition: transform 0.3s;
    }
    .card-vzla:hover, .card-rd:hover {
        transform: translateY(-5px);
    }
    .titulo-card {
        color: #1a237e;
        font-weight: 900;
        font-size: 22px;
        margin-bottom: 15px;
        letter-spacing: 1px;
    }
    .bandera-img {
        width: 100px;
        border-radius: 8px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.3);
        margin-bottom: 15px;
        border: 1px solid #e0e0e0;
    }
</style>
""", unsafe_allow_html=True)

# --- ESTADO DE SESIÓN ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# --- PANTALLA DE LOGIN DOBLE ---
def login():
    st.markdown("<h1 style='text-align: center; color: #1a237e; font-weight: 900;'>PCD GLOBAL</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666; font-size: 16px; margin-bottom: 30px;'>Selecciona tu región operativa para ingresar</p>", unsafe_allow_html=True)
    
    col_v, col_r = st.columns(2)
    
    # --- PANEL VENEZUELA ---
    with col_v:
        st.markdown('<div class="card-vzla">', unsafe_allow_html=True)
        st.markdown(f'<img src="{CONFIG_PAISES["admin_vzla"]["bandera_url"]}" class="bandera-img">', unsafe_allow_html=True)
        st.markdown('<div class="titulo-card">VENEZUELA</div>', unsafe_allow_html=True)
        
        with st.form("form_vzla"):
            usr_v = st.text_input("Usuario")
            pwd_v = st.text_input("Contraseña", type="password")
            btn_v = st.form_submit_button("Ingresar", use_container_width=True)
            
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
        st.markdown(f'<img src="{CONFIG_PAISES["admin_rd"]["bandera_url"]}" class="bandera-img">', unsafe_allow_html=True)
        st.markdown('<div class="titulo-card">DOMINICANA</div>', unsafe_allow_html=True)
        
        with st.form("form_rd"):
            usr_d = st.text_input("Usuario", key="usr_d")
            pwd_d = st.text_input("Contraseña", type="password", key="pwd_d")
            btn_d = st.form_submit_button("Ingresar", use_container_width=True)
            
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
    
    # 1. Menú Lateral (Bandera HD Gigante y Botón de Cerrar Sesión)
    st.sidebar.markdown(f"""
        <div style='text-align: center; margin-bottom: 10px;'>
            <img src='{u_data['bandera_url']}' style='width: 120px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.3); border: 2px solid #fff;'>
        </div>
    """, unsafe_allow_html=True)
    st.sidebar.markdown(f"<h2 style='text-align: center; color: #1a237e; font-weight: 900; margin-top: 0;'>{u_data['pais']}</h2>", unsafe_allow_html=True)
    st.sidebar.caption(f"<div style='text-align: center;'>Base de datos: <b>{u_data['sheet_name']}</b></div>", unsafe_allow_html=True)
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
