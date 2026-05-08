# ==========================================
# Archivo: cierre_semanal.py (Master Semanal Drotaca)
# ==========================================
import streamlit as st
import pandas as pd
import base64
import textwrap
from datetime import datetime
import streamlit.components.v1 as components
from google.oauth2.service_account import Credentials
import gspread

# ==========================================
# CONFIGURACIÓN DE CONEXIÓN
# ==========================================
def obtener_logo_base64():
    try:
        from pathlib import Path
        ruta_logo = Path(__file__).parent / "logo.png"
        if ruta_logo.exists():
            with open(ruta_logo, "rb") as image_file:
                return f"data:image/png;base64,{base64.b64encode(image_file.read()).decode()}"
        return None
    except: return None

# Parche de seguridad para la llave de Google
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"
CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def extraer_datos(nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        hoja = doc.worksheet(nombre_hoja)
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# INTERFAZ Y FILTRADO
# ==========================================
st.title("📊 Master Reporte Semanal de Gestión")
st.markdown("Consolidado detallado de las 10 áreas operativas de Drotaca.")

c1, c2 = st.columns(2)
with c1:
    ano_sel = st.selectbox("Año:", [2025, 2026], index=1)
with c2:
    semana_actual = datetime.now().isocalendar()[1]
    num_sem = st.number_input("Número de Semana:", 1, 53, value=semana_actual)

if st.button("⚡ GENERAR DETALLADO SEMANAL", type="primary", use_container_width=True):
    with st.spinner(f"Analizando las 10 áreas para la Semana {num_sem}..."):
        
        # 1. MAPEO DE HOJAS
        hojas = {
            "Apertura": "SEG_APERTURA",
            "M_Plan": "FLOTA_PLANIFICADO",
            "M_Real": "FLOTA_REALIZADO",
            "Despachos": "MONITOREO_DESPACHOS",
            "Trafico": "PIZARRA_TRAFICO",
            "Guardia": "SEG_ROL_GUARDIA",
            "Juanita": "SEG_CIERRE_JUANITA",
            "Cierre_Dro": "SEG_CIERRE_DROTACA",
            "Surtido": "SURTIDO_COMBUSTIBLE",
            "Reserva": "FLOTA_COMBUSTIBLE"
        }
        
        data = {k: extraer_datos(v) for k, v in hojas.items()}
        
        def filtrar_sem(df):
            if df.empty: return df
            col = 'Fecha' if 'Fecha' in df.columns else ('Fecha Sistema' if 'Fecha Sistema' in df.columns else None)
            if not col: return df
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
            return df[df[col].dt.isocalendar().week == num_sem]

        f = {k: filtrar_sem(v) for k, v in data.items()}

        # ==========================================
        # 2. PROCESAMIENTO POR ÁREA (DETALLADO)
        # ==========================================
        
        # SECCIÓN SEGURIDAD: Aperturas y Cierres
        avg_apertura = f['Apertura']['Hora'].mode()[0] if not f['Apertura'].empty else "N/A"
        avg_cierre_dro = f['Cierre_Dro']['Hora'].mode()[0] if not f['Cierre_Dro'].empty else "N/A"
        avg_juanita = f['Juanita']['Hora'].mode()[0] if not f['Juanita'].empty else "N/A"

        # SECCIÓN FLOTA: Cruce de Mantenimiento
        m_plan_list = set(f['M_Plan']['Unidad']) if 'Unidad' in f['M_Plan'].columns else set()
        m_real_list = set(f['M_Real']['Unidad']) if 'Unidad' in f['M_Real'].columns else set()
        unidades_pendientes = list(m_plan_list - m_real_list)
        efectividad = (len(m_real_list) / len(m_plan_list) * 100) if m_plan_list else 100

        # SECCIÓN OPERACIONES: Bultos y KMs
        t_bultos = pd.to_numeric(f['Despachos']['Bultos'], errors='coerce').sum()
        t_farmacias = pd.to_numeric(f['Despachos']['Farmacias'], errors='coerce').sum()
        t_kms = pd.to_numeric(f['Surtido']['Kms'], errors='coerce').sum()
        t_litros = pd.to_numeric(f['Surtido']['Litros'], errors='coerce').sum()

        # ==========================================
        # 3. GENERACIÓN VISUAL DEL REPORTE
        # ==========================================
        st.subheader("📋 Resumen Ejecutivo de KPIs")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📦 Bultos Totales", f"{t_bultos:,.0f}")
        k2.metric("🏥 Farmacias", f"{t_farmacias:,.0f}")
        k3.metric("🛣️ Kilometraje", f"{t_kms:,.0f} Km")
        k4.metric("🛠️ Efec. Mantenimiento", f"{efectividad:.1f}%")

        # ACORDEONES CON DETALLES (Lo que pediste)
        with st.expander("🔍 Detalle de Seguridad (Aperturas y Cierres)"):
            st.write(f"**Hora más frecuente de Apertura:** {avg_apertura}")
            st.write(f"**Hora más frecuente Cierre Drotaca:** {avg_cierre_dro}")
            st.write(f"**Hora más frecuente Cierre Juanita:** {avg_juanita}")
            st.dataframe(f['Guardia'], use_container_width=True, hide_index=True)

        with st.expander("🔍 Detalle de Flota y Combustible"):
            st.write(f"**Unidades Pendientes por Mantenimiento:** {', '.join(unidades_pendientes) if unidades_pendientes else 'Ninguna'}")
            st.write(f"**Consumo Total de Gasoil:** {t_litros:,.2f} Lts")
            st.dataframe(f['Reserva'], use_container_width=True, hide_index=True)

        with st.expander("🔍 Detalle de Tráfico y Despachos"):
            st.dataframe(f['Trafico'], use_container_width=True, hide_index=True)

        # GENERADOR PDF COMPLETO
        logo = obtener_logo_base64()
        html_pdf = f"""
        <div style="font-family: sans-serif; padding: 20px; border-top: 10px solid #0d47a1; background: white;">
            <div style="display: flex; justify-content: space-between;">
                <img src="{logo}" style="height: 60px;">
                <div style="text-align: right;">
                    <h2 style="margin:0; color:#0d47a1;">REPORTE SEMANAL MASTER</h2>
                    <p style="margin:0;">Semana {num_sem} - Año {ano_sel}</p>
                </div>
            </div>
            <hr>
            <h3 style="color:#0d47a1;">1. SEGURIDAD Y CIERRES</h3>
            <p>Apertura Promedio: {avg_apertura} | Cierre Drotaca: {avg_cierre_dro} | Cierre Juanita: {avg_juanita}</p>
            
            <h3 style="color:#0d47a1;">2. FLOTA Y MANTENIMIENTO</h3>
            <p>Efectividad: {efectividad:.1f}% | Kms: {t_kms:,.0f} | Gasoil: {t_litros:,.0f} Lts</p>
            <p><b>Pendientes de Mantenimiento:</b> {', '.join(unidades_pendientes) if unidades_pendientes else 'Ninguna'}</p>
            
            <h3 style="color:#0d47a1;">3. LOGÍSTICA Y DESPACHOS</h3>
            <p>Total Bultos: {t_bultos:,.0f} | Total Farmacias: {t_farmacias:,.0f}</p>
            <div style="text-align: center; margin-top: 50px;">
                <button onclick="window.print()" style="padding: 10px 20px; background: #e65100; color: white; border: none; cursor: pointer;">Descargar PDF Completo</button>
            </div>
        </div>
        """
        components.html(html_pdf, height=800)
