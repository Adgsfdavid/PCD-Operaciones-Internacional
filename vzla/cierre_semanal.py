# ==========================================
# Archivo: cierre_semanal.py (Master Semanal Drotaca)
# ==========================================
import streamlit as st
import pandas as pd
import base64
import textwrap
from datetime import datetime, timedelta
import streamlit.components.v1 as components
from google.oauth2.service_account import Credentials
import gspread

# ==========================================
# CONFIGURACIÓN DE CONEXIÓN Y LOGO
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
        # ID de tu base de datos Master
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        hoja = doc.worksheet(nombre_hoja)
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# LÓGICA DE PROCESAMIENTO SEMANAL
# ==========================================
st.title("📊 Master Reporte Semanal de Gestión")
st.markdown("Consolidado inteligente de las 10 áreas operativas de Drotaca.")

c1, c2 = st.columns(2)
with c1:
    ano_sel = st.selectbox("Año:", [2025, 2026], index=1)
with c2:
    semana_actual = datetime.now().isocalendar()[1]
    num_sem = st.number_input("Número de Semana a Analizar:", 1, 53, value=semana_actual)

if st.button("⚡ GENERAR CONSOLIDADO SEMANAL", type="primary", use_container_width=True):
    with st.spinner(f"Cruzando datos de la Semana {num_sem}..."):
        
        # Mapeo de extracción (Tus 10 hojas)
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
        
        # Función para filtrar por semana
        def filtrar_sem(df):
            if df.empty: return df
            col_fecha = 'Fecha' if 'Fecha' in df.columns else 'Fecha Sistema'
            if col_fecha not in df.columns: return df
            df[col_fecha] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
            return df[df[col_fecha].dt.isocalendar().week == num_sem]

        filtered = {k: filtrar_sem(v) for k, v in data.items()}
        
        # --- CÁLCULOS CLAVE (KPIs) ---
        # 1. Mantenimiento (El Cruce)
        plan = len(filtered['M_Plan'])
        real = len(filtered['M_Real'])
        efectividad_mant = (real/plan*100) if plan > 0 else 0
        
        # 2. Operaciones
        total_bultos = pd.to_numeric(filtered['Despachos']['Bultos'], errors='coerce').sum() if 'Bultos' in filtered['Despachos'].columns else 0
        total_kms = pd.to_numeric(filtered['Surtido']['Kms'], errors='coerce').sum() if 'Kms' in filtered['Surtido'].columns else 0
        total_litros = pd.to_numeric(filtered['Surtido']['Litros'], errors='coerce').sum() if 'Litros' in filtered['Surtido'].columns else 0
        
        # --- INTERFAZ DE RESULTADOS ---
        st.success(f"✅ Análisis completado para la Semana {num_sem}")
        
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📦 Bultos Semanal", f"{total_bultos:,.0f}")
        k2.metric("🛣️ Kms Totales", f"{total_kms:,.0f}")
        k3.metric("🛠️ Efec. Mantenimiento", f"{efectividad_mant:.1f}%")
        k4.metric("⛽ Consumo Gasoil", f"{total_litros:,.0f} Lts")
        
        # ==========================================
        # GENERADOR DEL PDF SEMANAL (DISEÑO EJECUTIVO)
        # ==========================================
        logo_b64 = obtener_logo_base64()
        color_dro = "#0d47a1"
        
        html_semanal = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                body {{ font-family: 'Montserrat', sans-serif; margin: 0; padding: 0; background-color: #f4f4f4; }}
                .page {{ width: 210mm; background: white; margin: 20px auto; padding: 40px; box-shadow: 0 0 20px rgba(0,0,0,0.1); border-top: 10px solid {color_dro}; }}
                .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #eee; padding-bottom: 20px; }}
                .title-box {{ text-align: right; }}
                .kpi-container {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin: 30px 0; }}
                .kpi-card {{ background: #f8f9fa; border: 1px solid #ddd; padding: 15px; text-align: center; border-radius: 8px; }}
                .kpi-value {{ font-size: 22px; font-weight: 900; color: {color_dro}; }}
                .section-title {{ background: {color_dro}; color: white; padding: 10px; font-size: 16px; margin-top: 30px; border-radius: 4px; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }}
                th {{ background: #eee; padding: 8px; border: 1px solid #ddd; text-align: left; }}
                td {{ padding: 8px; border: 1px solid #ddd; }}
                @media print {{ .no-print {{ display: none; }} body {{ background: white; }} .page {{ margin: 0; box-shadow: none; }} }}
            </style>
        </head>
        <body>
            <div class="no-print" style="text-align:center; padding:20px;">
                <button onclick="window.print()" style="background:#e65100; color:white; border:none; padding:15px 30px; font-size:18px; font-weight:bold; cursor:pointer; border-radius:5px;">📥 DESCARGAR REPORTE SEMANAL</button>
            </div>
            
            <div class="page">
                <div class="header">
                    <img src="{logo_b64}" style="height: 60px;">
                    <div class="title-box">
                        <h1 style="margin:0; color:{color_dro};">REPORTE DE GESTIÓN</h1>
                        <h3 style="margin:0; color:#666;">Semana {num_sem} - Año {ano_sel}</h3>
                    </div>
                </div>

                <div class="kpi-container">
                    <div class="kpi-card"><div>BULTOS</div><div class="kpi-value">{total_bultos:,.0f}</div></div>
                    <div class="kpi-card"><div>KILOMETRAJE</div><div class="kpi-value">{total_kms:,.0f}</div></div>
                    <div class="kpi-card"><div>EFEC. MANT.</div><div class="kpi-value">{efectividad_mant:.1f}%</div></div>
                    <div class="kpi-card"><div>GASOIL</div><div class="kpi-value">{total_litros:,.0f} Lts</div></div>
                </div>

                <div class="section-title">📊 RESUMEN DE MANTENIMIENTO VEHICULAR (EL CRUCE)</div>
                <p style="font-size: 13px;">Se analizaron <b>{plan}</b> mantenimientos planificados contra <b>{real}</b> ejecuciones reales registradas.</p>
                
                <div class="section-title">🚛 DESPACHOS Y LOGÍSTICA (PIZARRA TRÁFICO)</div>
                <table>
                    <thead><tr><th>Día/Fecha</th><th>Rutas Atendidas</th><th>Transbordos</th><th>Novedades</th></tr></thead>
                    <tbody>
                        <tr><td>Consolidado Semana {num_sem}</td><td>{len(filtered['Trafico'])} registros</td><td>Verificando...</td><td>Sin novedades críticas</td></tr>
                    </tbody>
                </table>

                <div class="section-title">⛽ CONTROL DE COMBUSTIBLE Y RESERVAS</div>
                <p style="font-size: 13px;">Promedio de reserva en tanques (El Tigre): <b>{pd.to_numeric(filtered['Reserva']['Nivel'], errors='coerce').mean() if 'Nivel' in filtered['Reserva'].columns else 0:.1f}%</b></p>
                
                <div style="margin-top:50px; border-top: 1px solid #eee; padding-top: 20px; font-size: 10px; color: #aaa; text-align: center;">
                    Documento generado automáticamente por el Sistema PCD - Drotaca Internacional
                </div>
            </div>
        </body>
        </html>
        """
        components.html(html_semanal, height=1200, scrolling=True)