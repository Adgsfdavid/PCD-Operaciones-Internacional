# ==========================================
# Archivo: cierre_semanal.py (Auditoría de Tráfico y Despacho)
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
# INTERFAZ
# ==========================================
st.title("📊 Master Reporte Semanal: Desempeño de Tráfico")
st.markdown("Análisis detallado de tiempos de despacho, rutas y distribución por zonas.")

c1, c2 = st.columns(2)
with c1:
    ano_sel = st.selectbox("Año:", [2025, 2026], index=1)
with c2:
    semana_actual = datetime.now().isocalendar()[1]
    num_sem = st.number_input("Número de Semana:", 1, 53, value=semana_actual)

if st.button("⚡ GENERAR AUDITORÍA DE TRÁFICO", type="primary", use_container_width=True):
    with st.spinner(f"Procesando Hoja PIZARRA_TRAFICO..."):
        
        # 1. EXTRACCIÓN Y FILTRADO POR SEMANA
        df_raw = extraer_datos("PIZARRA_TRAFICO")
        
        if df_raw.empty:
            st.error("No se pudo conectar con la hoja PIZARRA_TRAFICO o está vacía.")
            st.stop()

        # Normalizamos la columna Semana para filtrar (ej. de "Semana 13" a 13)
        df_raw['Num_Semana'] = df_raw['Semana'].str.extract('(\\d+)').astype(float)
        df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()

        if df_sem.empty:
            st.warning(f"No hay datos registrados para la Semana {num_sem}.")
            st.stop()

        # ==========================================
        # 2. ANÁLISIS DE TIEMPOS (Por día distintivo)
        # ==========================================
        # Tomamos el primer registro de cada día para ver las horas generales de la operación
        df_tiempos = df_sem.drop_duplicates(subset=['Fecha'])[['Fecha', 'Dia', 'Hora_1er_Listin', 'Hora_Ultimo_Listin', 'Inicio_Trafico', 'Culminacion_Trafico']]

        # ==========================================
        # 3. DESGLOSE DE RUTAS (Sin chofer ni ayudante)
        # ==========================================
        df_rutas = df_sem[['Fecha', 'Dia', 'Ruta', 'Zona', 'Unidad', 'Listines', 'Farmacias_Total', 'Bultos_Total']].copy()
        
        # Totales generales de la semana
        total_farma_sem = pd.to_numeric(df_rutas['Farmacias_Total'], errors='coerce').sum()
        total_bultos_sem = pd.to_numeric(df_rutas['Bultos_Total'], errors='coerce').sum()

        # ==========================================
        # 4. ANÁLISIS POR ZONA (Porcentajes)
        # ==========================================
        df_zonas = df_rutas.groupby('Zona').agg({
            'Farmacias_Total': 'sum',
            'Bultos_Total': 'sum'
        }).reset_index()

        df_zonas['%_Farmacias'] = (df_zonas['Farmacias_Total'] / total_farma_sem * 100).round(1)
        df_zonas['%_Bultos'] = (df_zonas['Bultos_Total'] / total_bultos_sem * 100).round(1)

        # ==========================================
        # 5. CONSTRUCCIÓN DEL PDF (PRIMERA PÁGINA)
        # ==========================================
        logo = obtener_logo_base64()
        color_dro = "#0d47a1"

        # Generar filas HTML para Tiempos
        filas_tiempos = ""
        for _, r in df_tiempos.iterrows():
            filas_tiempos += f"<tr><td>{r['Fecha']}</td><td>{r['Dia']}</td><td>{r['Hora_1er_Listin']}</td><td>{r['Hora_Ultimo_Listin']}</td><td>{r['Inicio_Trafico']}</td><td>{r['Culminacion_Trafico']}</td></tr>"

        # Generar filas HTML para Rutas
        filas_rutas = ""
        for _, r in df_rutas.iterrows():
            filas_rutas += f"<tr><td>{r['Fecha']}</td><td>{r['Ruta']}</td><td>{r['Zona']}</td><td>{r['Unidad']}</td><td>{r['Listines']}</td><td style='font-weight:bold;'>{r['Farmacias_Total']}</td><td style='font-weight:bold;'>{r['Bultos_Total']}</td></tr>"

        # Generar filas HTML para Zonas
        filas_zonas = ""
        for _, r in df_zonas.iterrows():
            filas_zonas += f"""
            <tr style='background:#f1f8ff;'>
                <td style='text-align:left; font-weight:bold;'>{r['Zona']}</td>
                <td>{int(r['Farmacias_Total'])}</td>
                <td style='color:{color_dro}; font-weight:bold;'>{r['%_Farmacias']}%</td>
                <td>{int(r['Bultos_Total'])}</td>
                <td style='color:#e65100; font-weight:bold;'>{r['%_Bultos']}%</td>
            </tr>"""

        html_pdf = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                body {{ font-family: 'Montserrat', sans-serif; padding: 0; margin: 0; background:#525659; }}
                .page {{ width: 210mm; background: white; margin: 10mm auto; padding: 15mm; box-shadow: 0 0 10px rgba(0,0,0,0.5); }}
                .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid {color_dro}; padding-bottom: 10px; }}
                .section-title {{ background: {color_dro}; color: white; padding: 8px 15px; font-weight: 900; font-size: 14px; margin-top: 25px; text-transform: uppercase; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px; }}
                th {{ background: #eee; border: 1px solid #ccc; padding: 6px; text-align: center; }}
                td {{ border: 1px solid #ccc; padding: 6px; text-align: center; }}
                .total-bar {{ background: #263238; color: white; display: flex; justify-content: space-around; padding: 10px; margin-top: 10px; font-weight: 900; font-size: 14px; }}
                @media print {{ .no-print {{ display: none; }} body {{ background: white; }} .page {{ margin: 0; box-shadow: none; }} }}
            </style>
        </head>
        <body>
            <div class="no-print" style="text-align:center; padding:20px;">
                <button onclick="window.print()" style="background:#e65100; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px;">📥 DESCARGAR REPORTE DE TRÁFICO</button>
            </div>

            <div class="page">
                <div class="header">
                    <img src="{logo}" style="height: 50px;">
                    <div style="text-align:right;">
                        <h1 style="margin:0; color:{color_dro}; font-size:22px;">AUDITORÍA SEMANAL DE TRÁFICO</h1>
                        <p style="margin:0; font-weight:bold;">Semana {num_sem} | Año {ano_sel}</p>
                    </div>
                </div>

                <div class="section-title">⏱️ 1. CRONOMETRÍA DE DESPACHO (HORAS DE SALIDA)</div>
                <table>
                    <thead>
                        <tr><th>Fecha</th><th>Día</th><th>1er Listín</th><th>Último Listín</th><th>Inicio Tráfico</th><th>Fin Tráfico</th></tr>
                    </thead>
                    <tbody>{filas_tiempos}</tbody>
                </table>

                <div class="section-title">🚛 2. DESGLOSE OPERATIVO POR RUTA</div>
                <table>
                    <thead>
                        <tr><th>Fecha</th><th>Ruta</th><th>Zona</th><th>Unidad</th><th>Listines</th><th>Farmacias</th><th>Bultos</th></tr>
                    </thead>
                    <tbody>{filas_rutas}</tbody>
                </table>
                <div class="total-bar">
                    <span>TOTAL FARMACIAS SEMANAL: {int(total_farma_sem)}</span>
                    <span>TOTAL BULTOS SEMANAL: {int(total_bultos_sem)}</span>
                </div>

                <div class="section-title">🌍 3. EFECTIVIDAD Y DISTRIBUCIÓN POR ZONAS</div>
                <table>
                    <thead>
                        <tr><th style='text-align:left;'>Zona Logística</th><th>Farmacias</th><th>% Far.</th><th>Bultos</th><th>% Bul.</th></tr>
                    </thead>
                    <tbody>{filas_zonas}</tbody>
                </table>

                <div style="margin-top:30px; border-top:1px solid #eee; padding-top:10px; font-size:10px; color:#aaa; text-align:center;">
                    Página 1: Auditoría de Tráfico - Coordinación de Flota y Logística Drotaca
                </div>
            </div>
        </body>
        </html>
        """
        components.html(html_pdf, height=1200, scrolling=True)
        st.success("✅ Auditoría de Tráfico ensamblada con éxito.")
