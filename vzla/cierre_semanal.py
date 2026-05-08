# ==========================================
# Archivo: cierre_semanal.py (Master Semanal de Auditoría Logística)
# ==========================================
import streamlit as st
import pandas as pd
import base64
import textwrap
import json
from datetime import datetime
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
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        hoja = doc.worksheet(nombre_hoja)
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# FUNCIONES DE ANÁLISIS INTELIGENTE
# ==========================================
def buscar_columna(df, palabras_clave):
    for col in df.columns:
        if any(p.lower() in str(col).lower() for p in palabras_clave):
            return col
    return None

def calcular_promedio_hora(df, palabras_clave):
    col = buscar_columna(df, palabras_clave)
    if col and not df.empty:
        # Intentamos sacar la moda (hora más frecuente) para reflejar la realidad operativa
        moda = df[col].astype(str).mode()
        return moda[0] if not moda.empty else "N/A"
    return "N/A"

# ==========================================
# INTERFAZ Y PROCESAMIENTO
# ==========================================
st.title("📊 Master Reporte Semanal de Gestión Operativa")
st.markdown("Auditoría integral basada en los 10 pilares de información de la sede El Tigre.")

c1, c2 = st.columns(2)
with c1:
    ano_sel = st.selectbox("Año Fiscal:", [2025, 2026], index=1)
with c2:
    semana_actual = datetime.now().isocalendar()[1]
    num_sem = st.number_input("Número de Semana:", 1, 53, value=semana_actual)

if st.button("⚡ GENERAR AUDITORÍA SEMANAL", type="primary", use_container_width=True):
    with st.spinner(f"Sincronizando con base de datos VZLA..."):
        
        # 1. MAPEO DE DATOS SEGÚN ESTRUCTURA SOLICITADA
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
            col = buscar_columna(df, ['fecha', 'sistema'])
            if not col: return df
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
            return df[df[col].dt.isocalendar().week == num_sem]

        f = {k: filtrar_sem(v) for k, v in data.items()}

        # --- CÁLCULO DE MÉTRICAS EJECUTIVAS ---
        
        # Seguridad y Tiempos (Priorizando CIERRE DE DROGUERÍA)
        hora_apertura = calcular_promedio_hora(f['Apertura'], ['hora'])
        hora_cierre_dro = calcular_promedio_hora(f['Cierre_Dro'], ['CIERRE DE DROGUERÍA', 'hora'])
        hora_juanita = calcular_promedio_hora(f['Juanita'], ['hora'])

        # Mantenimiento (El Cruce Solicitado)
        col_u_p = buscar_columna(f['M_Plan'], ['unidad', 'placa'])
        col_u_r = buscar_columna(f['M_Real'], ['unidad', 'placa'])
        unidades_plan = set(f['M_Plan'][col_u_p].dropna()) if col_u_p else set()
        unidades_real = set(f['M_Real'][col_u_r].dropna()) if col_u_r else set()
        unidades_pendientes = list(unidades_plan - unidades_real)
        cumplimiento_mantenimiento = (len(unidades_real) / len(unidades_plan) * 100) if unidades_plan else 100

        # Logística
        bultos_sem = pd.to_numeric(f['Despachos'][buscar_columna(f['Despachos'], ['bultos'])], errors='coerce').sum() if f['Despachos'].columns.any() else 0
        kms_sem = pd.to_numeric(f['Surtido'][buscar_columna(f['Surtido'], ['kms'])], errors='coerce').sum() if f['Surtido'].columns.any() else 0
        litros_sem = pd.to_numeric(f['Surtido'][buscar_columna(f['Surtido'], ['litros'])], errors='coerce').sum() if f['Surtido'].columns.any() else 0

        # ==========================================
        # CONSTRUCCIÓN DEL REPORTE MULTI-PÁGINA (PDF)
        # ==========================================
        logo = obtener_logo_base64()
        color_p = "#0d47a1"

        html_final = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700;900&display=swap');
                body {{ font-family: 'Roboto', sans-serif; margin: 0; padding: 0; background-color: #525659; }}
                .page {{ width: 210mm; height: 296mm; background: white; margin: 10mm auto; padding: 20mm; box-sizing: border-box; position: relative; box-shadow: 0 0 15px rgba(0,0,0,0.5); page-break-after: always; }}
                .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid {color_p}; padding-bottom: 10px; margin-bottom: 20px; }}
                .title-main {{ color: {color_p}; font-size: 28px; font-weight: 900; margin: 0; text-transform: uppercase; }}
                .section-box {{ background: #f4f6f9; border-left: 5px solid {color_p}; padding: 15px; margin-bottom: 20px; }}
                .section-title {{ color: {color_p}; font-weight: 700; font-size: 18px; margin-bottom: 10px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }}
                .kpi-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-top: 10px; }}
                .kpi-card {{ background: white; border: 1px solid #ddd; padding: 10px; text-align: center; border-radius: 5px; }}
                .kpi-val {{ font-size: 20px; font-weight: 900; color: #e65100; }}
                table {{ width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 10px; }}
                th {{ background: {color_p}; color: white; padding: 8px; border: 1px solid #ddd; }}
                td {{ padding: 6px; border: 1px solid #ddd; text-align: center; }}
                .footer {{ position: absolute; bottom: 20mm; left: 20mm; right: 20mm; border-top: 1px solid #eee; padding-top: 10px; font-size: 10px; color: #777; text-align: center; }}
                @media print {{ body {{ background: white; }} .page {{ margin: 0; box-shadow: none; }} .no-print {{ display: none; }} }}
            </style>
        </head>
        <body>
            <div class="no-print" style="text-align: center; padding: 20px;">
                <button onclick="window.print()" style="background: #e65100; color: white; padding: 15px 40px; border: none; border-radius: 5px; font-weight: bold; cursor: pointer;">🖨️ DESCARGAR REPORTE GERENCIAL (PDF)</button>
            </div>

            <div class="page">
                <div class="header">
                    <img src="{logo}" style="height: 60px;">
                    <div style="text-align: right;">
                        <h1 class="title-main">Reporte Semanal Master</h1>
                        <p style="margin:0; font-weight: 700;">Semana {num_sem} | Año {ano_sel}</p>
                    </div>
                </div>

                <div class="section-box">
                    <div class="section-title">⏱️ AUDITORÍA DE TIEMPOS Y CIERRES (SEGURIDAD)</div>
                    <div class="kpi-grid">
                        <div class="kpi-card"><div>Apertura Promedio</div><div class="kpi-val">{hora_apertura}</div></div>
                        <div class="kpi-card"><div>Cierre Droguería</div><div class="kpi-val">{hora_cierre_dro}</div></div>
                        <div class="kpi-card"><div>Cierre Juanita</div><div class="kpi-val">{hora_juanita}</div></div>
                    </div>
                    <p style="font-size: 12px; margin-top: 15px; color: #444;"><i>* El cierre de droguería se calcula bajo el campo de auditoría estricta para medir promedios de salida reales de todos los departamentos.</i></p>
                </div>

                <div class="section-box">
                    <div class="section-title">🛠️ CRUCE DE MANTENIMIENTO PREVENTIVO</div>
                    <div class="kpi-grid">
                        <div class="kpi-card"><div>Unidades Planificadas</div><div class="kpi-val">{len(unidades_plan)}</div></div>
                        <div class="kpi-card"><div>Unidades Realizadas</div><div class="kpi-val">{len(unidades_real)}</div></div>
                        <div class="kpi-card"><div>% Cumplimiento</div><div class="kpi-val">{cumplimiento_mantenimiento:.1f}%</div></div>
                    </div>
                    <div style="margin-top:10px;">
                        <strong>⚠️ Unidades Pendientes:</strong> 
                        <span style="color: #d32f2f;">{', '.join(unidades_pendientes) if unidades_pendientes else 'Todas las unidades al día'}</span>
                    </div>
                </div>

                <div class="footer">Documento de Uso Interno - Droguería Drotaca - El Tigre, Venezuela</div>
            </div>

            <div class="page">
                <div class="header">
                    <div class="section-title" style="border:none;">📦 DESEMPEÑO LOGÍSTICO Y TRÁFICO</div>
                </div>

                <div class="kpi-grid">
                    <div class="kpi-card"><div>Total Bultos</div><div class="kpi-val">{bultos_sem:,.0f}</div></div>
                    <div class="kpi-card"><div>Recorrido (Kms)</div><div class="kpi-val">{kms_sem:,.0f}</div></div>
                    <div class="kpi-card"><div>Surtido (Litros)</div><div class="kpi-val">{litros_sem:,.0f}</div></div>
                </div>

                <div class="section-box" style="margin-top:20px;">
                    <div class="section-title">🚛 DETALLE DE TRÁFICO (Transbordo + Oriente)</div>
                    <table>
                        <thead>
                            <tr><th>Fecha</th><th>Unidad</th><th>Ruta</th><th>Estatus</th></tr>
                        </thead>
                        <tbody>
                            {"".join([f"<tr><td>{r.get('Fecha','')}</td><td>{r.get('Unidad','')}</td><td>{r.get('Ruta','')}</td><td>{r.get('Observaciones','')}</td></tr>" for _,r in f['Trafico'].head(15).iterrows()])}
                        </tbody>
                    </table>
                </div>

                <div class="section-box">
                    <div class="section-title">⛽ RESERVAS Y CONTROL DE COMBUSTIBLE</div>
                    <p style="font-size: 13px;">Reserva promedio detectada en El Tigre: <b>{pd.to_numeric(f['Reserva']['Nivel'], errors='coerce').mean() if not f['Reserva'].empty else 0:.1f}%</b></p>
                </div>

                <div class="footer">Gerencia de Operaciones y Logística - Reporte Generado Semanalmente</div>
            </div>
        </body>
        </html>
        """
        components.html(html_final, height=1200, scrolling=True)
        st.success("✅ Auditoría Generada. Desliza hacia abajo para ver el reporte multi-página.")
