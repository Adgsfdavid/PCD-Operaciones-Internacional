# ==========================================
# Archivo: cierre_semanal.py (Auditoría Logística - Mejorado)
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

def extraer_datos_sheets(nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        # ID ORIGINAL PRESERVADO
        doc = cliente.open_by_key("1L0N2O82bLzT1fE6kX5k-f5713tW3oN_YgOQzR5tY1R0")
        hoja = doc.worksheet(nombre_hoja)
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

def extraer_datos_historicos():
    try:
        cliente = obtener_cliente_sheets()
        # ID ORIGINAL PRESERVADO
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        hoja = doc.worksheet("MONITOREO_DESPACHOS")
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# FUNCIONES DE FORMATEO
# ==========================================
def f_p(valor):
    """Formatea números con separador de miles por punto (ej. 1.806)"""
    try:
        return f"{int(float(valor)):,.0f}".replace(",", ".")
    except:
        return str(valor)

# ==========================================
# INTERFAZ
# ==========================================
st.title("📊 Master Reporte Semanal: Desempeño de Tráfico")

c1, c2 = st.columns(2)
with c1: f_ini = st.date_input("Fecha Inicio:", datetime.now() - timedelta(days=7))
with c2: f_fin = st.date_input("Fecha Fin:", datetime.now())

if st.button("🚀 Generar Auditoría Semanal", type="primary", use_container_width=True):
    with st.spinner("Compilando datos semanales..."):
        # 1. Cargar y Filtrar Datos
        df_t = extraer_datos_sheets("CIERRE_DIARIO")
        df_h = extraer_datos_historicos()
        
        if df_t.empty or df_h.empty:
            st.error("No se encontraron datos en las bases de datos de Google Sheets.")
            st.stop()

        df_t['Fecha_DT'] = pd.to_datetime(df_t['Fecha_Reporte'], format='%d/%m/%Y', errors='coerce')
        df_h['Fecha_DT'] = pd.to_datetime(df_h['Fecha'], format='%d/%m/%Y', errors='coerce')
        
        mask_t = (df_t['Fecha_DT'].dt.date >= f_ini) & (df_t['Fecha_DT'].dt.date <= f_fin)
        mask_h = (df_h['Fecha_DT'].dt.date >= f_ini) & (df_h['Fecha_DT'].dt.date <= f_fin)
        
        f_t = df_t[mask_t].sort_values('Fecha_DT')
        f_h = df_h[mask_h]

        # 2. Cálculos Agregados
        total_f = f_h['CUBIERTOS'].astype(float).sum()
        total_b = f_h['BULTOS'].astype(float).sum()
        total_rutas = len(f_h)
        
        # Consolidado por Ruta
        resumen_r = f_h.groupby(['Region', 'DESPACHOS']).agg({
            'UNIDAD': 'last',
            'CUBIERTOS': 'sum',
            'BULTOS': 'sum'
        }).reset_index()

        # Consolidado por Zona
        resumen_z = f_h.groupby('Region').agg({
            'CUBIERTOS': 'sum',
            'BULTOS': 'sum'
        }).reset_index()

        # 3. Construcción de Filas HTML
        filas_t = "".join([f"<tr><td>{r['Fecha_Reporte']}</td><td>{r['Dia_Semana']}</td><td>{r['Hora_Primer_Listin']}</td><td>{r['Hora_Ultimo_Listin']}</td><td style='color:green; font-weight:bold;'>{r['Hora_Inicio_Trafico']}</td><td style='color:red; font-weight:bold;'>{r['Hora_Fin_Trafico']}</td></tr>" for _,r in f_t.iterrows()])
        
        filas_r = ""
        for _, r in resumen_r.iterrows():
            filas_r += f"<tr><td style='text-align:left;'>{r['DESPACHOS']}</td><td>{r['Region']}</td><td>{r['UNIDAD']}</td><td>{f_p(r['CUBIERTOS'])}</td><td>{f_p(r['BULTOS'])}</td></tr>"
            
        filas_z = ""
        for _, r in resumen_z.iterrows():
            p_f = (r['CUBIERTOS'] / total_f * 100) if total_f > 0 else 0
            p_b = (r['BULTOS'] / total_b * 100) if total_b > 0 else 0
            filas_z += f"<tr><td style='text-align:left; font-weight:bold;'>{r['Region']}</td><td>{f_p(r['CUBIERTOS'])}</td><td>{p_f:.1f}%</td><td>{f_p(r['BULTOS'])}</td><td>{p_b:.1f}%</td></tr>"

        # 4. PIZARRA HTML FINAL
        logo_b64 = obtener_logo_base64()
        area_logo = f'<img src="{logo_b64}" style="max-height:60px;">' if logo_b64 else ''
        
        html_final = f"""
        <html>
        <head>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <style>
                body {{ font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px; }}
                #pizarra {{ background: white; width: 1000px; margin: auto; border: 3px solid #000; padding: 20px; box-shadow: 0 0 15px rgba(0,0,0,0.2); }}
                .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid #ffca28; padding-bottom: 10px; margin-bottom: 20px; text-align: center; }}
                .section-title {{ background: #0d47a1; color: white; padding: 8px; font-weight: bold; margin-top: 20px; text-align: center; font-size: 14px; text-transform: uppercase; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 5px; font-size: 11px; }}
                th {{ background: #eeeeee; border: 1px solid #000; padding: 6px; }}
                td {{ border: 1px solid #ddd; padding: 5px; text-align: center; color: #333; }}
                .total-bar {{ background: #ffe082; padding: 10px; display: flex; justify-content: space-around; font-weight: bold; font-size: 16px; border: 2px solid #000; margin-top: 10px; }}
                .btn-download {{ background: #d32f2f; color: white; border: none; padding: 10px 20px; font-weight: bold; border-radius: 5px; cursor: pointer; margin-bottom: 10px; }}
            </style>
        </head>
        <body>
            <div style="text-align: right;">
                <button class="btn-download" onclick="descargarFoto()">📸 DESCARGAR PIZARRA (FOTO)</button>
            </div>
            
            <div id="pizarra">
                <div class="header">
                    <div style="width:20%">{area_logo}</div>
                    <div style="width:60%">
                        <div style="font-size:24px; font-weight:bold;">AUDITORÍA SEMANAL DE TRÁFICO</div>
                        <div style="font-size:16px; font-weight:bold; margin-top:5px; color:#555;">DEPARTAMENTO DE TRÁFICO</div>
                        <div style="font-size:14px; color:#666;">Periodo: {f_ini.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}</div>
                    </div>
                    <div style="width:20%; font-size:12px; font-weight:bold;">CONFIDENCIAL<br>DROTACA 2026</div>
                </div>

                <div class="section-title">⏱️ 1. CRONOMETRÍA DE SALIDAS (CONTROL DE TIEMPOS)</div>
                <table><thead><tr><th>FECHA</th><th>DÍA</th><th>1ER LISTÍN</th><th>ÚLT. LISTÍN</th><th>INICIO TRÁFICO</th><th>FIN TRÁFICO</th></tr></thead>
                <tbody>{filas_t}</tbody></table>

                <div class="section-title">🚛 2. CONSOLIDADO DE RUTAS (GESTIÓN SEMANAL)</div>
                <table><thead><tr><th style='text-align:left;'>RUTA</th><th>ZONA</th><th>UNIDAD (ÚLT.)</th><th>FARMACIAS TOT.</th><th>BULTOS TOT.</th></tr></thead>
                <tbody>{filas_r}</tbody></table>
                <div class="total-bar">
                    <span>TOTAL FARMACIAS: {f_p(total_f)}</span>
                    <span>TOTAL BULTOS: {f_p(total_b)}</span>
                </div>
                
                <div class="section-title">🌍 3. DISTRIBUCIÓN POR ZONAS LOGÍSTICAS</div>
                <table><thead><tr><th style='text-align:left;'>ZONA</th><th>FARMACIAS</th><th>% FAR.</th><th>BULTOS</th><th>% BUL.</th></tr></thead>
                <tbody>{filas_z}</tbody></table>

                <div style="margin-top:40px; border-top:2px solid #000; padding-top:10px; font-size:9px; color:#000; text-align:center; font-weight:bold;">
                    REPORTING SYSTEM PCD - DROGUERÍA DROTACA VENEZUELA
                </div>
            </div>

            <script>
                function descargarFoto() {{
                    html2canvas(document.getElementById('pizarra'), {{ scale: 2 }}).then(canvas => {{
                        var link = document.createElement('a');
                        link.download = 'Reporte_Semanal_Trafico_{f_fin.strftime('%d%m%y')}.png';
                        link.href = canvas.toDataURL();
                        link.click();
                    }});
                }}
            </script>
        </body>
        </html>
        """
        components.html(html_final, height=1400, scrolling=True)

        # 5. WHATSAPP SEMANAL
        st.markdown("---")
        st.subheader("📱 Mensaje para WhatsApp (Copiado Rápido)")
        
        txt_ws = f"*Reporte Semanal de Tráfico Drotaca* 🚚\n"
        txt_ws += f"📅 Periodo: {f_ini.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}\n\n"
        txt_ws += f"*ESTADÍSTICA SEMANAL:*\n"
        txt_ws += f"📍 Total Despachos: {total_rutas}\n"
        txt_ws += f"🏥 Farmacias Atendidas: {f_p(total_f)}\n"
        txt_ws += f"📦 Total Bultos Procesados: {f_p(total_b)}\n\n"
        txt_ws += f"*DISTRIBUCIÓN POR REGIÓN:*\n"
        
        for _, r in resumen_z.iterrows():
            p_b = (r['BULTOS'] / total_b * 100) if total_b > 0 else 0
            txt_ws += f"▪️ {r['Region']}: {f_p(r['BULTOS'])} Bultos ({p_b:.1f}%)\n"
            
        txt_ws += "\n✅ *Pizarra de auditoría adjunta.*"
        
        st.code(txt_ws, language="markdown")
