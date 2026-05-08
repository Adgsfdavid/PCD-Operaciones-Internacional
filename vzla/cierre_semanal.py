# ==========================================
# Archivo: cierre_semanal.py (Auditoría Logística - Diseño Premium)
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
# FUNCIONES DE FORMATEO
# ==========================================
def limpiar_hora(hora_str):
    if not hora_str: return ""
    return str(hora_str).replace("*", "").strip()

def a_12h(hora_24):
    hora_limpia = limpiar_hora(hora_24)
    try:
        if "am" in hora_limpia.lower() or "pm" in hora_limpia.lower():
            return hora_limpia.lower()
        return datetime.strptime(hora_limpia, "%H:%M").strftime("%I:%M %p").lower()
    except:
        return hora_limpia

def buscar_columna(df, palabras_clave):
    if df.empty: return None
    for col in df.columns:
        if any(p.lower() in str(col).lower() for p in palabras_clave):
            return col
    return None

# ==========================================
# INTERFAZ
# ==========================================
st.set_page_config(page_title="Master Semanal Drotaca", layout="wide")
st.title("📊 Master Reporte Semanal: Desempeño de Tráfico")

c1, c2 = st.columns(2)
with c1:
    ano_sel = st.selectbox("Año Fiscal:", [2025, 2026], index=1)
with c2:
    semana_actual = datetime.now().isocalendar()[1]
    num_sem = st.number_input("Número de Semana:", 1, 53, value=semana_actual)

if st.button("⚡ GENERAR AUDITORÍA DE TRÁFICO", type="primary", use_container_width=True):
    with st.spinner("Sincronizando datos de Tráfico..."):
        
        df_raw = extraer_datos("PIZARRA_TRAFICO")
        if df_raw.empty:
            st.error("Error al conectar con la base de datos.")
            st.stop()

        df_raw['Num_Semana'] = df_raw['Semana'].astype(str).str.extract(r'(\d+)').astype(float)
        df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()

        if df_sem.empty:
            st.warning(f"No hay registros para la Semana {num_sem}.")
            st.stop()

        # --- CRONOMETRÍA ---
        c_fecha = buscar_columna(df_sem, ['fecha'])
        c_dia = buscar_columna(df_sem, ['dia', 'día'])
        c_h1 = buscar_columna(df_sem, ['1er'])
        c_hu = buscar_columna(df_sem, ['ultimo', 'último'])
        c_it = buscar_columna(df_sem, ['inicio'])
        c_ct = buscar_columna(df_sem, ['culminacion', 'fin'])

        df_t = df_sem.drop_duplicates(subset=[c_fecha]).copy()
        
        # --- CONSOLIDADO DE RUTAS ---
        c_ruta, c_zona, c_unidad, c_farma, c_bultos = buscar_columna(df_sem, ['ruta']), buscar_columna(df_sem, ['zona']), buscar_columna(df_sem, ['unidad']), buscar_columna(df_sem, ['farmacias']), buscar_columna(df_sem, ['bultos'])
        df_sem[c_farma] = pd.to_numeric(df_sem[c_farma], errors='coerce').fillna(0)
        df_sem[c_bultos] = pd.to_numeric(df_sem[c_bultos], errors='coerce').fillna(0)

        df_rutas = df_sem.groupby([c_ruta, c_zona, c_unidad], as_index=False).agg({c_farma: 'sum', c_bultos: 'sum'}).sort_values(by=[c_zona, c_ruta])
        total_f, total_b = df_rutas[c_farma].sum(), df_rutas[c_bultos].sum()

        # --- DISTRIBUCIÓN POR ZONAS ---
        df_zonas = df_rutas.groupby(c_zona).agg({c_farma: 'sum', c_bultos: 'sum'}).reset_index()
        df_zonas['%_Far'] = (df_zonas[c_farma] / total_f * 100).round(1).fillna(0)
        df_zonas['%_Bul'] = (df_zonas[c_bultos] / total_b * 100).round(1).fillna(0)

        # ==========================================
        # CONSTRUCCIÓN DEL PDF CON DISEÑO AZUL Y DORADO
        # ==========================================
        logo = obtener_logo_base64()
        color_azul = "#0d47a1"  # Azul Drotaca
        color_dorado = "#d4af37" # Dorado Premium

        filas_t = "".join([f"<tr><td>{r[c_fecha]}</td><td>{r[c_dia]}</td><td>{a_12h(r[c_h1])}</td><td>{a_12h(r[c_hu])}</td><td>{a_12h(r[c_it])}</td><td>{a_12h(r[c_ct])}</td></tr>" for _,r in df_t.iterrows()])
        filas_r = "".join([f"<tr><td style='text-align:left;'>{r[c_ruta]}</td><td>{r[c_zona]}</td><td>{r[c_unidad]}</td><td style='font-weight:bold;'>{int(r[c_farma])}</td><td style='font-weight:bold;'>{int(r[c_bultos])}</td></tr>" for _,r in df_rutas.iterrows()])
        filas_z = "".join([f"<tr><td style='text-align:left; font-weight:bold;'>{r[c_zona]}</td><td>{int(r[c_farma])}</td><td style='color:{color_azul}; font-weight:bold;'>{r['%_Far']}%</td><td>{int(r[c_bultos])}</td><td style='color:#e65100; font-weight:bold;'>{r['%_Bul']}%</td></tr>" for _,r in df_zonas.iterrows()])

        html_pdf = f"""
        <!DOCTYPE html><html><head><style>
            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
            body {{ font-family: 'Montserrat', sans-serif; background:#525659; margin:0; }}
            .page {{ width: 210mm; background: white; margin: 10mm auto; padding: 0; box-shadow: 0 0 10px rgba(0,0,0,0.5); overflow:hidden; }}
            
            /* ENCABEZADO AZUL Y DORADO */
            .header-master {{ background: {color_azul}; color: white; padding: 25px 40px; display: flex; justify-content: space-between; align-items: center; border-bottom: 6px solid {color_dorado}; }}
            .header-info {{ text-align: right; }}
            .header-info h2 {{ margin: 0; font-weight: 900; letter-spacing: 1px; font-size: 20px; }}
            .header-info p {{ margin: 0; font-size: 14px; opacity: 0.9; }}
            
            .content-padding {{ padding: 15mm; }}
            .section-title {{ border-left: 5px solid {color_dorado}; background: #f4f4f4; color: {color_azul}; padding: 8px 15px; font-weight: 900; font-size: 13px; margin-top: 20px; text-transform: uppercase; }}
            
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10px; }}
            th {{ background: {color_azul}; color: white; border: 1px solid #ddd; padding: 6px; text-transform: uppercase; }}
            td {{ border: 1px solid #ddd; padding: 6px; text-align: center; color: #333; }}
            
            .total-bar {{ background: #263238; color: white; display: flex; justify-content: space-around; padding: 10px; margin-top: 10px; font-weight: 900; font-size: 13px; border-bottom: 3px solid {color_dorado}; }}
            @media print {{ .no-print {{ display: none; }} body {{ background: white; }} .page {{ margin: 0; box-shadow: none; }} }}
        </style></head><body>
            <div class="no-print" style="text-align:center; padding:20px;">
                <button onclick="window.print()" style="background:#e65100; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px; box-shadow: 0 4px 6px rgba(0,0,0,0.2);">🖨️ IMPRIMIR REPORTE MASTER</button>
            </div>
            
            <div class="page">
                <div class="header-master">
                    <img src="{logo}" style="height: 55px; filter: drop-shadow(2px 2px 4px rgba(0,0,0,0.3));">
                    <div class="header-info">
                        <h2>AUDITORÍA SEMANAL DE TRÁFICO</h2>
                        <p>Semana {num_sem} | Año {ano_sel}</p>
                    </div>
                </div>
                
                <div class="content-padding">
                    <div class="section-title">⏱️ 1. CRONOMETRÍA DE SALIDAS (CONTROL DE TIEMPOS)</div>
                    <table><thead><tr><th>FECHA</th><th>DÍA</th><th>1ER LISTÍN</th><th>ÚLT. LISTÍN</th><th>INICIO TRÁFICO</th><th>FIN TRÁFICO</th></tr></thead>
                    <tbody>{filas_t}</tbody></table>

                    <div class="section-title">🚛 2. CONSOLIDADO DE RUTAS (GESTIÓN SEMANAL)</div>
                    <table><thead><tr><th style='text-align:left;'>RUTA</th><th>ZONA</th><th>UNIDAD</th><th>FARMACIAS</th><th>BULTOS</th></tr></thead>
                    <tbody>{filas_r}</tbody></table>
                    <div class="total-bar"><span>TOTAL FARMACIAS: {int(total_f)}</span><span>TOTAL BULTOS: {int(total_b)}</span></div>
                    
                    <div class="section-title">🌍 3. DISTRIBUCIÓN POR ZONAS LOGÍSTICAS</div>
                    <table><thead><tr><th style='text-align:left;'>ZONA</th><th>FARMACIAS</th><th>% FAR.</th><th>BULTOS</th><th>% BUL.</th></tr></thead>
                    <tbody>{filas_z}</tbody></table>

                    <div style="margin-top:40px; border-top:1px solid #eee; padding-top:10px; font-size:9px; color:#aaa; text-align:center; font-weight:bold;">
                        REPORTING SYSTEM PCD - DROGUERÍA DROTACA VENEZUELA
                    </div>
                </div>
            </div></body></html>
        """
        components.html(html_pdf, height=1200, scrolling=True)
