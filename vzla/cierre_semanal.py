# ==========================================
# Archivo: cierre_semanal.py (Auditoría Logística)
# ==========================================
import streamlit as st
import pandas as pd
import base64
import textwrap
import re
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
# FUNCIONES DE FORMATEO Y LIMPIEZA
# ==========================================
def limpiar_hora(hora_str):
    """Limpia asteriscos y espacios de las horas de WhatsApp"""
    if not hora_str: return ""
    return str(hora_str).replace("*", "").strip()

def a_12h(hora_24):
    """Convierte HH:MM a HH:MM AM/PM"""
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
    with st.spinner("Analizando Pizarras..."):
        
        df_raw = extraer_datos("PIZARRA_TRAFICO")
        if df_raw.empty:
            st.error("Error al conectar con la base de datos.")
            st.stop()

        # Filtrado por semana
        df_raw['Num_Semana'] = df_raw['Semana'].astype(str).str.extract(r'(\d+)').astype(float)
        df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()

        if df_sem.empty:
            st.warning(f"No hay registros para la Semana {num_sem}.")
            st.stop()

        # --- SECCIÓN 1: CRONOMETRÍA (12 HORAS) ---
        c_fecha = buscar_columna(df_sem, ['fecha'])
        c_dia = buscar_columna(df_sem, ['dia', 'día'])
        c_h1 = buscar_columna(df_sem, ['1er'])
        c_hu = buscar_columna(df_sem, ['ultimo', 'último'])
        c_it = buscar_columna(df_sem, ['inicio'])
        c_ct = buscar_columna(df_sem, ['culminacion', 'fin'])

        df_t = df_sem.drop_duplicates(subset=[c_fecha]).copy()
        
        # Aplicamos el formato 12h para la tabla
        for col in [c_h1, c_hu, c_it, c_ct]:
            if col:
                df_t[col + "_12h"] = df_t[col].apply(a_12h)

        # --- SECCIÓN 2: CONSOLIDADO DE RUTAS ---
        c_ruta = buscar_columna(df_sem, ['ruta'])
        c_zona = buscar_columna(df_sem, ['zona'])
        c_unidad = buscar_columna(df_sem, ['unidad'])
        c_farma = buscar_columna(df_sem, ['farmacias'])
        c_bultos = buscar_columna(df_sem, ['bultos'])
        
        df_sem[c_farma] = pd.to_numeric(df_sem[c_farma], errors='coerce').fillna(0)
        df_sem[c_bultos] = pd.to_numeric(df_sem[c_bultos], errors='coerce').fillna(0)

        df_rutas = df_sem.groupby([c_ruta, c_zona, c_unidad], as_index=False).agg({c_farma: 'sum', c_bultos: 'sum'}).sort_values(by=[c_zona, c_ruta])
        
        total_f = df_rutas[c_farma].sum()
        total_b = df_rutas[c_bultos].sum()

        # --- SECCIÓN 3: EFECTIVIDAD POR ZONAS ---
        if not df_rutas.empty:
            df_zonas = df_rutas.groupby(c_zona).agg({c_farma: 'sum', c_bultos: 'sum'}).reset_index()
            df_zonas['%_Farmacias'] = (df_zonas[c_farma] / total_f * 100).round(1).fillna(0)
            df_zonas['%_Bultos'] = (df_zonas[c_bultos] / total_b * 100).round(1).fillna(0)
        else:
            df_zonas = pd.DataFrame()

        # ==========================================
        # CONSTRUCCIÓN DEL PDF
        # ==========================================
        logo = obtener_logo_base64()
        color_p = "#0d47a1"

        filas_t = ""
        if not df_t.empty:
            for _, r in df_t.iterrows():
                filas_t += f"<tr><td>{r[c_fecha]}</td><td>{r[c_dia]}</td><td>{r.get(c_h1+'_12h', '')}</td><td>{r.get(c_hu+'_12h', '')}</td><td>{r.get(c_it+'_12h', '')}</td><td>{r.get(c_ct+'_12h', '')}</td></tr>"

        filas_r = ""
        if not df_rutas.empty:
            for _, r in df_rutas.iterrows():
                filas_r += f"<tr><td style='text-align:left;'>{r[c_ruta]}</td><td>{r[c_zona]}</td><td>{r[c_unidad]}</td><td>{int(r[c_farma])}</td><td>{int(r[c_bultos])}</td></tr>"

        filas_z = ""
        if not df_zonas.empty:
            for _, r in df_zonas.iterrows():
                filas_z += f"<tr style='background:#f1f8ff;'><td style='text-align:left; font-weight:bold;'>{r[c_zona]}</td><td>{int(r[c_farma])}</td><td style='color:{color_p}; font-weight:bold;'>{r['%_Farmacias']}%</td><td>{int(r[c_bultos])}</td><td style='color:#e65100; font-weight:bold;'>{r['%_Bultos']}%</td></tr>"

        html_pdf = f"""
        <!DOCTYPE html><html><head><style>
            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
            body {{ font-family: 'Montserrat', sans-serif; background:#525659; margin:0; }}
            .page {{ width: 210mm; background: white; margin: 10mm auto; padding: 15mm; box-shadow: 0 0 10px rgba(0,0,0,0.5); }}
            .header {{ display: flex; justify-content: space-between; border-bottom: 4px solid {color_p}; padding-bottom: 5px; }}
            .section-title {{ background: {color_p}; color: white; padding: 6px 12px; font-weight: 900; font-size: 13px; margin-top: 20px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 10px; }}
            th {{ background: #eee; border: 1px solid #ccc; padding: 5px; }}
            td {{ border: 1px solid #ccc; padding: 5px; text-align: center; }}
            .total-bar {{ background: #263238; color: white; display: flex; justify-content: space-around; padding: 8px; margin-top: 5px; font-weight: 900; font-size: 12px; }}
            @media print {{ .no-print {{ display: none; }} body {{ background: white; }} .page {{ margin: 0; box-shadow: none; }} }}
        </style></head><body>
            <div class="no-print" style="text-align:center; padding:20px;"><button onclick="window.print()" style="background:#e65100; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px;">📥 DESCARGAR REPORTE</button></div>
            <div class="page">
                <div class="header"><img src="{logo}" style="height: 45px;"><div><h2 style="margin:0; color:{color_p};">AUDITORÍA DE TRÁFICO SEMANAL</h2><p style="margin:0; text-align:right;">Semana {num_sem} | {ano_sel}</p></div></div>
                
                <div class="section-title">⏱️ 1. CRONOMETRÍA DE SALIDAS (FORMATO 12H)</div>
                <table><thead><tr><th>Fecha</th><th>Día</th><th>1er Listín</th><th>Último Listín</th><th>Inicio Tráfico</th><th>Fin Tráfico</th></tr></thead>
                <tbody>{filas_t}</tbody></table>

                <div class="section-title">🚛 2. CONSOLIDADO DE RUTAS ATENDIDAS</div>
                <table><thead><tr><th style='text-align:left;'>Ruta</th><th>Zona</th><th>Unidad</th><th>Farmacias Totales</th><th>Bultos Totales</th></tr></thead>
                <tbody>{filas_r}</tbody></table>
                <div class="total-bar"><span>TOTAL FARMACIAS: {int(total_f)}</span><span>TOTAL BULTOS: {int(total_b)}</span></div>
                
                <div class="section-title">🌍 3. EFECTIVIDAD Y DISTRIBUCIÓN POR ZONAS</div>
                <table><thead><tr><th style='text-align:left;'>Zona Logística</th><th>Farmacias</th><th>% Far.</th><th>Bultos</th><th>% Bul.</th></tr></thead>
                <tbody>{filas_z}</tbody></table>

                <div style="margin-top:30px; border-top:1px solid #eee; padding-top:10px; font-size:9px; color:#aaa; text-align:center;">Página 1: Tráfico Semanal - Drotaca</div>
            </div></body></html>
        """
        components.html(html_pdf, height=1200, scrolling=True)
