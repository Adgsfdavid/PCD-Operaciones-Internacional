# ==========================================
# Archivo: cierre_semanal.py (Auditoría Logística - Orden de Fecha Corregido)
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
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        hoja = doc.worksheet(nombre_hoja)
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# FUNCIONES DE FORMATEO Y CÁLCULO
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

def calcular_rango_semana(ano, semana):
    """Calcula el lunes y domingo de una semana ISO"""
    lunes = datetime.strptime(f'{int(ano)}-W{int(semana)}-1', "%G-W%V-%u")
    domingo = lunes + timedelta(days=6)
    return f"{lunes.strftime('%d/%m/%Y')} al {domingo.strftime('%d/%m/%Y')}"

# FORMATO DE MILES CON PUNTO
def f_p(valor):
    try:
        return f"{int(float(valor)):,.0f}".replace(",", ".")
    except:
        return str(valor)

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
    with st.spinner("Calculando periodos y consolidando rutas..."):
        
        df_raw = extraer_datos("PIZARRA_TRAFICO")
        if df_raw.empty:
            st.error("Error al conectar con la base de datos o la hoja está vacía.")
            st.stop()

        # Calculamos rango de fechas para el reporte
        rango_fechas = calcular_rango_semana(ano_sel, num_sem)

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
        
        # SOLUCIÓN: ORDENAR CRONOLÓGICAMENTE LAS FECHAS
        df_t['Fecha_Temp'] = pd.to_datetime(df_t[c_fecha], format='%d/%m/%Y', errors='coerce')
        df_t = df_t.sort_values('Fecha_Temp')
        
        # --- CONSOLIDADO DE RUTAS ---
        c_ruta, c_zona, c_unidad, c_farma, c_bultos = buscar_columna(df_sem, ['ruta']), buscar_columna(df_sem, ['zona']), buscar_columna(df_sem, ['unidad']), buscar_columna(df_sem, ['farmacias']), buscar_columna(df_sem, ['bultos'])
        df_sem[c_farma] = pd.to_numeric(df_sem[c_farma], errors='coerce').fillna(0)
        df_sem[c_bultos] = pd.to_numeric(df_sem[c_bultos], errors='coerce').fillna(0)

        # Agrupamos por Ruta para que aparezcan una sola vez con la última placa
        df_rutas = df_sem.groupby([c_ruta, c_zona], as_index=False).agg({
            c_unidad: 'last',
            c_farma: 'sum',
            c_bultos: 'sum'
        }).sort_values(by=[c_zona, c_ruta])
        
        total_f, total_b = df_rutas[c_farma].sum(), df_rutas[c_bultos].sum()

        # --- DISTRIBUCIÓN POR ZONAS ---
        df_zonas = df_rutas.groupby(c_zona).agg({c_farma: 'sum', c_bultos: 'sum'}).reset_index()
        df_zonas['%_Far'] = (df_zonas[c_farma] / total_f * 100).round(1).fillna(0)
        df_zonas['%_Bul'] = (df_zonas[c_bultos] / total_b * 100).round(1).fillna(0)

        # ==========================================
        # CONSTRUCCIÓN DEL PDF
        # ==========================================
        logo = obtener_logo_base64()
        color_azul = "#0d47a1"
        color_dorado = "#d4af37"

        # APLICANDO EL FORMATO DE PUNTOS f_p()
        filas_t = "".join([f"<tr><td>{r[c_fecha]}</td><td>{r[c_dia]}</td><td>{a_12h(r[c_h1])}</td><td>{a_12h(r[c_hu])}</td><td>{a_12h(r[c_it])}</td><td>{a_12h(r[c_ct])}</td></tr>" for _,r in df_t.iterrows()])
        filas_r = "".join([f"<tr><td style='text-align:left;'>{r[c_ruta]}</td><td>{r[c_zona]}</td><td>{r[c_unidad]}</td><td style='font-weight:bold;'>{f_p(r[c_farma])}</td><td style='font-weight:bold;'>{f_p(r[c_bultos])}</td></tr>" for _,r in df_rutas.iterrows()])
        filas_z = "".join([f"<tr><td style='text-align:left; font-weight:bold;'>{r[c_zona]}</td><td>{f_p(r[c_farma])}</td><td style='color:{color_azul}; font-weight:bold;'>{r['%_Far']}%</td><td>{f_p(r[c_bultos])}</td><td style='color:#e65100; font-weight:bold;'>{r['%_Bul']}%</td></tr>" for _,r in df_zonas.iterrows()])

        html_pdf = f"""
        <!DOCTYPE html><html><head>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <style>
            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
            body {{ font-family: 'Montserrat', sans-serif; background:#525659; margin:0; }}
            .page {{ width: 210mm; background: white; margin: 10mm auto; padding: 0; box-shadow: 0 0 10px rgba(0,0,0,0.5); overflow:hidden; border: 1px solid #000; }}
            
            .header-master {{ background: {color_azul}; color: white; padding: 25px 40px; display: flex; justify-content: space-between; align-items: center; border-bottom: 6px solid {color_dorado}; }}
            .header-info {{ text-align: right; }}
            .header-info h2 {{ margin: 0; font-weight: 900; font-size: 20px; text-transform: uppercase; }}
            .header-info p {{ margin: 0; font-size: 14px; font-weight: bold; color: {color_dorado}; }}
            
            .content-padding {{ padding: 12mm; }}
            .section-title {{ border-left: 6px solid {color_dorado}; background: #eee; color: #000; padding: 8px 15px; font-weight: 900; font-size: 13px; margin-top: 20px; border: 1px solid #000; }}
            
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10px; border: 1px solid #000; }}
            th {{ background: {color_azul}; color: white; border: 1px solid #000; padding: 8px; text-transform: uppercase; }}
            td {{ border: 1px solid #000; padding: 6px; text-align: center; color: #000; }}
            
            .total-bar {{ background: #000; color: white; display: flex; justify-content: space-around; padding: 12px; margin-top: 10px; font-weight: 900; font-size: 13px; border: 1px solid {color_dorado}; }}
            @media print {{ .no-print {{ display: none; }} body {{ background: white; }} .page {{ margin: 0; box-shadow: none; }} }}
        </style></head><body>
            <div class="no-print" style="text-align:center; padding:20px; display: flex; justify-content: center; gap: 15px;">
                <button onclick="window.print()" style="background:#e65100; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px;">🖨️ IMPRIMIR REPORTE MASTER</button>
                <button onclick="descargarFoto()" style="background:#d32f2f; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px;">📸 DESCARGAR PIZARRA (FOTO)</button>
            </div>
            
            <div class="page">
                <div class="header-master">
                    <img src="{logo}" style="height: 55px;">
                    <div class="header-info">
                        <h2>AUDITORÍA SEMANAL DE TRÁFICO</h2>
                        <h3 style="margin: 5px 0 0 0; color: #fff; font-size: 14px;">DEPARTAMENTO DE TRÁFICO</h3>
                        <p>Semana {int(num_sem)} ({rango_fechas})</p>
                    </div>
                </div>
                
                <div class="content-padding">
                    <div class="section-title">⏱️ 1. CRONOMETRÍA DE SALIDAS (CONTROL DE TIEMPOS)</div>
                    <table><thead><tr><th>FECHA</th><th>DÍA</th><th>1ER LISTÍN</th><th>ÚLT. LISTÍN</th><th>INICIO TRÁFICO</th><th>FIN TRÁFICO</th></tr></thead>
                    <tbody>{filas_t}</tbody></table>

                    <div class="section-title">🚛 2. CONSOLIDADO DE RUTAS (GESTIÓN SEMANAL)</div>
                    <table><thead><tr><th style='text-align:left;'>RUTA</th><th>ZONA</th><th>UNIDAD (ÚLT.)</th><th>FARMACIAS TOT.</th><th>BULTOS TOT.</th></tr></thead>
                    <tbody>{filas_r}</tbody></table>
                    <div class="total-bar"><span>TOTAL FARMACIAS: {f_p(total_f)}</span><span>TOTAL BULTOS: {f_p(total_b)}</span></div>
                    
                    <div class="section-title">🌍 3. DISTRIBUCIÓN POR ZONAS LOGÍSTICAS</div>
                    <table><thead><tr><th style='text-align:left;'>ZONA</th><th>FARMACIAS</th><th>% FAR.</th><th>BULTOS</th><th>% BUL.</th></tr></thead>
                    <tbody>{filas_z}</tbody></table>

                    <div style="margin-top:40px; border-top:2px solid #000; padding-top:10px; font-size:9px; color:#000; text-align:center; font-weight:bold;">
                        REPORTING SYSTEM PCD - DROGUERÍA DROTACA VENEZUELA
                    </div>
                </div>
            </div>
            <script>
                function descargarFoto() {{
                    html2canvas(document.querySelector('.page'), {{ scale: 2 }}).then(canvas => {{
                        var link = document.createElement('a');
                        link.download = 'Reporte_Semanal_Semana_{int(num_sem)}.png';
                        link.href = canvas.toDataURL();
                        link.click();
                    }});
                }}
            </script>
        </body></html>
        """
        components.html(html_pdf, height=1200, scrolling=True)

        # ==========================================
        # WHATSAPP SEMANAL REDACTADO
        # ==========================================
        st.markdown("---")
        st.subheader("📱 Mensaje para WhatsApp (Copiado Rápido)")
        
        txt_ws = f"*Reporte Semanal de Tráfico Drotaca* 🚚\n"
        txt_ws += f"📅 Semana: {int(num_sem)} ({rango_fechas})\n\n"
        txt_ws += f"*RESUMEN OPERATIVO:*\n"
        txt_ws += f"📍 Total Despachos: {len(df_rutas)}\n"
        txt_ws += f"🏥 Farmacias Atendidas: {f_p(total_f)}\n"
        txt_ws += f"📦 Total Bultos Procesados: {f_p(total_b)}\n\n"
        txt_ws += f"*PESO LOGÍSTICO POR ZONA:*\n"
        for _, r in df_zonas.iterrows():
            txt_ws += f"▪️ {r[c_zona]}: {f_p(r[c_bultos])} Bultos ({r['%_Bul']}%)\n"
        txt_ws += "\n✅ *Pizarra de auditoría adjunta.*"
        
        st.code(txt_ws, language="markdown")
