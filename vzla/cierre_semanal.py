# ==========================================
# Archivo: cierre_semanal.py (Master Semanal Multi-Reporte)
# ==========================================
import streamlit as st
import pandas as pd
import base64
import textwrap
import unicodedata
import re
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
# FUNCIONES DE FORMATEO Y BÚSQUEDA INTELIGENTE
# ==========================================
def f_p(valor):
    try: return f"{int(float(valor)):,.0f}".replace(",", ".")
    except: return str(valor)

def limpiar_hora(hora_str):
    if pd.isna(hora_str) or not hora_str: return "N/R"
    return str(hora_str).replace("*", "").strip()

def a_12h(hora_24):
    """Convierte cualquier formato de hora a un '05:30 PM' limpio para que el promedio no falle"""
    hora_limpia = limpiar_hora(hora_24)
    if hora_limpia == "N/R" or str(hora_limpia).lower() == "nan" or hora_limpia == "": return "N/R"
    try:
        match = re.match(r'(\d{1,2}:\d{2})\s*(AM|PM|am|pm)', hora_limpia)
        if match:
            time_part = match.group(1)
            ampm_part = match.group(2).upper()
            if len(time_part.split(':')[0]) == 1: time_part = "0" + time_part
            return f"{time_part} {ampm_part}"
        return datetime.strptime(hora_limpia, "%H:%M").strftime("%I:%M %p").upper()
    except: return str(hora_limpia).upper()

def buscar_columna_estricta(df, palabras_clave, evitar=None):
    if df.empty: return None
    evitar = evitar or []
    for col in df.columns:
        col_clean = str(col).lower().strip()
        if any(p in col_clean for p in palabras_clave) and not any(e in col_clean for e in evitar):
            return col
    return None

def norm_dia(d):
    if pd.isna(d): return ""
    d_str = str(d).strip().capitalize()
    return unicodedata.normalize('NFKD', d_str).encode('ASCII', 'ignore').decode('utf-8')

def calcular_promedio_horas(lista_horas):
    formato = "%I:%M %p"
    tiempos = []
    for h in lista_horas:
        try:
            h_clean = a_12h(h)
            if h_clean != "N/R":
                tiempos.append(datetime.strptime(h_clean, formato))
        except: continue
    
    if not tiempos: return "N/R"
    
    segundos_totales = sum(t.hour * 3600 + t.minute * 60 for t in tiempos) / len(tiempos)
    horas = int(segundos_totales // 3600)
    minutos = int((segundos_totales % 3600) // 60)
    
    temp_dt = datetime(2026, 1, 1, horas, minutos)
    return temp_dt.strftime("%I:%M %p").upper()

def calcular_rango_semana(ano, semana):
    lunes = datetime.strptime(f'{int(ano)}-W{int(semana)}-1', "%G-W%V-%u")
    domingo = lunes + timedelta(days=6)
    return f"{lunes.strftime('%d/%m/%Y')} al {domingo.strftime('%d/%m/%Y')}"

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.set_page_config(page_title="Master Semanal Drotaca", layout="wide")
st.title("📚 Master Reporte Semanal de Operaciones")

c1, c2 = st.columns(2)
with c1:
    ano_sel = st.selectbox("Año Fiscal:", [2025, 2026], index=1)
with c2:
    semana_actual = datetime.now().isocalendar()[1]
    num_sem = st.number_input("Número de Semana a Auditar:", 1, 53, value=semana_actual)

rango_fechas = calcular_rango_semana(ano_sel, num_sem)
t_trafico, t_cierres = st.tabs(["📈 Desempeño de Tráfico", "⏱️ Cronometría de Cierres"])

# ---------------------------------------------------------
# PESTAÑA 1: DESEMPEÑO DE TRÁFICO (Se mantiene igual)
# ---------------------------------------------------------
with t_trafico:
    st.info("Consolida la data de despachos diarios en una pizarra semanal.")
    if st.button("🚀 Generar Auditoría de Tráfico", type="primary", use_container_width=True):
        with st.spinner("Consolidando rutas de tráfico..."):
            df_raw = extraer_datos("PIZARRA_TRAFICO")
            if df_raw.empty:
                st.error("Error al conectar con la base de datos o la hoja está vacía.")
            else:
                c_sem = buscar_columna_estricta(df_raw, ['semana'])
                if not c_sem: st.error("No se encontró columna Semana en Tráfico."); st.stop()
                
                df_raw['Num_Semana'] = df_raw[c_sem].astype(str).str.extract(r'(\d+)').astype(float)
                df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()

                if df_sem.empty:
                    st.warning(f"No hay registros de tráfico para la Semana {int(num_sem)}.")
                else:
                    c_fecha = buscar_columna_estricta(df_sem, ['fecha'])
                    c_dia = buscar_columna_estricta(df_sem, ['dia', 'día'], evitar=['fecha'])
                    c_h1 = buscar_columna_estricta(df_sem, ['1er'])
                    c_hu = buscar_columna_estricta(df_sem, ['ultimo', 'último'])
                    c_it = buscar_columna_estricta(df_sem, ['inicio'])
                    c_ct = buscar_columna_estricta(df_sem, ['culminacion', 'fin'])
                    c_ruta = buscar_columna_estricta(df_sem, ['ruta'])
                    c_zona = buscar_columna_estricta(df_sem, ['zona'])
                    c_unidad = buscar_columna_estricta(df_sem, ['unidad'])
                    c_farma = buscar_columna_estricta(df_sem, ['farmacia']) 
                    c_bultos = buscar_columna_estricta(df_sem, ['bulto'])

                    df_t = df_sem.drop_duplicates(subset=[c_fecha]).copy()
                    df_t['Fecha_Temp'] = pd.to_datetime(df_t[c_fecha], format='%d/%m/%Y', errors='coerce')
                    df_t = df_t.sort_values('Fecha_Temp')

                    df_sem[c_farma] = pd.to_numeric(df_sem[c_farma], errors='coerce').fillna(0)
                    df_sem[c_bultos] = pd.to_numeric(df_sem[c_bultos], errors='coerce').fillna(0)

                    df_rutas = df_sem.groupby([c_ruta, c_zona], as_index=False).agg({
                        c_unidad: 'last', c_farma: 'sum', c_bultos: 'sum'
                    }).sort_values(by=[c_zona, c_ruta])
                    
                    total_f = df_rutas[c_farma].sum()
                    total_b = df_rutas[c_bultos].sum()

                    df_zonas = df_rutas.groupby(c_zona).agg({c_farma: 'sum', c_bultos: 'sum'}).reset_index()
                    df_zonas['%_Far'] = (df_zonas[c_farma] / total_f * 100).round(1).fillna(0)
                    df_zonas['%_Bul'] = (df_zonas[c_bultos] / total_b * 100).round(1).fillna(0)

                    logo = obtener_logo_base64()
                    color_azul = "#0d47a1"
                    color_dorado = "#d4af37"

                    filas_t = "".join([f"<tr><td>{r[c_fecha]}</td><td>{r[c_dia]}</td><td>{a_12h(r[c_h1])}</td><td>{a_12h(r[c_hu])}</td><td>{a_12h(r[c_it])}</td><td>{a_12h(r[c_ct])}</td></tr>" for _,r in df_t.iterrows()])
                    filas_r = "".join([f"<tr><td style='text-align:left;'>{r[c_ruta]}</td><td>{r[c_zona]}</td><td>{r[c_unidad]}</td><td style='font-weight:bold;'>{f_p(r[c_farma])}</td><td style='font-weight:bold;'>{f_p(r[c_bultos])}</td></tr>" for _,r in df_rutas.iterrows()])
                    filas_z = "".join([f"<tr><td style='text-align:left; font-weight:bold;'>{r[c_zona]}</td><td>{f_p(r[c_farma])}</td><td style='color:{color_azul}; font-weight:bold;'>{r['%_Far']}%</td><td>{f_p(r[c_bultos])}</td><td style='color:#e65100; font-weight:bold;'>{r['%_Bul']}%</td></tr>" for _,r in df_zonas.iterrows()])

                    html_pdf = f"""
                    <!DOCTYPE html><html><head>
                        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                        <style>
                        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                        body {{ font-family: 'Montserrat', sans-serif; background:#525659; margin:0; padding-bottom: 20px; }}
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
                    </style></head><body>
                        <div style="text-align:center; padding:20px; display: flex; justify-content: center; gap: 15px;">
                            <button onclick="descargarFoto()" style="background:#d32f2f; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px;">📸 DESCARGAR PIZARRA (FOTO)</button>
                        </div>
                        <div class="page" id="pizarra-trafico">
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
                            </div>
                        </div>
                        <script>
                            function descargarFoto() {{
                                html2canvas(document.getElementById('pizarra-trafico'), {{ scale: 2 }}).then(canvas => {{
                                    var link = document.createElement('a');
                                    link.download = 'Trafico_Semana_{int(num_sem)}.png';
                                    link.href = canvas.toDataURL();
                                    link.click();
                                }});
                            }}
                        </script>
                    </body></html>
                    """
                    components.html(html_pdf, height=1200, scrolling=True)

                    st.markdown("---")
                    st.subheader("📱 Mensaje para WhatsApp (Tráfico)")
                    txt_ws = f"*Reporte Semanal de Tráfico Drotaca* 🚚\n📅 Semana: {int(num_sem)} ({rango_fechas})\n\n"
                    txt_ws += f"*RESUMEN OPERATIVO:*\n📍 Total Despachos: {len(df_rutas)}\n🏥 Farmacias Atendidas: {f_p(total_f)}\n📦 Total Bultos Procesados: {f_p(total_b)}\n\n"
                    txt_ws += f"*PESO LOGÍSTICO POR ZONA:*\n"
                    for _, r in df_zonas.iterrows(): txt_ws += f"▪️ {r[c_zona]}: {f_p(r[c_bultos])} Bultos ({r['%_Bul']}%)\n"
                    txt_ws += "\n✅ *Pizarra de auditoría adjunta.*"
                    st.code(txt_ws, language="markdown")

# ---------------------------------------------------------
# PESTAÑA 2: CRONOMETRÍA DE CIERRES (DROTACA 2.0)
# ---------------------------------------------------------
with t_cierres:
    st.info("Análisis de Apertura y Cierres Drotaca 2.0 (Lunes a Viernes).")
    if st.button("🕒 Procesar Cronometría de Cierres", type="primary", use_container_width=True):
        with st.spinner("Extrayendo y estructurando bases de datos..."):
            
            df_a_raw = extraer_datos("SEG_APERTURA")
            df_j_raw = extraer_datos("SEG_CIERRE_JUANITA")
            df_d_raw = extraer_datos("SEG_CIERRE_DROTACA")
            
            if df_d_raw.empty:
                st.error("No se pudo acceder a la hoja principal SEG_CIERRE_DROTACA.")
            else:
                # ESTRUCTURA LUNES A VIERNES
                dias_base = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"]
                dias_display = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
                
                df_resumen = pd.DataFrame({"Dia_Norm": dias_base, "Día": dias_display})
                df_resumen['Apertura'] = "N/R"
                df_resumen['Juanita'] = "N/R"
                df_resumen['Drotaca'] = "N/R"
                df_resumen['Fecha'] = ""

                # 1. APERTURA
                if not df_a_raw.empty:
                    c_sem_a = buscar_columna_estricta(df_a_raw, ['semana'])
                    if c_sem_a:
                        df_a_raw['Num_Semana'] = df_a_raw[c_sem_a].astype(str).str.extract(r'(\d+)').astype(float)
                        f_a = df_a_raw[df_a_raw['Num_Semana'] == num_sem].copy()
                        c_dia_a = buscar_columna_estricta(f_a, ['dia', 'día'], evitar=['fecha'])
                        c_hora_a = buscar_columna_estricta(f_a, ['apertura', 'hora'], evitar=['fecha', 'dia'])
                        if c_dia_a and c_hora_a:
                            f_a['Dia_Norm'] = f_a[c_dia_a].apply(norm_dia)
                            dict_a = f_a.groupby('Dia_Norm')[c_hora_a].last().to_dict()
                            df_resumen['Apertura'] = df_resumen['Dia_Norm'].map(dict_a).fillna("N/R")

                # 2. JUANITA
                if not df_j_raw.empty:
                    c_sem_j = buscar_columna_estricta(df_j_raw, ['semana'])
                    if c_sem_j:
                        df_j_raw['Num_Semana'] = df_j_raw[c_sem_j].astype(str).str.extract(r'(\d+)').astype(float)
                        f_j = df_j_raw[df_j_raw['Num_Semana'] == num_sem].copy()
                        c_dia_j = buscar_columna_estricta(f_j, ['dia', 'día'], evitar=['fecha'])
                        c_hora_j = buscar_columna_estricta(f_j, ['juanita', 'hora', 'cierre'], evitar=['fecha', 'dia'])
                        if c_dia_j and c_hora_j:
                            f_j['Dia_Norm'] = f_j[c_dia_j].apply(norm_dia)
                            dict_j = f_j.groupby('Dia_Norm')[c_hora_j].last().to_dict()
                            df_resumen['Juanita'] = df_resumen['Dia_Norm'].map(dict_j).fillna("N/R")

                # 3. DROTACA (Extracción en Fila y Matriz)
                pivot_deps = pd.DataFrame()
                if not df_d_raw.empty:
                    c_sem_d = buscar_columna_estricta(df_d_raw, ['semana'])
                    if c_sem_d:
                        df_d_raw['Num_Semana'] = df_d_raw[c_sem_d].astype(str).str.extract(r'(\d+)').astype(float)
                        f_d = df_d_raw[df_d_raw['Num_Semana'] == num_sem].copy()
                        
                        if not f_d.empty:
                            c_dia_d = buscar_columna_estricta(f_d, ['dia', 'día'], evitar=['fecha'])
                            c_fecha_d = buscar_columna_estricta(f_d, ['fecha'])
                            c_dep = buscar_columna_estricta(f_d, ['departamento', 'area', 'área'], evitar=['fecha', 'hora'])
                            c_hora_sal = buscar_columna_estricta(f_d, ['hora salida', 'salida', 'hora', 'cierre'], evitar=['fecha'])
                            
                            f_d['Dia_Norm'] = f_d[c_dia_d].apply(norm_dia) if c_dia_d else ""
                            
                            if c_dep and c_hora_sal and c_dia_d:
                                # A. Extraer el Cierre de Droguería (Fila específica)
                                filtro_cierre_gen = f_d[c_dep].astype(str).str.upper().str.contains("CIERRE DE DROG|CIERRE GENERAL|CIERRE DROTACA", na=False)
                                df_cierre_gral = f_d[filtro_cierre_gen]
                                
                                if not df_cierre_gral.empty:
                                    dict_d = df_cierre_gral.groupby('Dia_Norm')[c_hora_sal].last().to_dict()
                                    df_resumen['Drotaca'] = df_resumen['Dia_Norm'].map(dict_d).fillna("N/R")
                                
                                if c_fecha_d:
                                    dict_f = f_d.dropna(subset=[c_fecha_d]).groupby('Dia_Norm')[c_fecha_d].last().to_dict()
                                    df_resumen['Fecha'] = df_resumen['Dia_Norm'].map(dict_f).fillna("")

                                # B. Extraer Departamentos (Resto de las Filas)
                                df_deps = f_d[~filtro_cierre_gen].dropna(subset=[c_dep, c_hora_sal]).copy()
                                df_deps = df_deps[df_deps[c_dep].str.strip() != ""]
                                df_deps = df_deps[df_deps[c_dep].astype(str).str.lower() != "nan"]
                                
                                if not df_deps.empty:
                                    pivot_deps = df_deps.pivot_table(index=c_dep, columns='Dia_Norm', values=c_hora_sal, aggfunc='last')
                                    for d in dias_base:
                                        if d not in pivot_deps.columns: pivot_deps[d] = "N/R"
                                    pivot_deps = pivot_deps[dias_base].fillna("N/R")
                                    pivot_deps['Promedio'] = pivot_deps.apply(lambda row: calcular_promedio_horas(row.tolist()), axis=1)

                # ==========================================
                # HTML RENDERING - PIZARRA GENERAL
                # ==========================================
                prom_juanita = calcular_promedio_horas(df_resumen['Juanita'].tolist())
                prom_drotaca = calcular_promedio_horas(df_resumen['Drotaca'].tolist())

                logo_b64 = obtener_logo_base64()
                color_azul = "#0d47a1"
                
                filas_gral_html = ""
                for _, r in df_resumen.iterrows():
                    hora_ap = a_12h(r['Apertura'])
                    hora_ju = a_12h(r['Juanita'])
                    hora_dr = a_12h(r['Drotaca'])
                    
                    color_ap = "#2e7d32" if "06:" in hora_ap or "07:00" in hora_ap else "#000"
                    color_ju = "#e65100" if hora_ju != "N/R" else "#777"
                    
                    filas_gral_html += f"""
                    <tr style="text-align: center; border-bottom: 1px solid #ddd;">
                        <td style="padding: 15px; font-weight: bold; background-color: #f8f9fa;">{r['Día'].upper()}<br><small style="color:#666;">{r['Fecha']}</small></td>
                        <td style="padding: 15px; color: {color_ap}; font-weight: bold; font-size: 16px;">{hora_ap}</td>
                        <td style="padding: 15px; color: {color_ju}; font-weight: bold; font-size: 16px;">{hora_ju}</td>
                        <td style="padding: 15px; font-weight: 900; font-size: 16px;">{hora_dr}</td>
                    </tr>
                    """

                html_pizarra_general = f"""
                <html><head>
                    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                    <style>
                        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                        body {{ font-family: 'Montserrat', sans-serif; padding: 20px; background-color: #f0f2f6; }}
                        .pizarra {{ background: white; width: 900px; margin: auto; border-radius: 15px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: 2px solid {color_azul}; margin-bottom: 40px; }}
                        .header {{ background: {color_azul}; color: white; padding: 30px; display: flex; justify-content: space-between; align-items: center; }}
                        table {{ width: 100%; border-collapse: collapse; }}
                        th {{ background: #f1f4f9; color: {color_azul}; padding: 15px; text-transform: uppercase; font-size: 12px; border-bottom: 2px solid {color_azul}; }}
                        .footer-promedios {{ background: #f1f4f9; padding: 20px; display: flex; justify-content: space-around; border-top: 2px solid {color_azul}; }}
                        .promedio-box {{ text-align: center; }}
                        .promedio-label {{ font-size: 12px; font-weight: bold; color: #555; text-transform: uppercase; }}
                        .promedio-val {{ font-size: 20px; font-weight: 900; color: {color_azul}; }}
                    </style>
                </head><body>
                    <div style="text-align: center; margin-bottom: 15px;">
                        <button onclick="capturar('pizarra-general', 'Cierres_Generales_Semana_{int(num_sem)}.png')" style="background: #2e7d32; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR REPORTE GENERAL</button>
                    </div>
                    <div class="pizarra" id="pizarra-general">
                        <div class="header">
                            <img src="{logo_b64}" style="height: 50px;">
                            <div style="text-align: right;">
                                <div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">REPORTE SEMANAL DE GESTIÓN</div>
                                <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} | CIERRES GENERALES</div>
                            </div>
                        </div>
                        <table>
                            <thead><tr><th>DÍA</th><th>APERTURA DROTACA</th><th>CIERRE JUANITA</th><th>CIERRE DROTACA</th></tr></thead>
                            <tbody>{filas_gral_html}</tbody>
                        </table>
                        <div class="footer-promedios">
                            <div class="promedio-box">
                                <div class="promedio-label">📊 Promedio Cierre Drotaca</div>
                                <div class="promedio-val" style="color: #2e7d32;">{prom_drotaca}</div>
                            </div>
                            <div class="promedio-box">
                                <div class="promedio-label">📊 Promedio Cierre Juanita</div>
                                <div class="promedio-val" style="color: #e65100;">{prom_juanita}</div>
                            </div>
                        </div>
                    </div>
                """

                # ==========================================
                # HTML RENDERING - PIZARRA DEPARTAMENTOS
                # ==========================================
                html_pizarra_deps = ""
                if not pivot_deps.empty:
                    filas_deps_html = ""
                    for dep_nombre, row in pivot_deps.iterrows():
                        tds = ""
                        for dia in dias_base:
                            val = a_12h(row[dia])
                            tds += f"<td style='padding: 10px; border: 1px solid #ddd; font-weight: bold; font-size: 13px;'>{val}</td>"
                        
                        prom = row['Promedio']
                        filas_deps_html += f"""
                        <tr style="text-align: center;">
                            <td style="padding: 10px; border: 1px solid #ddd; font-weight: 900; background-color: #f8f9fa; text-align: left; color: {color_azul}; font-size: 12px;">{str(dep_nombre).upper()}</td>
                            {tds}
                            <td style="padding: 10px; border: 1px solid #ddd; font-weight: 900; font-size: 14px; color: #d32f2f; background-color: #ffebee;">{prom}</td>
                        </tr>
                        """

                    html_pizarra_deps = f"""
                        <div style="text-align: center; margin-bottom: 15px; margin-top: 30px;">
                            <button onclick="capturar('pizarra-departamentos', 'Cierres_Departamentos_Semana_{int(num_sem)}.png')" style="background: #e65100; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR DEPARTAMENTOS</button>
                        </div>
                        <div class="pizarra" id="pizarra-departamentos" style="width: 1050px;">
                            <div class="header" style="background: #1565c0;">
                                <img src="{logo_b64}" style="height: 50px;">
                                <div style="text-align: right;">
                                    <div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">MATRIZ DE DEPARTAMENTOS</div>
                                    <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} | HORARIOS DE SALIDA</div>
                                </div>
                            </div>
                            <table>
                                <thead>
                                    <tr>
                                        <th style="text-align: left;">DEPARTAMENTO</th>
                                        <th>LUNES</th>
                                        <th>MARTES</th>
                                        <th>MIÉRCOLES</th>
                                        <th>JUEVES</th>
                                        <th>VIERNES</th>
                                        <th style="background: #ffcdd2; color: #b71c1c;">PROMEDIO</th>
                                    </tr>
                                </thead>
                                <tbody>{filas_deps_html}</tbody>
                            </table>
                            <div style="padding: 15px; text-align: center; font-size: 11px; color: #666; font-weight: bold; background: #f1f4f9;">
                                REPORTE GENERADO POR CONTROL TOWER LOGÍSTICA - DROTACA VENEZUELA
                            </div>
                        </div>
                        <script>
                            function capturar(id, filename) {{
                                html2canvas(document.getElementById(id), {{ scale: 2 }}).then(canvas => {{
                                    var link = document.createElement('a');
                                    link.download = filename;
                                    link.href = canvas.toDataURL();
                                    link.click();
                                }});
                            }}
                        </script>
                    </body></html>
                    """

                # Renderizamos ambas pizarras
                components.html(html_pizarra_general + html_pizarra_deps, height=1600, scrolling=True)

                # --- WHATSAPP CIERRES ---
                st.markdown("---")
                st.subheader("📱 Resumen para WhatsApp (Cierres y Departamentos)")
                msg_w = f"⏱️ *Reporte de Cierres Semanal - Drotaca 2.0*\n📅 Semana: {int(num_sem)}\n\n"
                msg_w += f"📍 *Cronometría de la Droguería:*\n"
                msg_w += f"🔹 Promedio Cierre General: *{prom_drotaca}*\n"
                msg_w += f"🔹 Promedio Cierre Juanita: *{prom_juanita}*\n\n"
                
                if not pivot_deps.empty:
                    msg_w += f"📍 *Top 3 Áreas que salieron más tarde (Promedio):*\n"
                    # Calculamos el top 3 para el Whatsapp (ordenando las horas de manera simple para el texto)
                    try:
                        pivot_deps['Para_Ordenar'] = pd.to_datetime(pivot_deps['Promedio'], format="%I:%M %p", errors='coerce')
                        top_3 = pivot_deps.sort_values(by='Para_Ordenar', ascending=False).head(3)
                        for dep, row in top_3.iterrows():
                            msg_w += f"🔹 {str(dep).title()}: *{row['Promedio']}*\n"
                    except:
                        pass
                        
                msg_w += f"\n✅ Tablas de auditoría detalladas adjuntas en imagen."
                st.code(msg_w, language="markdown")
