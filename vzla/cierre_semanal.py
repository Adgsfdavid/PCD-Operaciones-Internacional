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

def calcular_rango_lunes_viernes(ano, semana):
    lunes = datetime.strptime(f'{int(ano)}-W{int(semana)}-1', "%G-W%V-%u")
    viernes = lunes + timedelta(days=4)
    return f"{lunes.strftime('%d/%m/%Y')} al {viernes.strftime('%d/%m/%Y')}"

def extraer_fecha_limpia(fecha_str):
    if pd.isna(fecha_str): return pd.NaT
    match = re.search(r'(\d{2}/\d{2}/\d{4})', str(fecha_str))
    if match:
        try: return datetime.strptime(match.group(1), "%d/%m/%Y")
        except: return pd.NaT
    return pd.NaT

def filtrar_ultima_carga(df, num_sem):
    """Filtro inteligente para hojas que solo tienen Timestamp en vez de Semana"""
    if df.empty: return df
    c_sem = buscar_columna_estricta(df, ['semana'])
    if c_sem:
        df['Num_Semana'] = df[c_sem].astype(str).str.extract(r'(\d+)').astype(float)
        df_sem = df[df['Num_Semana'] == num_sem].copy()
        if not df_sem.empty: return df_sem
        
    c_f = buscar_columna_estricta(df, ['fecha', 'timestamp'])
    if not c_f: return pd.DataFrame()
    
    df['F_DT'] = pd.to_datetime(df[c_f], dayfirst=True, errors='coerce')
    df['Sem_Calc'] = df['F_DT'].dt.isocalendar().week
    df_sem = df[df['Sem_Calc'] == num_sem].copy()
    
    if df_sem.empty: return df_sem
    max_dt = df_sem['F_DT'].max()
    return df_sem[df_sem['F_DT'] == max_dt].copy()

def agrupar_rol_compacto(df_seg, num_sem):
    """Comprime el Rol de Guardia expandido en bloques visuales compactos"""
    c_fecha = buscar_columna_estricta(df_seg, ['fecha'])
    c_area = buscar_columna_estricta(df_seg, ['area', 'área'])
    c_diu = buscar_columna_estricta(df_seg, ['diurno'])
    c_noc = buscar_columna_estricta(df_seg, ['nocturno'])
    c_cant = buscar_columna_estricta(df_seg, ['cant'])
    
    if not c_fecha or not c_area: return []
    
    df_seg['DT'] = df_seg[c_fecha].apply(extraer_fecha_limpia)
    df_sem = df_seg[df_seg['DT'].dt.isocalendar().week == num_sem].copy()
    df_sem = df_sem.dropna(subset=['DT']).sort_values('DT')
    if df_sem.empty: return []
    
    grupos = []
    df_sat = df_sem[df_sem['DT'].dt.weekday == 5]
    if not df_sat.empty:
        fecha_str = f"SÁBADO {df_sat['DT'].iloc[0].strftime('%d/%m/%Y')}"
        grupos.append((fecha_str, df_sat.drop_duplicates(subset=[c_area, c_diu, c_noc])))
        
    df_sun = df_sem[df_sem['DT'].dt.weekday == 6]
    if not df_sun.empty:
        fecha_str = f"DOMINGO {df_sun['DT'].iloc[0].strftime('%d/%m/%Y')}"
        grupos.append((fecha_str, df_sun.drop_duplicates(subset=[c_area, c_diu, c_noc])))
        
    df_lv = df_sem[df_sem['DT'].dt.weekday < 5]
    if not df_lv.empty:
        min_dt = df_lv['DT'].min()
        max_dt = df_lv['DT'].max()
        if min_dt != max_dt:
            fecha_str = f"LUNES {min_dt.strftime('%d/%m/%Y')} AL VIERNES {max_dt.strftime('%d/%m/%Y')}"
        else:
            dias_es = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"]
            fecha_str = f"{dias_es[min_dt.weekday()]} {min_dt.strftime('%d/%m/%Y')}"
        grupos.append((fecha_str, df_lv.drop_duplicates(subset=[c_area, c_diu, c_noc])))
        
    return grupos, c_area, c_diu, c_noc, c_cant

def asignar_subregion(ruta, macro_default):
    r = str(ruta).upper()
    if any(x in r for x in ["ANACO", "CANTAURA", "CARUPANO", "GUIRIA", "CUMANA", "MATURIN", "PUNTA DE MATA", "ESPARTA", "ARAGUA DE BARCELONA", "CARIPITO", "EL TIGRE"]): return "ORIENTE NORTE"
    if any(x in r for x in ["BARCELONA", "CLARINES", "BOLIVAR", "DELTA", "TUMEREMO", "GUARICO", "PARIAGUAN", "ORDAZ", "FELIX", "UPATA", "PIAR", "PARAGUA", "VALLE LA PASCUA", "ZARAZA", "MERCEDES"]): return "ORIENTE SUR"
    
    if any(x in r for x in ["CARACAS", "ARAGUA", "SAN JUAN"]): return "CENTRO"
    if any(x in r for x in ["CARABOBO", "COJEDES"]): return "CENTRO OCCIDENTE"
    
    if any(x in r for x in ["LARA 1", "LARA 01", "PORTUGUESA 1", "PORTUGUESA 01", "LARA 2", "LARA 02", "YARACUY"]): return "OCCIDENTE SUR"
    if any(x in r for x in ["PORTUGUESA 2", "PORTUGUESA 02", "BARINAS", "APURE"]): return "LOS LLANOS"
    if any(x in r for x in ["CORO", "PUNTO FIJO", "MARACAIBO", "CABIMAS", "OJEDA"]): return "OCCIDENTE NORTE"
    if any(x in r for x in ["MERIDA", "TRUJILLO", "PORTUGUESA 3", "PORTUGUESA 03", "TACHIRA"]): return "LOS ANDES / TACHIRA"
    
    if str(macro_default).upper() == "ORIENTE": return "ORIENTE SUR"
    if str(macro_default).upper() == "CENTRO": return "CENTRO"
    if str(macro_default).upper() == "OCCIDENTE": return "OCCIDENTE NORTE"
    return "OTRAS REGIONES"

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
rango_fechas_lv = calcular_rango_lunes_viernes(ano_sel, num_sem)
color_azul = "#0d47a1"
logo_b64 = obtener_logo_base64()

t_trafico, t_cierres, t_comensales, t_guardias, t_despachos, t_combustible, t_surtido = st.tabs([
    "📈 Desempeño de Tráfico", 
    "⏱️ Cronometría de Cierres", 
    "🍽️ Pizarra Comensales",
    "🛡️ Guardias Semanales",
    "🚚 Resumen de Despachos",
    "⛽ Reserva Combustible",
    "⛽ Resultados de Surtido"
])

# (Las Pestañas 1, 2, 3 y 4 se mantienen exactamente iguales...)
with t_trafico:
    st.info("Consolida la data de despachos diarios en una pizarra semanal de rutas.")
    if st.button("🚀 Generar Auditoría de Tráfico", type="primary", use_container_width=True):
        with st.spinner("Consolidando rutas de tráfico..."):
            df_raw = extraer_datos("PIZARRA_TRAFICO")
            if df_raw.empty: st.error("Error al conectar con la base de datos o la hoja está vacía.")
            else:
                c_sem = buscar_columna_estricta(df_raw, ['semana'])
                if not c_sem: st.error("No se encontró columna Semana en Tráfico."); st.stop()
                df_raw['Num_Semana'] = df_raw[c_sem].astype(str).str.extract(r'(\d+)').astype(float)
                df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()
                if df_sem.empty: st.warning(f"No hay registros de tráfico para la Semana {int(num_sem)}.")
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
                    df_rutas = df_sem.groupby([c_ruta, c_zona], as_index=False).agg({c_unidad: 'last', c_farma: 'sum', c_bultos: 'sum'}).sort_values(by=[c_zona, c_ruta])
                    total_f = df_rutas[c_farma].sum()
                    total_b = df_rutas[c_bultos].sum()
                    df_zonas = df_rutas.groupby(c_zona).agg({c_farma: 'sum', c_bultos: 'sum'}).reset_index()
                    df_zonas['%_Far'] = (df_zonas[c_farma] / total_f * 100).round(1).fillna(0)
                    df_zonas['%_Bul'] = (df_zonas[c_bultos] / total_b * 100).round(1).fillna(0)
                    color_dorado = "#d4af37"
                    filas_t = "".join([f"<tr><td>{r[c_fecha]}</td><td>{r[c_dia]}</td><td>{a_12h(r[c_h1])}</td><td>{a_12h(r[c_hu])}</td><td>{a_12h(r[c_it])}</td><td>{a_12h(r[c_ct])}</td></tr>" for _,r in df_t.iterrows()])
                    filas_r = "".join([f"<tr><td style='text-align:left;'>{r[c_ruta]}</td><td>{r[c_zona]}</td><td>{r[c_unidad]}</td><td style='font-weight:bold;'>{f_p(r[c_farma])}</td><td style='font-weight:bold;'>{f_p(r[c_bultos])}</td></tr>" for _,r in df_rutas.iterrows()])
                    filas_z = "".join([f"<tr><td style='text-align:left; font-weight:bold;'>{r[c_zona]}</td><td>{f_p(r[c_farma])}</td><td style='color:{color_azul}; font-weight:bold;'>{r['%_Far']}%</td><td>{f_p(r[c_bultos])}</td><td style='color:#e65100; font-weight:bold;'>{r['%_Bul']}%</td></tr>" for _,r in df_zonas.iterrows()])
                    html_pdf = f"""<!DOCTYPE html><html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><style>@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');body {{ font-family: 'Montserrat', sans-serif; background:#525659; margin:0; padding-bottom: 20px; }} .page {{ width: 210mm; background: white; margin: 10mm auto; padding: 0; box-shadow: 0 0 10px rgba(0,0,0,0.5); overflow:hidden; border: 1px solid #000; }} .header-master {{ background: {color_azul}; color: white; padding: 25px 40px; display: flex; justify-content: space-between; align-items: center; border-bottom: 6px solid {color_dorado}; }} .header-info {{ text-align: right; }} .header-info h2 {{ margin: 0; font-weight: 900; font-size: 20px; text-transform: uppercase; }} .header-info p {{ margin: 0; font-size: 14px; font-weight: bold; color: {color_dorado}; }} .content-padding {{ padding: 12mm; }} .section-title {{ border-left: 6px solid {color_dorado}; background: #eee; color: #000; padding: 8px 15px; font-weight: 900; font-size: 13px; margin-top: 20px; border: 1px solid #000; }} table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10px; border: 1px solid #000; }} th {{ background: {color_azul}; color: white; border: 1px solid #000; padding: 8px; text-transform: uppercase; }} td {{ border: 1px solid #000; padding: 6px; text-align: center; color: #000; }} .total-bar {{ background: #000; color: white; display: flex; justify-content: space-around; padding: 12px; margin-top: 10px; font-weight: 900; font-size: 13px; border: 1px solid {color_dorado}; }}</style></head><body><div style="text-align:center; padding:20px; display: flex; justify-content: center; gap: 15px;"><button onclick="descargarFoto()" style="background:#d32f2f; color:white; border:none; padding:12px 30px; font-weight:bold; cursor:pointer; border-radius:5px;">📸 DESCARGAR PIZARRA (FOTO)</button></div><div class="page" id="pizarra-trafico"><div class="header-master"><img src="{logo_b64}" style="height: 55px;"><div class="header-info"><h2>AUDITORÍA SEMANAL DE TRÁFICO</h2><h3 style="margin: 5px 0 0 0; color: #fff; font-size: 14px;">DEPARTAMENTO DE TRÁFICO</h3><p>Semana {int(num_sem)} ({rango_fechas})</p></div></div><div class="content-padding"><div class="section-title">⏱️ 1. CRONOMETRÍA DE SALIDAS (CONTROL DE TIEMPOS)</div><table><thead><tr><th>FECHA</th><th>DÍA</th><th>1ER LISTÍN</th><th>ÚLT. LISTÍN</th><th>INICIO TRÁFICO</th><th>FIN TRÁFICO</th></tr></thead><tbody>{filas_t}</tbody></table><div class="section-title">🚛 2. CONSOLIDADO DE RUTAS (GESTIÓN SEMANAL)</div><table><thead><tr><th style='text-align:left;'>RUTA</th><th>ZONA</th><th>UNIDAD (ÚLT.)</th><th>FARMACIAS TOT.</th><th>BULTOS TOT.</th></tr></thead><tbody>{filas_r}</tbody></table><div class="total-bar"><span>TOTAL FARMACIAS: {f_p(total_f)}</span><span>TOTAL BULTOS: {f_p(total_b)}</span></div><div class="section-title">🌍 3. DISTRIBUCIÓN POR ZONAS LOGÍSTICAS</div><table><thead><tr><th style='text-align:left;'>ZONA</th><th>FARMACIAS</th><th>% FAR.</th><th>BULTOS</th><th>% BUL.</th></tr></thead><tbody>{filas_z}</tbody></table></div></div><script>function descargarFoto() {{ html2canvas(document.getElementById('pizarra-trafico'), {{ scale: 2 }}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Trafico_Semana_{int(num_sem)}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script></body></html>"""
                    components.html(html_pdf, height=1200, scrolling=True)
                    st.markdown("---")
                    st.subheader("📱 Mensaje para WhatsApp (Tráfico)")
                    txt_ws = f"*Reporte Semanal de Tráfico Drotaca* 🚚\n📅 Semana: {int(num_sem)} ({rango_fechas})\n\n*RESUMEN OPERATIVO:*\n📍 Total Despachos: {len(df_rutas)}\n🏥 Farmacias Atendidas: {f_p(total_f)}\n📦 Total Bultos Procesados: {f_p(total_b)}\n\n*PESO LOGÍSTICO POR ZONA:*\n"
                    for _, r in df_zonas.iterrows(): txt_ws += f"▪️ {r[c_zona]}: {f_p(r[c_bultos])} Bultos ({r['%_Bul']}%)\n"
                    txt_ws += "\n✅ *Pizarra de auditoría adjunta.*"
                    st.code(txt_ws, language="markdown")

with t_cierres:
    st.info("Análisis de Apertura y Cierres Drotaca 2.0 (Lunes a Viernes).")
    if st.button("🕒 Procesar Cronometría de Cierres", type="primary", use_container_width=True):
        with st.spinner("Extrayendo y estructurando bases de datos..."):
            df_a_raw = extraer_datos("SEG_APERTURA")
            df_j_raw = extraer_datos("SEG_CIERRE_JUANITA")
            df_d_raw = extraer_datos("SEG_CIERRE_DROTACA")
            if df_d_raw.empty: st.error("No se pudo acceder a la hoja principal SEG_CIERRE_DROTACA.")
            else:
                dias_base = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"]
                dias_display = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
                df_resumen = pd.DataFrame({"Dia_Norm": dias_base, "Día": dias_display})
                df_resumen['Apertura'] = "N/R"; df_resumen['Juanita'] = "N/R"; df_resumen['Drotaca'] = "N/R"; df_resumen['Fecha'] = ""
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
                                filtro_cierre_gen = f_d[c_dep].astype(str).str.upper().str.contains(r"CIERRE DE DROG|CIERRE GENERAL|CIERRE DROTACA", na=False)
                                df_cierre_gral = f_d[filtro_cierre_gen]
                                if not df_cierre_gral.empty:
                                    dict_d = df_cierre_gral.groupby('Dia_Norm')[c_hora_sal].last().to_dict()
                                    df_resumen['Drotaca'] = df_resumen['Dia_Norm'].map(dict_d).fillna("N/R")
                                if c_fecha_d:
                                    dict_f = f_d.dropna(subset=[c_fecha_d]).groupby('Dia_Norm')[c_fecha_d].last().to_dict()
                                    df_resumen['Fecha'] = df_resumen['Dia_Norm'].map(dict_f).fillna("")
                                df_deps = f_d[~filtro_cierre_gen].dropna(subset=[c_dep, c_hora_sal]).copy()
                                df_deps = df_deps[df_deps[c_dep].str.strip() != ""]; df_deps = df_deps[df_deps[c_dep].astype(str).str.lower() != "nan"]
                                if not df_deps.empty:
                                    pivot_deps = df_deps.pivot_table(index=c_dep, columns='Dia_Norm', values=c_hora_sal, aggfunc='last')
                                    for d in dias_base:
                                        if d not in pivot_deps.columns: pivot_deps[d] = "N/R"
                                    pivot_deps = pivot_deps[dias_base].fillna("N/R")
                                    pivot_deps['Promedio'] = pivot_deps.apply(lambda row: calcular_promedio_horas(row.tolist()), axis=1)
                prom_juanita = calcular_promedio_horas(df_resumen['Juanita'].tolist())
                prom_drotaca = calcular_promedio_horas(df_resumen['Drotaca'].tolist())
                filas_gral_html = ""
                for _, r in df_resumen.iterrows():
                    hora_ap = a_12h(r['Apertura']); hora_ju = a_12h(r['Juanita']); hora_dr = a_12h(r['Drotaca'])
                    color_ap = "#2e7d32" if "06:" in hora_ap or "07:00" in hora_ap else "#000"
                    color_ju = "#e65100" if hora_ju != "N/R" else "#777"
                    filas_gral_html += f"""<tr style="text-align: center;"><td style="padding: 15px; font-weight: bold; background-color: #f8f9fa; border: 1px solid #000;">{r['Día'].upper()}<br><small style="color:#666;">{r['Fecha']}</small></td><td style="padding: 15px; color: {color_ap}; font-weight: bold; font-size: 16px; border: 1px solid #000;">{hora_ap}</td><td style="padding: 15px; color: {color_ju}; font-weight: bold; font-size: 16px; border: 1px solid #000;">{hora_ju}</td><td style="padding: 15px; font-weight: 900; font-size: 16px; border: 1px solid #000;">{hora_dr}</td></tr>"""
                html_pizarra_general = f"""<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><style>@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');body {{ font-family: 'Montserrat', sans-serif; padding: 20px; background-color: #f0f2f6; }} .pizarra {{ background: white; width: 900px; margin: auto; border-radius: 15px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: 2px solid {color_azul}; margin-bottom: 40px; }} .header {{ background: {color_azul}; color: white; padding: 30px; display: flex; justify-content: space-between; align-items: center; }} table {{ width: 100%; border-collapse: collapse; }} th {{ background: #f1f4f9; color: {color_azul}; padding: 15px; text-transform: uppercase; font-size: 12px; border: 1px solid #000; }} .footer-promedios {{ background: #f1f4f9; padding: 20px; display: flex; justify-content: space-around; border-top: 2px solid {color_azul}; }} .promedio-box {{ text-align: center; }} .promedio-label {{ font-size: 12px; font-weight: bold; color: #555; text-transform: uppercase; }} .promedio-val {{ font-size: 20px; font-weight: 900; color: {color_azul}; }}</style></head><body><div style="text-align: center; margin-bottom: 15px;"><button onclick="capturar('pizarra-general', 'Cierres_Generales_Semana_{int(num_sem)}.png')" style="background: #2e7d32; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR REPORTE GENERAL</button></div><div class="pizarra" id="pizarra-general"><div class="header"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">REPORTE SEMANAL DE GESTIÓN</div><div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} ({rango_fechas_lv}) | CIERRES GENERALES</div></div></div><table><thead><tr><th>DÍA</th><th>APERTURA DROTACA</th><th>CIERRE JUANITA</th><th>CIERRE DROTACA</th></tr></thead><tbody>{filas_gral_html}</tbody></table><div class="footer-promedios"><div class="promedio-box"><div class="promedio-label">📊 Promedio Cierre Drotaca</div><div class="promedio-val" style="color: #2e7d32;">{prom_drotaca}</div></div><div class="promedio-box"><div class="promedio-label">📊 Promedio Cierre Juanita</div><div class="promedio-val" style="color: #e65100;">{prom_juanita}</div></div></div></div>"""
                html_pizarra_deps = ""
                if not pivot_deps.empty:
                    filas_deps_html = ""
                    for dep_nombre, row in pivot_deps.iterrows():
                        tds = "".join([f"<td style='padding: 10px; border: 1px solid #000; font-weight: bold; font-size: 13px;'>{a_12h(row[dia])}</td>" for dia in dias_base])
                        prom = row['Promedio']
                        filas_deps_html += f"""<tr style="text-align: center;"><td style="padding: 10px; border: 1px solid #000; font-weight: 900; background-color: #f8f9fa; text-align: left; color: {color_azul}; font-size: 12px;">{str(dep_nombre).upper()}</td>{tds}<td style="padding: 10px; border: 1px solid #000; font-weight: 900; font-size: 14px; color: #1b5e20; background-color: #e8f5e9;">{prom}</td></tr>"""
                    html_pizarra_deps = f"""<div style="text-align: center; margin-bottom: 15px; margin-top: 30px;"><button onclick="capturar('pizarra-departamentos', 'Cierres_Departamentos_Semana_{int(num_sem)}.png')" style="background: #e65100; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR DEPARTAMENTOS</button></div><div class="pizarra" id="pizarra-departamentos" style="width: 1050px;"><div class="header" style="background: #1565c0;"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">MATRIZ DE DEPARTAMENTOS</div><div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} ({rango_fechas_lv}) | HORARIOS DE SALIDA</div></div></div><table><thead><tr><th style="text-align: left; border: 1px solid #000;">DEPARTAMENTO</th><th style="border: 1px solid #000;">LUNES</th><th style="border: 1px solid #000;">MARTES</th><th style="border: 1px solid #000;">MIÉRCOLES</th><th style="border: 1px solid #000;">JUEVES</th><th style="border: 1px solid #000;">VIERNES</th><th style="background: #c8e6c9; color: #1b5e20; border: 1px solid #000;">PROMEDIO</th></tr></thead><tbody>{filas_deps_html}</tbody></table><div style="padding: 15px; text-align: center; font-size: 11px; color: #666; font-weight: bold; background: #f1f4f9;">REPORTE GENERADO POR CONTROL TOWER LOGÍSTICA - DROTACA VENEZUELA</div></div><script>function capturar(id, filename) {{ html2canvas(document.getElementById(id), {{ scale: 2 }}).then(canvas => {{ var link = document.createElement('a'); link.download = filename; link.href = canvas.toDataURL(); link.click(); }}); }}</script></body></html>"""
                components.html(html_pizarra_general + html_pizarra_deps, height=1600, scrolling=True)
                st.markdown("---")
                st.subheader("📱 Resumen para WhatsApp (Cierres y Departamentos)")
                msg_w = f"⏱️ *Reporte de Cierres Semanal - Drotaca 2.0*\n📅 Semana: {int(num_sem)} ({rango_fechas_lv})\n\n📍 *Cronometría de la Droguería:*\n🔹 Promedio Cierre General: *{prom_drotaca}*\n🔹 Promedio Cierre Juanita: *{prom_juanita}*\n\n"
                if not pivot_deps.empty:
                    msg_w += f"📍 *Top 10 Áreas que salieron más tarde (Promedio):*\n"
                    try:
                        pivot_deps['Para_Ordenar'] = pd.to_datetime(pivot_deps['Promedio'], format="%I:%M %p", errors='coerce')
                        top_10 = pivot_deps.sort_values(by='Para_Ordenar', ascending=False).head(10)
                        for dep, row in top_10.iterrows():
                            if pd.notna(row['Para_Ordenar']): msg_w += f"🔹 {str(dep).title()}: *{row['Promedio']}*\n"
                    except: pass
                msg_w += f"\n---\n✅ Tablas de auditoría detalladas adjuntas en imagen."
                st.code(msg_w, language="markdown")

with t_comensales:
    st.info("Consolida el consumo de comedor por departamento de Lunes a Viernes.")
    if st.button("🍽️ Generar Auditoría de Comensales", type="primary", use_container_width=True):
        with st.spinner("Procesando datos de comedor..."):
            df_com_raw = extraer_datos("PIZARRA_COMENSALES")
            if df_com_raw.empty: st.error("No se pudo acceder a la hoja PIZARRA_COMENSALES.")
            else:
                c_fecha_c = buscar_columna_estricta(df_com_raw, ['fecha', 'timestamp'])
                if c_fecha_c:
                    df_com_raw['Fecha_DT'] = df_com_raw[c_fecha_c].apply(extraer_fecha_limpia)
                    df_com_raw['Num_Semana'] = df_com_raw['Fecha_DT'].dt.isocalendar().week
                else:
                    c_sem_c = buscar_columna_estricta(df_com_raw, ['semana'])
                    if c_sem_c: df_com_raw['Num_Semana'] = df_com_raw[c_sem_c].astype(str).str.extract(r'(\d+)').astype(float)
                    else: df_com_raw['Num_Semana'] = num_sem
                df_com = df_com_raw[df_com_raw['Num_Semana'] == num_sem].copy()
                if df_com.empty: st.warning(f"No hay registros de comensales para la Semana {int(num_sem)}.")
                else:
                    if c_fecha_c:
                        df_com = df_com.dropna(subset=['Fecha_DT'])
                        df_com = df_com[df_com['Fecha_DT'].dt.weekday < 5].copy()
                    c_dep = buscar_columna_estricta(df_com, ['departamento', 'area'])
                    c_des = buscar_columna_estricta(df_com, ['desayuno']); c_alm = buscar_columna_estricta(df_com, ['almuerzo']); c_cen = buscar_columna_estricta(df_com, ['cena'])
                    if not c_dep: st.error("No se detectó columna 'Departamento'."); st.stop()
                    df_com[c_dep] = df_com[c_dep].astype(str).str.upper().str.strip()
                    for c in [c_des, c_alm, c_cen]:
                        if c: df_com[c] = pd.to_numeric(df_com[c], errors='coerce').fillna(0)
                    df_grp = df_com.groupby(c_dep).agg({c_des: 'sum' if c_des else lambda x: 0, c_alm: 'sum' if c_alm else lambda x: 0, c_cen: 'sum' if c_cen else lambda x: 0}).reset_index()
                    df_grp['Total_Servicios'] = df_grp[c_des] + df_grp[c_alm] + df_grp[c_cen]
                    df_grp = df_grp[df_grp['Total_Servicios'] > 0].sort_values('Total_Servicios', ascending=False)
                    gran_total = df_grp['Total_Servicios'].sum()
                    df_grp['%'] = (df_grp['Total_Servicios'] / gran_total * 100).fillna(0)
                    tot_des = df_grp[c_des].sum() if c_des else 0; tot_alm = df_grp[c_alm].sum() if c_alm else 0; tot_cen = df_grp[c_cen].sum() if c_cen else 0
                    filas_com_html = ""
                    for _, r in df_grp.iterrows():
                        filas_com_html += f"<tr style='text-align: center;'><td style='padding: 12px; font-weight: bold; text-align: left; background-color: #e3f2fd; color: #333; border: 1px solid #000;'>{r[c_dep]}</td><td style='padding: 12px; color: #555; border: 1px solid #000;'>{f_p(r[c_des]) if c_des else '0'}</td><td style='padding: 12px; color: #555; border: 1px solid #000;'>{f_p(r[c_alm]) if c_alm else '0'}</td><td style='padding: 12px; color: #555; border: 1px solid #000;'>{f_p(r[c_cen]) if c_cen else '0'}</td><td style='padding: 12px; font-weight: 900; font-size: 15px; color: #000; border: 1px solid #000;'>{f_p(r['Total_Servicios'])}</td><td style='padding: 12px; font-weight: bold; color: #555; border: 1px solid #000;'>{r['%']:.1f}%</td></tr>"
                    html_pizarra_comensales = f"""<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head><body><div style="text-align: center; margin-bottom: 15px;"><button onclick="capturarComensales()" style="background: {color_azul}; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR PIZARRA COMENSALES</button></div><div id="pizarra-comensales" style="background: white; width: 900px; margin: auto; border-radius: 15px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: 2px solid {color_azul}; font-family: 'Montserrat', sans-serif;"><div style="background: {color_azul}; color: white; padding: 30px; display: flex; justify-content: space-between; align-items: center;"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">AUDITORÍA DE COMENSALES</div><div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} ({rango_fechas_lv})</div></div></div><table style="width: 100%; border-collapse: collapse;"><thead><tr><th style="background: #e3f2fd; color: {color_azul}; padding: 15px; text-align: left; border: 1px solid #000;">DEPARTAMENTO</th><th style="background: #e3f2fd; color: {color_azul}; padding: 15px; border: 1px solid #000;">DESAYUNOS</th><th style="background: #e3f2fd; color: {color_azul}; padding: 15px; border: 1px solid #000;">ALMUERZOS</th><th style="background: #e3f2fd; color: {color_azul}; padding: 15px; border: 1px solid #000;">CENAS</th><th style="background: #bbdefb; color: {color_azul}; padding: 15px; border: 1px solid #000;">TOTAL PLATOS</th><th style="background: #e3f2fd; color: {color_azul}; padding: 15px; border: 1px solid #000;">% CONSUMO</th></tr></thead><tbody>{filas_com_html}</tbody></table><div style="background: #bbdefb; padding: 15px; display: flex; justify-content: space-around; border-top: 3px solid {color_azul}; font-weight: 900; color: #0d47a1;"><div style="text-align: center;"><span style="font-size:12px; color:#555;">TOT. DESAYUNOS</span><br><span style="font-size:20px;">{f_p(tot_des)}</span></div><div style="text-align: center;"><span style="font-size:12px; color:#555;">TOT. ALMUERZOS</span><br><span style="font-size:20px;">{f_p(tot_alm)}</span></div><div style="text-align: center;"><span style="font-size:12px; color:#555;">TOT. CENAS</span><br><span style="font-size:20px;">{f_p(tot_cen)}</span></div><div style="text-align: center; color:{color_azul};"><span style="font-size:12px;">GRAN TOTAL</span><br><span style="font-size:24px; color:#000;">{f_p(gran_total)}</span></div></div></div><script>function capturarComensales() {{ html2canvas(document.getElementById('pizarra-comensales'), {{ scale: 2 }}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Comensales_Semana_{int(num_sem)}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script></body></html>"""
                    components.html(html_pizarra_comensales, height=1200, scrolling=True)
                    st.markdown("---")
                    st.subheader("📱 Resumen para WhatsApp (Comedor)")
                    msg_c = f"🍽️ *Reporte Semanal de Comensales - Drotaca*\n📅 Semana: {int(num_sem)} ({rango_fechas_lv})\n\n*RESUMEN GENERAL:*\n🍳 Desayunos Servidos: *{f_p(tot_des)}*\n🍲 Almuerzos Servidos: *{f_p(tot_alm)}*\n🍝 Cenas Servidas: *{f_p(tot_cen)}*\n📊 Gran Total de Platos: *{f_p(gran_total)}*\n\n*TOP 5 DEPARTAMENTOS DE MAYOR CONSUMO:*\n"
                    for idx, r in df_grp.head(5).reset_index().iterrows(): msg_c += f"{idx+1}. {r[c_dep].title()} - {f_p(r['Total_Servicios'])} Platos ({r['%']:.1f}%)\n"
                    msg_c += f"\n✅ Pizarra detallada de comedor adjunta."
                    st.code(msg_c, language="markdown")

with t_guardias:
    st.info("Genera las pizarras de Guardia (Seguridad, Flota y Monitoreo).")
    if st.button("🛡️ Generar Pizarras de Guardias", type="primary", use_container_width=True):
        with st.spinner("Extrayendo bases de datos de Guardias..."):
            df_seg_raw = extraer_datos("SEG_ROL_GUARDIA")
            if df_seg_raw.empty: st.error("No se encontraron registros en SEG_ROL_GUARDIA.")
            else:
                grupos_seg, c_area, c_diu, c_noc, c_cant = agrupar_rol_compacto(df_seg_raw, num_sem)
                if not grupos_seg: st.warning(f"No hay registros de Seguridad para la semana {int(num_sem)}.")
                else:
                    for fecha_str, df_grupo in grupos_seg:
                        filas_seg_html = ""
                        for _, r in df_grupo.iterrows():
                            d_list = [f"✓ {x.strip()}" for x in str(r.get(c_diu, '')).split('\n') if x.strip()]
                            n_list = [f"✓ {x.strip()}" for x in str(r.get(c_noc, '')).split('\n') if x.strip()]
                            d_html = "<br>".join(d_list); n_html = "<br>".join(n_list)
                            cant = str(r.get(c_cant, '')) if c_cant and pd.notna(r.get(c_cant, '')) else ""
                            filas_seg_html += f"<tr><td style='padding:10px; border:1px solid #000; font-weight:bold; color:{color_azul}; vertical-align:top;'>{r.get(c_area, '')}</td><td style='padding:10px; border:1px solid #000; text-align:center; font-weight:bold; vertical-align:top;'>{cant}</td><td style='padding:10px; border:1px solid #000; vertical-align:top;'>{d_html}</td><td style='padding:10px; border:1px solid #000; vertical-align:top;'>{n_html}</td></tr>"
                        html_piz_seg = f"""<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head><body style="font-family: Arial, sans-serif; background-color: #f0f2f6; padding: 20px;"><div style="text-align: center; margin-bottom: 15px;"><button onclick="capSeg('{fecha_str.split()[0]}')" style="background: {color_azul}; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR {fecha_str.split()[0]}</button></div><div id="piz-seg-{fecha_str.split()[0]}" style="background:white; width:800px; margin:auto; border:2px solid {color_azul}; border-radius:12px; overflow:hidden; margin-bottom: 40px;"><div style="background-color: {color_azul}; color: white; padding: 25px; display: flex; align-items: center; justify-content: space-between;"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><h2 style="margin:0; font-size:22px;">ROL GUARDIA SEMANAL</h2><p style="margin:0; font-size:14px;">SEGURIDAD INTEGRAL - SEMANA {int(num_sem)}</p></div></div><div style="padding: 0;"><table style="width:100%; border-collapse:collapse; font-size:14px;"><thead><tr><th colspan="4" style="background:#e3f2fd; color:{color_azul}; padding:10px; border:1px solid #000; font-size:15px;">OFICIALES DE SEGURIDAD - {fecha_str}</th></tr><tr style="background:#e8eaf6; color:{color_azul};"><th style="padding:10px; border:1px solid #000; text-align:left; width:25%;">ÁREA ASIGNADA</th><th style="padding:10px; border:1px solid #000; text-align:center; width:10%;">CANT.</th><th style="padding:10px; border:1px solid #000; text-align:left; width:32%;">TURNO DIURNO</th><th style="padding:10px; border:1px solid #000; text-align:left; width:33%;">TURNO NOCTURNO</th></tr></thead><tbody>{filas_seg_html}</tbody></table></div></div><script>function capSeg(id) {{ html2canvas(document.getElementById('piz-seg-'+id), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Seguridad_'+id+'.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script></body></html>"""
                        components.html(html_piz_seg, height=450 + (len(df_grupo)*45), scrolling=True)
                    msg_seg = f"🛡️ *ROL DE GUARDIA DE SEGURIDAD*\n📅 Semana: {int(num_sem)}\n\nSi el reporte de personal tiene guardias duplicadas, recuerda que el sistema comprime las planificaciones en bloques limpios por área.\n📸 Ver distribución detallada en las imágenes adjuntas."
                    st.code(msg_seg, language="markdown")
            st.markdown("---")
            df_flo_raw = extraer_datos("GUARDIA_FLOTA")
            if df_flo_raw.empty: st.error("No se encontraron registros en GUARDIA_FLOTA.")
            else:
                df_flo = filtrar_ultima_carga(df_flo_raw, num_sem)
                if df_flo.empty: st.warning(f"No hay registros de Flota para la semana {int(num_sem)}.")
                else:
                    c_nom = buscar_columna_estricta(df_flo, ['supervisor', 'nombre', 'personal']); c_car = buscar_columna_estricta(df_flo, ['turno', 'cargo']); c_dia = buscar_columna_estricta(df_flo, ['dia', 'dias', 'día', 'días'], evitar=['hora', 'horario']); c_hor = buscar_columna_estricta(df_flo, ['hora', 'horario', 'horas'])
                    if c_nom and c_car:
                        filas_flo_html = ""
                        for _, r in df_flo.iterrows():
                            filas_flo_html += f"<tr><td style='padding:12px; border:1px solid #000; font-weight:bold; font-size:15px;'>{r.get(c_nom, '')}</td><td style='padding:12px; border:1px solid #000; font-size:14px;'>{r.get(c_car, '')}</td><td style='padding:12px; border:1px solid #000; text-align:center; font-size:14px; font-weight:bold; color:{color_azul};'>{r.get(c_dia, '')}</td><td style='padding:12px; border:1px solid #000; text-align:center; font-size:14px; font-weight:bold; color:#000000;'>{r.get(c_hor, '')}</td></tr>"
                        html_piz_flo = f"""<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head><body style="font-family: Arial, sans-serif; background-color: #f0f2f6;"><div style="text-align: center; margin-bottom: 15px;"><button onclick="capFlo()" style="background: {color_azul}; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR GUARDIA FLOTA</button></div><div id="piz-flo" style="background:white; width:800px; margin:auto; border:2px solid {color_azul}; border-radius:12px; overflow:hidden;"><div style="background-color: {color_azul}; color: white; padding: 25px; display: flex; align-items: center; justify-content: space-between;"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><h2 style="margin:0; font-size:22px;">GUARDIA DE FLOTA</h2><p style="margin:0; font-size:14px;">SEMANA {int(num_sem)} ({rango_fechas})</p></div></div><div style="padding: 20px;"><table style="width:100%; border-collapse:collapse; font-size:14px;"><thead><tr style="background:{color_azul}; color:white;"><th style="padding:10px; border:1px solid #000; text-align:left;">PERSONAL</th><th style="padding:10px; border:1px solid #000; text-align:left;">CARGO Y TURNO</th><th style="padding:10px; border:1px solid #000; text-align:center;">DÍAS</th><th style="padding:10px; border:1px solid #000; text-align:center;">HORARIO</th></tr></thead><tbody>{filas_flo_html}</tbody></table></div></div><script>function capFlo() {{ html2canvas(document.getElementById('piz-flo'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Flota_Semana_{int(num_sem)}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script></body></html>"""
                        components.html(html_piz_flo, height=600, scrolling=True)
                        msg_flo = f"🚛 *GUARDIA DE FLOTA*\n📅 Semana: {int(num_sem)}\n\n"
                        for _, r in df_flo.iterrows(): msg_flo += f"👤 *{r.get(c_nom, '')}*\n🔹 Rol: {r.get(c_car, '')}\n🗓️ Días: {r.get(c_dia, '')}\n⏰ Horario: {r.get(c_hor, '')}\n\n"
                        st.code(msg_flo, language="markdown")
            st.markdown("---")
            df_mon_raw = extraer_datos("GUARDIA_MONITOREO")
            if df_mon_raw.empty: st.error("No se encontraron registros en GUARDIA_MONITOREO.")
            else:
                df_mon = filtrar_ultima_carga(df_mon_raw, num_sem)
                if df_mon.empty: st.warning(f"No hay registros de Monitoreo para la semana {int(num_sem)}.")
                else:
                    c_nom = buscar_columna_estricta(df_mon, ['analista', 'nombre', 'turno']); c_hor = buscar_columna_estricta(df_mon, ['horario', 'hora']); c_ram = buscar_columna_estricta(df_mon, ['unidades señor ramon', 'ramon', 'ramón', 'responsable', 'unidades'])
                    if c_nom and c_hor:
                        filas_mon_html = ""
                        uni_ramon = ""
                        for _, r in df_mon.iterrows():
                            filas_mon_html += f"<tr><td style='padding:12px; border:1px solid #000; font-weight:bold; font-size:15px; color:{color_azul};'>{r.get(c_nom, '')}</td><td style='padding:12px; border:1px solid #000; font-size:14px; font-weight:bold;'>{r.get(c_hor, '')}</td></tr>"
                            if c_ram and not uni_ramon and pd.notna(r.get(c_ram, '')): uni_ramon = str(r.get(c_ram, ''))
                        box_ramon = f"<div style='background-color:#f8f9fa; padding:15px; border-left:5px solid #e65100; margin-top:20px; border-radius:5px;'><h4 style='margin:0 0 5px 0; color:#333;'>🚛 Unidades del Sr. Ramón</h4><span style='font-size:15px; font-weight:bold; color:{color_azul};'>Responsable: {uni_ramon}</span></div>" if uni_ramon else ""
                        html_piz_mon = f"""<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head><body style="font-family: Arial, sans-serif; background-color: #f0f2f6;"><div style="text-align: center; margin-bottom: 15px;"><button onclick="capMon()" style="background: {color_azul}; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR GUARDIA MONITOREO</button></div><div id="piz-mon" style="background:white; width:750px; margin:auto; border:2px solid {color_azul}; border-radius:12px; overflow:hidden;"><div style="background-color: {color_azul}; color: white; padding: 25px; display: flex; align-items: center; justify-content: space-between;"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><h2 style="margin:0; font-size:22px;">GUARDIA DE MONITOREO</h2><p style="margin:0; font-size:14px;">SEMANA {int(num_sem)} ({rango_fechas})</p></div></div><div style="padding: 20px;"><table style="width:100%; border-collapse:collapse; font-size:14px;"><thead><tr style="background:{color_azul}; color:white;"><th style="padding:10px; border:1px solid #000; text-align:left;">PERSONAL (TURNO)</th><th style="padding:10px; border:1px solid #000; text-align:left;">HORARIO Y DÍAS</th></tr></thead><tbody>{filas_mon_html}</tbody></table>{box_ramon}</div></div><script>function capMon() {{ html2canvas(document.getElementById('piz-mon'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Monitoreo_Semana_{int(num_sem)}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script></body></html>"""
                        components.html(html_piz_mon, height=600, scrolling=True)
                        msg_mon = f"🖥️ *GUARDIA DE MONITOREO*\n📅 Semana: {int(num_sem)}\n\n🕒 *Guardias Programadas:*\n\n"
                        for _, r in df_mon.iterrows(): msg_mon += f"👤 *{r.get(c_nom, '')}*\n⏰ {r.get(c_hor, '')}\n\n"
                        if uni_ramon: msg_mon += f"🚛 *Unidades del Sr. Ramón*\nResponsable: {uni_ramon}"
                        st.code(msg_mon, language="markdown")

# ---------------------------------------------------------
# PESTAÑA 5: RESUMEN DE DESPACHOS (MIGRADO NETAMENTE A EXCEL)
# ---------------------------------------------------------
with t_despachos:
    st.info("Genera el Resumen Logístico de Lunes a Domingo leyendo directamente de los Excels.")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        archivos_despacho = st.file_uploader("📂 Cargar Excels de Despacho (Oriente, Centro, Occidente)", type=["xlsx", "xlsm"], accept_multiple_files=True)
    with col_up2:
        archivo_flota = st.file_uploader("🚛 Cargar Excel de Kilometraje (Flota)", type=["xlsx", "xlsm"])

    if st.button("🚚 Generar Pizarra de Despachos", type="primary", use_container_width=True):
        if not archivos_despacho or not archivo_flota:
            st.warning("⚠️ Debes cargar al menos un archivo de despacho y el archivo de flota para generar el reporte.")
        else:
            with st.spinner("Procesando histórico de Excels..."):
                registros_log = []
                registros_log.append(f"ℹ️ Iniciando auditoría para el Año {ano_sel} - Semana {int(num_sem)}.")
                
                # --- 1. PROCESAR DESPACHOS (ESCANEO DINÁMICO) ---
                dfs_despacho = []
                for file in archivos_despacho:
                    filename = file.name.upper()
                    try:
                        if 'ORIENTE' in filename: region_macro = 'ORIENTE'
                        elif 'CENTRO' in filename: region_macro = 'CENTRO'
                        elif 'OCCIDENTE' in filename: region_macro = 'OCCIDENTE'
                        else:
                            registros_log.append(f"⚠️ Archivo omitido (no se identificó región en el nombre): {file.name}")
                            continue

                        df_raw = pd.read_excel(file, sheet_name="FARMACIAS", header=None)
                        header_idx = 0
                        for i, row in df_raw.head(20).iterrows():
                            row_vals = [str(val).upper().strip() for val in row.values]
                            if 'BULTOS' in row_vals and 'FECHA DE ENTREGA' in row_vals:
                                header_idx = i
                                break
                        
                        df = df_raw.iloc[header_idx+1:].copy()
                        df.columns = [str(c).upper().strip() for c in df_raw.iloc[header_idx].values]
                        
                        col_fecha = 'FECHA DE ENTREGA'
                        col_bultos = 'BULTOS'
                        col_status = 'TIPO DE ENTREGA'
                        col_subregion = 'RUTAS' if 'RUTAS' in df.columns else 'DESPACHO'
                        col_region = 'DISTRIBUCION' if 'DISTRIBUCION' in df.columns else 'ZONA'
                        
                        missing = [c for c in [col_fecha, col_bultos, col_status, col_subregion, col_region] if c not in df.columns]
                        if missing:
                            registros_log.append(f"⚠️ {file.name}: Falló la lectura. Faltan las columnas: {missing}.")
                            continue
                        
                        df = df.rename(columns={
                            col_fecha: 'Fecha', col_bultos: 'Bultos', col_status: 'Status',
                            col_subregion: 'SubRegion', col_region: 'Region'
                        })
                        
                        df['Region_Macro'] = region_macro
                        cols_to_keep = ['Fecha', 'Bultos', 'Status', 'SubRegion', 'Region', 'Region_Macro']
                        df = df[cols_to_keep]
                        filas_brutas = len(df)
                        
                        df['Fecha'] = df['Fecha'].replace(r'^\s*$', pd.NA, regex=True)
                        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
                        df['Fecha'] = df['Fecha'].ffill()
                        df = df.dropna(subset=['Fecha'])
                        
                        registros_log.append(f"✅ Archivo leído: {file.name} | Filas procesadas: {filas_brutas} -> Con Fecha Válida: {len(df)}")
                        dfs_despacho.append(df)
                    except Exception as e:
                        registros_log.append(f"❌ Error procesando el archivo de despacho {file.name}: {str(e)}")
                
                if not dfs_despacho:
                    st.error("No se pudo extraer información válida de los archivos de despacho cargados. Revisa el log abajo.")
                    for log_msg in registros_log: st.write(log_msg)
                    st.stop()

                df_desp = pd.concat(dfs_despacho, ignore_index=True)
                df_desp['Num_Semana'] = df_desp['Fecha'].dt.isocalendar().week
                df_desp['Ano_Calc'] = df_desp['Fecha'].dt.isocalendar().year
                
                df_sem = df_desp[(df_desp['Num_Semana'] == num_sem) & (df_desp['Ano_Calc'] == ano_sel)].copy()
                registros_log.append(f"📊 Registros totales de despacho (Semana {int(num_sem)}): {len(df_sem)}")
                
                df_sem['Status'] = df_sem['Status'].astype(str).str.upper().str.strip()
                filas_antes_status = len(df_sem)
                df_sem = df_sem[df_sem['Status'] == 'ENTREGADO']
                registros_log.append(f"🎯 Filtrados con Status 'ENTREGADO': {len(df_sem)} (Omitidos: {filas_antes_status - len(df_sem)})")
                
                df_sem['Bultos'] = pd.to_numeric(df_sem['Bultos'], errors='coerce').fillna(0)
                df_sem['Pedidos'] = 1 
                df_sem['SubRegion_Clean'] = df_sem.apply(lambda x: asignar_subregion(x['SubRegion'], x['Region_Macro']), axis=1)

                # --- 2. PROCESAR MAESTRO 'FLOTA ACTUAL' ---
                placas_despacho_validas = None
                try:
                    df_flota_actual = pd.read_excel(archivo_flota, sheet_name="FLOTA ACTUAL")
                    df_flota_actual.columns = df_flota_actual.columns.astype(str).str.strip().str.upper()
                    
                    if 'PLACA' in df_flota_actual.columns and 'RUTA' in df_flota_actual.columns:
                        df_flota_actual['PLACA'] = df_flota_actual['PLACA'].astype(str).str.strip().str.upper()
                        df_flota_actual['RUTA'] = df_flota_actual['RUTA'].astype(str).str.strip().str.upper()
                        
                        rutas_operativas = ['ORIENTE', 'CENTRO', 'OCCIDENTE']
                        df_solo_despacho = df_flota_actual[df_flota_actual['RUTA'].isin(rutas_operativas)]
                        placas_despacho_validas = set(df_solo_despacho['PLACA'].unique())
                        
                        registros_log.append(f"✅ 'FLOTA ACTUAL' procesada. Vehículos autorizados para Despacho: {len(placas_despacho_validas)}.")
                except: pass

                # --- 3. PROCESAR KILOMETRAJE ---
                try:
                    df_flota_raw = pd.read_excel(archivo_flota, sheet_name="BASE DE DATOS")
                    df_flota_raw.columns = df_flota_raw.columns.astype(str).str.strip().str.upper() 
                    df_flota_raw = df_flota_raw.loc[:, ~df_flota_raw.columns.str.contains('^UNNAMED')]
                    
                    df_flota_raw['FECHAS'] = pd.to_datetime(df_flota_raw['FECHAS'], dayfirst=True, errors='coerce')
                    df_flota_raw['FECHAS'] = df_flota_raw['FECHAS'].ffill() 
                    df_flota_raw = df_flota_raw.dropna(subset=['FECHAS'])
                    
                    columnas_id = ['FECHAS', 'DIA', 'MES']
                    columnas_id_existentes = [col for col in columnas_id if col in df_flota_raw.columns]
                    columnas_placas = [col for col in df_flota_raw.columns if col not in columnas_id_existentes]
                    
                    df_flota = df_flota_raw.melt(id_vars=columnas_id_existentes, value_vars=columnas_placas, var_name='PLACA', value_name='KILOMETROS')
                    df_flota['KILOMETROS'] = pd.to_numeric(df_flota['KILOMETROS'], errors='coerce').fillna(0)
                    df_flota['Num_Semana'] = df_flota['FECHAS'].dt.isocalendar().week
                    df_flota['Ano_Calc'] = df_flota['FECHAS'].dt.isocalendar().year
                    
                    df_km_sem = df_flota[(df_flota['Num_Semana'] == num_sem) & (df_flota['Ano_Calc'] == ano_sel)].copy()
                    
                    if placas_despacho_validas is not None:
                        km_brutos = df_km_sem['KILOMETROS'].sum()
                        df_km_sem['PLACA'] = df_km_sem['PLACA'].astype(str).str.strip().str.upper()
                        df_km_sem = df_km_sem[df_km_sem['PLACA'].isin(placas_despacho_validas)].copy()
                        registros_log.append(f"🎯 KMs Netos de Despacho calculados: {df_km_sem['KILOMETROS'].sum():,.2f} Kms.")
                except Exception as e:
                    st.error(f"Error procesando el archivo de Kilometraje: {e}")
                    st.stop()

                if df_sem.empty and df_km_sem.empty:
                    st.warning(f"No hay registros de despachos ni kilometraje para la Semana {int(num_sem)} en el Año {ano_sel}.")
                else:
                    # --- 4. CONSOLIDACIÓN DE DATOS DIARIOS ---
                    df_diario_desp = df_sem.groupby(df_sem['Fecha'].dt.date).agg({'Pedidos': 'sum', 'Bultos': 'sum'}).reset_index()
                    df_diario_desp.rename(columns={'Fecha': 'Date'}, inplace=True)
                    df_diario_desp['Date'] = pd.to_datetime(df_diario_desp['Date']).dt.date
                    
                    df_diario_km = df_km_sem.groupby(df_km_sem['FECHAS'].dt.date).agg({'KILOMETROS': 'sum'}).reset_index()
                    df_diario_km.rename(columns={'FECHAS': 'Date'}, inplace=True)
                    df_diario_km['Date'] = pd.to_datetime(df_diario_km['Date']).dt.date

                    df_diario = pd.merge(df_diario_desp, df_diario_km, on='Date', how='outer').fillna(0).sort_values('Date')
                    
                    total_pedidos = df_diario['Pedidos'].sum()
                    total_bultos = df_diario['Bultos'].sum()
                    total_kms = df_diario['KILOMETROS'].sum()
                    dias_activos = max(len(df_diario[df_diario['Pedidos'] > 0]), 1)
                    
                    prom_ped = total_pedidos / dias_activos
                    prom_bul = total_bultos / dias_activos
                    prom_kms = total_kms / dias_activos
                    df_diario['% Diario'] = (df_diario['Pedidos'] / total_pedidos * 100).fillna(0) if total_pedidos > 0 else 0
                    
                    def get_arrow(val, prom):
                        if val > prom * 1.05: return "<span style='color:#2e7d32; font-size:18px;'>⬆️</span>", "#2e7d32"
                        elif val < prom * 0.95: return "<span style='color:#d32f2f; font-size:18px;'>⬇️</span>", "#d32f2f"
                        return "<span style='color:#f57c00; font-size:18px;'>➡️</span>", "#f57c00"

                    filas_diarias_html = ""
                    dias_semana_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                    meses_es = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

                    for _, r in df_diario.iterrows():
                        fecha_obj = pd.to_datetime(r['Date'])
                        str_fecha = f"{dias_semana_es[fecha_obj.weekday()]} {fecha_obj.day} de {meses_es[fecha_obj.month-1]}"
                        arr_ped, c_ped = get_arrow(r['Pedidos'], prom_ped)
                        arr_por, c_por = get_arrow(r['% Diario'], (100/dias_activos))
                        
                        filas_diarias_html += f"""
                        <tr style="text-align: center; border-bottom: 1px solid #000; background: white;">
                            <td style="padding: 8px; border: 1px solid #000; font-weight: bold; text-align: left;">{str_fecha}</td>
                            <td style="padding: 8px; border: 1px solid #000; font-weight: 900; font-size: 16px; color: {c_ped};">{arr_ped} {f_p(r['Pedidos'])}</td>
                            <td style="padding: 8px; border: 1px solid #000; font-weight: 900; font-size: 16px; color: {c_por};">{arr_por} {r['% Diario']:.2f}%</td>
                            <td style="padding: 8px; border: 1px solid #000; font-weight: 900; font-size: 14px;">{f_p(r['Bultos'])} BULTOS</td>
                            <td style="padding: 8px; border: 1px solid #000; font-weight: 900; font-size: 14px;">{f_p(r['KILOMETROS'])} Kms</td>
                        </tr>
                        """

                    df_reg = df_sem.groupby(['Region_Macro', 'SubRegion_Clean']).agg({'Pedidos': 'sum', 'Bultos': 'sum'}).reset_index()

                    def generar_bloque_region(nombre_macro, color_macro, df_f):
                        if df_f.empty: return ""
                        t_ped = df_f['Pedidos'].sum()
                        t_bul = df_f['Bultos'].sum()
                        
                        html_bloque = f"""
                        <div style="display:flex; justify-content:space-between; margin-bottom: 10px;">
                            <div style="width: 55%;">
                                <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; font-size: 13px;">
                                    <tr>
                                        <th style="padding: 8px; background: #e0e0e0; border: 1px solid #000; text-align: left;">PEDIDOS {nombre_macro} (GENERAL)</th>
                                        <th style="padding: 8px; background: {color_macro}; color: white; border: 1px solid #000;">{f_p(t_ped)} PEDIDOS</th>
                                        <th style="padding: 8px; background: {color_macro}; color: white; border: 1px solid #000;">{f_p(t_bul)} BULTOS</th>
                                    </tr>
                        """
                        barras_html = ""
                        colores = ["#1976d2", "#e65100", "#388e3c", "#fbc02d", "#8e24aa", "#e91e63", "#00bcd4", "#ff9800", "#4caf50"]
                        
                        for i, r in df_f.iterrows():
                            c_p, c_c = get_arrow(r['Pedidos'], t_ped/len(df_f) if len(df_f)>0 else 0)
                            html_bloque += f"""
                                    <tr style="background: white;">
                                        <td style="padding: 6px; border: 1px solid #000; font-weight: bold; text-align: right;">{r['SubRegion_Clean']}</td>
                                        <td style="padding: 6px; border: 1px solid #000; font-weight: 900; text-align: right;">{c_p} {f_p(r['Pedidos'])}</td>
                                        <td style="padding: 6px; border: 1px solid #000; font-weight: 900; text-align: right;">{f_p(r['Bultos'])}</td>
                                    </tr>
                            """
                            pct = (r['Pedidos'] / t_ped * 100) if t_ped > 0 else 0
                            col_bar = colores[i % len(colores)]
                            if pct > 0: barras_html += f"<div style='width: {pct}%; background-color: {col_bar}; color: white; font-size:10px; text-align:center; font-weight:bold; padding: 5px 0; border-right: 1px solid white;'>{int(pct)}%</div>"
                            
                        html_bloque += f"""
                                </table>
                            </div>
                            <div style="width: 43%; background: #333; color: white; padding: 10px; border-radius: 5px; border: 1px solid #000; display:flex; flex-direction:column; justify-content:center;">
                                <div style="text-align: center; font-size: 12px; font-weight: bold; margin-bottom: 10px; text-transform: uppercase;">PEDIDOS ENTREGADOS {nombre_macro}</div>
                                <div style="display:flex; width: 100%; border-radius: 4px; overflow: hidden; border: 1px solid #fff;">{barras_html}</div>
                                <div style="display:flex; flex-wrap: wrap; justify-content: center; margin-top: 10px;">
                        """
                        for i, r in df_f.iterrows():
                            col_bar = colores[i % len(colores)]
                            html_bloque += f"<div style='font-size: 9px; margin-right: 10px;'><span style='display:inline-block; width:10px; height:10px; background:{col_bar}; margin-right:3px;'></span>{r['SubRegion_Clean']}</div>"
                            
                        html_bloque += "</div></div></div>"
                        return html_bloque

                    bloque_oriente = generar_bloque_region("ORIENTE", "#2e7d32", df_reg[df_reg['Region_Macro'].str.upper() == 'ORIENTE'])
                    bloque_centro = generar_bloque_region("CENTRO", "#1565c0", df_reg[df_reg['Region_Macro'].str.upper() == 'CENTRO'])
                    bloque_occidente = generar_bloque_region("OCCIDENTE", "#e65100", df_reg[df_reg['Region_Macro'].str.upper() == 'OCCIDENTE'])

                    # AQUI AGREGAMOS EL HEADER CON LOGO Y TÍTULO A LA PIZARRA DE DESPACHOS
                    html_pizarra_despachos = f"""
                    <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head>
                    <body style="font-family: Arial, sans-serif; background-color: #f0f2f6; padding: 20px;">
                    <div style="text-align: center; margin-bottom: 15px;">
                        <button onclick="capDespachos()" style="background: #1565c0; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR RESUMEN DESPACHOS</button>
                    </div>
                    <div id="piz-despachos" style="background:white; width:950px; margin:auto; border:2px solid #1565c0; border-radius:12px; overflow:hidden;">
                        
                        <div style="background-color: #1565c0; color: white; padding: 25px 30px; display: flex; align-items: center; justify-content: space-between; border-bottom: 5px solid #d4af37;">
                            <img src="{logo_b64}" style="height: 50px;">
                            <div style="text-align: right;">
                                <h2 style="margin:0; font-size:24px; font-weight: 900;">DESPACHOS A NIVEL NACIONAL</h2>
                                <p style="margin:5px 0 0 0; font-size:14px; font-weight:bold; color:#d4af37; letter-spacing: 1px;">SEMANA {int(num_sem)} ({rango_fechas.upper()})</p>
                            </div>
                        </div>

                        <div style="padding: 15px;">
                            <table style="width: 100%; border-collapse: collapse; font-size: 13px; border: 2px solid #000;">
                                <thead>
                                    <tr style="background: #e3f2fd; color: #1565c0;">
                                        <th style="padding: 10px; border: 1px solid #000;">DÍA DE LA SEMANA</th>
                                        <th style="padding: 10px; border: 1px solid #000;">PEDIDOS ENTREGADOS</th>
                                        <th style="padding: 10px; border: 1px solid #000;">% DIARIO</th>
                                        <th style="padding: 10px; border: 1px solid #000;">BULTOS ENTREGADOS</th>
                                        <th style="padding: 10px; border: 1px solid #000;">RECORRIDO LOGÍSTICO</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filas_diarias_html}
                                    <tr style="background: #1565c0; color: white; font-weight: 900; font-size: 16px;">
                                        <td style="padding: 10px; border: 1px solid #000; text-align: center;">Total general</td>
                                        <td style="padding: 10px; border: 1px solid #000; text-align: center;">{f_p(total_pedidos)}</td>
                                        <td style="padding: 10px; border: 1px solid #000; text-align: center;">100,00%</td>
                                        <td style="padding: 10px; border: 1px solid #000; text-align: center;">{f_p(total_bultos)} BULTOS</td>
                                        <td style="padding: 10px; border: 1px solid #000; text-align: center;">{f_p(total_kms)} Kms</td>
                                    </tr>
                                </tbody>
                            </table>

                            <div style="display: flex; justify-content: space-between; margin-top: 15px;">
                                <table style="width: 42%; border-collapse: collapse; border: 2px solid #000; font-size: 13px;">
                                    <tr><th colspan="2" style="background: #1565c0; color: white; padding: 8px;">PROMEDIO DIARIO</th></tr>
                                    <tr style="background: white;"><td style="padding: 8px; border: 1px solid #000; font-weight: bold;">PEDIDOS ENTREGADOS</td><td style="padding: 8px; border: 1px solid #000; font-weight: 900; text-align: center;">{f_p(prom_ped)} PEDIDOS</td></tr>
                                    <tr style="background: white;"><td style="padding: 8px; border: 1px solid #000; font-weight: bold;">BULTOS ENTREGADOS</td><td style="padding: 8px; border: 1px solid #000; font-weight: 900; text-align: center;">{f_p(prom_bul)} BULTOS</td></tr>
                                    <tr style="background: white;"><td style="padding: 8px; border: 1px solid #000; font-weight: bold;">KILOMETRAJE DIARIO</td><td style="padding: 8px; border: 1px solid #000; font-weight: 900; text-align: center;">{f_p(prom_kms)} Kms</td></tr>
                                </table>
                                
                                <div style="width: 14%; display: flex; align-items: center; justify-content: center;">
                                    <img src="{logo_b64}" style="width: 100%;">
                                </div>
                                
                                <table style="width: 42%; border-collapse: collapse; border: 2px solid #000; font-size: 13px;">
                                    <tr><th colspan="2" style="background: #1565c0; color: white; padding: 8px;">RESULTADO SEMANAL</th></tr>
                                    <tr style="background: white;"><td style="padding: 8px; border: 1px solid #000; font-weight: bold;">PEDIDOS ENTREGADOS</td><td style="padding: 8px; border: 1px solid #000; font-weight: 900; text-align: center;">{f_p(total_pedidos)} PEDIDOS</td></tr>
                                    <tr style="background: white;"><td style="padding: 8px; border: 1px solid #000; font-weight: bold;">BULTOS ENTREGADOS</td><td style="padding: 8px; border: 1px solid #000; font-weight: 900; text-align: center;">{f_p(total_bultos)} BULTOS</td></tr>
                                    <tr style="background: white;"><td style="padding: 8px; border: 1px solid #000; font-weight: bold;">KILOMETRAJE SEMANAL</td><td style="padding: 8px; border: 1px solid #000; font-weight: 900; text-align: center;">{f_p(total_kms)} Kms</td></tr>
                                </table>
                            </div>

                            <div style="background: #1565c0; color: white; text-align: center; font-size: 24px; font-weight: 900; padding: 10px; margin-top: 15px; border: 2px solid #000;">
                                ESTADÍSTICA POR REGIONES
                            </div>
                            <div style="margin-top: 10px;">
                                {bloque_oriente}
                                {bloque_centro}
                                {bloque_occidente}
                            </div>
                        </div>
                    </div>
                    <script>function capDespachos() {{ html2canvas(document.getElementById('piz-despachos'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Resumen_Despachos_{int(num_sem)}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>
                    </body></html>
                    """
                    components.html(html_pizarra_despachos, height=1350, scrolling=True)

                    st.markdown("---")
                    st.subheader("📱 Mensaje para WhatsApp (Resumen Ejecutivo)")
                    msg_d = f"📊 *RESUMEN SEMANAL DE DESPACHOS*\n📅 Semana: {int(num_sem)} ({rango_fechas})\n\n"
                    msg_d += f"*📍 RESULTADO GLOBAL (SOLO UNIDADES DE DESPACHO):*\n"
                    msg_d += f"🏥 Pedidos Entregados: *{f_p(total_pedidos)}*\n"
                    msg_d += f"📦 Bultos Entregados: *{f_p(total_bultos)}*\n"
                    msg_d += f"📏 Recorrido Total: *{f_p(total_kms)} Kms*\n\n"
                    msg_d += f"*⏱️ PROMEDIO DIARIO OPERATIVO:*\n"
                    msg_d += f"🔹 Pedidos: {f_p(prom_ped)}\n🔹 Bultos: {f_p(prom_bul)}\n🔹 KMs: {f_p(prom_kms)}\n\n"
                    msg_d += "✅ *Dashboard estadístico adjunto en imagen.*"
                    st.code(msg_d, language="markdown")
                    
                    st.markdown("---")
                    st.subheader("📋 Registro de Auditoría (Log de Procesamiento)")
                    for log_msg in registros_log:
                        if "❌" in log_msg: st.error(log_msg)
                        elif "⚠️" in log_msg: st.warning(log_msg)
                        elif "✅" in log_msg: st.success(log_msg)
                        else: st.info(log_msg)

# ---------------------------------------------------------
# PESTAÑA 6: CONTROL DE RESERVA DE COMBUSTIBLE
# ---------------------------------------------------------
with t_combustible:
    st.info("Monitoreo de Extracción y Recarga de Tanques y Bidones de Combustible (Base El Tigre).")
    
    if st.button("⛽ Generar Pizarra de Combustible", type="primary", use_container_width=True):
        with st.spinner("Conectando con Google Sheets y calculando variaciones..."):
            df_comb_raw = extraer_datos("FLOTA_COMBUSTIBLE")
            if df_comb_raw.empty:
                st.error("No se pudo conectar con la hoja 'FLOTA_COMBUSTIBLE' o está vacía.")
            else:
                # Estandarizar columnas y buscar num_sem
                df_comb_raw.columns = df_comb_raw.columns.str.strip()
                df_comb_raw['Fecha_DT'] = pd.to_datetime(df_comb_raw['Fecha'], format='%d/%m/%Y', errors='coerce')
                df_comb_raw = df_comb_raw.dropna(subset=['Fecha_DT'])
                
                # Filtrar la semana seleccionada
                df_comb_raw['Num_Semana'] = df_comb_raw['Fecha_DT'].dt.isocalendar().week
                df_comb = df_comb_raw[df_comb_raw['Num_Semana'] == num_sem].sort_values('Fecha_DT').copy()
                
                if df_comb.empty:
                    st.warning(f"No existen registros de combustible para la Semana {int(num_sem)}.")
                else:
                    # Función robusta para limpiar errores de tipeo como "15.000.0" o "15,000"
                    def limpiar_volumen(v):
                        if pd.isna(v) or str(v).strip() == "": return 0.0
                        s = str(v).strip().replace(',', '.')
                        # Si teclearon más de un punto (ej. 15.000.0), borramos los primeros y dejamos el último como decimal
                        if s.count('.') > 1:
                            partes = s.split('.')
                            s = "".join(partes[:-1]) + "." + partes[-1]
                        try:
                            return float(s)
                        except:
                            return 0.0

                    # Columnas numéricas aplicándoles el filtro limpiador antibasura
                    cols_tanques = ['Tanque_1_50K', 'Tanque_2_12K', 'Tanque_3_7K', 'Total_Tanques']
                    cols_bidones = ['Gasolina_Bidones', 'Gasoil_Bidones']
                    for col in cols_tanques + cols_bidones:
                        if col in df_comb.columns:
                            df_comb[col] = df_comb[col].apply(limpiar_volumen)
                        else:
                            df_comb[col] = 0.0

                    # 1. Apertura (Lunes/Primer día) y Cierre (Viernes/Último día)
                    row_apertura = df_comb.iloc[0]
                    row_cierre = df_comb.iloc[-1]
                    fecha_apertura = row_apertura['Fecha_DT'].strftime('%d/%m/%Y')
                    fecha_cierre = row_cierre['Fecha_DT'].strftime('%d/%m/%Y')

                    # 2. Función para calcular Consumo (bajadas) y Recargas (subidas)
                    def calcular_variacion(col_nombre):
                        consumo = 0
                        recarga = 0
                        valores = df_comb[col_nombre].tolist()
                        for i in range(1, len(valores)):
                            delta = valores[i] - valores[i-1]
                            if delta > 0: recarga += delta      # Si sube, llegó cisterna/llenaron
                            elif delta < 0: consumo += abs(delta) # Si baja, se extrajo para uso
                        return consumo, recarga

                    cons_50k, rec_50k = calcular_variacion('Tanque_1_50K')
                    cons_12k, rec_12k = calcular_variacion('Tanque_2_12K')
                    cons_7k, rec_7k = calcular_variacion('Tanque_3_7K')
                    cons_tot, rec_tot = calcular_variacion('Total_Tanques')

                    # Para los bidones nos interesa el stock de apertura, cierre y la suma de consumos reales
                    cons_gasolina, rec_gasolina = calcular_variacion('Gasolina_Bidones')
                    cons_gasoil, rec_gasoil = calcular_variacion('Gasoil_Bidones')

                    # --- CONSTRUCCIÓN DEL DASHBOARD HTML ---
                    
                    # Generador de tarjetas para los Tanques
                    def crear_card_tanque(titulo, apertura, cierre, consumo, recarga, max_cap):
                        pct_apertura = min((apertura / max_cap) * 100, 100) if max_cap > 0 else 0
                        pct_cierre = min((cierre / max_cap) * 100, 100) if max_cap > 0 else 0
                        
                        color_tanque = "#2e7d32" if pct_cierre > 50 else ("#f57c00" if pct_cierre > 20 else "#d32f2f")

                        return f"""
                        <div style="background: white; border: 1px solid #ccc; border-radius: 8px; padding: 15px; width: 31%; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                            <h3 style="margin: 0 0 10px 0; color: #1565c0; font-size: 16px; border-bottom: 2px solid #1565c0; padding-bottom: 5px;">🛢️ {titulo}</h3>
                            <div style="display: flex; justify-content: space-between; font-size: 13px; font-weight: bold; margin-bottom: 5px;">
                                <span>Apertura: <span style="color:#555;">{f_p(apertura)} L</span></span>
                                <span>Cierre: <span style="color:{color_tanque};">{f_p(cierre)} L</span></span>
                            </div>
                            <div style="width: 100%; background: #e0e0e0; height: 10px; border-radius: 5px; margin-bottom: 15px; overflow: hidden;">
                                <div style="width: {pct_cierre}%; background: {color_tanque}; height: 100%;"></div>
                            </div>
                            <table style="width: 100%; font-size: 12px; border-collapse: collapse;">
                                <tr>
                                    <td style="padding: 5px; border-top: 1px solid #eee;">🔻 <b>Extracción:</b></td>
                                    <td style="padding: 5px; border-top: 1px solid #eee; text-align: right; color: #d32f2f; font-weight: bold;">{f_p(consumo)} L</td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px; border-top: 1px solid #eee;">⬆️ <b>Recargas:</b></td>
                                    <td style="padding: 5px; border-top: 1px solid #eee; text-align: right; color: #2e7d32; font-weight: bold;">{f_p(recarga)} L</td>
                                </tr>
                            </table>
                        </div>
                        """

                    card_50k = crear_card_tanque("TANQUE 50.000", row_apertura['Tanque_1_50K'], row_cierre['Tanque_1_50K'], cons_50k, rec_50k, 50000)
                    card_12k = crear_card_tanque("TANQUE 12.000", row_apertura['Tanque_2_12K'], row_cierre['Tanque_2_12K'], cons_12k, rec_12k, 12000)
                    card_7k = crear_card_tanque("TANQUE 7.000", row_apertura['Tanque_3_7K'], row_cierre['Tanque_3_7K'], cons_7k, rec_7k, 7000)

                    # --- GRÁFICA DE BARRAS DE RENDIMIENTO DIARIO ---
                    filas_rendimiento = ""
                    dias_semana_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                    
                    for i, row in df_comb.iterrows():
                        fecha_str = row['Fecha_DT'].strftime('%d/%m')
                        dia_nombre = dias_semana_es[row['Fecha_DT'].weekday()]
                        
                        # Barras visuales
                        l_50 = row['Tanque_1_50K']
                        l_12 = row['Tanque_2_12K']
                        l_7 = row['Tanque_3_7K']
                        t_tot = row['Total_Tanques']
                        
                        pct_tot = min((t_tot / 69000) * 100, 100) # Capacidad maxima base 69000
                        
                        filas_rendimiento += f"""
                        <tr>
                            <td style="padding: 8px; border: 1px solid #ccc; font-weight: bold; width: 15%;">{dia_nombre} {fecha_str}</td>
                            <td style="padding: 8px; border: 1px solid #ccc; width: 50%;">
                                <div style="width: 100%; background: #e0e0e0; height: 16px; border-radius: 4px; overflow: hidden; position: relative;">
                                    <div style="width: {pct_tot}%; background: #1565c0; height: 100%;"></div>
                                </div>
                            </td>
                            <td style="padding: 8px; border: 1px solid #ccc; font-weight: 900; text-align: center; color: #1565c0; width: 15%;">{f_p(t_tot)} L</td>
                            <td style="padding: 8px; border: 1px solid #ccc; font-size: 11px; color: #555; text-align: center;">50k: {f_p(l_50)} | 12k: {f_p(l_12)} | 7k: {f_p(l_7)}</td>
                        </tr>
                        """

                    html_pizarra_combustible = f"""
                    <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head>
                    <body style="font-family: Arial, sans-serif; background-color: #f0f2f6; padding: 20px;">
                    <div style="text-align: center; margin-bottom: 15px;">
                        <button onclick="capCombustible()" style="background: #e65100; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR REPORTE COMBUSTIBLE</button>
                    </div>
                    <div id="piz-combustible" style="background:white; width:950px; margin:auto; border:2px solid #e65100; border-radius:12px; overflow:hidden;">
                        
                        <div style="background-color: #e65100; color: white; padding: 25px 30px; display: flex; align-items: center; justify-content: space-between; border-bottom: 5px solid #333;">
                            <img src="{logo_b64}" style="height: 50px;">
                            <div style="text-align: right;">
                                <h2 style="margin:0; font-size:24px; font-weight: 900;">CONTROL DE RESERVA COMBUSTIBLE</h2>
                                <p style="margin:5px 0 0 0; font-size:14px; font-weight:bold; letter-spacing: 1px;">SEMANA {int(num_sem)} ({fecha_apertura} - {fecha_cierre})</p>
                            </div>
                        </div>

                        <div style="padding: 20px;">
                            <h3 style="margin-top: 0; color: #333; font-size: 18px; text-transform: uppercase;">1. Estatus de Tanques Principales (Gasoil)</h3>
                            <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                                {card_50k}
                                {card_12k}
                                {card_7k}
                            </div>

                            <h3 style="color: #333; font-size: 18px; text-transform: uppercase; margin-top: 30px;">2. Rendimiento y Extracción Semanal</h3>
                            <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 20px;">
                                <thead>
                                    <tr style="background: #f5f5f5; color: #333;">
                                        <th style="padding: 10px; border: 1px solid #ccc; text-align: left;">DÍA</th>
                                        <th style="padding: 10px; border: 1px solid #ccc; text-align: left;">NIVEL GLOBAL DE TANQUES</th>
                                        <th style="padding: 10px; border: 1px solid #ccc;">LITROS TOTALES</th>
                                        <th style="padding: 10px; border: 1px solid #ccc;">DESGLOSE POR TANQUE</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filas_rendimiento}
                                </tbody>
                            </table>

                            <div style="display: flex; justify-content: space-between; margin-top: 30px;">
                                <table style="width: 48%; border-collapse: collapse; border: 2px solid #333; font-size: 14px;">
                                    <tr><th colspan="2" style="background: #333; color: white; padding: 10px; font-size: 15px;">TOTALES BASE (TANQUES)</th></tr>
                                    <tr style="background: white;">
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">🔻 Extracción Total (Consumo):</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: 900; text-align: right; color: #d32f2f;">{f_p(cons_tot)} L</td>
                                    </tr>
                                    <tr style="background: #f9f9f9;">
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">⬆️ Llenado Total (Cisternas):</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: 900; text-align: right; color: #2e7d32;">{f_p(rec_tot)} L</td>
                                    </tr>
                                    <tr style="background: white;">
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">📊 Balance Final de Cierre:</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: 900; text-align: right; color: #1565c0; font-size: 16px;">{f_p(row_cierre['Total_Tanques'])} L</td>
                                    </tr>
                                </table>

                                <table style="width: 48%; border-collapse: collapse; border: 2px solid #e65100; font-size: 14px;">
                                    <tr><th colspan="3" style="background: #e65100; color: white; padding: 10px; font-size: 15px;">ESTATUS DE BIDONES (RESERVA MOVIL)</th></tr>
                                    <tr style="background: #ffe0b2; color: #d84315;">
                                        <th style="padding: 8px; border: 1px solid #ccc;">Tipo</th>
                                        <th style="padding: 8px; border: 1px solid #ccc;">Lunes (Ap.)</th>
                                        <th style="padding: 8px; border: 1px solid #ccc;">Viernes (Ci.)</th>
                                    </tr>
                                    <tr style="background: white; text-align: center;">
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">⛽ Gasolina</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">{f_p(row_apertura['Gasolina_Bidones'])} L</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: 900; color: #1565c0;">{f_p(row_cierre['Gasolina_Bidones'])} L</td>
                                    </tr>
                                    <tr style="background: white; text-align: center;">
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">🛢️ Gasoil</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: bold;">{f_p(row_apertura['Gasoil_Bidones'])} L</td>
                                        <td style="padding: 10px; border: 1px solid #ccc; font-weight: 900; color: #1565c0;">{f_p(row_cierre['Gasoil_Bidones'])} L</td>
                                    </tr>
                                </table>
                            </div>
                        </div>
                    </div>
                    <script>function capCombustible() {{ html2canvas(document.getElementById('piz-combustible'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Combustible_Semana_{int(num_sem)}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>
                    </body></html>
                    """
                    components.html(html_pizarra_combustible, height=1100, scrolling=True)

                    st.markdown("---")
                    st.subheader("📱 Mensaje para WhatsApp (Resumen Ejecutivo)")
                    
                    msg_comb = f"⛽ *REPORTE SEMANAL DE COMBUSTIBLE*\n📅 Semana: {int(num_sem)}\n\n"
                    msg_comb += f"*ESTATUS DE TANQUES (BASE EL TIGRE):*\n"
                    msg_comb += f"🛢️ Tanque 50K: {f_p(row_apertura['Tanque_1_50K'])}L ➡️ {f_p(row_cierre['Tanque_1_50K'])}L\n"
                    msg_comb += f"🛢️ Tanque 12K: {f_p(row_apertura['Tanque_2_12K'])}L ➡️ {f_p(row_cierre['Tanque_2_12K'])}L\n"
                    msg_comb += f"🛢️ Tanque 7K: {f_p(row_apertura['Tanque_3_7K'])}L ➡️ {f_p(row_cierre['Tanque_3_7K'])}L\n\n"
                    msg_comb += f"*BIDONES EN BASE (CIERRE):*\n"
                    msg_comb += f"⛽ Gasolina: *{f_p(row_cierre['Gasolina_Bidones'])}L*\n"
                    msg_comb += f"🛢️ Gasoil: *{f_p(row_cierre['Gasoil_Bidones'])}L*\n\n"
                    msg_comb += f"*MOVIMIENTOS DE LA SEMANA (TANQUES):*\n"
                    msg_comb += f"🔻 Total Extracción/Consumo: *{f_p(cons_tot)} Litros*\n"
                    msg_comb += f"⬆️ Total Recargas (Cisternas): *{f_p(rec_tot)} Litros*\n\n"
                    msg_comb += "✅ *Dashboard visual adjunto.*"
                    
                    st.code(msg_comb, language="markdown")
                    
# ---------------------------------------------------------
# PESTAÑA 7: PIZARRA DE RESULTADO DE SURTIDO Y EXTRACCIÓN (VERSIÓN SIMPLE Y AL PUNTO)
# ---------------------------------------------------------
with t_surtido:
    st.info("Resumen simplificado del consumo de combustible por Rutas, Método de Surtido y Extracción Interna.")
    
    if st.button("⛽ Calcular Resumen Semanal de Surtido", type="primary", use_container_width=True):
        with st.spinner("Procesando transacciones de surtido desde la base de datos..."):
            df_surt_raw = extraer_datos("SURTIDO_COMBUSTIBLE")
            
            if df_surt_raw.empty:
                st.error("No se pudo acceder a la hoja 'SURTIDO_COMBUSTIBLE' o la tabla está vacía.")
            else:
                # Estandarizar columnas y fechas
                df_surt_raw.columns = df_surt_raw.columns.str.strip()
                df_surt_raw['Fecha_DT'] = pd.to_datetime(df_surt_raw['Fecha'], format='%d/%m/%Y', errors='coerce')
                df_surt_raw = df_surt_raw.dropna(subset=['Fecha_DT'])
                
                df_surt_raw['Num_Semana'] = df_surt_raw['Fecha_DT'].dt.isocalendar().week
                df_surt_raw['Ano_Calc'] = df_surt_raw['Fecha_DT'].dt.isocalendar().year
                
                df_surt = df_surt_raw[(df_surt_raw['Num_Semana'] == num_sem) & (df_surt_raw['Ano_Calc'] == ano_sel)].sort_values('Fecha_DT').copy()
                
                if df_surt.empty:
                    st.warning(f"No se encontraron transacciones de surtido para la Semana {int(num_sem)} del Año {ano_sel}.")
                else:
                    # Limpiador numérico
                    def limpiar_litros(l):
                        if pd.isna(l) or str(l).strip() == "": return 0.0
                        s = str(l).strip().replace(',', '.')
                        if s.count('.') > 1:
                            partes = s.split('.')
                            s = "".join(partes[:-1]) + "." + partes[-1]
                        try: return float(s)
                        except: return 0.0

                    df_surt['LITROS'] = df_surt['LITROS'].apply(limpiar_litros)
                    df_surt['COMBUSTIBLE'] = df_surt['COMBUSTIBLE'].astype(str).str.strip().str.upper()
                    df_surt['GRUPO'] = df_surt['GRUPO'].astype(str).str.strip().str.upper()
                    df_surt['TIPO_SURTIDO'] = df_surt['TIPO_SURTIDO'].astype(str).str.strip().str.upper()

                    # Renombrar grupos para mayor claridad
                    df_surt['GRUPO'] = df_surt['GRUPO'].replace({
                        'RUTA CORTA': 'ORIENTE (RUTA CORTA)',
                        'RUTA CENTRO': 'CENTRO',
                        'RUTA OCCIDENTE': 'OCCIDENTE'
                    })

                    # --- KPIs GENERALES ---
                    total_gasolina = df_surt[df_surt['COMBUSTIBLE'] == 'GASOLINA']['LITROS'].sum()
                    total_gasoil = df_surt[df_surt['COMBUSTIBLE'] == 'GASOIL']['LITROS'].sum()
                    gran_total_surtido = total_gasolina + total_gasoil

                    fecha_inicio_surt = df_surt['Fecha_DT'].min().strftime('%d/%m/%Y')
                    fecha_fin_surt = df_surt['Fecha_DT'].max().strftime('%d/%m/%Y')

                    # --- SEPARAR OPERATIVO VS EXTRACCIÓN INTERNA ---
                    df_operativo = df_surt[df_surt['GRUPO'] != 'EXTRACCION']
                    df_extraccion = df_surt[df_surt['GRUPO'] == 'EXTRACCION']

                    total_operativo = df_operativo['LITROS'].sum()
                    total_extraccion = df_extraccion['LITROS'].sum()

                    # 1. TABLA: RUTAS (Operativo)
                    df_rutas = df_operativo.groupby('GRUPO')['LITROS'].sum().reset_index()
                    filas_rutas_html = ""
                    for _, r in df_rutas.iterrows():
                        pct = (r['LITROS'] / total_operativo * 100) if total_operativo > 0 else 0
                        filas_rutas_html += f"""
                        <tr style="background: white;">
                            <td style="padding: 10px; border: 2px solid #000; font-weight: bold;">{r['GRUPO']}</td>
                            <td style="padding: 10px; border: 2px solid #000; font-weight: 900; text-align: center; color: #0d47a1;">{f_p(r['LITROS'])} L</td>
                            <td style="padding: 10px; border: 2px solid #000; text-align: center; font-weight: bold;">{pct:.1f}%</td>
                        </tr>
                        """

                    # 2. TABLA: ORIGEN DEL SURTIDO (Operativo)
                    df_origen = df_operativo.groupby('TIPO_SURTIDO')['LITROS'].sum().reset_index()
                    filas_origen_html = ""
                    for _, r in df_origen.iterrows():
                        pct = (r['LITROS'] / total_operativo * 100) if total_operativo > 0 else 0
                        filas_origen_html += f"""
                        <tr style="background: white;">
                            <td style="padding: 10px; border: 2px solid #000; font-weight: bold;">{r['TIPO_SURTIDO']}</td>
                            <td style="padding: 10px; border: 2px solid #000; font-weight: 900; text-align: center; color: #2e7d32;">{f_p(r['LITROS'])} L</td>
                            <td style="padding: 10px; border: 2px solid #000; text-align: center; font-weight: bold;">{pct:.1f}%</td>
                        </tr>
                        """

                    # 3. TABLA: EXTRACCIONES INTERNAS (Consumo de la Base)
                    # Agrupamos por Sitio/Tipo y Combustible para que quede súper claro
                    df_ext_g = df_extraccion.groupby(['TIPO_SURTIDO', 'COMBUSTIBLE'])['LITROS'].sum().reset_index()
                    filas_ext_html = ""
                    for _, r in df_ext_g.iterrows():
                        pct = (r['LITROS'] / total_extraccion * 100) if total_extraccion > 0 else 0
                        nombre_ext = f"CIUDAD DROTACA ({r['COMBUSTIBLE']})" if r['TIPO_SURTIDO'] == 'TANQUE RESERVA' else f"BIDÓN 2.0 ({r['COMBUSTIBLE']})"
                        filas_ext_html += f"""
                        <tr style="background: white;">
                            <td style="padding: 10px; border: 2px solid #000; font-weight: bold;">{nombre_ext}</td>
                            <td style="padding: 10px; border: 2px solid #000; font-weight: 900; text-align: center; color: #c62828;">{f_p(r['LITROS'])} L</td>
                            <td style="padding: 10px; border: 2px solid #000; text-align: center; font-weight: bold;">{pct:.1f}%</td>
                        </tr>
                        """
                    if filas_ext_html == "":
                        filas_ext_html = "<tr><td colspan='3' style='padding:10px; border:2px solid #000; text-align:center;'>Sin extracciones en la semana</td></tr>"

                    # --- ESTRUCTURA HTML (SIMPLE Y AL PUNTO) ---
                    html_pizarra_surtido = f"""
                    <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head>
                    <body style="font-family: Arial, sans-serif; background-color: #f0f2f6; padding: 20px;">
                    <div style="text-align: center; margin-bottom: 15px;">
                        <button onclick="capSurtido()" style="background: #000; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR REPORTE</button>
                    </div>
                    <div id="piz-surtido" style="background:white; width:1000px; margin:auto; border:3px solid #000; overflow:hidden;">
                        
                        <div style="background-color: #000; color: white; padding: 25px 30px; display: flex; align-items: center; justify-content: space-between; border-bottom: 5px solid #d4af37;">
                            <img src="{logo_b64}" style="height: 50px;">
                            <div style="text-align: right;">
                                <h2 style="margin:0; font-size:24px; font-weight: 900; text-transform: uppercase;">REPORTE DE SURTIDO Y EXTRACCIÓN</h2>
                                <p style="margin:5px 0 0 0; font-size:14px; font-weight:bold; color:#d4af37;">SEMANA {int(num_sem)} ({fecha_inicio_surt} AL {fecha_fin_surt})</p>
                            </div>
                        </div>

                        <div style="padding: 20px;">
                            <div style="display: flex; justify-content: space-between; border: 2px solid #000; background: #eee; margin-bottom: 25px;">
                                <div style="padding: 15px; width: 33%; text-align: center; border-right: 2px solid #000;">
                                    <span style="font-size: 13px; font-weight: bold; color: #333;">⛽ TOTAL GASOLINA</span>
                                    <h2 style="margin: 5px 0 0 0; font-size: 26px; font-weight: 900; color: #000;">{f_p(total_gasolina)} L</h2>
                                </div>
                                <div style="padding: 15px; width: 33%; text-align: center; border-right: 2px solid #000;">
                                    <span style="font-size: 13px; font-weight: bold; color: #333;">🛢️ TOTAL GASOIL</span>
                                    <h2 style="margin: 5px 0 0 0; font-size: 26px; font-weight: 900; color: #000;">{f_p(total_gasoil)} L</h2>
                                </div>
                                <div style="padding: 15px; width: 33%; text-align: center; background: #e0f2f1;">
                                    <span style="font-size: 13px; font-weight: bold; color: #004d40;">📊 GRAN TOTAL MOVILIZADO</span>
                                    <h2 style="margin: 5px 0 0 0; font-size: 26px; font-weight: 900; color: #004d40;">{f_p(gran_total_surtido)} L</h2>
                                </div>
                            </div>

                            <div style="display: flex; justify-content: space-between;">
                                
                                <div style="width: 32%;">
                                    <h3 style="margin: 0 0 10px 0; font-size: 14px; background: #006064; color: white; padding: 8px; text-align: center; border: 2px solid #000; text-transform: uppercase;">1. DESPACHO POR RUTAS</h3>
                                    <table style="width: 100%; border-collapse: collapse; font-size: 12px; border: 2px solid #000;">
                                        <thead>
                                            <tr style="background: #e0f7fa;">
                                                <th style="padding: 8px; border: 2px solid #000; text-align: left;">RUTAS</th>
                                                <th style="padding: 8px; border: 2px solid #000;">LITROS</th>
                                                <th style="padding: 8px; border: 2px solid #000;">%</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filas_rutas_html}
                                            <tr style="background: #b2ebf2;">
                                                <td style="padding: 8px; border: 2px solid #000; font-weight: bold; text-align: right;">TOTAL:</td>
                                                <td style="padding: 8px; border: 2px solid #000; font-weight: 900; text-align: center;">{f_p(total_operativo)} L</td>
                                                <td style="padding: 8px; border: 2px solid #000;">100%</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>

                                <div style="width: 32%;">
                                    <h3 style="margin: 0 0 10px 0; font-size: 14px; background: #2e7d32; color: white; padding: 8px; text-align: center; border: 2px solid #000; text-transform: uppercase;">2. TIPO DE SURTIDO (FLOTA)</h3>
                                    <table style="width: 100%; border-collapse: collapse; font-size: 12px; border: 2px solid #000;">
                                        <thead>
                                            <tr style="background: #e8f5e9;">
                                                <th style="padding: 8px; border: 2px solid #000; text-align: left;">MÉTODO</th>
                                                <th style="padding: 8px; border: 2px solid #000;">LITROS</th>
                                                <th style="padding: 8px; border: 2px solid #000;">%</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filas_origen_html}
                                            <tr style="background: #c8e6c9;">
                                                <td style="padding: 8px; border: 2px solid #000; font-weight: bold; text-align: right;">TOTAL:</td>
                                                <td style="padding: 8px; border: 2px solid #000; font-weight: 900; text-align: center;">{f_p(total_operativo)} L</td>
                                                <td style="padding: 8px; border: 2px solid #000;">100%</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>

                                <div style="width: 32%;">
                                    <h3 style="margin: 0 0 10px 0; font-size: 14px; background: #c62828; color: white; padding: 8px; text-align: center; border: 2px solid #000; text-transform: uppercase;">3. EXTRACCIÓN (INTERNA)</h3>
                                    <table style="width: 100%; border-collapse: collapse; font-size: 12px; border: 2px solid #000;">
                                        <thead>
                                            <tr style="background: #ffebee;">
                                                <th style="padding: 8px; border: 2px solid #000; text-align: left;">ORIGEN</th>
                                                <th style="padding: 8px; border: 2px solid #000;">LITROS</th>
                                                <th style="padding: 8px; border: 2px solid #000;">%</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filas_ext_html}
                                            <tr style="background: #ffcdd2;">
                                                <td style="padding: 8px; border: 2px solid #000; font-weight: bold; text-align: right;">TOTAL:</td>
                                                <td style="padding: 8px; border: 2px solid #000; font-weight: 900; text-align: center;">{f_p(total_extraccion)} L</td>
                                                <td style="padding: 8px; border: 2px solid #000;">100%</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>

                            </div>
                        </div>
                    </div>
                    <script>
                    function capSurtido() {{ 
                        html2canvas(document.getElementById('piz-surtido'), {{scale: 2}}).then(canvas => {{ 
                            var link = document.createElement('a'); 
                            link.download = 'Resumen_Surtido_Semana_{int(num_sem)}.png'; 
                            link.href = canvas.toDataURL(); 
                            link.click(); 
                        }}); 
                    }}
                    </script>
                    </body></html>
                    """
                    components.html(html_pizarra_surtido, height=550, scrolling=True)

                    # --- WHATSAPP (CORTO Y DIRECTO) ---
                    st.markdown("---")
                    st.subheader("📱 Resumen Ejecutivo para WhatsApp")
                    
                    msg_surt = f"⛽ *RESULTADOS DE SURTIDO Y EXTRACCIÓN*\n📅 Semana: {int(num_sem)} ({fecha_inicio_surt} al {fecha_fin_surt})\n\n"
                    msg_surt += f"📊 *Total Movilizado:* {f_p(gran_total_surtido)} Litros\n"
                    msg_surt += f"   ⛽ Gasolina: {f_p(total_gasolina)} L\n"
                    msg_surt += f"   🛢️ Gasoil: {f_p(total_gasoil)} L\n\n"
                    
                    msg_surt += f"*📍 1. DESPACHO POR RUTAS (Flota)*\n"
                    for _, r in df_rutas.iterrows():
                        msg_surt += f"▪️ {r['GRUPO']}: *{f_p(r['LITROS'])} L*\n"
                        
                    msg_surt += f"\n*🏪 2. MÉTODO DE SURTIDO (Flota)*\n"
                    for _, r in df_origen.iterrows():
                        msg_surt += f"▪️ {r['TIPO_SURTIDO'].title()}: *{f_p(r['LITROS'])} L*\n"
                        
                    msg_surt += f"\n*🚫 3. EXTRACCIÓN INTERNA (Auto-consumo)*\n"
                    msg_surt += f"▪️ Total extraído de la base: *{f_p(total_extraccion)} L*\n\n"
                    
                    msg_surt += "✅ *Pizarra simplificada adjunta.*"
                    st.code(msg_surt, language="markdown")
