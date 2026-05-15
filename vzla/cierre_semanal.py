# ==========================================
# Archivo: cierre_semanal.py (Master Semanal Multi-Reporte)
# ==========================================
import streamlit as st
import pandas as pd
import base64
import textwrap
import unicodedata
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
    if hora_limpia == "N/R" or hora_limpia.lower() == "nan": return "N/R"
    try:
        if "am" in hora_limpia.lower() or "pm" in hora_limpia.lower():
            return hora_limpia.upper()
        return datetime.strptime(hora_limpia, "%H:%M").strftime("%I:%M %p").upper()
    except: return hora_limpia

def buscar_columna_estricta(df, palabras_clave, evitar=None):
    if df.empty: return None
    evitar = evitar or []
    # Primero buscamos coincidencias exactas para casos como *CIERRE DE DROGUERÍA*
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
            if h_clean != "N/R" and h_clean != "":
                tiempos.append(datetime.strptime(h_clean, formato))
        except: continue
    
    if not tiempos: return "N/R"
    
    # Promedio matemático de los segundos transcurridos en el día
    segundos_totales = sum(t.hour * 3600 + t.minute * 60 for t in tiempos) / len(tiempos)
    horas = int(segundos_totales // 3600)
    minutos = int((segundos_totales % 3600) // 60)
    
    # Convertir de vuelta a formato 12h
    temp_dt = datetime(2026, 1, 1, horas, minutos)
    return temp_dt.strftime("%I:%M %p").upper()

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

t_trafico, t_cierres = st.tabs(["📈 Desempeño de Tráfico", "⏱️ Cronometría de Cierres"])

# --- PESTAÑA 1 (TRÁFICO) SE MANTIENE IGUAL ---
with t_trafico:
    # (Lógica de tráfico omitida para no alargar el mensaje, se mantiene la funcional)
    st.info("Consolida la data de despachos diarios en una pizarra semanal.")
    # ... código de tráfico ...

# ---------------------------------------------------------
# PESTAÑA 2: CRONOMETRÍA DE CIERRES (REDISEÑO MATRIZ)
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
                # SOLO LUNES A VIERNES
                dias_base = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"]
                dias_display = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
                
                df_resumen = pd.DataFrame({"Dia_Norm": dias_base, "Día": dias_display})
                df_resumen['Apertura'] = "N/R"
                df_resumen['Juanita'] = "N/R"
                df_resumen['Drotaca'] = "N/R"
                df_resumen['Fecha'] = ""

                # 1. APERTURA
                if not df_a_raw.empty:
                    df_a_raw['Num_Semana'] = df_a_raw[buscar_columna_estricta(df_a_raw, ['semana'])].astype(str).str.extract(r'(\d+)').astype(float)
                    f_a = df_a_raw[df_a_raw['Num_Semana'] == num_sem].copy()
                    c_dia_a = buscar_columna_estricta(f_a, ['dia', 'día'], evitar=['fecha'])
                    c_hora_a = buscar_columna_estricta(f_a, ['apertura', 'hora'], evitar=['fecha', 'dia'])
                    if c_dia_a and c_hora_a:
                        f_a['Dia_Norm'] = f_a[c_dia_a].apply(norm_dia)
                        dict_a = f_a.groupby('Dia_Norm')[c_hora_a].last().to_dict()
                        df_resumen['Apertura'] = df_resumen['Dia_Norm'].map(dict_a).fillna("N/R")

                # 2. JUANITA
                if not df_j_raw.empty:
                    df_j_raw['Num_Semana'] = df_j_raw[buscar_columna_estricta(df_j_raw, ['semana'])].astype(str).str.extract(r'(\d+)').astype(float)
                    f_j = df_j_raw[df_j_raw['Num_Semana'] == num_sem].copy()
                    c_dia_j = buscar_columna_estricta(f_j, ['dia', 'día'], evitar=['fecha'])
                    c_hora_j = buscar_columna_estricta(f_j, ['juanita', 'hora', 'cierre'], evitar=['fecha', 'dia'])
                    if c_dia_j and c_hora_j:
                        f_j['Dia_Norm'] = f_j[c_dia_j].apply(norm_dia)
                        dict_j = f_j.groupby('Dia_Norm')[c_hora_j].last().to_dict()
                        df_resumen['Juanita'] = df_resumen['Dia_Norm'].map(dict_j).fillna("N/R")

                # 3. DROTACA (Búsqueda específica de *CIERRE DE DROGUERÍA*)
                pivot_deps = pd.DataFrame()
                if not df_d_raw.empty:
                    df_d_raw['Num_Semana'] = df_d_raw[buscar_columna_estricta(df_d_raw, ['semana'])].astype(str).str.extract(r'(\d+)').astype(float)
                    f_d = df_d_raw[df_d_raw['Num_Semana'] == num_sem].copy()
                    
                    if not f_d.empty:
                        c_dia_d = buscar_columna_estricta(f_d, ['dia', 'día'], evitar=['fecha'])
                        c_fecha_d = buscar_columna_estricta(f_d, ['fecha'])
                        # Ajuste clave: busca la columna con asteriscos
                        c_drog = buscar_columna_estricta(f_d, ['cierre de drogueria', 'cierre de droguería', 'drogueria'])
                        c_dep = buscar_columna_estricta(f_d, ['departamento', 'area'], evitar=['fecha', 'hora'])
                        c_hora_sal = buscar_columna_estricta(f_d, ['hora salida', 'salida'], evitar=['fecha', 'drogueria'])
                        
                        f_d['Dia_Norm'] = f_d[c_dia_d].apply(norm_dia) if c_dia_d else ""
                        
                        if c_drog:
                            dict_d = f_d.dropna(subset=[c_drog]).groupby('Dia_Norm')[c_drog].last().to_dict()
                            df_resumen['Drotaca'] = df_resumen['Dia_Norm'].map(dict_d).fillna("N/R")
                        
                        if c_fecha_d:
                            dict_f = f_d.dropna(subset=[c_fecha_d]).groupby('Dia_Norm')[c_fecha_d].last().to_dict()
                            df_resumen['Fecha'] = df_resumen['Dia_Norm'].map(dict_f).fillna("")

                        # MATRIZ DE DEPARTAMENTOS
                        if c_dep and c_hora_sal:
                            df_deps = f_d.dropna(subset=[c_dep, c_hora_sal]).copy()
                            df_deps = df_deps[~df_deps[c_dep].str.upper().str.contains("CIERRE DROTACA|CIERRE GENERAL", na=False)]
                            if not df_deps.empty:
                                pivot_deps = df_deps.pivot_table(index=c_dep, columns='Dia_Norm', values=c_hora_sal, aggfunc='last')
                                for d in dias_base:
                                    if d not in pivot_deps.columns: pivot_deps[d] = "N/R"
                                pivot_deps = pivot_deps[dias_base].fillna("N/R")
                                pivot_deps['Promedio'] = pivot_deps.apply(lambda row: calcular_promedio_horas(row.tolist()), axis=1)

                # ==========================================
                # HTML RENDERING
                # ==========================================
                prom_juanita = calcular_promedio_horas(df_resumen['Juanita'].tolist())
                prom_drotaca = calcular_promedio_horas(df_resumen['Drotaca'].tolist())
                logo_b64 = obtener_logo_base64()

                filas_gral = ""
                for _, r in df_resumen.iterrows():
                    h_ap = a_12h(r['Apertura'])
                    h_ju = a_12h(r['Juanita'])
                    h_dr = a_12h(r['Drotaca'])
                    filas_gral += f"""
                    <tr style="text-align: center; border-bottom: 1px solid #ddd;">
                        <td style="padding: 15px; font-weight: bold; background-color: #f8f9fa;">{r['Día'].upper()}<br><small style="color:#666;">{r['Fecha']}</small></td>
                        <td style="padding: 15px; color: #2e7d32; font-weight: bold; font-size: 16px;">{h_ap}</td>
                        <td style="padding: 15px; color: #e65100; font-weight: bold; font-size: 16px;">{h_ju}</td>
                        <td style="padding: 15px; font-weight: 900; font-size: 16px;">{h_dr}</td>
                    </tr>
                    """

                html_general = f"""
                <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                <style>@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                body {{ font-family: 'Montserrat', sans-serif; padding: 20px; background-color: #f0f2f6; }}
                .pizarra {{ background: white; width: 900px; margin: auto; border-radius: 15px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: 2px solid #0d47a1; margin-bottom: 40px; }}
                .header {{ background: #0d47a1; color: white; padding: 30px; display: flex; justify-content: space-between; align-items: center; }}
                table {{ width: 100%; border-collapse: collapse; }}
                th {{ background: #f1f4f9; color: #0d47a1; padding: 15px; text-transform: uppercase; font-size: 12px; border-bottom: 2px solid #0d47a1; }}
                .footer-promedios {{ background: #f1f4f9; padding: 20px; display: flex; justify-content: space-around; border-top: 2px solid #0d47a1; }}
                .promedio-val {{ font-size: 20px; font-weight: 900; color: #0d47a1; }}
                </style></head><body>
                <div style="text-align: center; margin-bottom: 15px;">
                    <button onclick="capturar('pizarra-general', 'Cierres_Generales_Semana_{int(num_sem)}.png')" style="background: #2e7d32; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR REPORTE GENERAL</button>
                </div>
                <div class="pizarra" id="pizarra-general">
                    <div class="header"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><div style="font-size: 22px; font-weight: 900;">REPORTE SEMANAL DE GESTIÓN</div><div style="font-size: 14px; font-weight: bold;">SEMANA {int(num_sem)} | CIERRES GENERALES</div></div></div>
                    <table><thead><tr><th>DÍA</th><th>APERTURA DROTACA</th><th>CIERRE JUANITA</th><th>CIERRE DROTACA</th></tr></thead><tbody>{filas_gral}</tbody></table>
                    <div class="footer-promedios"><div style="text-align:center;"><b>Promedio Cierre Drotaca:</b><br><span class="promedio-val">{prom_drotaca}</span></div><div style="text-align:center;"><b>Promedio Cierre Juanita:</b><br><span class="promedio-val" style="color:#e65100;">{prom_juanita}</span></div></div>
                </div>
                """

                filas_deps = ""
                if not pivot_deps.empty:
                    for dep, row in pivot_deps.iterrows():
                        tds = "".join([f"<td style='padding:10px; border:1px solid #ddd; font-weight:bold;'>{a_12h(row[d])}</td>" for d in dias_base])
                        filas_deps += f"<tr style='text-align:center;'><td style='text-align:left; padding:10px; font-weight:900; background:#f8f9fa; color:#0d47a1;'>{str(dep).upper()}</td>{tds}<td style='background:#ffebee; color:#d32f2f; font-weight:900;'>{row['Promedio']}</td></tr>"

                html_deps = f"""
                <div style="text-align: center; margin-bottom: 15px;"><button onclick="capturar('pizarra-deps', 'Departamentos_Semana_{int(num_sem)}.png')" style="background: #e65100; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR DEPARTAMENTOS</button></div>
                <div class="pizarra" id="pizarra-deps" style="width: 1050px;">
                    <div class="header" style="background:#1565c0;"><img src="{logo_b64}" style="height: 50px;"><div style="text-align: right;"><div style="font-size: 22px; font-weight: 900;">MATRIZ DE DEPARTAMENTOS</div><div style="font-size: 14px; font-weight: bold;">HORARIOS DE SALIDA</div></div></div>
                    <table><thead><tr><th style='text-align:left;'>DEPARTAMENTO</th><th>LUNES</th><th>MARTES</th><th>MIÉRCOLES</th><th>JUEVES</th><th>VIERNES</th><th style='background:#ffcdd2;'>PROMEDIO</th></tr></thead><tbody>{filas_deps}</tbody></table>
                </div>
                <script>function capturar(id, filename) {{ html2canvas(document.getElementById(id), {{ scale: 2 }}).then(canvas => {{ var link = document.createElement('a'); link.download = filename; link.href = canvas.toDataURL(); link.click(); }}); }}</script>
                </body></html>
                """

                components.html(html_general + html_deps, height=1600, scrolling=True)

                st.markdown("---")
                st.subheader("📱 WhatsApp Cierres")
                msg_w = f"⏱️ *Reporte de Cierres Semanal - Drotaca 2.0*\n📅 Semana: {int(num_sem)}\n\n"
                msg_w += f"📍 *Promedios de la Semana:*\n🔹 Cierre General: *{prom_drotaca}*\n🔹 Cierre Juanita: *{prom_juanita}*\n\n✅ Imágenes con detalle por departamento adjuntas."
                st.code(msg_w, language="markdown")
