# ==========================================
# Archivo: cierre_semanal.py (Master Semanal Multi-Reporte)
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
# FUNCIONES DE FORMATEO Y BÚSQUEDA INTELIGENTE
# ==========================================
def f_p(valor):
    try: return f"{int(float(valor)):,.0f}".replace(",", ".")
    except: return str(valor)

def limpiar_hora(hora_str):
    if not hora_str: return ""
    return str(hora_str).replace("*", "").strip()

def a_12h(hora_24):
    hora_limpia = limpiar_hora(hora_24)
    try:
        if "am" in hora_limpia.lower() or "pm" in hora_limpia.lower():
            return hora_limpia.lower()
        return datetime.strptime(hora_limpia, "%H:%M").strftime("%I:%M %p").lower()
    except: return hora_limpia

def buscar_columna(df, palabras_clave):
    if df.empty: return None
    for col in df.columns:
        if any(p.lower() in str(col).lower() for p in palabras_clave):
            return col
    return df.columns[0] 

def calcular_rango_semana(ano, semana):
    lunes = datetime.strptime(f'{int(ano)}-W{int(semana)}-1', "%G-W%V-%u")
    domingo = lunes + timedelta(days=6)
    return f"{lunes.strftime('%d/%m/%Y')} al {domingo.strftime('%d/%m/%Y')}"

def calcular_promedio_horas(lista_horas):
    formato = "%I:%M %p"
    tiempos = []
    for h in lista_horas:
        try:
            if pd.notna(h) and str(h).strip().upper() != "N/R" and str(h).strip() != "":
                tiempos.append(datetime.strptime(str(h).strip().upper(), formato))
        except: continue
    if not tiempos: return "N/R"
    segundos_totales = sum(t.hour * 3600 + t.minute * 60 for t in tiempos) / len(tiempos)
    return datetime.fromtimestamp(segundos_totales).strftime("%I:%M %p")

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

# ---------------------------------------------------------
# PESTAÑA 1: DESEMPEÑO DE TRÁFICO
# ---------------------------------------------------------
with t_trafico:
    st.info("Consolida la data de despachos diarios en una pizarra semanal.")
    if st.button("🚀 Generar Auditoría de Tráfico", type="primary", use_container_width=True):
        with st.spinner("Consolidando rutas de tráfico..."):
            df_raw = extraer_datos("PIZARRA_TRAFICO")
            if df_raw.empty:
                st.error("Error al conectar con la base de datos o la hoja está vacía.")
            else:
                rango_fechas = calcular_rango_semana(ano_sel, num_sem)
                
                c_sem = buscar_columna(df_raw, ['semana'])
                df_raw['Num_Semana'] = df_raw[c_sem].astype(str).str.extract(r'(\d+)').astype(float)
                df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()

                if df_sem.empty:
                    st.warning(f"No hay registros de tráfico para la Semana {int(num_sem)}.")
                else:
                    c_fecha = buscar_columna(df_sem, ['fecha'])
                    c_dia = buscar_columna(df_sem, ['dia', 'día'])
                    c_h1 = buscar_columna(df_sem, ['1er'])
                    c_hu = buscar_columna(df_sem, ['ultimo', 'último'])
                    c_it = buscar_columna(df_sem, ['inicio'])
                    c_ct = buscar_columna(df_sem, ['culminacion', 'fin'])
                    c_ruta = buscar_columna(df_sem, ['ruta'])
                    c_zona = buscar_columna(df_sem, ['zona'])
                    c_unidad = buscar_columna(df_sem, ['unidad'])
                    c_farma = buscar_columna(df_sem, ['farmacia']) # Búsqueda inteligente
                    c_bultos = buscar_columna(df_sem, ['bulto'])

                    df_t = df_sem.drop_duplicates(subset=[c_fecha]).copy()
                    df_t['Fecha_Temp'] = pd.to_datetime(df_t[c_fecha], format='%d/%m/%Y', errors='coerce')
                    df_t = df_t.sort_values('Fecha_Temp')

                    # Conversión segura a numérico
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
                        <div style="text-align:center; padding:20px;">
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
# PESTAÑA 2: CRONOMETRÍA DE CIERRES (DROTACA 2.0 & JUANITA)
# ---------------------------------------------------------
with t_cierres:
    st.info("Cruza la data de SEG_CIERRE_DROTACA y SEG_CIERRE_JUANITA para auditar los tiempos.")
    if st.button("🕒 Procesar Cronometría de Cierres", type="primary", use_container_width=True):
        with st.spinner("Analizando y cruzando registros de Cierres..."):
            # EXTRAEMOS AMBAS HOJAS
            df_d_raw = extraer_datos("SEG_CIERRE_DROTACA")
            df_j_raw = extraer_datos("SEG_CIERRE_JUANITA")
            
            if df_d_raw.empty:
                st.error("No se pudo acceder a la hoja SEG_CIERRE_DROTACA.")
            else:
                # Filtrar DROTACA
                c_sem_d = buscar_columna(df_d_raw, ['semana'])
                df_d_raw['Num_Semana'] = df_d_raw[c_sem_d].astype(str).str.extract(r'(\d+)').astype(float)
                f_d = df_d_raw[df_d_raw['Num_Semana'] == num_sem].copy()
                
                if f_d.empty:
                    st.warning(f"No hay registros en SEG_CIERRE_DROTACA para la Semana {int(num_sem)}.")
                else:
                    c_fecha_d = buscar_columna(f_d, ['fecha'])
                    c_dia_d = buscar_columna(f_d, ['dia', 'día'])
                    c_aper = buscar_columna(f_d, ['apertura'])
                    c_drog = buscar_columna(f_d, ['droguer', 'drotaca'])

                    # Filtrar JUANITA
                    f_j_clean = pd.DataFrame(columns=['Fecha_Join', 'Hora_Juanita'])
                    if not df_j_raw.empty:
                        c_sem_j = buscar_columna(df_j_raw, ['semana'])
                        df_j_raw['Num_Semana'] = df_j_raw[c_sem_j].astype(str).str.extract(r'(\d+)').astype(float)
                        f_j = df_j_raw[df_j_raw['Num_Semana'] == num_sem].copy()
                        if not f_j.empty:
                            c_fecha_j = buscar_columna(f_j, ['fecha'])
                            c_juan_hora = buscar_columna(f_j, ['cierre', 'hora', 'juanita'])
                            f_j_clean = f_j[[c_fecha_j, c_juan_hora]].rename(columns={c_fecha_j: 'Fecha_Join', c_juan_hora: 'Hora_Juanita'})

                    # CRUZAR DATOS (MERGE por FECHA)
                    if not f_j_clean.empty and c_fecha_d:
                        f_c = pd.merge(f_d, f_j_clean, left_on=c_fecha_d, right_on='Fecha_Join', how='left')
                    else:
                        f_c = f_d.copy()
                        f_c['Hora_Juanita'] = "N/R"

                    # ORDENAR POR DÍAS
                    dias_orden = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                    f_c['Dia_Orden'] = pd.Categorical(f_c[c_dia_d].str.capitalize(), categories=dias_orden, ordered=True)
                    f_c = f_c.sort_values('Dia_Orden')
                    
                    # PROMEDIOS
                    prom_juanita = calcular_promedio_horas(f_c['Hora_Juanita'].tolist())
                    prom_drotaca = calcular_promedio_horas(f_c[c_drog].tolist()) if c_drog else "N/R"

                    logo_b64 = obtener_logo_base64()
                    
                    filas_cierres_html = ""
                    for _, r in f_c.iterrows():
                        hora_ap = str(r[c_aper]) if c_aper and pd.notna(r[c_aper]) else 'N/R'
                        hora_ju = str(r['Hora_Juanita']) if pd.notna(r['Hora_Juanita']) else 'N/R'
                        hora_dr = str(r[c_drog]) if c_drog and pd.notna(r[c_drog]) else 'N/R'
                        
                        color_ap = "#2e7d32" if "06:" in hora_ap or "07:00" in hora_ap else "#000"
                        color_ju = "#e65100" if hora_ju != "N/R" and hora_ju != "nan" else "#777"
                        
                        filas_cierres_html += f"""
                        <tr style="text-align: center; border-bottom: 1px solid #ddd;">
                            <td style="padding: 15px; font-weight: bold; background-color: #f8f9fa;">{str(r[c_dia_d]).upper()}<br><small style="color:#666;">{r[c_fecha_d] if c_fecha_d else ''}</small></td>
                            <td style="padding: 15px; color: {color_ap}; font-weight: bold; font-size: 16px;">{hora_ap}</td>
                            <td style="padding: 15px; color: {color_ju}; font-weight: bold; font-size: 16px;">{hora_ju if hora_ju != 'nan' else 'N/R'}</td>
                            <td style="padding: 15px; font-weight: 900; font-size: 16px;">{hora_dr}</td>
                        </tr>
                        """

                    html_pizarra_cierres = f"""
                    <html><head>
                        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                        <style>
                            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                            body {{ font-family: 'Montserrat', sans-serif; padding: 20px; background-color: #f0f2f6; }}
                            #pizarra-cierres {{ background: white; width: 900px; margin: auto; border-radius: 15px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: 2px solid #0d47a1; }}
                            .header {{ background: #0d47a1; color: white; padding: 30px; display: flex; justify-content: space-between; align-items: center; }}
                            .table-cierres {{ width: 100%; border-collapse: collapse; }}
                            .table-cierres th {{ background: #f1f4f9; color: #0d47a1; padding: 15px; text-transform: uppercase; font-size: 12px; border-bottom: 2px solid #0d47a1; }}
                            .footer-promedios {{ background: #f1f4f9; padding: 20px; display: flex; justify-content: space-around; border-top: 2px solid #0d47a1; }}
                            .promedio-box {{ text-align: center; }}
                            .promedio-label {{ font-size: 12px; font-weight: bold; color: #555; text-transform: uppercase; }}
                            .promedio-val {{ font-size: 20px; font-weight: 900; color: #0d47a1; }}
                            .btn-capture {{ background: #2e7d32; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer; margin-bottom: 15px; }}
                        </style>
                    </head><body>
                        <div style="text-align: center;">
                            <button class="btn-capture" onclick="capturarCierres()">📸 DESCARGAR REPORTE DE HORAS</button>
                        </div>
                        <div id="pizarra-cierres">
                            <div class="header">
                                <img src="{logo_b64}" style="height: 50px;">
                                <div style="text-align: right;">
                                    <div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">REPORTE SEMANAL DE GESTIÓN</div>
                                    <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} | DROTACA 2.0</div>
                                </div>
                            </div>
                            <table class="table-cierres">
                                <thead><tr><th>DÍA</th><th>APERTURA</th><th>CIERRE JUANITA</th><th>CIERRE DROTACA</th></tr></thead>
                                <tbody>{filas_cierres_html}</tbody>
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
                        <script>
                            function capturarCierres() {{
                                html2canvas(document.getElementById('pizarra-cierres'), {{ scale: 2 }}).then(canvas => {{
                                    var link = document.createElement('a');
                                    link.download = 'Gestion_Horas_Semana_{int(num_sem)}.png';
                                    link.href = canvas.toDataURL();
                                    link.click();
                                }});
                            }}
                        </script>
                    </body></html>
                    """
                    components.html(html_pizarra_cierres, height=800, scrolling=True)

                    st.markdown("---")
                    st.subheader("📱 Resumen para WhatsApp (Cierres)")
                    msg_w = f"⏱️ *Reporte de Cierres Semanal - Drotaca 2.0*\n📅 Semana: {int(num_sem)}\n\n"
                    msg_w += f"📍 *Cronometría de la Droguería:*\n🔹 Promedio Cierre General: *{prom_drotaca}*\n🔹 Promedio Cierre Juanita: *{prom_juanita}*\n\n✅ Detalle de apertura y cierres diarios adjunto en imagen."
                    st.code(msg_w, language="markdown")

                    with st.expander("🔍 Ver Cierre por Departamentos"):
                        st.write("Horas de culminación registradas por área:")
                        # Buscamos columnas extra (Departamentos) que no sean base
                        cols_basicas = [c_sem_d, 'Num_Semana', 'Dia_Orden', c_dia_d, c_fecha_d, c_aper, c_drog, 'Fecha_Join', 'Hora_Juanita']
                        deps = [c for c in f_c.columns if c not in cols_basicas]
                        if deps: st.dataframe(f_c[deps], use_container_width=True, hide_index=True)
                        else: st.info("No se encontraron columnas adicionales de departamentos.")
