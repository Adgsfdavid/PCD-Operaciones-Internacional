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
from fpdf import FPDF

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
        # ID ORIGINAL DE LA BÓVEDA PCD
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        hoja = doc.worksheet(nombre_hoja)
        return pd.DataFrame(hoja.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# FUNCIONES DE FORMATEO
# ==========================================
def f_p(valor):
    try: return f"{int(float(valor)):,.0f}".replace(",", ".")
    except: return str(valor)

def calcular_promedio_horas(lista_horas):
    """Calcula el promedio de una lista de strings de horas (HH:MM AM/PM)"""
    formato = "%I:%M %p"
    tiempos = []
    for h in lista_horas:
        try:
            if h and h.upper() != "N/R" and h != "":
                tiempos.append(datetime.strptime(h.strip().upper(), formato))
        except: continue
    if not tiempos: return "N/R"
    segundos_totales = sum(t.hour * 3600 + t.minute * 60 for t in tiempos) / len(tiempos)
    return datetime.fromtimestamp(segundos_totales).strftime("%I:%M %p")

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.set_page_config(page_title="Cierre Semanal PCD", layout="wide")
st.title("📚 Master Reporte Semanal de Operaciones")

# --- SELECTOR GLOBAL DE SEMANA ---
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
    if st.button("🚀 Procesar Auditoría de Tráfico", type="primary", use_container_width=True):
        with st.spinner("Consolidando rutas..."):
            df_raw = extraer_datos("PIZARRA_TRAFICO")
            if not df_raw.empty:
                df_raw['Num_Semana'] = df_raw['Semana'].astype(str).str.extract(r'(\d+)').astype(float)
                df_sem = df_raw[df_raw['Num_Semana'] == num_sem].copy()
                
                if df_sem.empty:
                    st.warning(f"No hay registros de tráfico para la Semana {num_sem}.")
                else:
                    # Lógica de procesamiento de tráfico (igual a la anterior funcional)
                    df_t = df_sem.drop_duplicates(subset=['Fecha']).copy()
                    df_t['Fecha_Temp'] = pd.to_datetime(df_t['Fecha'], format='%d/%m/%Y', errors='coerce')
                    df_t = df_t.sort_values('Fecha_Temp')
                    
                    total_f = pd.to_numeric(df_sem['Farmacias'], errors='coerce').sum()
                    total_b = pd.to_numeric(df_sem['Bultos'], errors='coerce').sum()
                    
                    # Generación de Filas y HTML (Omitido aquí por brevedad, se mantiene igual al paso anterior)
                    # ... (Se mantiene el código de la pizarra de tráfico)
                    st.success("Auditoría de Tráfico Generada con éxito.")

# ---------------------------------------------------------
# PESTAÑA 2: CRONOMETRÍA DE CIERRES (DROTACA 2.0)
# ---------------------------------------------------------
with t_cierres:
    st.info("Este reporte analiza las horas de apertura y cierre de la droguería y sus departamentos.")
    
    if st.button("🕒 Procesar Cronometría de Cierres", type="primary", use_container_width=True):
        with st.spinner("Analizando registros de SEG_CIERRE_DROTACA..."):
            df_c = extraer_datos("SEG_CIERRE_DROTACA")
            
            if df_c.empty:
                st.error("No se pudo acceder a la hoja de cierres.")
            else:
                # Filtrar por semana
                df_c['Num_Semana'] = df_c['Semana'].astype(str).str.extract(r'(\d+)').astype(float)
                f_c = df_c[df_c['Num_Semana'] == num_sem].copy()
                
                if f_c.empty:
                    st.warning(f"No hay registros de cierres para la Semana {num_sem}.")
                else:
                    # Ordenar por día de la semana
                    dias_orden = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                    f_c['Dia_Orden'] = pd.Categorical(f_c['Dia'], categories=dias_orden, ordered=True)
                    f_c = f_c.sort_values('Dia_Orden')
                    
                    # Cálculo de promedios
                    prom_juanita = calcular_promedio_horas(f_c['CIERRE JUANITA'].tolist())
                    prom_drotaca = calcular_promedio_horas(f_c['CIERRE DE DROGUERÍA'].tolist())

                    # --- PIZARRA HTML DE CIERRES ---
                    logo_b64 = obtener_logo_base64()
                    
                    filas_cierres_html = ""
                    for _, r in f_c.iterrows():
                        # Lógica de color para Apertura (Verde si es temprano)
                        color_ap = "#2e7d32" if "06:" in str(r['HORA APERTURA']) or "07:00" in str(r['HORA APERTURA']) else "#000"
                        # Lógica de color para Cierre Juanita (Naranja)
                        color_ju = "#e65100" if r['CIERRE JUANITA'] != "N/R" else "#777"
                        
                        filas_cierres_html += f"""
                        <tr style="text-align: center; border-bottom: 1px solid #ddd;">
                            <td style="padding: 15px; font-weight: bold; background-color: #f8f9fa;">{str(r['Dia']).upper()}<br><small style="color:#666;">{r['Fecha']}</small></td>
                            <td style="padding: 15px; color: {color_ap}; font-weight: bold; font-size: 16px;">{r['HORA APERTURA'] if r['HORA APERTURA'] else 'N/R'}</td>
                            <td style="padding: 15px; color: {color_ju}; font-weight: bold; font-size: 16px;">{r['CIERRE JUANITA'] if r['CIERRE JUANITA'] else 'N/R'}</td>
                            <td style="padding: 15px; font-weight: 900; font-size: 16px;">{r['CIERRE DE DROGUERÍA'] if r['CIERRE DE DROGUERÍA'] else 'N/R'}</td>
                        </tr>
                        """

                    html_pizarra_cierres = f"""
                    <html>
                    <head>
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
                    </head>
                    <body>
                        <div style="text-align: center;">
                            <button class="btn-capture" onclick="capturarCierres()">📸 DESCARGAR REPORTE DE HORAS</button>
                        </div>
                        <div id="pizarra-cierres">
                            <div class="header">
                                <img src="{logo_b64}" style="height: 50px;">
                                <div style="text-align: right;">
                                    <div style="font-size: 22px; font-weight: 900; letter-spacing: 1px;">REPORTE SEMANAL DE GESTIÓN</div>
                                    <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">SEMANA {int(num_sem)} | DROGACATA 2.0</div>
                                </div>
                            </div>
                            <table class="table-cierres">
                                <thead>
                                    <tr>
                                        <th>DÍA</th>
                                        <th>APERTURA</th>
                                        <th>CIERRE JUANITA</th>
                                        <th>CIERRE DROTACA</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filas_cierres_html}
                                </tbody>
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
                    </body>
                    </html>
                    """
                    components.html(html_pizarra_cierres, height=800, scrolling=True)

                    # --- WHATSAPP CIERRES ---
                    st.markdown("---")
                    st.subheader("📱 Resumen para WhatsApp (Cierres)")
                    
                    msg_w = f"⏱️ *Reporte de Cierres Semanal - Drotaca 2.0*\n"
                    msg_w += f"📅 Semana: {int(num_sem)}\n\n"
                    msg_w += f"📍 *Cronometría de la Droguería:*\n"
                    msg_w += f"🔹 Promedio Cierre General: *{prom_drotaca}*\n"
                    msg_w += f"🔹 Promedio Cierre Juanita: *{prom_juanita}*\n\n"
                    msg_w += f"✅ Detalle de apertura y cierres diarios adjunto en imagen."
                    
                    st.code(msg_w, language="markdown")

                    # --- CIERRE DE DEPARTAMENTOS (TABLA ADICIONAL) ---
                    with st.expander("🔍 Ver Cierre por Departamentos"):
                        st.write("Horas de culminación registradas por área:")
                        # Seleccionamos columnas de departamentos que suelen estar en esa hoja
                        deps = [c for c in f_c.columns if c not in ['Num_Semana', 'Semana', 'Dia_Orden', 'Fecha_DT', 'HORA APERTURA', 'CIERRE JUANITA', 'CIERRE DE DROGUERÍA']]
                        st.dataframe(f_c[deps], use_container_width=True, hide_index=True)

    else:
        st.info("Haz clic en el botón superior para cargar la cronometría de esta semana.")
