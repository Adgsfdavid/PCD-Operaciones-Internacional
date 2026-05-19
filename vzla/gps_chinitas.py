# ==========================================
# Archivo: gps_chinitas.py (Análisis de Rutas Tracksolid)
# ==========================================
import streamlit as st
import pandas as pd
from geopy.distance import great_circle
from datetime import datetime, timedelta
import os
import json
import traceback
import streamlit.components.v1 as components
import gspread
import textwrap
from google.oauth2.service_account import Credentials

# ==========================================
# CONFIGURACIÓN DE CONEXIÓN A GOOGLE SHEETS
# ==========================================
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"
CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def guardar_en_googlesheets(datos_lista):
    cliente = obtener_cliente_sheets()
    doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
    sheet = doc.worksheet("Historial_GPS")
    for fila in datos_lista:
        sheet.append_row(fila)

# ==========================================
# CONSTANTES
# ==========================================
VELOCIDAD_MINIMA_MOVIMIENTO = 5
DISTANCIA_MAXIMA_METROS = 300
DESPACHOS_DB_FILE = "despachos_guardados.json"

# Placa A38CV4G (FORD RANGER) Eliminada
PLACAS_AUTORIZADAS = {
    'A36AC9X', 'A37CS2D', 'A38BA4N', 'A48AU5T',
    'A71EB8P', 'A72EB0P', 'A88BD0J', 'A84EZ6P', 'A87EZ8P',
    'AB893NB', 'A00DS2V', 'A73CL7D'
}

MASTER_VEHICULOS = {
    'A71EB8P': {'modelo': 'DFSK D1', 'color': 'BLANCO'},
    'A72EB0P': {'modelo': 'DFSK D1', 'color': 'BLANCO'},
    'A36AC9X': {'modelo': 'CHANGAN HUNTER 4X2', 'color': 'BLANCO'},
    'A37CS2D': {'modelo': 'CHANGAN HUNTER 4X2', 'color': 'VERDE'},
    'A88BD0J': {'modelo': 'CHANGAN HUNTER 4X2', 'color': 'BLANCO'},
    'A38BA4N': {'modelo': 'CHANGAN KAICENE F70', 'color': 'AZUL'},
    'A48AU5T': {'modelo': 'CHANGAN KAICENE F70', 'color': 'AZUL'},
    'A00DS2V': {'modelo': 'ENCAVA', 'color': 'BLANCO'},
    'AB893NB': {'modelo': 'MITSUBISHI LANCER', 'color': 'GRIS'},
    'A73CL7D': {'modelo': 'REY CAMION', 'color': 'PLATA'},
    'A84EZ6P': {'modelo': 'RICH P11', 'color': 'BLANCO'},
    'A87EZ8P': {'modelo': 'RICH P11', 'color': 'BLANCO'},
}

# ==========================================
# FUNCIONES DE CÁLCULO Y LIMPIEZA
# ==========================================
def _normalize_df_columns(df):
    df.columns = [str(col).strip().title().replace(" De ", " De ").replace(" Y ", " Y ").replace(" (Km)", " (Km)").replace(" (Km/H)", " (Km/H)") for col in df.columns]
    return df

def obtener_direccion_cardinal(lat_prev, lon_prev, lat_actual, lon_actual):
    TOLERANCIA = 0.00001
    d_lat = lat_actual - lat_prev
    d_lon = lon_actual - lon_prev
    if abs(d_lat) < TOLERANCIA and abs(d_lon) < TOLERANCIA: return None
    norte, sur, este, oeste = d_lat > 0, d_lat < 0, d_lon > 0, d_lon < 0
    if norte and este: return "NE"
    if sur and este: return "SE"
    if sur and oeste: return "SW"
    if norte and oeste: return "NW"
    if norte: return "N"
    if sur: return "S"
    if este: return "E"
    if oeste: return "W"
    return None

def encontrar_punto_mas_cercano(lat, lon, df_base_loc, distancia_maxima):
    if df_base_loc is None or df_base_loc.empty: return None
    coordenada_actual = (lat, lon)
    distancia_minima = float('inf')
    ubicacion_cercana = None
    for _, punto in df_base_loc.iterrows():
        try:
            lat_punto = float(str(punto.get('Latitud', 0)).replace(',', '.'))
            lon_punto = float(str(punto.get('Longitud', 0)).replace(',', '.'))
        except: continue
        distancia = great_circle(coordenada_actual, (lat_punto, lon_punto)).meters
        if distancia < distancia_minima:
            distancia_minima, ubicacion_cercana = distancia, punto.get('Localizacion', 'Ubicación sin nombre')
    if ubicacion_cercana and distancia_minima <= distancia_maxima:
        return ubicacion_cercana
    return None

def cargar_json_local(file_name):
    if os.path.exists(file_name):
        with open(file_name, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def guardar_json_local(file_name, data):
    with open(file_name, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def formatear_km(valor):
    try:
        return f"{int(float(valor)):,.0f}".replace(",", ".")
    except:
        return str(valor)

def obtener_logo_base64():
    try:
        from pathlib import Path
        ruta_logo = Path(__file__).parent / "logo.png"
        if ruta_logo.exists():
            with open(ruta_logo, "rb") as image_file:
                return f"data:image/png;base64,{base64.b64encode(image_file.read()).decode()}"
        return None
    except: return None

# ==========================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==========================================
st.set_page_config(page_title="Análisis GPS Chinitas", layout="wide")
st.title("🛰️ Análisis de Rutas y Paradas GPS (Tracksolid)")

# Inicializar estados
if 'despachos_guardados' not in st.session_state:
    st.session_state['despachos_guardados'] = cargar_json_local(DESPACHOS_DB_FILE)
if 'datos_resumen' not in st.session_state:
    st.session_state['datos_resumen'] = []
if 'reportes_texto' not in st.session_state:
    st.session_state['reportes_texto'] = {}
if 'chofer_defecto' not in st.session_state:
    st.session_state['chofer_defecto'] = ""

t_config, t_resumen, t_reportes, t_historico = st.tabs(["⚙️ Configuración", "📊 Pizarra Ejecutiva", "📝 Reportes Individuales", "💾 Guardar en Nube"])

# ---------------------------------------------------------
# PESTAÑA 1: CONFIGURACIÓN
# ---------------------------------------------------------
with t_config:
    st.subheader("1. Carga de Archivos Base")
    col1, col2, col3 = st.columns(3)
    with col1:
        archivos_historial = st.file_uploader("1. Historial(es) Tracksolid (.xlsx)", type=['xlsx', 'xls'], accept_multiple_files=True)
    with col2:
        archivo_odometro = st.file_uploader("2. Odómetro (.xlsx)", type=['xlsx', 'xls'])
    with col3:
        archivo_geocercas = st.file_uploader("3. Base de Localizaciones (.xlsx)", type=['xlsx', 'xls'])
    
    st.markdown("---")
    st.subheader("2. Parámetros Globales")
    c1, c2, c3 = st.columns(3)
    st.session_state['chofer_defecto'] = c1.text_input("Chofer por defecto:", value="").upper()
    despacho_defecto = c2.text_input("Despacho por defecto:", value="EL TIGRITO").upper()
    auto_resguardo = c3.checkbox("Detectar Hora de Resguardo Automáticamente", value=True)
    hora_manual = c3.time_input("Hora de Resguardo (Manual):", value=datetime.strptime("18:00", "%H:%M").time(), disabled=auto_resguardo)

    if st.button("🚀 Procesar Cruce GPS vs Geocercas", type="primary", use_container_width=True):
        if not archivos_historial or not archivo_geocercas:
            st.error("⚠️ Faltan archivos por cargar. (Obligatorio: Tracksolid y Base de Localizaciones)")
        else:
            with st.spinner("Masticando datos de coordenadas con Geopy..."):
                try:
                    df_base_loc = pd.read_excel(archivo_geocercas)
                    df_base_loc = _normalize_df_columns(df_base_loc)
                    
                    df_odometro = None
                    if archivo_odometro:
                        df_odometro = pd.read_excel(archivo_odometro, header=8)
                        df_odometro = _normalize_df_columns(df_odometro)

                    all_dfs = []
                    for f in archivos_historial:
                        df_h = pd.read_excel(f, header=8)
                        df_h = _normalize_df_columns(df_h)
                        all_dfs.append(df_h)
                    df_historial = pd.concat(all_dfs, ignore_index=True)

                    datos_resumen = []
                    reportes = {}
                    placas_a_procesar = sorted([p for p in MASTER_VEHICULOS.keys() if p in PLACAS_AUTORIZADAS])

                    for placa in placas_a_procesar:
                        total_km = 0
                        if df_odometro is not None and 'Placa' in df_odometro.columns:
                            odometro_placa = df_odometro[df_odometro['Placa'] == placa]
                            if not odometro_placa.empty and 'Odómetro (Km)' in odometro_placa.columns:
                                total_km = pd.to_numeric(odometro_placa['Odómetro (Km)'], errors='coerce').max()
                                if pd.isna(total_km): total_km = 0
                        
                        modelo = MASTER_VEHICULOS[placa]['modelo']
                        color = MASTER_VEHICULOS[placa]['color']
                        
                        chofer_para_esta_placa = "YONNER TAMOY" if placa == 'A72EB0P' else st.session_state['chofer_defecto']
                        despacho_guardado_manual = st.session_state['despachos_guardados'].get(placa)
                        despacho_actual = despacho_guardado_manual if despacho_guardado_manual else despacho_defecto
                        
                        ubicacion_final_gps = "N/A"
                        reporte_texto = ""

                        if 'Placa' in df_historial.columns and placa in df_historial['Placa'].values:
                            historial_placa = df_historial[df_historial['Placa'] == placa].copy()
                            historial_placa['Fecha De Reporte'] = pd.to_datetime(historial_placa['Fecha De Reporte'], errors='coerce')
                            historial_placa.dropna(subset=['Fecha De Reporte', 'Latitud', 'Longitud'], inplace=True)
                            
                            if not historial_placa.empty:
                                dia_mas_reciente = historial_placa['Fecha De Reporte'].dt.date.max()
                                hist_dia = historial_placa[historial_placa['Fecha De Reporte'].dt.date == dia_mas_reciente].sort_values('Fecha De Reporte')
                                mov_dia = hist_dia[pd.to_numeric(hist_dia['Velocidad (Km/H)'], errors='coerce').fillna(0) > VELOCIDAD_MINIMA_MOVIMIENTO]
                                
                                if mov_dia.empty:
                                    reporte_texto = (f"🚛 {placa} - {chofer_para_esta_placa}\n*DESPACHO:* {despacho_actual}\n\n"
                                                     f"LA UNIDAD NO REGISTRÓ MOVIMIENTO EL {dia_mas_reciente.strftime('%d/%m/%Y')}.\n\n*KM Total:* {total_km:,.0f} Kms".replace(',', '.'))
                                    ubicacion_final_gps = "SIN MOVIMIENTO"
                                else:
                                    hora_salida = mov_dia.iloc[0]['Fecha De Reporte']
                                    if auto_resguardo:
                                        fecha_resguardo = mov_dia.iloc[-1]['Fecha De Reporte']
                                    else:
                                        fecha_resguardo = datetime.combine(hora_salida.date(), hora_manual)
                                    
                                    ultimo_punto = hist_dia.iloc[-1]
                                    ubicacion_final_calculada = encontrar_punto_mas_cercano(ultimo_punto['Latitud'], ultimo_punto['Longitud'], df_base_loc, DISTANCIA_MAXIMA_METROS) or "Ubicación Desconocida"

                                    historial_ruta = []
                                    ultimo_lugar = None
                                    recorrido = hist_dia[(hist_dia['Fecha De Reporte'] >= hora_salida) & (hist_dia['Fecha De Reporte'] <= fecha_resguardo)]
                                    
                                    punto_anterior = None
                                    for _, p in recorrido.iterrows():
                                        df_busq = df_base_loc
                                        if punto_anterior is not None and 'Posicion' in df_base_loc.columns:
                                            dir_act = obtener_direccion_cardinal(punto_anterior['Latitud'], punto_anterior['Longitud'], p['Latitud'], p['Longitud'])
                                            if dir_act:
                                                df_busq = pd.concat([df_base_loc[df_base_loc['Posicion'] == dir_act], df_base_loc[df_base_loc['Posicion'].isnull() | (df_base_loc['Posicion'] == '')]])
                                        
                                        lugar = encontrar_punto_mas_cercano(p['Latitud'], p['Longitud'], df_busq, DISTANCIA_MAXIMA_METROS)
                                        if lugar and lugar != ultimo_lugar:
                                            historial_ruta.append(f"{p['Fecha De Reporte'].strftime('%I:%M:%S %p')} - {str(lugar).upper()}")
                                            ultimo_lugar = lugar
                                        punto_anterior = p
                                    
                                    ubi_inicial = historial_ruta[0].split(' - ')[1] if historial_ruta else "N/A"
                                    if ubicacion_final_calculada != "Ubicación Desconocida": ubicacion_final_gps = ubicacion_final_calculada.upper()
                                    elif historial_ruta: ubicacion_final_gps = historial_ruta[-1].split(' - ')[-1].upper()
                                    
                                    horas, rem = divmod((fecha_resguardo - hora_salida).total_seconds(), 3600)
                                    minutos = (rem % 3600) // 60
                                    
                                    hist_text = '\n'.join(historial_ruta) if historial_ruta else "No se detectaron puntos de ruta conocidos."
                                    reporte_texto = (f"🚛 {placa} - {chofer_para_esta_placa}\n*DESPACHO:* {despacho_actual}\n\n"
                                                     f"*UBICACIÓN INICIAL:*\n{hora_salida.strftime('%I:%M:%S %p')} - UNIDAD REPORTA: {ubi_inicial}\n\n"
                                                     f"*UBICACIÓN FINAL:*\n{fecha_resguardo.strftime('%I:%M:%S %p')} - UNIDAD REPORTA: {ubicacion_final_gps}\n\n"
                                                     f"----------------------------------------------------\n*HISTORIAL DE RUTA*\n----------------------------------------------------\n{hist_text}\n\n"
                                                     f"----------------------------------------------------\n*RESUMEN DEL DIA - {placa} - {hora_salida.strftime('%d/%m/%Y')}*\n----------------------------------------------------\n\n"
                                                     f"*TOTAL KILOMETRAJE:* {total_km:,.0f} Kms\n\n*SALIDA:* {hora_salida.strftime('%I:%M:%S %p')}\n*RESGUARDO:* {fecha_resguardo.strftime('%I:%M:%S %p')}\n\n*TOTAL HORAS:* {int(horas)} Horas y {int(minutos)} minutos".replace(',', '.'))
                        
                        if not reporte_texto:
                            reporte_texto = f"🚛 {placa} - {chofer_para_esta_placa}\n*DESPACHO:* {despacho_actual}\n\n[SIN DATOS GPS]"
                            ubicacion_final_gps = "Plantilla Manual"

                        reportes[placa] = reporte_texto
                        ruta_resumen = despacho_guardado_manual if despacho_guardado_manual else ubicacion_final_gps

                        datos_resumen.append({
                            'PLACA': placa,
                            'MODELO': modelo,
                            'COLOR': color,
                            'RUTA': str(ruta_resumen).upper(),
                            'KM': total_km
                        })

                    st.session_state['datos_resumen'] = datos_resumen
                    st.session_state['reportes_texto'] = reportes
                    st.success("✅ Procesamiento completado. Revisa las pestañas de Resumen y Reportes.")

                except Exception as e:
                    st.error(f"Error fatal: {e}")
                    st.code(traceback.format_exc())

# ---------------------------------------------------------
# PESTAÑA 2: RESUMEN DE VEHÍCULOS (PIZARRA FOTO HD Y WHATSAPP)
# ---------------------------------------------------------
with t_resumen:
    if st.session_state['datos_resumen']:
        df_res = pd.DataFrame(st.session_state['datos_resumen'])
        km_total_gral = df_res['KM'].sum()
        
        st.subheader("1. Edición y Actualización de Rutas")
        st.info("Puedes editar la columna **RUTA** directamente en la tabla. El texto de WhatsApp se actualizará automáticamente.")
        df_editado = st.data_editor(
            df_res[['PLACA', 'MODELO', 'COLOR', 'RUTA', 'KM']], 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "PLACA": st.column_config.TextColumn(disabled=True),
                "MODELO": st.column_config.TextColumn(disabled=True),
                "COLOR": st.column_config.TextColumn(disabled=True),
                "KM": st.column_config.NumberColumn("Kilometraje", format="%.0f Kms", disabled=True),
            }
        )
        
        if st.button("💾 Guardar Rutas Definitivas"):
            for _, r in df_editado.iterrows():
                st.session_state['despachos_guardados'][r['PLACA']] = r['RUTA']
            guardar_json_local(DESPACHOS_DB_FILE, st.session_state['despachos_guardados'])
            st.success("Rutas guardadas para próximos reportes.")
        
        st.markdown("---")
        
        # --- GENERADOR DE WHATSAPP DINÁMICO ---
        st.subheader("2. Resumen para WhatsApp (Listo para Copiar)")
        msg_w = f"🛰️ *REPORTE EJECUTIVO DE FLOTA PCD - GPS*\n📅 Fecha: {datetime.now().strftime('%d/%m/%Y')}\n\n"
        for _, r in df_editado.iterrows():
            msg_w += f"🚙 *{r['MODELO']}* ({r['PLACA']})\n📍 Ruta/Status: {r['RUTA']}\n📏 Odómetro: {formatear_km(r['KM'])} Kms\n\n"
        msg_w += f"📊 *TOTAL VEHÍCULOS:* {len(df_editado)} Unidades\n"
        msg_w += f"🛣️ *GRAN TOTAL RECORRIDO:* {formatear_km(km_total_gral)} Kms"
        
        st.code(msg_w, language="markdown")
        
        st.markdown("---")
        st.subheader("3. Pizarra Ejecutiva (Imagen HD)")
        
        # --- GENERADOR HTML PARA FOTO (DISEÑO SANS-SERIF Y ORO) ---
        logo_b64 = obtener_logo_base64()
        filas_html = ""
        fondo = 0
        df_editado_ordenado = df_editado.sort_values(by=['MODELO', 'PLACA'])
        
        for _, r in df_editado_ordenado.iterrows():
            bg_color = "#f2f2f2" if fondo == 1 else "white"
            fondo = 1 if fondo == 0 else 0
            
            filas_html += f"""
            <tr style="background-color: {bg_color}; border-bottom: 1px solid #ddd; text-align: center;">
                <td style="padding: 10px; font-weight: bold; border-right: 1px solid #ddd;">{r['PLACA']}</td>
                <td style="padding: 10px; text-align: left; border-right: 1px solid #ddd;">{r['MODELO']}</td>
                <td style="padding: 10px; border-right: 1px solid #ddd;">{r['COLOR']}</td>
                <td style="padding: 10px; text-align: left; font-weight: bold; color: #0D47A1; border-right: 1px solid #ddd;">{r['RUTA']}</td>
                <td style="padding: 10px; font-weight: 900; text-align: right;">{formatear_km(r['KM'])} Kms</td>
            </tr>
            """

        html_pizarra_completa = f"""
        <html><head>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        </head><body style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background: #f0f2f6; padding: 20px;">
            <div style="text-align: center; margin-bottom: 15px;">
                <button onclick="capResumen()" style="background: #0d47a1; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">📸 DESCARGAR PIZARRA (FOTO HD)</button>
            </div>
            
            <div id="pizarra-resumen" style="background: white; width: 950px; margin: auto; border: 2px solid #0d47a1; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 20px rgba(0,0,0,0.15);">
                <div style="background-color: #0d47a1; padding: 20px 30px; display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid #F57F17;">
                    <div style="display: flex; align-items: center;">
                        <img src="{logo_b64}" style="height: 55px; margin-right: 15px;">
                        <div style="color: white; font-size: 24px; font-weight: 900; letter-spacing: 1px;">RESUMEN LOGÍSTICO DE FLOTA GPS</div>
                    </div>
                    <div style="color: #F57F17; font-size: 16px; font-weight: bold;">
                        {datetime.now().strftime("%d/%m/%Y - %I:%M %p")}
                    </div>
                </div>
                
                <div style="padding: 20px;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 14px; border: 1px solid #ddd;">
                        <thead>
                            <tr style="background-color: #F57F17; color: white;">
                                <th style="padding: 12px; border-right: 1px solid #ddd;">PLACA</th>
                                <th style="padding: 12px; text-align: left; border-right: 1px solid #ddd;">MODELO</th>
                                <th style="padding: 12px; border-right: 1px solid #ddd;">COLOR</th>
                                <th style="padding: 12px; text-align: left; border-right: 1px solid #ddd;">RUTA / UBICACIÓN FINAL GPS</th>
                                <th style="padding: 12px; text-align: right;">ODÓMETRO (KM)</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filas_html}
                        </tbody>
                    </table>
                </div>
                
                <div style="background-color: #333; color: white; padding: 15px 30px; display: flex; justify-content: space-between; align-items: center;">
                    <div style="font-size: 18px; font-weight: bold;">TOTAL DE VEHÍCULOS: {len(df_editado)}</div>
                    <div style="font-size: 20px; font-weight: 900; color: #FFD54F;">TOTAL RECORRIDO: {formatear_km(km_total_gral)} Kms</div>
                </div>
            </div>
            
            <script>
                function capResumen() {{
                    // Usando scale: 3 para asegurar una calidad de imagen altísima (HD)
                    html2canvas(document.getElementById('pizarra-resumen'), {{scale: 3}}).then(canvas => {{
                        var link = document.createElement('a');
                        link.download = 'Pizarra_GPS_{datetime.now().strftime('%Y%m%d')}.png';
                        link.href = canvas.toDataURL('image/png', 1.0);
                        link.click();
                    }});
                }}
            </script>
        </body></html>
        """
        components.html(html_pizarra_completa.replace(',', '.'), height=1100, scrolling=True)

    else:
        st.info("Sube los archivos y presiona Procesar en la pestaña Configuración.")

# ---------------------------------------------------------
# PESTAÑA 3: REPORTES INDIVIDUALES
# ---------------------------------------------------------
with t_reportes:
    if st.session_state['reportes_texto']:
        placa_sel = st.selectbox("Seleccione el Vehículo:", list(st.session_state['reportes_texto'].keys()))
        texto = st.text_area("Reporte Generado (Editable):", value=st.session_state['reportes_texto'][placa_sel], height=400)
        
        import urllib.parse
        texto_url = urllib.parse.quote(texto)
        url_wa = f"https://wa.me/584127969408?text={texto_url}"
        
        st.markdown(f"""
        <a href="{url_wa}" target="_blank" style="text-decoration:none;">
            <button style="background-color:#25D366; color:white; padding:10px 20px; border:none; border-radius:5px; font-weight:bold; cursor:pointer; width:100%;">
                📲 ENVIAR REPORTE A WHATSAPP
            </button>
        </a>
        """, unsafe_allow_html=True)
    else:
        st.info("Los reportes aparecerán aquí después de procesar los datos.")

# ---------------------------------------------------------
# PESTAÑA 4: GUARDAR EN LA NUBE (GOOGLE SHEETS)
# ---------------------------------------------------------
with t_historico:
    st.subheader("💾 Guardar en Google Sheets")
    st.info("Al presionar el botón, se enviará el resumen del día a tu hoja 'Historial_GPS' en la nube.")
    
    if st.button("🚀 Enviar Datos a Google Sheets", type="primary"):
        if st.session_state['datos_resumen']:
            try:
                with st.spinner("Conectando con la nube (Google Sheets)..."):
                    datos_a_enviar = []
                    fecha_actual = datetime.now()
                    
                    semana = fecha_actual.strftime("%W")
                    
                    dias_es = {"Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles", "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"}
                    meses_es = {"January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril", "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto", "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"}
                    
                    dia_nombre = dias_es.get(fecha_actual.strftime("%A"), fecha_actual.strftime("%A"))
                    mes_nombre = meses_es.get(fecha_actual.strftime("%B"), fecha_actual.strftime("%B"))
                    
                    for d in st.session_state['datos_resumen']:
                        chofer = "YONNER TAMOY" if d['PLACA'] == 'A72EB0P' else st.session_state['chofer_defecto']
                        if not chofer: chofer = "POR DEFINIR"
                        
                        fila = [
                            fecha_actual.strftime("%d/%m/%Y"),
                            dia_nombre,
                            semana,
                            mes_nombre,
                            d['PLACA'],
                            chofer,
                            d['RUTA'],
                            d['KM']
                        ]
                        datos_a_enviar.append(fila)
                    
                    guardar_en_googlesheets(datos_a_enviar)
                    st.success("✅ ¡Datos registrados en Google Sheets exitosamente!")
            except Exception as e:
                st.error(f"Error al conectar con Sheets: {e}")
                st.code(traceback.format_exc())
        else:
            st.warning("No hay datos procesados en la pestaña de Resumen para enviar a la nube.")
