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

# --- CONSTANTES ---
VELOCIDAD_MINIMA_MOVIMIENTO = 5
DISTANCIA_MAXIMA_METROS = 300
HISTORICO_CSV_FILE = "historico_chinitas.csv"
DESPACHOS_DB_FILE = "despachos_guardados.json"

PLACAS_AUTORIZADAS = {
    'A36AC9X', 'A37CS2D', 'A38BA4N', 'A38CV4G', 'A48AU5T',
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
    'A38CV4G': {'modelo': 'FORD RANGER', 'color': 'PLATA'},
    'A00DS2V': {'modelo': 'ENCAVA', 'color': 'BLANCO'},
    'AB893NB': {'modelo': 'MITSUBISHI LANCER', 'color': 'GRIS'},
    'A73CL7D': {'modelo': 'REY CAMION', 'color': 'PLATA'},
    'A84EZ6P': {'modelo': 'RICH P11', 'color': 'BLANCO'},
    'A87EZ8P': {'modelo': 'RICH P11', 'color': 'BLANCO'},
}

# --- FUNCIONES DE CÁLCULO ---
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

# --- INTERFAZ DE STREAMLIT ---
st.set_page_config(page_title="Análisis GPS Chinitas", layout="wide")
st.title("🛰️ Análisis de Rutas y Paradas GPS (Tracksolid)")

# Inicializar estados
if 'despachos_guardados' not in st.session_state:
    st.session_state['despachos_guardados'] = cargar_json_local(DESPACHOS_DB_FILE)
if 'datos_resumen' not in st.session_state:
    st.session_state['datos_resumen'] = []
if 'reportes_texto' not in st.session_state:
    st.session_state['reportes_texto'] = {}

t_config, t_resumen, t_reportes, t_historico = st.tabs(["⚙️ Configuración", "📊 Resumen de Vehículos", "📝 Reportes Individuales", "💾 Historial"])

# ==========================================
# PESTAÑA 1: CONFIGURACIÓN
# ==========================================
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
    chofer_defecto = c1.text_input("Chofer por defecto:", value="").upper()
    despacho_defecto = c2.text_input("Despacho por defecto:", value="EL TIGRITO").upper()
    auto_resguardo = c3.checkbox("Detectar Hora de Resguardo Automáticamente", value=True)
    hora_manual = c3.time_input("Hora de Resguardo (Manual):", value=datetime.strptime("18:00", "%H:%M").time(), disabled=auto_resguardo)

    if st.button("🚀 ProcesAR CRUCE GPS VS GEOCERCAS", type="primary", use_container_width=True):
        if not archivos_historial or not archivo_geocercas:
            st.error("⚠️ Faltan archivos por cargar. (Obligatorio: Tracksolid y Base de Localizaciones)")
        else:
            with st.spinner("Masticando datos de coordenadas con Geopy..."):
                try:
                    # Leer y limpiar Geocercas
                    df_base_loc = pd.read_excel(archivo_geocercas)
                    df_base_loc = _normalize_df_columns(df_base_loc)
                    
                    # Leer y limpiar Odómetro
                    df_odometro = None
                    if archivo_odometro:
                        df_odometro = pd.read_excel(archivo_odometro, header=8)
                        df_odometro = _normalize_df_columns(df_odometro)

                    # Leer y limpiar Historial GPS
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
                        
                        # Lógica Chofer
                        chofer_para_esta_placa = "YONNER TAMOY" if placa == 'A72EB0P' else chofer_defecto
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

# ==========================================
# PESTAÑA 2: RESUMEN DE VEHÍCULOS (PIZARRA)
# ==========================================
with t_resumen:
    if st.session_state['datos_resumen']:
        df_res = pd.DataFrame(st.session_state['datos_resumen'])
        km_total_gral = df_res['KM'].sum()
        
        # Opciones de Guardado de Edición de Rutas
        st.info("Puedes editar la columna **RUTA** directamente en la tabla. Los cambios se guardarán para el próximo reporte.")
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
        
        # Guardar cambios en el JSON
        if st.button("💾 Guardar Rutas Editadas"):
            for _, r in df_editado.iterrows():
                st.session_state['despachos_guardados'][r['PLACA']] = r['RUTA']
            guardar_json_local(DESPACHOS_DB_FILE, st.session_state['despachos_guardados'])
            st.success("Rutas actualizadas exitosamente.")
            
        st.markdown("---")
        
        # --- GENERACIÓN DE PIZARRA HTML EXACTA AL ORIGINAL ---
        modelos_grp = df_editado.groupby('MODELO')
        tarjetas_html = ""
        
        for mod, grp in modelos_grp:
            km_mod = grp['KM'].sum()
            cant_mod = len(grp)
            
            filas_vehiculos = "".join([f"<tr><td style='padding:5px; border-bottom:1px solid #ddd;'>{r['PLACA']}</td><td style='padding:5px; border-bottom:1px solid #ddd;'>{r['COLOR']}</td><td style='padding:5px; border-bottom:1px solid #ddd; font-weight:bold; color:#0d47a1;'>{r['RUTA']}</td><td style='padding:5px; border-bottom:1px solid #ddd; text-align:right;'>{r['KM']:,.0f} Kms</td></tr>".replace(',', '.') for _, r in grp.iterrows()])
            
            tarjetas_html += f"""
            <div style="background: white; border: 1px solid #ccc; border-radius: 8px; padding: 15px; margin-bottom: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0d47a1; padding-bottom: 10px; margin-bottom: 10px;">
                    <div>
                        <div style="font-size: 18px; font-weight: 900; color: #333;">{mod} <span style="background: #eee; padding: 2px 8px; border-radius: 10px; font-size: 14px; margin-left: 10px; border: 1px solid #aaa;">{cant_mod} UND</span></div>
                    </div>
                    <div style="text-align: right;">
                        <div style="font-size: 11px; font-weight: bold; color: #666;">KILOMETRAJE</div>
                        <div style="background: #fff9c4; color: #f57f17; font-weight: 900; font-size: 18px; padding: 5px 15px; border-radius: 5px; border: 1px solid #fbc02d;">{km_mod:,.0f} Kms</div>
                    </div>
                </div>
                <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                    <thead><tr style="background: #f8f9fa; color: #555; text-align: left;"><th style="padding:5px;">PLACA</th><th style="padding:5px;">COLOR</th><th style="padding:5px;">RUTA ASIGNADA</th><th style="padding:5px; text-align:right;">ODÓMETRO</th></tr></thead>
                    <tbody>{filas_vehiculos.replace(',', '.')}</tbody>
                </table>
            </div>
            """

        html_pizarra_completa = f"""
        <html><head>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        </head><body style="font-family: 'Arial', sans-serif; background: #f0f2f6; padding: 20px;">
            <div style="text-align: center; margin-bottom: 15px;">
                <button onclick="capResumen()" style="background: #0d47a1; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer;">📸 DESCARGAR PIZARRA COMPLETA</button>
            </div>
            <div id="pizarra-resumen" style="background: #f4f6f9; width: 900px; margin: auto; border: 2px solid #0d47a1; border-radius: 12px; overflow: hidden;">
                <div style="background: white; padding: 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #0d47a1;">
                    <div style="font-size: 22px; font-weight: 900; color: #0d47a1;">FLOTA MAZIAD / FLAMINGO 2026</div>
                    <div style="font-size: 16px; font-weight: bold; color: #555;">{datetime.now().strftime("%d/%m/%Y")}</div>
                </div>
                <div style="padding: 20px;">
                    {tarjetas_html}
                </div>
                <div style="background: #333; color: white; padding: 15px 25px; display: flex; justify-content: space-between; align-items: center;">
                    <div style="font-size: 16px; font-weight: bold;">TOTAL DE VEHÍCULOS: {len(df_res)}</div>
                    <div style="font-size: 18px; font-weight: 900; color: #ffd54f;">TOTAL RECORRIDO: {km_total_gral:,.0f} Kms</div>
                </div>
            </div>
            <script>function capResumen() {{ html2canvas(document.getElementById('pizarra-resumen'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Pizarra_GPS.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>
        </body></html>
        """
        components.html(html_pizarra_completa.replace(',', '.'), height=1500, scrolling=True)

    else:
        st.info("Sube los archivos y presiona Procesar en la pestaña Configuración.")

# ==========================================
# PESTAÑA 3: REPORTES INDIVIDUALES
# ==========================================
with t_reportes:
    if st.session_state['reportes_texto']:
        placa_sel = st.selectbox("Seleccione el Vehículo:", list(st.session_state['reportes_texto'].keys()))
        texto = st.text_area("Reporte Generado (Editable):", value=st.session_state['reportes_texto'][placa_sel], height=400)
        
        # Enlace rápido para WhatsApp (Abre web.whatsapp.com con el texto pre-escrito)
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

# ==========================================
# PESTAÑA 4: HISTORIAL DE KILOMETRAJES
# ==========================================
with t_historico:
    st.subheader("Gestión de Historial Local")
    if st.button("💾 Guardar Datos Actuales en el Historial", type="primary"):
        if st.session_state['datos_resumen']:
            nuevos = []
            fecha_hoy = datetime.now().strftime("%Y-%m-%d")
            for d in st.session_state['datos_resumen']:
                nuevos.append({'FECHA': fecha_hoy, 'PLACA': d['PLACA'], 'MODELO': d['MODELO'], 'RUTA': d['RUTA'], 'KMS': d['KM']})
            
            df_nuevos = pd.DataFrame(nuevos)
            if os.path.exists(HISTORICO_CSV_FILE):
                df_hist = pd.read_csv(HISTORICO_CSV_FILE)
                df_final = pd.concat([df_hist, df_nuevos], ignore_index=True)
                df_final.drop_duplicates(subset=['FECHA', 'PLACA'], keep='last', inplace=True)
            else:
                df_final = df_nuevos
            
            df_final.to_csv(HISTORICO_CSV_FILE, index=False)
            st.success("Historial guardado exitosamente.")
        else:
            st.warning("No hay datos en el resumen para guardar.")
            
    if os.path.exists(HISTORICO_CSV_FILE):
        df_historico_view = pd.read_csv(HISTORICO_CSV_FILE)
        st.dataframe(df_historico_view, use_container_width=True)
        
        with open(HISTORICO_CSV_FILE, "rb") as file:
            st.download_button(label="📥 Descargar CSV Histórico", data=file, file_name="historico_chinitas.csv", mime="text/csv")
    else:
        st.info("Aún no se ha guardado ningún historial.")