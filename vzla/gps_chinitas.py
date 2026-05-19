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
import textwrap
from google.oauth2.service_account import Credentials
import gspread
from fpdf import FPDF
import io

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

# ==========================================
# GENERADOR DE PDF SERVER-SIDE
# ==========================================
class GPS_PDF(FPDF):
    def header(self):
        # Color Oro PCD para el header
        self.set_fill_color(13, 71, 161) # Azul PCD
        self.set_text_color(255, 255, 255)
        self.set_font('helvetica', 'B', 16)
        self.cell(0, 15, 'RESUMEN LOGÍSTICO DE FLOTA GPS', 1, 1, 'C', 1)
        
        # Detalles bajo el header
        self.set_text_color(50, 50, 50)
        self.set_font('helvetica', 'B', 10)
        fecha_str = datetime.now().strftime("%d/%m/%Y - %I:%M %p")
        self.cell(0, 10, f'Fecha de Generación: {fecha_str}', 0, 1, 'L')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f'Reporte Generado por Streamlit-Server v2.1 | Pág. {self.page_no()}', 0, 0, 'C')

def generar_resumen_pdf(df_datos, total_vehiculos, total_km):
    pdf = GPS_PDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font('helvetica', '', 10)

    # --- TABLA PRINCIPAL DE VEHÍCULOS ---
    
    # Encabezados con color Oro (F57F17 es el oro de Chinitas/PCD anterior)
    pdf.set_fill_color(245, 127, 23) # Color ORO PCD
    pdf.set_text_color(255, 255, 255) # Texto Blanco
    pdf.set_font('helvetica', 'B', 11)
    
    ancho_col = {'placa': 30, 'modelo': 60, 'color': 30, 'ruta': 120, 'km': 35}
    h_encabezado = 10
    
    pdf.cell(ancho_col['placa'], h_encabezado, 'PLACA', 1, 0, 'C', 1)
    pdf.cell(ancho_col['modelo'], h_encabezado, 'MODELO', 1, 0, 'C', 1)
    pdf.cell(ancho_col['color'], h_encabezado, 'COLOR', 1, 0, 'C', 1)
    pdf.cell(ancho_col['ruta'], h_encabezado, 'RUTA / UBICACIÓN FINAL GPS', 1, 0, 'C', 1)
    pdf.cell(ancho_col['km'], h_encabezado, 'ODÓMETRO (KM)', 1, 1, 'C', 1)

    # Filas de datos
    pdf.set_text_color(0, 0, 0) # Texto Negro
    pdf.set_font('helvetica', '', 10)
    h_fila = 8
    
    # Alternar color de fondo
    fondo = 0
    
    # Ordenar por modelo para mejor visualización
    df_datos = df_datos.sort_values(by=['MODELO', 'PLACA'])

    for _, row in df_datos.iterrows():
        # Fondo alterno
        if fondo == 1: pdf.set_fill_color(240, 240, 240); fondo=0
        else: pdf.set_fill_color(255, 255, 255); fondo=1
        
        pdf.cell(ancho_col['placa'], h_fila, row['PLACA'], 1, 0, 'C', 1)
        pdf.cell(ancho_col['modelo'], h_fila, row['MODELO'], 1, 0, 'L', 1)
        pdf.cell(ancho_col['color'], h_fila, row['COLOR'], 1, 0, 'C', 1)
        
        # Truncar ruta si es muy larga para evitar desbordamiento
        ruta_corta = str(row['RUTA'])
        if len(ruta_corta) > 60: ruta_corta = ruta_corta[:57] + "..."
        pdf.cell(ancho_col['ruta'], h_fila, ruta_corta, 1, 0, 'L', 1)
        
        pdf.cell(ancho_col['km'], h_fila, formatear_km(row['KM']), 1, 1, 'R', 1)

    # --- BARRA DE TOTALES ---
    pdf.ln(10)
    pdf.set_fill_color(13, 71, 161) # Azul PCD
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('helvetica', 'B', 12)
    pdf.cell(140, 12, f'TOTAL DE VEHÍCULOS: {total_vehiculos}', 1, 0, 'C', 1)
    pdf.set_fill_color(255, 213, 79) # Amarillo/Oro claro para contraste
    pdf.set_text_color(0, 0, 0)
    pdf.cell(135, 12, f'TOTAL RECORRIDO (GRAN TOTAL): {formatear_km(total_km)} Kms', 1, 1, 'C', 1)

    return pdf.output(dest='S') # Retorna bytes

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

t_config, t_resumen, t_reportes, t_historico = st.tabs(["⚙️ Configuración", "📊 Resumen de Vehículos (Pizarra)", "📝 Reportes Individuales", "💾 Guardar en Nube"])

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
# PESTAÑA 2: RESUMEN DE VEHÍCULOS (PIZARRA MODERNA)
# ---------------------------------------------------------
with t_resumen:
    if st.session_state['datos_resumen']:
        df_res = pd.DataFrame(st.session_state['datos_resumen'])
        km_total_gral = df_res['KM'].sum()
        
        st.subheader("1. Edición de Rutas")
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
        
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            if st.button("💾 Guardar Rutas"):
                for _, r in df_editado.iterrows():
                    st.session_state['despachos_guardados'][r['PLACA']] = r['RUTA']
                guardar_json_local(DESPACHOS_DB_FILE, st.session_state['despachos_guardados'])
                st.success("Rutas actualizadas.")
        
        st.markdown("---")
        
        st.subheader("2. Generación de Pizarra Ejecutiva (PDF)")
        
        # Botón para descargar PDF (Server-side)
        with st.spinner("Preparando documento PDF profesional..."):
            pdf_bytes = generar_resumen_pdf(df_editado, len(df_editado), km_total_gral)
            st.download_button(
                label="📥 DESCARGAR REPORTE PROFESIONAL (PDF)",
                data=pdf_bytes,
                file_name=f"Pizarra_GPS_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        
        st.markdown("### Vista previa del diseño (Formato Sans-Serif & Oro PCD)")
        
        # O Oro (F57F17), Azul PCD (0D47A1)
        st.markdown(f"""
        <style>
            .pizarra-cont {{ font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; border: 1px solid #ccc; border-radius: 8px; overflow: hidden; }}
            .pizarra-header {{ background-color: white; padding: 15px 25px; border-bottom: 3px solid #0D47A1; display: flex; justify-content: space-between; align-items: center; }}
            .pizarra-header-title {{ font-size: 20px; font-weight: 900; color: #0D47A1; }}
            .pizarra-body {{ padding: 20px; background-color: #F8F9FA; }}
            .pizarra-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
            .pizarra-table th {{ background-color: #F57F17; color: white; padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold; text-transform: uppercase; }}
            .pizarra-table td {{ padding: 8px; border: 1px solid #ddd; background-color: white; }}
            .pizarra-table tr:nth-child(even) td {{ background-color: #f2f2f2; }}
            .pizarra-ruta {{ font-weight: bold; color: #0D47A1; }}
            .pizarra-total {{ background-color: #333; color: white; padding: 15px 25px; display: flex; justify-content: space-between; align-items: center; }}
        </style>
        <div class="pizarra-cont">
            <div class="pizarra-header">
                <div class="pizarra-header-title">FLOTA PCD / VENEZUELA</div>
                <div style="font-weight: bold; color: #555;">{datetime.now().strftime("%d/%m/%Y")}</div>
            </div>
            <div class="pizarra-body">
                <table class="pizarra-table">
                    <thead><tr><th>PLACA</th><th>MODELO</th><th>COLOR</th><th>RUTA / UBICACIÓN FINAL</th><th style="text-align:right;">ODÓMETRO</th></tr></thead>
                    <tbody>
                        {"".join([f"<tr><td style='text-align:center;'>{r['PLACA']}</td><td>{r['MODELO']}</td><td style='text-align:center;'>{r['COLOR']}</td><td class='pizarra-ruta'>{r['RUTA']}</td><td style='text-align:right;'>{formatear_km(r['KM'])} Kms</td></tr>" for _, r in df_editado.iterrows()])}
                    </tbody>
                </table>
            </div>
            <div class="pizarra-total">
                <div style="font-weight: bold;">VEHÍCULOS: {len(df_editado)}</div>
                <div style="font-size: 18px; font-weight: 900; color: #FFD54F;">GRAN TOTAL: {formatear_km(km_total_gral)} Kms</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

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
# PESTAÑA 4: GUARDAR EN LA NUBE (NUEVO GOOGLE SHEETS)
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
                    
                    # Calcular semana y mes
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
