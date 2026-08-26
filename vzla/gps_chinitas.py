# ==========================================
# Archivo: gps_chinitas.py (Análisis de Rutas Tracksolid)
# ==========================================
import streamlit as st
import pandas as pd
from geopy.distance import great_circle
from datetime import datetime, timedelta
import os
import json
import base64
import io
import traceback
import urllib.parse
import streamlit.components.v1 as components
import gspread
import textwrap
from pathlib import Path
from google.oauth2.service_account import Credentials

try:
    import folium
    FOLIUM_DISPONIBLE = True
except ImportError:
    FOLIUM_DISPONIBLE = False

try:
    from staticmap import StaticMap, Line, CircleMarker
    STATICMAP_DISPONIBLE = True
except ImportError:
    STATICMAP_DISPONIBLE = False

try:
    from fotos_vehiculos_data import FOTOS_BASE64
except ImportError:
    FOTOS_BASE64 = {}

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
NUMERO_WHATSAPP_DEFECTO = "584127969408"

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

# Mapa MODELO -> nombre de archivo esperado dentro de la carpeta "fotos/".
# Deja caer ahí una foto del vehículo con ese nombre exacto (jpg o png) y la
# Pizarra Ejecutiva la usará automáticamente. Si no existe, se usa un ícono
# genérico y la app sigue funcionando igual.
FOTOS_VEHICULOS = {
    'CHANGAN HUNTER 4X2': 'changan_hunter.png',
    'CHANGAN KAICENE F70': 'changan_kaicene.png',
    'DFSK D1': 'dfsk_d1.png',
    'ENCAVA': 'encava.png',
    'MITSUBISHI LANCER': 'mitsubishi_lancer.png',
    'REY CAMION': 'rey_camion.png',
    'RICH P11': 'rich_p11.png',
}

ICONO_GENERICO_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="100%" height="100%">
<rect width="64" height="64" fill="#e3e8f0"/>
<g fill="#0d47a1">
<path d="M4 38h34v10H4z"/>
<path d="M38 30h12l8 8v10H38z"/>
<circle cx="16" cy="50" r="6" fill="#333"/>
<circle cx="46" cy="50" r="6" fill="#333"/>
</g>
</svg>
"""
ICONO_GENERICO_B64 = "data:image/svg+xml;base64," + base64.b64encode(ICONO_GENERICO_SVG.encode()).decode()

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
        ruta_logo = Path(__file__).parent / "logo.png"
        if ruta_logo.exists():
            with open(ruta_logo, "rb") as image_file:
                return f"data:image/png;base64,{base64.b64encode(image_file.read()).decode()}"
        return None
    except: return None

def obtener_foto_vehiculo(modelo):
    """
    Orden de prioridad:
    1) Foto propia en la carpeta fotos/ (si algún día quieres reemplazar la de catálogo
       por la foto real de TU vehículo, solo colócala ahí con el nombre de FOTOS_VEHICULOS).
    2) Foto de catálogo ya integrada en fotos_vehiculos_data.py (no requiere subir nada).
    3) Ícono genérico, por si el modelo no está contemplado.
    """
    nombre_archivo = FOTOS_VEHICULOS.get(modelo)
    if nombre_archivo:
        for carpeta in [Path(__file__).parent / "fotos", Path(__file__).parent]:
            ruta = carpeta / nombre_archivo
            if ruta.exists():
                ext = ruta.suffix.lower().replace('.', '') or 'png'
                mime = 'jpeg' if ext in ('jpg', 'jpeg') else ext
                with open(ruta, "rb") as f:
                    return f"data:image/{mime};base64,{base64.b64encode(f.read()).decode()}"
    if modelo in FOTOS_BASE64:
        return FOTOS_BASE64[modelo]
    return ICONO_GENERICO_B64

def generar_mapa_folium(placa, puntos, altura=420):
    """puntos: lista de dicts {lat, lon, hora, lugar}. Devuelve HTML del mapa Leaflet/OSM."""
    if not FOLIUM_DISPONIBLE or not puntos:
        return None
    coords = [(p['lat'], p['lon']) for p in puntos]
    centro = coords[len(coords) // 2]
    m = folium.Map(location=centro, zoom_start=14, tiles="OpenStreetMap")
    folium.PolyLine(coords, color="#0d47a1", weight=4, opacity=0.85).add_to(m)
    folium.Marker(
        coords[0], tooltip=f"Salida: {puntos[0]['hora'].strftime('%I:%M:%S %p')} - {puntos[0].get('lugar','')}",
        icon=folium.Icon(color="green", icon="play")
    ).add_to(m)
    folium.Marker(
        coords[-1], tooltip=f"Resguardo: {puntos[-1]['hora'].strftime('%I:%M:%S %p')} - {puntos[-1].get('lugar','')}",
        icon=folium.Icon(color="red", icon="stop")
    ).add_to(m)
    m.fit_bounds(coords)
    return m._repr_html_()

@st.cache_data(show_spinner=False)
def generar_mapa_estatico_png(puntos, ancho=900, alto=600):
    """
    Genera el trazado como una imagen PNG real (no un mapa interactivo), dibujada del
    lado del servidor con la librería `staticmap`. Esto es lo que permite el botón de
    descarga: capturar un mapa de Leaflet con html2canvas NO funciona de forma confiable
    porque el mapa vive dentro de iframes anidados (el propio folium genera un iframe, y
    components.html de Streamlit agrega otro) — por eso salía el texto de aviso de
    "Trust Notebook" en vez del mapa. Generando el PNG en Python nos evitamos ese problema
    por completo.
    """
    if not STATICMAP_DISPONIBLE or not puntos:
        return None
    coords = [(p['lon'], p['lat']) for p in puntos]  # staticmap espera (lon, lat)
    m = StaticMap(
        ancho, alto,
        url_template='https://a.tile.openstreetmap.org/{z}/{x}/{y}.png',
        headers={"User-Agent": "FlotaGPS-Chinitas/1.0"},  # OSM exige un User-Agent identificable
    )
    m.add_line(Line(coords, '#0d47a1', 4))
    m.add_marker(CircleMarker(coords[0], '#2e7d32', 14))   # salida (verde)
    m.add_marker(CircleMarker(coords[-1], '#c62828', 14))  # resguardo (rojo)
    imagen = m.render()
    buffer = io.BytesIO()
    imagen.save(buffer, format="PNG")
    return buffer.getvalue()

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
if 'rutas_gps' not in st.session_state:
    st.session_state['rutas_gps'] = {}
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
        archivos_historial = st.file_uploader(
            "1. Historial(es) Tracksolid (.xlsx)", type=['xlsx', 'xls'], accept_multiple_files=True,
            help="Puedes subir varios Excel a la vez (por ejemplo 5 archivos con 5 placas cada uno). "
                 "El sistema los une automáticamente antes de procesar, siempre que todos tengan el mismo formato."
        )
        if archivos_historial:
            st.caption(f"📎 {len(archivos_historial)} archivo(s) cargado(s).")
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

                    placas_en_historial = set(df_historial['Placa'].unique()) if 'Placa' in df_historial.columns else set()
                    faltantes = PLACAS_AUTORIZADAS - placas_en_historial
                    if faltantes:
                        st.warning(f"⚠️ No se encontraron datos GPS para: {', '.join(sorted(faltantes))} (quedarán como 'Plantilla Manual').")

                    datos_resumen = []
                    reportes = {}
                    rutas_gps = {}
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

                                    puntos_mapa = []
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

                                        try:
                                            lat_f = float(str(p['Latitud']).replace(',', '.'))
                                            lon_f = float(str(p['Longitud']).replace(',', '.'))
                                            puntos_mapa.append({'lat': lat_f, 'lon': lon_f, 'hora': p['Fecha De Reporte'], 'lugar': lugar or ''})
                                        except (ValueError, TypeError):
                                            pass

                                    if puntos_mapa:
                                        rutas_gps[placa] = puntos_mapa

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
                    st.session_state['rutas_gps'] = rutas_gps
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
        st.info("Puedes editar la columna **RUTA** directamente en la tabla. La pizarra se actualizará automáticamente.")
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

        # --- GENERADOR DE WHATSAPP DINÁMICO (EJECUTIVO / RESUMIDO) ---
        st.subheader("2. Resumen para WhatsApp (Listo para Copiar)")
        msg_w = f"🛰️ *REPORTE EJECUTIVO DE FLOTA PCD - GPS*\n📅 Fecha: {datetime.now().strftime('%d/%m/%Y')}\n\n"
        msg_w += f"📊 *TOTAL VEHÍCULOS:* {len(df_editado)} Unidades\n"
        msg_w += f"🛣️ *GRAN TOTAL RECORRIDO:* {formatear_km(km_total_gral)} Kms\n\n"
        msg_w += f"✅ *Pizarra detallada de rutas adjunta en imagen.*"

        st.code(msg_w, language="markdown")

        st.markdown("---")
        st.subheader("3. Pizarra Ejecutiva (Imagen HD)")

        # --- GENERADOR HTML PARA FOTO: TARJETAS POR MODELO CON FOTOS ---
        fecha_pizarra = datetime.now().strftime('%d/%m/%Y')
        df_agrupado = df_editado.copy()
        df_agrupado['MODELO'] = df_agrupado['MODELO'].astype(str)
        orden_modelos = list(dict.fromkeys(df_agrupado.sort_values('MODELO')['MODELO']))

        tarjetas_html = ""
        for modelo in orden_modelos:
            grupo = df_agrupado[df_agrupado['MODELO'] == modelo].sort_values('PLACA')
            unidades = len(grupo)
            km_modelo = grupo['KM'].sum()
            foto_b64 = obtener_foto_vehiculo(modelo)

            filas_modelo = ""
            for _, r in grupo.iterrows():
                filas_modelo += f"""
                <tr style="border-bottom:1px solid #dde3ec;">
                    <td style="padding:6px 8px; font-weight:700;">{r['PLACA']}</td>
                    <td style="padding:6px 8px;">{r['COLOR']}</td>
                    <td style="padding:6px 8px; text-align:right;">{formatear_km(r['KM'])}</td>
                    <td style="padding:6px 8px; color:#0D47A1; font-weight:600;">{r['RUTA']}</td>
                </tr>"""

            tarjetas_html += f"""
            <div style="background:white; border:1px solid #d5dce8; border-radius:10px; overflow:hidden; box-shadow:0 3px 8px rgba(0,0,0,0.08);">
                <div style="background:#0d47a1; color:white; padding:10px 14px; display:flex; align-items:center; gap:10px;">
                    <img src="{foto_b64}" style="width:46px; height:46px; object-fit:cover; border-radius:6px; background:white;">
                    <div style="font-weight:800; font-size:15px; letter-spacing:.3px;">{modelo}</div>
                </div>
                <div style="display:flex; border-bottom:1px solid #eee;">
                    <div style="flex:1; text-align:center; padding:8px; border-right:1px solid #eee;">
                        <div style="font-size:11px; color:#666; font-weight:600;">UNIDADES</div>
                        <div style="font-size:20px; font-weight:900; color:#0d47a1;">{unidades}</div>
                    </div>
                    <div style="flex:1; text-align:center; padding:8px; background:#f4f6fa;">
                        <div style="font-size:11px; color:#666; font-weight:600;">TOTAL KM</div>
                        <div style="font-size:20px; font-weight:900; color:#F57F17;">{formatear_km(km_modelo)}</div>
                    </div>
                </div>
                <table style="width:100%; border-collapse:collapse; font-size:12.5px;">
                    <thead>
                        <tr style="background:#eef1f7; text-align:left;">
                            <th style="padding:6px 8px;">PLACA</th>
                            <th style="padding:6px 8px;">COLOR</th>
                            <th style="padding:6px 8px; text-align:right;">KM</th>
                            <th style="padding:6px 8px;">RUTA</th>
                        </tr>
                    </thead>
                    <tbody>{filas_modelo}</tbody>
                </table>
            </div>
            """

        logo_b64 = obtener_logo_base64() or ICONO_GENERICO_B64

        html_pizarra_completa = f"""
        <html><head>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        </head><body style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background: #f0f2f6; padding: 20px;">
            <div style="text-align: center; margin-bottom: 15px;">
                <button onclick="capResumen()" style="background: #0d47a1; color: white; border: none; padding: 12px 25px; border-radius: 8px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">📸 DESCARGAR PIZARRA (FOTO HD)</button>
            </div>

            <div id="pizarra-resumen" style="background: #eef1f7; width: 1000px; margin: auto; border: 2px solid #0d47a1; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 20px rgba(0,0,0,0.15); padding-bottom: 4px;">
                <div style="background-color: #0d47a1; padding: 18px 28px; display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid #F57F17;">
                    <div style="display: flex; align-items: center;">
                        <img src="{logo_b64}" style="height: 40px; margin-right: 12px; border-radius:4px;">
                        <div style="color: white; font-size: 20px; font-weight: 900; letter-spacing: 1px;">FLOTA MAZIAD / FLAMINGO / ENCOMIENDAS</div>
                    </div>
                    <div style="color: #F57F17; font-size: 14px; font-weight: bold;">FECHA: {fecha_pizarra}</div>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:14px; padding:18px;">
                    {tarjetas_html}
                </div>

                <div style="background-color: #333; color: white; padding: 15px 30px; display: flex; justify-content: space-between; align-items: center; margin: 0 18px 18px 18px; border-radius: 8px;">
                    <div style="font-size: 16px; font-weight: bold;">TOTAL UNIDADES FLOTA: {len(df_editado)}</div>
                    <div style="font-size: 18px; font-weight: 900; color: #FFD54F;">RECORRIDO TOTAL FLOTA: {formatear_km(km_total_gral)} Kms</div>
                </div>
            </div>

            <script>
                function capResumen() {{
                    html2canvas(document.getElementById('pizarra-resumen'), {{scale: 2, useCORS: true}}).then(canvas => {{
                        canvas.toBlob(function(blob) {{
                            var url = URL.createObjectURL(blob);
                            var link = document.createElement('a');
                            link.download = 'Pizarra_GPS_{datetime.now().strftime('%Y%m%d')}.png';
                            link.href = url;
                            document.body.appendChild(link);
                            link.click();
                            document.body.removeChild(link);
                            URL.revokeObjectURL(url);
                        }}, 'image/png');
                    }});
                }}
            </script>
        </body></html>
        """
        # OJO: no se aplica .replace(',', '.') aquí (como en versiones anteriores) porque
        # rompería la coma del "data:image/png;base64,..." de las fotos y del CSS/JS.
        # Los números ya salen formateados en español porque pasan por formatear_km().
        components.html(html_pizarra_completa, height=1150, scrolling=True)
        modelos_sin_foto = [m for m in orden_modelos if m not in FOTOS_BASE64
                             and not (FOTOS_VEHICULOS.get(m) and (Path(__file__).parent / "fotos" / FOTOS_VEHICULOS[m]).exists())]
        if modelos_sin_foto:
            st.caption(f"ℹ️ Sin foto de referencia para: {', '.join(modelos_sin_foto)} — se usó un ícono genérico. "
                       "Coloca un archivo en `fotos/` con el nombre indicado en `FOTOS_VEHICULOS` para agregarla.")

    else:
        st.info("Sube los archivos y presiona Procesar en la pestaña Configuración.")

# ---------------------------------------------------------
# PESTAÑA 3: REPORTES INDIVIDUALES
# ---------------------------------------------------------
with t_reportes:
    if st.session_state['reportes_texto']:
        if not FOLIUM_DISPONIBLE:
            st.warning("El paquete `folium` no está instalado, así que no se puede dibujar el trazado GPS. "
                       "Instálalo con `pip install folium` para activar los mapas automáticos.")

        numero_destino = st.text_input("Número de WhatsApp destino (con código de país, sin '+'):", value=NUMERO_WHATSAPP_DEFECTO)

        placas = list(st.session_state['reportes_texto'].keys())

        # --- Botón para enviar TODOS los reportes de una vez ---
        st.markdown("### 📤 Envío masivo")
        st.caption("Abre un chat de WhatsApp con el texto ya cargado por cada placa, uno detrás de otro. "
                   "El navegador puede pedirte permitir las ventanas emergentes (pop-ups) la primera vez — solo dale 'Permitir'. "
                   "WhatsApp no permite adjuntar imágenes automáticamente por este método: el mapa de cada unidad lo descargas aparte y lo adjuntas tú en el chat que se abre.")

        reportes_js = json.dumps([
            {"placa": p, "texto": st.session_state['reportes_texto'][p]}
            for p in placas
        ])

        html_enviar_todos = f"""
        <div style="text-align:left;">
            <button id="btn-enviar-todos" style="background:#25D366; color:white; padding:12px 22px; border:none;
                border-radius:8px; font-weight:bold; cursor:pointer; font-size:15px;">
                📲 ENVIAR LOS {len(placas)} REPORTES POR WHATSAPP
            </button>
        </div>
        <script>
            const reportes = {reportes_js};
            const numero = "{numero_destino.strip()}";
            document.getElementById('btn-enviar-todos').addEventListener('click', function() {{
                reportes.forEach(function(r, i) {{
                    setTimeout(function() {{
                        const url = "https://wa.me/" + numero + "?text=" + encodeURIComponent(r.texto);
                        window.open(url, '_blank');
                    }}, i * 1200);
                }});
            }});
        </script>
        """
        components.html(html_enviar_todos, height=70)

        st.markdown("---")
        st.markdown("### 📝 Reportes por placa")

        tabs_placas = st.tabs(placas)
        for i, placa in enumerate(placas):
            with tabs_placas[i]:
                texto = st.text_area("Reporte Generado (Editable):", value=st.session_state['reportes_texto'][placa], height=350, key=f"txt_{placa}")

                puntos_ruta = st.session_state['rutas_gps'].get(placa)
                mapa_html = generar_mapa_folium(placa, puntos_ruta) if puntos_ruta else None

                col_mapa, col_envio = st.columns([2, 1])
                with col_mapa:
                    if mapa_html:
                        st.markdown("**Trazado GPS (generado automáticamente desde las coordenadas del Excel):**")
                        components.html(mapa_html, height=420, scrolling=False)

                        if STATICMAP_DISPONIBLE:
                            # Se genera bajo demanda (no automáticamente en cada pestaña) para no
                            # disparar 12 descargas de tiles de OpenStreetMap en cada recarga.
                            cache_key = f"mapa_png_{placa}"
                            if cache_key not in st.session_state:
                                if st.button("🖼️ Generar imagen del mapa", key=f"gen_mapa_{placa}", use_container_width=True):
                                    with st.spinner("Generando imagen del mapa (puede tardar unos segundos)..."):
                                        try:
                                            st.session_state[cache_key] = generar_mapa_estatico_png(puntos_ruta)
                                        except Exception as e:
                                            st.warning(f"No se pudo generar la imagen del mapa: {e}")
                            if st.session_state.get(cache_key):
                                st.download_button(
                                    "📸 Descargar mapa (PNG)",
                                    data=st.session_state[cache_key],
                                    file_name=f"Mapa_{placa}_{datetime.now().strftime('%Y%m%d')}.png",
                                    mime="image/png",
                                    key=f"dl_mapa_{placa}",
                                    use_container_width=True,
                                )
                        else:
                            st.caption("Instala `pip install staticmap` para poder descargar el mapa como imagen PNG.")
                    elif FOLIUM_DISPONIBLE:
                        st.info("No hay puntos de ruta suficientes para dibujar el trazado de esta unidad (posible 'sin movimiento').")

                with col_envio:
                    texto_url = urllib.parse.quote(texto)
                    url_wa = f"https://wa.me/{numero_destino.strip()}?text={texto_url}"
                    st.markdown(f"""
                    <a href="{url_wa}" target="_blank" style="text-decoration:none;">
                        <button style="background-color:#25D366; color:white; padding:10px 20px; border:none; border-radius:5px; font-weight:bold; cursor:pointer; width:100%; margin-top:10px;">
                            📲 ENVIAR ESTE REPORTE
                        </button>
                    </a>
                    """, unsafe_allow_html=True)
                    if mapa_html:
                        st.caption("1️⃣ Descarga el mapa → 2️⃣ Envía este reporte (se abre el chat) → 3️⃣ Adjunta la imagen descargada en WhatsApp.")
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
