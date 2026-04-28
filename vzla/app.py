# ==========================================
# Archivo: app.py (Pizarra Central Drotaca - PCD)
# ==========================================
import streamlit as st
import pandas as pd
import re
import json
import difflib # <-- LIBRERÍA NATIVA NUEVA PARA COMPARAR NOMBRES
from datetime import datetime
import streamlit.components.v1 as components

# LIBRERÍAS PARA GOOGLE SHEETS (MODERNIZADAS)
import gspread
import textwrap
import traceback
from google.oauth2.service_account import Credentials

# ==========================================
# CREDENCIALES Y CONEXIÓN DINÁMICA A GOOGLE SHEETS
# ==========================================
# 1. Cargamos el diccionario de la "Caja Fuerte"
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

# 2. RECONSTRUCTOR BLINDADO DE LLAVE
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"

CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

# 3. Función de conexión moderna
def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

# ==========================================
# FUNCION PARA GUARDAR EN GOOGLE SHEETS DIRECTO
# ==========================================
def guardar_en_google_sheets_directo(df_para_guardar):
    try:
        cliente = obtener_cliente_sheets()
        
        # Conexión directa a prueba de balas usando tu ID
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        
        try:
            hoja = doc.worksheet("PIZARRA_TRAFICO")
        except Exception:
            hoja = doc.add_worksheet(title="PIZARRA_TRAFICO", rows=1000, cols=20)
            hoja.append_row(list(df_para_guardar.columns))

        # Convertir DataFrame a lista de listas (para Google Sheets)
        df_clean = df_para_guardar.fillna('').astype(str)
        valores = df_clean.values.tolist()
        
        try:
            hoja.append_rows(valores, value_input_option='USER_ENTERED')
        except Exception as error_interno:
            if "200" in str(error_interno):
                pass # Silenciador del falso error 200
            else:
                raise error_interno
                
        return True, "✅ Datos guardados en PCD_BaseDatos correctamente."
        
    except Exception as e:
        st.error(f"🛑 Falla en la conexión: {e}")
        st.code(traceback.format_exc(), language="python")
        return False, f"❌ Error al conectar con Google Sheets: {e}"

# ==========================================
# MAPEO REFERENCIAL DE ZONAS
# ==========================================
MAPEO_ZONAS_REFERENCIAL = {
    'CARACAS 1': 'CENTRO', 
    'CARACAS 2': 'CENTRO', 
    'CARACAS 3': 'CENTRO', 
    'CARACAS 4-5': 'CENTRO',
    'ARAGUA 1': 'CENTRO', 
    'ARAGUA 2-SAN JUAN': 'CENTRO',
    'CARABOBO 1': 'CENTRO-OCCIDENTE', 
    'CARABOBO 2 - COJEDES': 'CENTRO-OCCIDENTE', 
    'CARABOBO 3': 'CENTRO-OCCIDENTE',
    'LARA 1 - PORTUGUESA 1': 'OCCIDENTE', 
    'LARA 2 - YARACUY': 'OCCIDENTE',
    'PORTUGUESA 2 - BARINAS 1-2': 'OCCIDENTE', 
    'PORTUGUESA 2': 'OCCIDENTE', 
    'BARINAS 3': 'OCCIDENTE', 
    'FALCON': 'OCCIDENTE',
    'MARACAIBO 1': 'OCCIDENTE', 
    'MARACAIBO 2': 'OCCIDENTE', 
    'CABIMAS-CIUDAD OJEDA': 'OCCIDENTE',
    'MERIDA 1-TRUJILLO-PORTUGUESA 3': 'OCCIDENTE', 
    'MERIDA 2-TACHIRA': 'OCCIDENTE',
    'MATURIN-PUNTA DE MATA': 'ORIENTE', 
    'BARCELONA': 'ORIENTE', 
    'BARCELONA - CLARINES': 'ORIENTE',
    'PUERTO ORDAZ-SAN FELIX-UPATA': 'ORIENTE',
    'PUERTO ORDAZ - SAN FELIX UPATA': 'ORIENTE', 
    'PUERTO ORDAZ': 'ORIENTE', 
    'CARÚPANO-GUIRIA': 'ORIENTE',
    'TUMEREMO': 'ORIENTE',
    'CUMANÁ': 'ORIENTE',
    'BOLIVAR': 'ORIENTE',
    'PTO ORDAZ': 'ORIENTE',
    'DELTA AMACURO': 'ORIENTE',
    'ANACO-CANTAURA': 'ORIENTE',
    'MATURIN': 'ORIENTE',
    'GUARICO': 'ORIENTE',
    'CUMANA - CUMANACOA': 'ORIENTE',
    'CARUPANO - GUIRIA': 'ORIENTE',
    'NUEVA ESPARTA': 'ORIENTE', 
    'ARAGUA 2 - SAN JUAN DE LOS MORROS': 'CENTRO',
    'BARINAS 03': 'OCCIDENTE',
    'CORO - PUNTO FIJO': 'OCCIDENTE',
    'CABIMAS - CIUDAD OJEDA': 'OCCIDENTE',
    'MERIDA 1 - TRUJILLO - PORTUGUESA 3': 'OCCIDENTE',
    'MERIDA 2 - TACHIRA': 'OCCIDENTE'
}

def mapear_zona(ruta):
    ruta_limpia = str(ruta).upper().strip()
    ruta_limpia = ruta_limpia.replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
    
    if ruta_limpia in MAPEO_ZONAS_REFERENCIAL: 
        return MAPEO_ZONAS_REFERENCIAL[ruta_limpia]
        
    for key, val in MAPEO_ZONAS_REFERENCIAL.items():
        if key in ruta_limpia: 
            return val
            
    return 'NO CLASIFICADO'

# ==========================================
# 1. EXTRACCIÓN ULTRA RÁPIDA (PYTHON REGEX)
# ==========================================
def procesar_trafico_python(raw_text):
    try:
        raw_text = raw_text.replace('\u200e', '').replace('\u200f', '')
        
        tiempos = {
            "Fecha_Reporte": "", 
            "Hora_1er_listin": "", 
            "Hora_ultimo_listin": "", 
            "Inicio_trafico": "", 
            "Culminacion_trafico": ""
        }
        
        m_fecha = re.search(r'Reporte de Rutas:\s*(\d{2}/\d{2}/\d{4})', raw_text, re.IGNORECASE)
        if m_fecha: 
            tiempos['Fecha_Reporte'] = m_fecha.group(1)
            
        m_1er = re.search(r'1er list[ií]n:\s*(.*?)(?=\n|$)', raw_text, re.IGNORECASE)
        if m_1er: 
            tiempos['Hora_1er_listin'] = m_1er.group(1).strip()
            
        m_ult = re.search(r'[uú]ltimo list[ií]n:\s*(.*?)(?=\n|$)', raw_text, re.IGNORECASE)
        if m_ult: 
            tiempos['Hora_ultimo_listin'] = m_ult.group(1).strip()
            
        m_ini = re.search(r'Inicio de tr[aá]fico:\s*(.*?)(?=\n|$)', raw_text, re.IGNORECASE)
        if m_ini: 
            tiempos['Inicio_trafico'] = m_ini.group(1).strip()
            
        m_cul = re.search(r'Culminaci[oó]n de Tr[aá]fico:\s*(.*?)(?=\n|$)', raw_text, re.IGNORECASE)
        if m_cul: 
            tiempos['Culminacion_trafico'] = m_cul.group(1).strip()

        lineas = [l.strip() for l in raw_text.split('\n') if l.strip()]
        rutas_extraidas = []
        
        for i, linea in enumerate(lineas):
            if "listin" in linea.lower() and ":" in linea and "hora" not in linea.lower():
                if i > 0:
                    nombre_ruta = lineas[i-1].replace(':', '').replace('*', '').strip()
                    nombre_ruta = re.sub(r'^[^\w]+', '', nombre_ruta).strip() 
                    
                    listines_raw = linea.split(':', 1)[-1].strip()
                    listines = re.split(r'[-,\s]+', listines_raw)
                    
                    bultos, farmacias = None, None
                    chofer, ayudante, unidad = None, None, None
                    encomiendas, reposiciones = None, None
                    
                    for j in range(i+1, min(i+12, len(lineas))):
                        l_lower = lineas[j].lower()
                        
                        if "listin" in l_lower and ":" in l_lower and "hora" not in l_lower:
                            break
                        
                        if "bultos" in l_lower and "total" not in l_lower and bultos is None:
                            num = re.search(r'\d+', lineas[j])
                            if num: bultos = int(num.group())
                        elif "farmacia" in l_lower and "total" not in l_lower and farmacias is None:
                            num = re.search(r'\d+', lineas[j])
                            if num: farmacias = int(num.group())
                        elif ("chofer" in l_lower or "chófer" in l_lower) and chofer is None:
                            chofer = lineas[j].split(':', 1)[-1].replace('*', '').strip()
                        elif "ayudante" in l_lower and ayudante is None:
                            ayudante = lineas[j].split(':', 1)[-1].replace('*', '').strip()
                        elif "unidad" in l_lower and unidad is None:
                            unidad = lineas[j].split(':', 1)[-1].replace('*', '').strip()
                        elif "encomienda" in l_lower and encomiendas is None:
                            encomiendas = lineas[j].split(':', 1)[-1].replace('*', '').strip()
                        elif ("reposicion" in l_lower or "reposición" in l_lower) and reposiciones is None:
                            reposiciones = lineas[j].split(':', 1)[-1].replace('*', '').strip()
                            
                    rutas_extraidas.append({
                        "Ruta": nombre_ruta,
                        "Listines": [l for l in listines if l],
                        "Bultos_Total": bultos if bultos is not None else 0,
                        "Farmacias_Total": farmacias if farmacias is not None else 0,
                        "Chofer": chofer if chofer is not None else "",
                        "Ayudante": ayudante if ayudante is not None else "",
                        "Unidad": unidad if unidad is not None else "",
                        "Encomiendas": encomiendas if encomiendas is not None else "",
                        "Reposiciones": reposiciones if reposiciones is not None else ""
                    })
                    
        df = pd.DataFrame(rutas_extraidas)
        if not df.empty:
            df['Zona'] = df['Ruta'].apply(mapear_zona)
            df['Bultos_Total'] = pd.to_numeric(df['Bultos_Total'], errors='coerce').fillna(0).astype(int)
            df['Farmacias_Total'] = pd.to_numeric(df['Farmacias_Total'], errors='coerce').fillna(0).astype(int)
            return df, tiempos, "OK"
        else:
            return None, None, "No se encontraron rutas en el texto."
            
    except Exception as e:
        return None, None, str(e)

# ==========================================
# 2. ANÁLISIS MATEMÁTICO INMEDIATO
# ==========================================
def generar_analisis_python(df_datos, tiempos, texto_original):
    t_1 = tiempos.get('Hora_1er_listin', 'N/A')
    t_2 = tiempos.get('Hora_ultimo_listin', 'N/A')
    t_3 = tiempos.get('Inicio_trafico', 'N/A')
    t_4 = tiempos.get('Culminacion_trafico', 'N/A')

    top3 = df_datos.nlargest(3, 'Bultos_Total')
    top_rutas_str = ""
    for i, row in enumerate(top3.itertuples(), 1):
        top_rutas_str += f"{i}. {row.Ruta} con {row.Bultos_Total} bultos.\n"

    total_calculado = int(df_datos['Bultos_Total'].sum())
    match_total = re.search(r'Cargados[^\d]*(\d+)', texto_original, re.IGNORECASE)
    
    if match_total:
        total_reportado = int(match_total.group(1))
        if total_calculado == total_reportado:
            auditoria = f"✅ *AUDITORÍA EXITOSA:* El total sumado ruta por ruta ({total_calculado} bultos) cuadra perfectamente con el reporte del supervisor."
        else:
            auditoria = f"🚨 *ALERTA DE DESCUADRE:* La suma de la tabla da {total_calculado} bultos, pero el supervisor reportó {total_reportado} al final. ¡Revisar!"
    else:
        auditoria = f"⚠️ *ATENCIÓN:* El sistema sumó {total_calculado} bultos. No se encontró el 'Total Cargados' al final del mensaje para comparar."

    seccion_extras = ""
    extras_list = []
    for i, row in df_datos.iterrows():
        enc = str(row.get('Encomiendas', '')).strip()
        rep = str(row.get('Reposiciones', '')).strip()
        
        has_enc = bool(enc) and enc.lower() not in ['no', 'ninguna', 'n/a', 'na', 'null', '']
        has_rep = bool(rep) and rep.lower() not in ['no', 'ninguna', 'n/a', 'na', 'null', '']
        
        if has_enc or has_rep:
            item_str = f"📍 *{row['Ruta']}* (Chófer: {row['Chofer']} | Unidad: {row['Unidad']})"
            if has_enc: 
                item_str += f"\n  📦 Encomienda: {enc}"
            if has_rep: 
                item_str += f"\n  🔄 Reposición: {rep}"
            extras_list.append(item_str)
            
    if extras_list:
        seccion_extras = "\n\n⚠️ *ENCOMIENDAS Y REPOSICIONES:*\n" + "\n\n".join(extras_list)

    mensaje = f"""Buenas tardes Gerente, adjunto el reporte de salidas de hoy.

⏱️ *TIEMPOS OPERATIVOS:*
* Primer listín registrado: {t_1}
* Último listín registrado: {t_2}
* Inicio de tráfico: {t_3}
* Culminación de tráfico: {t_4}

{auditoria}

🏆 *DATOS RELEVANTES:*
Las rutas con mayor volumen de bultos son:
{top_rutas_str.strip()}{seccion_extras}"""

    return mensaje

# ==========================================
# GENERADORES HTML (PIZARRA CORPORATIVA)
# ==========================================
def formatear_fecha_extraida(fecha_str):
    if not fecha_str: return datetime.now().strftime("%d/%m/%Y")
    try:
        match = re.search(r'\d{2}/\d{2}/\d{4}', fecha_str)
        if match:
            fecha_limpia = match.group(0)
            obj_fecha = datetime.strptime(fecha_limpia, "%d/%m/%Y")
            dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            return f"{dias[obj_fecha.weekday()]}, {fecha_limpia}"
        return fecha_str
    except: 
        return fecha_str

def obtener_html_base(df, tiempos, filas_html):
    total_bultos = df['Bultos_Total'].sum()
    total_farmacias = df['Farmacias_Total'].sum()
    total_rutas = len(df)
    fecha_reporte = formatear_fecha_extraida(tiempos.get('Fecha_Reporte', '') if isinstance(tiempos, dict) else '')

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div id="reporte-drotaca" style="font-family: Arial, sans-serif; width: 1000px; margin: auto; background-color: #fff; border: 1px solid #000; padding-bottom: 20px; height: auto;">
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 15px 20px; border-bottom: 3px solid #004080;">
            <div><h1 style="color: #004080; margin: 0; font-size: 24px;">Drotaca</h1><span style="font-size: 11px; color: #000;">¡Conectamos con la salud!</span></div>
            <div style="text-align: center;"><h2 style="margin: 0; color: #004080; font-size: 18px;">REPORTE DE TRÁFICO Y SALIDAS</h2><p style="margin: 5px 0 0; font-size: 13px; color: #000;"><b>PERÍODO:</b> {fecha_reporte}</p></div>
        </div>
        <div style="background-color: #e6f2ff; border-bottom: 1px solid #000; padding: 8px 20px; font-size: 13px; color: #004080; display: flex; justify-content: space-between; font-weight: bold;">
            <span>🕜 Hora 1er listín: {tiempos.get('Hora_1er_listin', 'N/A')}</span>
            <span>🕜 Hora de último listín: {tiempos.get('Hora_ultimo_listin', 'N/A')}</span>
            <span>🕜 Inicio de tráfico: {tiempos.get('Inicio_trafico', 'N/A')}</span>
            <span>🕜 Culminación de Tráfico: {tiempos.get('Culminacion_trafico', 'N/A')}</span>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 15px 20px; background-color: #f9f9f9; border-bottom: 1px solid #000;">
            <div style="width: 23%; text-align: center; border: 1px solid #000; padding: 10px;"><b>BULTOS</b><h3 style="margin:0; font-size: 22px; color: #004080;">{total_bultos}</h3></div>
            <div style="width: 23%; text-align: center; border: 1px solid #000; padding: 10px;"><b>FARMACIAS</b><h3 style="margin:0; font-size: 22px;">{total_farmacias}</h3></div>
            <div style="width: 23%; text-align: center; border: 1px solid #000; padding: 10px;"><b>RUTAS</b><h3 style="margin:0; font-size: 22px;">{total_rutas}</h3></div>
            <div style="width: 23%; text-align: center; border: 1px solid #000; padding: 10px;"><b>PROM. BULTOS</b><h3 style="margin:0; font-size: 22px;">{round(total_bultos/total_rutas,1) if total_rutas>0 else 0}</h3></div>
        </div>
        <div style="padding: 15px 20px 0 20px;">
            <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
                <thead><tr style="background-color: #004080; color: white; font-size: 12px;">
                    <th style="padding: 8px; border: 1px solid #000; width: 18%;">RUTA / ZONA</th>
                    <th style="padding: 8px; border: 1px solid #000; width: 30%;">INFO UNIDAD/PERSONAL</th>
                    <th style="padding: 8px; border: 1px solid #000; width: 32%;">LISTINES</th>
                    <th style="padding: 8px; border: 1px solid #000; width: 10%; text-align: center;">FARMA.</th>
                    <th style="padding: 8px; border: 1px solid #000; width: 10%; text-align: center; background-color: #003366;">BULTOS</th>
                </tr></thead>
                <tbody>{filas_html}</tbody>
            </table>
        </div>
    </div>
    <script>
    function descargarImagen() {{
        html2canvas(document.getElementById('reporte-drotaca'), {{scale: 2, useCORS: true}}).then(canvas => {{
            var link = document.createElement('a');
            link.download = 'Reporte_Trafico.png'; link.href = canvas.toDataURL(); link.click();
        }});
    }}
    </script>
    """

def generar_filas_agrupadas(df):
    html = ""
    for zona in sorted(df['Zona'].unique()):
        df_z = df[df['Zona'] == zona]
        html += f'<tr style="background-color: #00264d; color: white; font-weight: bold;"><td colspan="5" style="padding: 8px; border: 1px solid #000;">📍 ZONA {zona} <span style="float:right;">Rutas: {len(df_z)} | Bultos: {df_z["Bultos_Total"].sum()}</span></td></tr>'
        for i, r in df_z.reset_index().iterrows():
            bg = "#ffffff" if i % 2 == 0 else "#f2f2f2"
            listines_str = ", ".join(r["Listines"]) if isinstance(r["Listines"], list) else str(r["Listines"])
            
            html += f'<tr style="background-color: {bg}; color: #000; font-size: 11px;">'
            html += f'<td style="padding: 6px; border: 1px solid #000;"><b>{r["Ruta"]}</b></td>'
            html += f'<td style="padding: 6px; border: 1px solid #000;"><b>Chófer:</b> {r["Chofer"]} <span style="color:#888;">|</span> <b>Ayudante:</b> {r["Ayudante"]}<br><b>Unidad:</b> {r["Unidad"]}</td>'
            html += f'<td style="padding: 6px; border: 1px solid #000;">{listines_str}</td>'
            html += f'<td style="padding: 6px; border: 1px solid #000; text-align: center; font-size: 16px; font-weight: bold; color: #004080;">{r["Farmacias_Total"]}</td>'
            html += f'<td style="padding: 6px; border: 1px solid #000; text-align: center; font-size: 16px; font-weight: bold; color: #000; background-color: #ffffcc;">{r["Bultos_Total"]}</td></tr>'
    return html

def generar_filas_plana(df):
    html = ""
    for i, r in df.iterrows():
        bg = "#ffffff" if i % 2 == 0 else "#f2f2f2"
        listines_str = ", ".join(r["Listines"]) if isinstance(r["Listines"], list) else str(r["Listines"])
        
        html += f'<tr style="background-color: {bg}; color: #000; font-size: 11px;">'
        html += f'<td style="padding: 6px; border: 1px solid #000;"><b>{r["Ruta"]}</b><br><b>{r["Zona"]}</b></td>'
        html += f'<td style="padding: 6px; border: 1px solid #000;"><b>Chófer:</b> {r["Chofer"]} <span style="color:#888;">|</span> <b>Ayudante:</b> {r["Ayudante"]}<br><b>Unidad:</b> {r["Unidad"]}</td>'
        html += f'<td style="padding: 6px; border: 1px solid #000;">{listines_str}</td>'
        html += f'<td style="padding: 6px; border: 1px solid #000; text-align: center; font-size: 16px; font-weight: bold; color: #004080;">{r["Farmacias_Total"]}</td>'
        html += f'<td style="padding: 6px; border: 1px solid #000; text-align: center; font-size: 16px; font-weight: bold; color: #000; background-color: #ffffcc;">{r["Bultos_Total"]}</td></tr>'
    return html


# ==========================================
# FUNCIONES: MÓDULO 2 (EXCEL PARSER Y TABLA VERDE)
# ==========================================
def obtener_listines_unicos(serie):
    conjunto_listines = set()
    for texto in serie.dropna():
        txt = str(texto).replace('.0', '').strip()
        partes = re.split(r'[,\-;\n]+', txt)
        for p in partes:
            p_limpio = p.strip()
            if p_limpio:
                conjunto_listines.add(p_limpio)
    return sorted(list(conjunto_listines))

def procesar_excel_base(df, fecha_buscada):
    # 1. LIMPIEZA BRUTAL Y AUTODETECCIÓN DE CABECERAS
    def limpiar_cols(columnas):
        cols = [str(c).upper().replace('\n', ' ').replace('\r', ' ').replace('\xa0', ' ').strip() for c in columnas]
        cols = [" ".join(c.split()) for c in cols]
        cols = [c.replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U') for c in cols]
        return cols

    df.columns = limpiar_cols(df.columns)
    
    # El Sabueso
    if 'FECHA DE LISTIN' not in df.columns and 'CHOFER' not in df.columns:
        for i, row in df.iterrows():
            fila_str = " ".join(row.astype(str).str.upper())
            if 'CHOFER' in fila_str and 'BULTOS' in fila_str:
                df.columns = limpiar_cols(row.values)
                df = df.iloc[i+1:].reset_index(drop=True)
                break

    if 'FECHA DE LISTIN' not in df.columns:
        cols_detectadas = ", ".join(df.columns.tolist()[:10])
        return None, f"Error: No encontré la columna 'FECHA DE LISTIN'. Columnas detectadas: {cols_detectadas}..."

    # 2. Filtrar por fecha exacta
    df['FECHA DE LISTIN'] = pd.to_datetime(df['FECHA DE LISTIN'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
    fecha_str = fecha_buscada.strftime('%d/%m/%Y')
    df_filtrado = df[df['FECHA DE LISTIN'] == fecha_str].copy()

    if df_filtrado.empty:
        return None, f"No se encontraron despachos para la fecha {fecha_str} en el archivo."

    col_chofer = 'CHOFER' if 'CHOFER' in df_filtrado.columns else [c for c in df_filtrado.columns if 'CHOFER' in c][0]
    col_bultos = 'BULTOS' if 'BULTOS' in df_filtrado.columns else [c for c in df_filtrado.columns if 'BULTOS' in c][0]
    col_codigo = 'CODIGO' if 'CODIGO' in df_filtrado.columns else [c for c in df_filtrado.columns if 'CODIGO' in c or 'FARMACIA' in c][0]
    
    try: 
        col_listin = [c for c in df_filtrado.columns if 'LISTIN' in c and 'FECHA' not in c][0]
    except: 
        col_listin = None

    # 3. Exclusiones obligatorias
    df_filtrado = df_filtrado[~df_filtrado[col_chofer].astype(str).str.contains('RETIRO POR OFICINA|RETIRAR|GOTICA', case=False, na=False)]
    mask_tigre = df_filtrado.astype(str).apply(lambda x: x.str.contains('EL TIGRE', case=False, na=False)).any(axis=1)
    df_filtrado = df_filtrado[~mask_tigre]

    df_filtrado[col_bultos] = pd.to_numeric(df_filtrado[col_bultos], errors='coerce').fillna(0)

    # 4. Agrupar la Tabla Dinámica
    if col_listin:
        resumen = df_filtrado.groupby(col_chofer).agg(
            Cant_Listines=(col_listin, lambda x: len(obtener_listines_unicos(x))),
            Num_Listines=(col_listin, lambda x: ", ".join(obtener_listines_unicos(x))),
            Farmacias=(col_codigo, 'nunique'),
            Bultos=(col_bultos, 'sum')
        ).reset_index()
        resumen.columns = ['CHOFER', 'Cant_Listines', 'Num_Listines', 'Farmacias', 'Bultos']
    else:
        resumen = df_filtrado.groupby(col_chofer).agg(
            Farmacias=(col_codigo, 'nunique'),
            Bultos=(col_bultos, 'sum')
        ).reset_index()
        resumen['Cant_Listines'] = 0
        resumen['Num_Listines'] = "-"
        resumen = resumen[['CHOFER', 'Cant_Listines', 'Num_Listines', 'Farmacias', 'Bultos']]
    
    resumen = resumen.sort_values(by='CHOFER').reset_index(drop=True)
    return resumen, "OK"

def generar_html_tabla_verde(df_resumen):
    total_listines = df_resumen['Cant_Listines'].sum()
    total_farmacias = df_resumen['Farmacias'].sum()
    total_bultos = df_resumen['Bultos'].sum()
    
    filas_html = ""
    for i, row in df_resumen.iterrows():
        bg_color = "#92d050" if i % 2 == 0 else "#a9d18e" 
        filas_html += f"""
        <tr style="background-color: {bg_color}; color: #000; font-family: Arial, sans-serif; font-size: 14px;">
            <td style="padding: 8px; border: 1px solid #000; text-transform: uppercase;">{row['CHOFER']}</td>
            <td style="padding: 8px; border: 1px solid #000; text-align: center;">{row['Cant_Listines']}</td>
            <td style="padding: 8px; border: 1px solid #000; text-align: left; font-size: 12px; font-weight: normal;">{row['Num_Listines']}</td>
            <td style="padding: 8px; border: 1px solid #000; text-align: right;">{row['Farmacias']}</td>
            <td style="padding: 8px; border: 1px solid #000; text-align: right;">{row['Bultos']}</td>
        </tr>
        """
        
    html = f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div id="tabla-verde-drotaca" style="width: 800px; margin: auto; background-color: #fff; padding-bottom: 5px;">
        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
            <thead>
                <tr style="background-color: #add8e6; color: #000; font-family: Arial, sans-serif; font-size: 14px; font-weight: bold;">
                    <th style="padding: 8px; border: 1px solid #000; text-align: left; width: 20%;">Etiquetas de fila 🔽</th>
                    <th style="padding: 8px; border: 1px solid #000; text-align: center; width: 10%;">Cant. Listines</th>
                    <th style="padding: 8px; border: 1px solid #000; text-align: left; width: 40%;">Números de Listines</th>
                    <th style="padding: 8px; border: 1px solid #000; text-align: right; width: 15%;">Cuenta de CODIGO</th>
                    <th style="padding: 8px; border: 1px solid #000; text-align: right; width: 15%;">Suma de BULTOS</th>
                </tr>
            </thead>
            <tbody>
                {filas_html}
                <tr style="background-color: #add8e6; color: #000; font-family: Arial, sans-serif; font-size: 15px; font-weight: bold;">
                    <td style="padding: 8px; border: 1px solid #000;">Total general</td>
                    <td style="padding: 8px; border: 1px solid #000; text-align: center;">{total_listines}</td>
                    <td style="padding: 8px; border: 1px solid #000; text-align: center;">-</td>
                    <td style="padding: 8px; border: 1px solid #000; text-align: right;">{total_farmacias}</td>
                    <td style="padding: 8px; border: 1px solid #000; text-align: right;">{total_bultos}</td>
                </tr>
            </tbody>
        </table>
    </div>
    <script>
    function descargarTablaVerde() {{
        html2canvas(document.getElementById('tabla-verde-drotaca'), {{scale: 2, useCORS: true}}).then(canvas => {{
            var link = document.createElement('a');
            link.download = 'Resumen_Previo_Despacho.png'; link.href = canvas.toDataURL(); link.click();
        }});
    }}
    </script>
    """
    return html

# ==========================================
# 3. AUDITORÍA CRUZADA (MEJORADA CON FUZZY MATCHING)
# ==========================================
def fuzzy_match_names(name_wa, list_excel_names):
    def prep(n):
        n = str(n).upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U')
        return re.sub(r'[^A-Z\s]', '', n).strip()
        
    n1 = prep(name_wa)
    words1 = set(n1.split())
    
    best_match = None
    best_score = 0
    
    for name_ex in list_excel_names:
        n2 = prep(name_ex)
        words2 = set(n2.split())
        
        interseccion = words1.intersection(words2)
        
        n1_sorted = " ".join(sorted(n1.split()))
        n2_sorted = " ".join(sorted(n2.split()))
        score_sorted = difflib.SequenceMatcher(None, n1_sorted, n2_sorted).ratio()
        
        score_direct = difflib.SequenceMatcher(None, n1, n2).ratio()
        
        score_final = max(score_sorted, score_direct)
        
        if len(interseccion) >= 2:
            score_final = max(score_final, 0.85)
            
        if score_final > best_score:
            best_score = score_final
            best_match = name_ex
            
    if best_score >= 0.75: 
        return best_match
    return None

def realizar_auditoria_cruzada(df_wa, df_ex):
    wa_grouped = df_wa.groupby('Chofer').agg({
        'Farmacias_Total': 'sum', 
        'Bultos_Total': 'sum',
        'Ruta': lambda x: " + ".join(x.unique())
    }).reset_index()
    
    ex_grouped = df_ex.copy()
    excel_names_list = ex_grouped['CHOFER'].tolist()
    
    wa_to_ex_map = {}
    matched_ex_names = set()
    
    for _, row in wa_grouped.iterrows():
        match = fuzzy_match_names(row['Chofer'], excel_names_list)
        if match:
            wa_to_ex_map[row['Chofer']] = match
            matched_ex_names.add(match)

    resultados = []
    
    for _, row in wa_grouped.iterrows():
        wa_name = row['Chofer']
        wa_b = int(row['Bultos_Total'])
        wa_f = int(row['Farmacias_Total'])
        
        ex_name = wa_to_ex_map.get(wa_name)
        
        if ex_name:
            ex_row = ex_grouped[ex_grouped['CHOFER'] == ex_name].iloc[0]
            ex_b = int(ex_row['Bultos'])
            ex_f = int(ex_row['Farmacias'])
            
            diff_bultos = wa_b - ex_b
            diff_farm = wa_f - ex_f
            
            if diff_bultos == 0 and diff_farm == 0:
                estado = "✅ Cuadre Perfecto"
                detalle = f"Ambos coinciden exactamente: {ex_b} bultos y {ex_f} farmacias."
            else:
                estado = "❌ Discrepancia"
                errores = ["🚨 ALERTA: Excel manda."]
                if diff_bultos != 0:
                    errores.append(f"📦 Excel dice: {ex_b} Bultos | WhatsApp dice: {wa_b}")
                if diff_farm != 0:
                    errores.append(f"🏥 Excel dice: {ex_f} Farmacias | WhatsApp dice: {wa_f}")
                
                detalle = " - ".join(errores)
            
            resultados.append({
                "Chófer": f"{wa_name} (WA) -> {ex_name} (Excel)",
                "Estado": estado,
                "Detalles de Auditoría": detalle
            })
        else:
            resultados.append({
                "Chófer": wa_name,
                "Estado": "🚨 Falta en Excel",
                "Detalles de Auditoría": f"Asignado a la ruta {row['Ruta']}. WA indica {wa_b} bultos pero no aparece en Excel."
            })
            
    for _, row in ex_grouped.iterrows():
        ex_name = row['CHOFER']
        if ex_name not in matched_ex_names:
            resultados.append({
                "Chófer": ex_name,
                "Estado": "⚠️ Solo en Excel",
                "Detalles de Auditoría": f"El Excel reporta {row['Bultos']} bultos y {row['Farmacias']} farmacias. No aparece en el mensaje de WhatsApp."
            })
            
    df_resultado = pd.DataFrame(resultados)
    orden_prioridad = {"❌ Discrepancia": 1, "🚨 Falta en Excel": 2, "⚠️ Solo en Excel": 3, "✅ Cuadre Perfecto": 4}
    df_resultado['Prioridad'] = df_resultado['Estado'].map(orden_prioridad)
    return df_resultado.sort_values(by='Prioridad').drop(columns=['Prioridad'])

# ==========================================
# INTERFAZ PRINCIPAL (TABS)
# ==========================================
st.title("📊 PCD - Centro de Control Integrado")

tab1, tab2 = st.tabs(["📲 Módulo 1: Pizarra Rápida (WhatsApp)", "📊 Módulo 2: Auditoría Avanzada (Excel)"])

# ---------------------------------------------------------
# PESTAÑA 1: EL FLUJO QUE YA TIENES (INTACTO Y EDITABLE)
# ---------------------------------------------------------
with tab1:
    if 'df_trafico' not in st.session_state:
        st.session_state['df_trafico'] = None

    col_in, col_out = st.columns([1, 2])

    with col_in:
        texto = st.text_area("Pega el texto de WhatsApp aquí:", height=300)
        procesar = st.button("⚡ Procesar Datos (Instantáneo / 100% Gratis)", type="primary", use_container_width=True)

    if procesar and texto:
        with st.spinner("Procesando datos instantáneamente..."):
            df, t, status = procesar_trafico_python(texto)
            if df is not None:
                st.session_state['df_trafico'] = df
                st.session_state['tiempos'] = t
                st.session_state['texto_original'] = texto
            else:
                st.error(f"❌ Error leyendo el texto.\n\nDetalle: {status}")

    if st.session_state.get('df_trafico') is not None:
        st.markdown("---")
        st.subheader("✏️ 1. Verifica y Edita los Datos")
        st.caption("Si ves algún error en el texto de WhatsApp, corrígelo en esta tabla antes de generar la imagen o guardar en Sheets.")
        
        df_editado = st.data_editor(
            st.session_state['df_trafico'], 
            num_rows="dynamic", 
            use_container_width=True,
            hide_index=True
        )
        
        msg_actual = generar_analisis_python(df_editado, st.session_state['tiempos'], st.session_state['texto_original'])
        
        col_msg, col_img = st.columns([1, 1.5])
        
        with col_msg:
            st.subheader("📝 2. Mensaje WhatsApp")
            st.text_area("Copia esto para tu gerente:", msg_actual, height=500)
            
            st.markdown("---")
            if st.button("💾 Guardar en Base de Datos (Sheets)", use_container_width=True):
                try:
                    df_s = df_editado.copy()
                    
                    if 'Listines' in df_s.columns:
                        df_s['Listines'] = df_s['Listines'].apply(lambda x: ", ".join(x) if isinstance(x, list) else str(x))
                    
                    fecha_raw = st.session_state['tiempos'].get('Fecha_Reporte', '')
                    obj_fecha = datetime.now()
                    if fecha_raw:
                        match = re.search(r'\d{2}/\d{2}/\d{4}', fecha_raw)
                        if match:
                            try:
                                obj_fecha = datetime.strptime(match.group(0), "%d/%m/%Y")
                            except:
                                pass
                    
                    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                    meses_ano = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    
                    df_s['Fecha'] = obj_fecha.strftime("%d/%m/%Y")
                    df_s['Dia'] = dias_semana[obj_fecha.weekday()]
                    df_s['Semana'] = f"Semana {obj_fecha.isocalendar()[1]}"
                    df_s['Mes'] = meses_ano[obj_fecha.month - 1]

                    df_s['Hora_1er_Listin'] = str(st.session_state['tiempos'].get('Hora_1er_listin',''))
                    df_s['Hora_Ultimo_Listin'] = str(st.session_state['tiempos'].get('Hora_ultimo_listin',''))
                    df_s['Inicio_Trafico'] = str(st.session_state['tiempos'].get('Inicio_trafico',''))
                    df_s['Culminacion_Trafico'] = str(st.session_state['tiempos'].get('Culminacion_trafico',''))
                    
                    df_s = df_s.fillna("")
                    
                    columnas_esperadas = [
                        'Fecha', 'Dia', 'Semana', 'Mes', 
                        'Ruta', 'Zona', 'Chofer', 'Ayudante', 'Unidad', 'Listines', 
                        'Farmacias_Total', 'Bultos_Total', 
                        'Hora_1er_Listin', 'Hora_Ultimo_Listin', 'Inicio_Trafico', 'Culminacion_Trafico'
                    ]
                    
                    for col in columnas_esperadas:
                        if col not in df_s.columns:
                            df_s[col] = "" 
                            
                    df_para_guardar = df_s[columnas_esperadas]
                    
                    with st.spinner("Guardando en Google Sheets..."):
                        guardar_en_google_sheets_directo(df_para_guardar)
                    
                    st.success("✅ ¡Datos guardados exitosamente en tu Google Sheets!")
                except Exception as e:
                    st.error(f"❌ Error al guardar en Sheets. Detalle: {str(e)}")
        
        with col_img:
            st.subheader("📸 3. Pizarra Corporativa")
            vista = st.radio("Selecciona el formato visual:", ["Agrupada", "Plana"], horizontal=True)
            
            st.components.v1.html("""
            <script>
            function d(){
                var f=window.parent.document.querySelector('iframe[title="streamlit_components.v1.components.html"]'); 
                if(f) { f.contentWindow.descargarImagen(); }
                else { 
                    var iframes = window.parent.document.getElementsByTagName('iframe'); 
                    for(var i=0; i<iframes.length; i++){ 
                        if(iframes[i].contentWindow.descargarImagen) { iframes[i].contentWindow.descargarImagen(); } 
                    } 
                } 
            }
            </script>
            <div style="text-align:right; margin-bottom:10px;">
                <button onclick="d()" style="background:#28a745;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;">⬇️ Descargar Imagen</button>
            </div>
            """, height=50)
            
            filas = generar_filas_agrupadas(df_editado) if vista == "Agrupada" else generar_filas_plana(df_editado)
            margen_extra_por_ruta = 75
            alt = 400 + (len(df_editado) * margen_extra_por_ruta) + (len(df_editado['Zona'].unique()) * 50 if vista=="Agrupada" else 0)
            
            components.html(obtener_html_base(df_editado, st.session_state['tiempos'], filas), height=alt, scrolling=True)

# ---------------------------------------------------------
# PESTAÑA 2: EL NUEVO MÓDULO DIOS (EXCEL PREVIO)
# ---------------------------------------------------------
with tab2:
    st.header("Construcción Proactiva desde Excel")
    st.info("Sube el archivo general del mes, selecciona la fecha, y genera la tabla verde referencial para los supervisores.")
    
    col_excel1, col_excel2 = st.columns(2)
    with col_excel1:
        fecha_filtro = st.date_input("📅 ¿Qué fecha de listín deseas auditar?")
    with col_excel2:
        archivo_subido = st.file_uploader("📂 Sube el Excel de Despachos", type=["xlsx", "xls"])
        
    if archivo_subido is not None:
        with st.spinner("Procesando archivo blindado..."):
            try:
                if archivo_subido.name.endswith('.xls'):
                    try:
                        df_base = pd.read_excel(archivo_subido, engine='xlrd')
                    except Exception as e_xls:
                        if "b'<html>" in str(e_xls) or "bof record" in str(e_xls).lower() or "html" in str(e_xls).lower():
                            archivo_subido.seek(0)
                            df_base = pd.read_html(archivo_subido)[0]
                        else:
                            raise e_xls
                else:
                    df_base = pd.read_excel(archivo_subido)
                
                resumen_verde, estado_excel = procesar_excel_base(df_base, fecha_filtro)
                
                if resumen_verde is not None:
                    st.success("¡Datos filtrados y agrupados con éxito!")
                    
                    col_resumen1, col_resumen2 = st.columns([1, 1.5])
                    
                    with col_resumen1:
                        st.subheader("Visualización Rápida")
                        st.dataframe(resumen_verde, hide_index=True, use_container_width=True)
                        
                    with col_resumen2:
                        st.subheader("📸 Imagen para el Equipo (Tabla Dinámica)")
                        st.components.v1.html("""
                        <script>
                        function dbv(){
                            var iframes = window.parent.document.getElementsByTagName('iframe'); 
                            for(var i=0; i<iframes.length; i++){ 
                                if(iframes[i].contentWindow.descargarTablaVerde) { iframes[i].contentWindow.descargarTablaVerde(); } 
                            } 
                        }
                        </script>
                        <div style="text-align:right; margin-bottom:10px;">
                            <button onclick="dbv()" style="background:#28a745;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;">⬇️ Descargar Tabla Verde</button>
                        </div>
                        """, height=50)
                        
                        html_verde = generar_html_tabla_verde(resumen_verde)
                        altura_verde = 200 + (len(resumen_verde) * 35)
                        components.html(html_verde, height=altura_verde, scrolling=True)
                        
                    if st.session_state.get('df_trafico') is not None:
                        st.markdown("---")
                        st.subheader("🔍 Auditoría Cruzada (WhatsApp vs Excel)")
                        st.caption("Verificando si hay discrepancias entre los datos procesados en el Módulo 1 y este archivo Excel.")
                        
                        df_auditoria = realizar_auditoria_cruzada(st.session_state['df_trafico'], resumen_verde)
                        st.dataframe(df_auditoria, use_container_width=True, hide_index=True)
                    else:
                        st.markdown("---")
                        st.info("💡 Consejo: Si procesas primero el texto de WhatsApp en el Módulo 1, aquí verás una comparación automática de discrepancias detectando al instante quién no cuadra.")
                        
                else:
                    st.warning(estado_excel)
                    
            except Exception as e:
                error_msg = str(e).lower()
                if "xlrd" in error_msg:
                    st.error("🚨 Falta la librería xlrd. Ve a tu terminal y ejecuta:\n\n`pip install xlrd`")
                elif "lxml" in error_msg or "html5lib" in error_msg:
                    st.error("🚨 Falta la librería de HTML. Ve a tu terminal y ejecuta:\n\n`pip install lxml html5lib`")
                else:
                    st.error(f"Error procesando el archivo: {e}")
