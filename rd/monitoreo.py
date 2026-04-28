# ==========================================
# Archivo: monitoreo.py (Módulo de Monitoreo de Despachos, KMs, Surtidos e Informes PDF)
# ==========================================
import streamlit as st
import pandas as pd
import re
import base64
import calendar
from datetime import datetime, timedelta
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import traceback
from pathlib import Path
import plotly.graph_objects as go

# ==========================================
# FUNCIONES GLOBALES
# ==========================================
def obtener_logo_base64():
    try:
        ruta_logo = Path(__file__).parent.parent / "logo.png"
        if ruta_logo.exists():
            with open(ruta_logo, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode()
                return f"data:image/png;base64,{encoded_string}"
        return None
    except Exception:
        return None

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

def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def extraer_datos_sheets(nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        
        # Conexión directa a prueba de balas usando tu ID
        doc = cliente.open_by_key("1ourNW6VifjXiJFsyVKamjeBL7iACKEpH0ozdWo8rCMc")
        
        hoja = doc.worksheet(nombre_hoja)
        data = hoja.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()

def guardar_en_google_sheets(df_para_guardar, nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        
        # Conexión directa a prueba de balas usando tu ID
        doc = cliente.open_by_key("1ourNW6VifjXiJFsyVKamjeBL7iACKEpH0ozdWo8rCMc")
        
        try:
            hoja = doc.worksheet(nombre_hoja)
        except Exception:
            hoja = doc.add_worksheet(title=nombre_hoja, rows=1000, cols=20)
            hoja.append_row(list(df_para_guardar.columns))
            
        df_clean = df_para_guardar.fillna('').astype(str)
        valores = df_clean.values.tolist()
        
        try:
            hoja.append_rows(valores, value_input_option='USER_ENTERED')
        except Exception as error_interno:
            if "200" in str(error_interno):
                pass # Silenciador del falso error 200
            else:
                raise error_interno
                
        # Limpiamos la caché de Streamlit para que los gráficos carguen la data nueva al instante
        st.cache_data.clear() 
        return True
        
    except Exception as e:
        st.error(f"🛑 Falla en la conexión: {e}")
        st.code(traceback.format_exc(), language="python")
        return False

# ==========================================
# CREADORES DE WHATSAPP
# ==========================================
def generar_ws_nacional(dict_dfs, g_rutas, g_cubiertos, g_bultos, g_kms, fecha_str):
    msg = f"*Reporte Nacional de Despachos Drotaca* 🚚\n📅 Fecha: {fecha_str}\n\n"
    
    msg += "*ESTADÍSTICA GENERAL:*\n"
    msg += f"📍 Total Rutas Activas: {g_rutas}\n"
    msg += f"🏥 Farmacias Entregadas: {g_cubiertos:,.0f}\n"
    msg += f"📦 Total Bultos Entregados: {g_bultos:,.0f}\n"
    msg += f"📏 Kilómetros Recorridos: {g_kms:,.2f} Km\n\n"
    
    msg += "🏆 *TOP RUTAS (POR BULTOS ENTREGADOS):*\n"
    
    orden_esperado = ["ORIENTE", "CENTRO", "OCCIDENTE"]
    for reg in orden_esperado:
        if reg in dict_dfs:
            df = dict_dfs[reg]
            nombre_reg = "CENTRO OCCIDENTE" if reg == "CENTRO" else reg
            msg += f"\n*{nombre_reg}:*\n"
            
            top_df = df.sort_values(by='BULTOS', ascending=False).head(3)
            for i, r in top_df.iterrows():
                chofer = str(r['CHOFER']).title()
                bultos = r['BULTOS']
                ruta = str(r['DESPACHOS']).title()
                msg += f"▪️ {chofer} - *{bultos} Bultos* ({ruta})\n"
                
    msg += "\n✅ *Pizarra visual de cobertura adjunta.*"
    return msg

# --- WHATSAPP PIZARRA 1 (Corta y Extracción) ---
def generar_ws_surtido_p1(df, fecha_str):
    df_p1 = df[df['GRUPO'].isin(['RUTA CORTA', 'EXTRACCION'])]
    if df_p1.empty:
        return f"*Reporte Surtido (1/3): Corta y Extracciones* ⛽\n📅 Fecha: {fecha_str}\n\nNo hay registros."

    df_surtido = df_p1[df_p1['GRUPO'] == 'RUTA CORTA']
    df_extrac = df_p1[df_p1['GRUPO'] == 'EXTRACCION']

    t_gasoil = df_surtido[df_surtido['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
    t_gasolina = df_surtido[df_surtido['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()
    t_extraccion = df_extrac['LITROS'].sum()

    msg = f"*Reporte Surtido (1/3): Corta y Extracciones* ⛽\n📅 Fecha: {fecha_str}\n\n"
    msg += f"▪️ *Gasoil (Ruta Corta):* {t_gasoil:,.0f} Lts\n"
    msg += f"▪️ *Gasolina (Ruta Corta):* {t_gasolina:,.0f} Lts\n"
    msg += f"▪️ *Extracciones (Drotaca 2.0 / Ciudad Drotaca):* {t_extraccion:,.0f} Lts\n\n"

    if not df_surtido.empty:
        msg += "🏆 *TOP CONSUMO RUTA CORTA:*\n"
        top_df = df_surtido.sort_values(by='LITROS', ascending=False).head(3)
        for _, r in top_df.iterrows():
            msg += f"🚚 {r['UNIDAD']} ({str(r['CHOFER']).title()}): {r['LITROS']:,.0f} Lts\n"

    return msg

# --- WHATSAPP PIZARRA 2 (Centro y Larga) ---
def generar_ws_surtido_p2(df, fecha_str):
    df_p2 = df[df['GRUPO'].str.contains('CENTRO|LARGA|OCCIDENTE', na=False)]
    if df_p2.empty:
        return f"*Reporte Surtido (2/3): Centro y Occidente* ⛽\n📅 Fecha: {fecha_str}\n\nNo hay registros."

    t_gasoil = df_p2[df_p2['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
    t_gasolina = df_p2[df_p2['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()

    msg = f"*Reporte Surtido (2/3): Centro y Occidente* ⛽\n📅 Fecha: {fecha_str}\n\n"
    msg += f"▪️ *Gasoil Surtido:* {t_gasoil:,.0f} Lts\n"
    msg += f"▪️ *Gasolina Surtida:* {t_gasolina:,.0f} Lts\n\n"

    if not df_p2.empty:
        msg += "🏆 *TOP CONSUMO CENTRO / OCCIDENTE:*\n"
        top_df = df_p2.sort_values(by='LITROS', ascending=False).head(3)
        for _, r in top_df.iterrows():
            msg += f"🚚 {r['UNIDAD']} ({str(r['CHOFER']).title()}): {r['LITROS']:,.0f} Lts\n"

    return msg

# --- WHATSAPP PIZARRA 3 (Resumen General y Tipos) ---
def generar_ws_surtido_p3(df, fecha_str):
    if df.empty:
        return f"*Reporte Ejecutivo de Surtido (3/3)* ⛽\n📅 Fecha: {fecha_str}\n\nNo hay registros."

    def std_tipo(t):
        t = str(t).upper()
        if 'ESTACION' in t or 'E/S' in t or 'E / S' in t: return 'ESTACION DE SERVICIO'
        if 'BIDON' in t: return 'BIDON'
        if 'RESERVA' in t or 'TANQUE' in t or 'BASE' in t or 'PLANTA' in t: return 'TANQUE RESERVA'
        return 'OTROS'
    
    df_copy = df.copy()
    df_copy['TIPO_STD'] = df_copy['TIPO_SURTIDO'].apply(std_tipo)
    
    df_surtido = df_copy[df_copy['GRUPO'] != 'EXTRACCION']
    df_extrac = df_copy[df_copy['GRUPO'] == 'EXTRACCION']
    
    t_gasoil = df_surtido[df_surtido['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
    t_gasolina = df_surtido[df_surtido['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()
    t_extraccion = df_extrac['LITROS'].sum()
    t_total = df_surtido['LITROS'].sum()
    
    msg = f"*Reporte Ejecutivo de Surtido (3/3)* ⛽\n📅 Fecha: {fecha_str}\n\n"
    msg += f"▪️ *Total Gasoil:* {t_gasoil:,.0f} Lts\n"
    msg += f"▪️ *Total Gasolina:* {t_gasolina:,.0f} Lts\n"
    msg += f"▪️ *Extracciones (Drotaca 2.0 / Ciudad Drotaca):* {t_extraccion:,.0f} Lts\n\n"
    
    if t_total > 0:
        msg += "📊 *DISTRIBUCIÓN ESTRATÉGICA:*\n"
        for tipo in ['ESTACION DE SERVICIO', 'BIDON', 'TANQUE RESERVA', 'OTROS']:
            lts_tipo = df_surtido[df_surtido['TIPO_STD'] == tipo]['LITROS'].sum()
            if lts_tipo > 0:
                pct = (lts_tipo / t_total) * 100
                icon = "⛽" if tipo == 'ESTACION DE SERVICIO' else ("🛢️" if tipo == 'BIDON' else "🏭")
                nombre_mostrar = "TANQUE RESERVA CIUDAD DROTACA" if tipo == 'TANQUE RESERVA' else tipo.title()
                msg += f"{icon} *{nombre_mostrar}:* {lts_tipo:,.0f} Lts ({pct:.1f}%)\n"
    
    return msg

# ==========================================
# MOTOR EXTRAECTOR DE KILOMETRAJE
# ==========================================
def procesar_excel_km(file, target_date):
    try:
        df_km_raw = pd.read_excel(file, sheet_name="KM - ODOMETRO", header=None)
        
        header_idx = -1
        for i in range(min(20, len(df_km_raw))):
            fila_str = " ".join([str(x).upper() for x in df_km_raw.iloc[i].values])
            if "UNIDAD" in fila_str or "TIPO" in fila_str:
                header_idx = i
                break
                
        if header_idx != -1:
            df_km_raw.columns = df_km_raw.iloc[header_idx]
            df_km = df_km_raw.iloc[header_idx+1:].reset_index(drop=True)
        else:
            df_km = df_km_raw 
            
        target_col = None
        for col in df_km.columns:
            if isinstance(col, datetime):
                if col.month == target_date.month and col.day == target_date.day:
                    target_col = col
                    break
            else:
                c_str = str(col).lower()
                meses = {1:['ene','jan'], 2:['feb'], 3:['mar'], 4:['abr','apr'], 5:['may'], 6:['jun'], 
                         7:['jul'], 8:['ago','aug'], 9:['sep'], 10:['oct'], 11:['nov'], 12:['dic','dec']}
                
                day_pattern = rf'(?<!\d)0?{target_date.day}(?!\d)'
                if re.search(day_pattern, c_str):
                    for m_str in meses[target_date.month]:
                        if m_str in c_str:
                            target_col = col
                            break
            if target_col is not None: break
            
        if not target_col:
            return {}, f"No se encontró la fecha {target_date.strftime('%d/%m')} en el archivo de Kilometraje."
            
        unidad_col = None
        for c in df_km.columns:
            if 'UNIDAD' in str(c).upper() or 'PLACA' in str(c).upper():
                unidad_col = c
                break
        if not unidad_col:
            unidad_col = df_km.columns[0] 
            
        km_dict = {}
        for _, row in df_km.iterrows():
            unidad_raw = str(row[unidad_col]).strip().upper().replace('"', '').replace(' ', '')
            if unidad_raw in ['NAN', 'NONE', 'NAT', '']: continue
            
            val_str = str(row[target_col]).upper().replace('KMS', '').replace('KM', '').replace(',', '.').strip()
            try: km_val = float(val_str)
            except: km_val = 0.0
            
            km_dict[unidad_raw] = km_val
            
        return km_dict, "OK"
    except Exception as e:
        return {}, f"Error procesando KMs: {str(e)}"

if 'log_cruces_km' not in st.session_state:
    st.session_state['log_cruces_km'] = []

def aplicar_kilometraje(df, dict_km, region):
    if dict_km and not df.empty:
        kms_list = []
        for _, r in df.iterrows():
            unidad_original = str(r['UNIDAD']).upper().strip()
            placa_base = unidad_original.split()[0].replace('-', '')
            placa_limpia = re.sub(r'\s*(C-?3500|FOTON|NPR|CARGO|L300|CANTER|PEUGEOT|PANEL|VITARA|MOTO|DYNA|SUPER DUTY|SUPER DUTTY|HIACE|DFSK|DONGFENG|GRAND VITARA)\s*', '', unidad_original).strip().replace('-', '')
            
            km_asignado = 0.0
            placa_cruzada = "NO ENCONTRADA"
            
            if placa_base in dict_km:
                km_asignado = dict_km[placa_base]
                placa_cruzada = placa_base
            elif placa_limpia in dict_km:
                km_asignado = dict_km[placa_limpia]
                placa_cruzada = placa_limpia
            else:
                for k_placa, val in dict_km.items():
                    if (len(placa_base) >= 3 and k_placa.startswith(placa_base)) or (len(placa_limpia) >= 3 and k_placa.startswith(placa_limpia)): 
                        km_asignado = val
                        placa_cruzada = k_placa
                        break
                        
            kms_list.append(km_asignado)
            
            st.session_state['log_cruces_km'].append({
                "Región": region,
                "Placa Pizarra": unidad_original,
                "Placa Detectada": placa_base,
                "Cruzó con (Odómetro)": placa_cruzada,
                "KMs Asignados": km_asignado
            })
            
        df['KILOMETROS'] = kms_list
    else:
        df['KILOMETROS'] = 0.0
    return df

# ==========================================
# MOTOR EXTRAECTOR DE TEXTO Y EXCEL (DESPACHOS)
# ==========================================
def procesar_texto_oriente(texto):
    try:
        lineas = [l.strip() for l in texto.split('\n') if l.strip()]
        datos = []
        for linea in lineas:
            if "ITEM" in linea.upper() or "CHOFER" in linea.upper() or "TRANSBORDO" in linea.upper():
                continue
            partes = linea.split('\t')
            if len(partes) < 8:
                partes = [p.strip() for p in re.split(r'\t|\s{2,}', linea) if p.strip()]
            if len(partes) >= 8:
                try:
                    bultos = int(float(str(partes[-1]).replace(',', '').strip()))
                    pendientes = int(float(str(partes[-2]).replace(',', '').strip()))
                    cubiertos = int(float(str(partes[-3]).replace(',', '').strip()))
                    cubrir = int(float(str(partes[-4]).replace(',', '').strip()))
                    despachos = str(partes[-5]).strip().upper()
                    ayudante = str(partes[-6]).strip().title()
                    chofer = str(partes[-7]).strip().title()
                    unidad = str(partes[-8]).strip().upper()
                    
                    if cubrir > 0 and 'TRANSBORDO' not in chofer.upper():
                        datos.append({
                            "UNIDAD": unidad, "CHOFER": chofer, "AYUDANTE": ayudante,
                            "DESPACHOS": despachos, "CUBRIR": cubrir, "CUBIERTOS": cubiertos,
                            "PENDIENTES": pendientes, "BULTOS": bultos
                        })
                except Exception:
                    pass 
        df = pd.DataFrame(datos)
        if not df.empty:
            df['ITEM'] = range(1, len(df)+1)
            df = df[['ITEM', 'UNIDAD', 'CHOFER', 'AYUDANTE', 'DESPACHOS', 'CUBRIR', 'CUBIERTOS', 'PENDIENTES', 'BULTOS']]
            return df, "OK"
        else:
            return None, "No se detectaron datos válidos. Asegúrate de copiar la tabla directamente desde el Excel."
    except Exception as e:
        return None, f"Error al procesar texto: {str(e)}"

def procesar_excel_region(file, hojas_posibles, region_nombre):
    try:
        xl = pd.ExcelFile(file)
        hoja_target = xl.sheet_names[0] 
        for hp in hojas_posibles:
            for s in xl.sheet_names:
                if hp.upper() in s.upper():
                    hoja_target = s
                    break
        df_raw = pd.read_excel(file, sheet_name=hoja_target, header=None)
        mapa_idx = {'ITEM': -1, 'UNIDAD': -1, 'CHOFER': -1, 'AYUDANTE': -1, 'DESPACHOS': -1, 
                    'CUBRIR': -1, 'CUBIERTOS': -1, 'PENDIENTES': -1, 'BULTOS': -1}
        row_start = -1
        for r in range(min(50, len(df_raw))):
            for c in range(len(df_raw.columns)):
                val = str(df_raw.iloc[r, c]).upper().strip()
                if val in ['NAN', 'NONE', 'NAT', '']: continue
                if ('ITEM' in val or 'Nº' in val or 'N°' in val) and mapa_idx['ITEM'] == -1: mapa_idx['ITEM'] = c
                elif ('UNIDAD' in val or 'PLACA' in val or 'VEHICUL' in val or 'CAMION' in val or 'TRANSPORTE' in val) and mapa_idx['UNIDAD'] == -1: mapa_idx['UNIDAD'] = c
                elif ('CHOFER' in val or 'CHÓFER' in val or 'CONDUCTOR' in val) and mapa_idx['CHOFER'] == -1: 
                    mapa_idx['CHOFER'] = c
                    if row_start == -1: row_start = r  
                elif ('AYUDANTE' in val or 'ESCOLTA' in val) and mapa_idx['AYUDANTE'] == -1: mapa_idx['AYUDANTE'] = c
                elif ('DESPACHO' in val or 'RUTA' in val or 'DESTINO' in val or 'ZONA' in val) and 'HORA' not in val and mapa_idx['DESPACHOS'] == -1: mapa_idx['DESPACHOS'] = c
                elif ('CUBRIR' in val or 'ASIGNAD' in val or 'PROGRAMAD' in val) and mapa_idx['CUBRIR'] == -1: mapa_idx['CUBRIR'] = c
                elif ('CUBIERT' in val or 'ENTREGAD' in val or 'VISITAD' in val) and mapa_idx['CUBIERTOS'] == -1: mapa_idx['CUBIERTOS'] = c
                elif ('PENDIENT' in val or 'RECHAZO' in val or 'NO ENTREGAD' in val) and mapa_idx['PENDIENTES'] == -1: mapa_idx['PENDIENTES'] = c
                elif ('BULTO' in val or 'CAJA' in val or 'PIEZA' in val) and mapa_idx['BULTOS'] == -1: mapa_idx['BULTOS'] = c

        if mapa_idx['CUBRIR'] == -1:
            for r in range(min(50, len(df_raw))):
                for c in range(len(df_raw.columns)):
                    val = str(df_raw.iloc[r, c]).upper().strip()
                    if 'CLIENTES' in val and 'PENDIENT' not in val and 'CUBIERT' not in val and 'ENTREG' not in val:
                        mapa_idx['CUBRIR'] = c
                        break
        if mapa_idx['CHOFER'] == -1: return None, f"No se encontró la palabra 'CHOFER' en la hoja {hoja_target}."
            
        df_data = df_raw.iloc[row_start+1:].reset_index(drop=True)
        idx_chofer = mapa_idx['CHOFER']
        df_data = df_data.dropna(subset=[df_raw.columns[idx_chofer]])
        df_data = df_data[~df_data.iloc[:, idx_chofer].astype(str).str.upper().str.contains('TOTAL', na=False)]
        df_data = df_data[~df_data.iloc[:, idx_chofer].astype(str).str.upper().str.contains('TRANSBORDO', na=False)]
        df_data = df_data[df_data.iloc[:, idx_chofer].astype(str).str.strip() != '']
        df_data = df_data[df_data.iloc[:, idx_chofer].astype(str).str.upper().str.strip() != 'NAN']
        
        def obtener_columna(idx, default_val):
            if idx != -1 and idx < len(df_data.columns): return df_data.iloc[:, idx]
            return pd.Series([default_val] * len(df_data))
        
        df_clean = pd.DataFrame()
        df_clean['ITEM'] = obtener_columna(mapa_idx['ITEM'], 0)
        df_clean['UNIDAD'] = obtener_columna(mapa_idx['UNIDAD'], '-').fillna('-')
        df_clean['CHOFER'] = obtener_columna(mapa_idx['CHOFER'], '-').fillna('-').astype(str).str.title()
        df_clean['AYUDANTE'] = obtener_columna(mapa_idx['AYUDANTE'], '-').fillna('-').astype(str).str.title()
        df_clean['DESPACHOS'] = obtener_columna(mapa_idx['DESPACHOS'], '-').fillna('-').astype(str).str.upper()
        
        def to_num(s):
            try: 
                val = str(s).upper().replace(',', '').replace(' ', '').replace('NAN', '')
                if val == '-' or val == '': return 0
                return int(float(val))
            except: return 0
            
        df_clean['CUBRIR'] = obtener_columna(mapa_idx['CUBRIR'], 0).apply(to_num)
        df_clean['CUBIERTOS'] = obtener_columna(mapa_idx['CUBIERTOS'], 0).apply(to_num)
        df_clean['PENDIENTES'] = obtener_columna(mapa_idx['PENDIENTES'], 0).apply(to_num)
        df_clean['BULTOS'] = obtener_columna(mapa_idx['BULTOS'], 0).apply(to_num)
        
        for idx, r in df_clean.iterrows():
            if r['CUBRIR'] == 0 and (r['CUBIERTOS'] > 0 or r['PENDIENTES'] > 0):
                df_clean.at[idx, 'CUBRIR'] = r['CUBIERTOS'] + r['PENDIENTES']
            elif r['PENDIENTES'] == 0 and r['CUBIERTOS'] > 0 and r['CUBRIR'] > 0:
                df_clean.at[idx, 'PENDIENTES'] = max(0, r['CUBRIR'] - r['CUBIERTOS'])
        
        df_clean = df_clean[df_clean['CUBRIR'] > 0].reset_index(drop=True)
        df_clean['ITEM'] = range(1, len(df_clean)+1)
        
        return df_clean, "OK"
    except Exception as e:
        return None, str(e)


# ==========================================
# CREADOR DEL MEGA-HTML NACIONAL DESPACHOS
# ==========================================
def html_pizarra_nacional(dict_dfs, fecha_str):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 70px; max-width: 200px; object-fit: contain; display: block;">' if logo else '<span style="font-size:30px; color:white;">🚚 DROTACA</span>'
    color_header = "#1a4685" 
    
    g_rutas = 0; g_cubrir = 0; g_cubiertos = 0; g_pendientes = 0; g_bultos = 0; g_kms = 0.0
    html_tablas_regiones = ""
    
    orden_esperado = ["ORIENTE", "CENTRO", "OCCIDENTE"]
    regiones_presentes = [r for r in orden_esperado if r in dict_dfs]
    
    for region in regiones_presentes:
        df = dict_dfs[region]
        
        t_rutas = len(df); t_cubrir = df['CUBRIR'].sum(); t_cubiertos = df['CUBIERTOS'].sum()
        t_pendientes = df['PENDIENTES'].sum(); t_bultos = df['BULTOS'].sum(); t_kms = df['KILOMETROS'].sum()
        
        g_rutas += t_rutas; g_cubrir += t_cubrir; g_cubiertos += t_cubiertos
        g_pendientes += t_pendientes; g_bultos += t_bultos; g_kms += t_kms
        
        nombre_mostrar = "CENTRO OCCIDENTE" if region == "CENTRO" else region
        
        filas_html = ""
        for i, r in df.iterrows():
            bg = "#e9edf4" if i % 2 != 0 else "#ffffff"
            filas_html += f"""
            <tr style="background-color: {bg}; text-align: center; font-size: 14px; color: #000000;">
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold;">{r['ITEM']}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['UNIDAD']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold;">{r['CHOFER']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold;">{r['AYUDANTE']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-size: 12px;">{r['DESPACHOS']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; font-size: 16px;">{r['CUBRIR']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; font-size: 16px;">{r['CUBIERTOS']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; font-size: 16px;">{r['PENDIENTES']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; font-size: 16px;">{r['BULTOS']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; font-size: 16px;">{r['KILOMETROS']:,.2f}</td>
            </tr>
            """
            
        html_tablas_regiones += f"""
        <div style="margin-top: 30px;">
            <h3 style="color: {color_header}; border-bottom: 3px solid {color_header}; padding-bottom: 5px; margin-bottom: 10px; font-size: 20px; padding-left: 10px;">📍 RUTA {nombre_mostrar}</h3>
            <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
                <thead>
                    <tr style="background-color: #f2f2f2; font-size: 12px; text-align: center; color: #000;">
                        <th style="border: 1px solid #000; padding: 8px;">ITEM</th>
                        <th style="border: 1px solid #000; padding: 8px;">UNIDAD</th>
                        <th style="border: 1px solid #000; padding: 8px;">CHOFER</th>
                        <th style="border: 1px solid #000; padding: 8px;">AYUDANTE</th>
                        <th style="border: 1px solid #000; padding: 8px;">DESPACHOS</th>
                        <th style="border: 1px solid #000; padding: 8px;">Cubrir</th>
                        <th style="border: 1px solid #000; padding: 8px;">Cubiertos</th>
                        <th style="border: 1px solid #000; padding: 8px;">Pendientes</th>
                        <th style="border: 1px solid #000; padding: 8px;">BULTOS</th>
                        <th style="border: 1px solid #000; padding: 8px;">KMS</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_html}
                    <tr style="background-color: #ffffff; text-align: center; font-size: 16px; color: #000;">
                        <td colspan="5" style="border: 1px solid #000; padding: 10px; text-align: right; font-weight: bold;">Subtotal {nombre_mostrar.title()}:</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold;">{t_cubrir}</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold; background-color: #c8e6c9; color: #1b5e20;">{t_cubiertos}</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold; background-color: #ffcdd2; color: #b71c1c;">{t_pendientes}</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold; background-color: #fffb00;">{t_bultos}</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold; background-color: #e8f5e9; color: #2e7d32;">{t_kms:,.2f}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    bloque_estadistica_general = f"""
    <div style="margin-top: 40px; padding: 25px; border: 3px solid {color_header}; border-radius: 10px; background-color: #f8f9fa;">
        <h2 style="text-align: center; color: {color_header}; margin-top: 0; margin-bottom: 20px; font-size: 26px; text-transform: uppercase; letter-spacing: 1px;">🌎 Estadística General de Despacho</h2>
        <div style="display: flex; justify-content: space-around; align-items: center;">
            <div style="text-align: center; flex: 1; border-right: 2px solid #ddd;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #666; text-transform: uppercase;">Total Rutas</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #333;">{g_rutas}</p>
            </div>
            <div style="text-align: center; flex: 1; border-right: 2px solid #ddd;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #1565c0; text-transform: uppercase;">Farmacias Entregadas</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #1565c0;">{g_cubiertos:,.0f}</p>
            </div>
            <div style="text-align: center; flex: 1; border-right: 2px solid #ddd;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #e65100; text-transform: uppercase;">Total Bultos</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #e65100;">{g_bultos:,.0f}</p>
            </div>
            <div style="text-align: center; flex: 1;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #2e7d32; text-transform: uppercase;">Kilómetros Recorridos</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #2e7d32;">{g_kms:,.2f} <span style="font-size: 16px;">Km</span></p>
            </div>
        </div>
    </div>
    """

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;">
        <button onclick="descargar_nacional()" style="background:#d32f2f;color:#fff;border:none;padding:12px 25px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:16px;box-shadow:0 4px 6px rgba(0,0,0,0.2);">⬇️ DESCARGAR MEGA-PIZARRA NACIONAL</button>
    </div>
    
    <div id="pizarra-nacional" style="font-family: Arial, sans-serif; width: 1200px; margin: auto; background-color: #fff; border: 3px solid #000; box-sizing: border-box; padding-bottom: 20px;">
        <table style="width: 100%; border-collapse: collapse; background-color: {color_header}; color: white;">
            <tr>
                <td style="padding: 20px; width: 25%; text-align: center;">{area_logo}</td>
                <td style="padding: 20px; font-size: 28px; font-weight: bold; text-align: center; width: 50%; text-transform: uppercase; letter-spacing: 2px;">Reporte Nacional de Despachos</td>
                <td style="padding: 20px; font-size: 18px; text-align: center; width: 25%; font-weight: bold;">Fecha: {fecha_str}</td>
            </tr>
        </table>
        <div style="padding: 0 20px;">
            {html_tablas_regiones}
            {bloque_estadistica_general}
        </div>
        <div style="margin-top: 30px; text-align: center; font-size: 14px; color: #666; font-weight: bold;">Departamento de Flota y Logística</div>
    </div>
    <script>
    function descargar_nacional() {{ 
        html2canvas(document.getElementById('pizarra-nacional'), {{scale: 2}}).then(canvas => {{ 
            var link = document.createElement('a'); 
            link.download = 'Pizarra_Nacional_{fecha_str.replace("/","-")}.png'; 
            link.href = canvas.toDataURL(); link.click(); 
        }}); 
    }}
    </script>
    """


# ==========================================
# MOTOR ESCÁNER DE SURTIDOS DE COMBUSTIBLE
# ==========================================
def procesar_excel_surtidos(file, grupo_default, target_date):
    try:
        file.seek(0)
        xl = pd.ExcelFile(file)
        datos_totales = []
        t_date = target_date.date() if hasattr(target_date, 'date') else target_date

        for sheet in xl.sheet_names:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            mapa = None
            current_grupo = grupo_default
            
            if 'EXTRAC' in str(sheet).upper(): current_grupo = 'EXTRACCION'
            elif 'CORTA' in str(sheet).upper(): current_grupo = 'RUTA CORTA'
            elif 'CENTRO' in str(sheet).upper(): current_grupo = 'RUTA CENTRO'
            elif 'LARGA' in str(sheet).upper(): current_grupo = 'RUTA OCCIDENTE' 
                
            for i in range(len(df_raw)):
                fila_valores = df_raw.iloc[i].values
                fila_str = " ".join([str(x).upper().replace('\n', ' ') for x in fila_valores])
                
                if ("SURTIDO" in fila_str or "LITROS" in fila_str) and ("UNIDAD" in fila_str or "PLACA" in fila_str or "CHOFER" in fila_str):
                    mapa = {'FECHA': -1, 'UNIDAD': -1, 'CHOFER': -1, 'LITROS': -1, 'COMBUSTIBLE': -1, 'TIPO_SURTIDO': -1, 'SITIO': -1, 'HORA': -1, 'RUTA': -1}
                    for idx, col in enumerate(fila_valores):
                        c_str = str(col).upper().replace('\n', ' ').strip()
                        if ('FECHA' in c_str or 'DATE' in c_str) and mapa['FECHA'] == -1: mapa['FECHA'] = idx
                        elif ('UNIDAD' in c_str or 'PLACA' in c_str) and mapa['UNIDAD'] == -1: mapa['UNIDAD'] = idx
                        elif 'CHOFER' in c_str and mapa['CHOFER'] == -1: mapa['CHOFER'] = idx
                        elif ('TOTAL SURTIDO' in c_str or 'LITROS' in c_str or ('SURTIDO' in c_str and 'TIPO' not in c_str)) and mapa['LITROS'] == -1: mapa['LITROS'] = idx
                        elif 'COMBUSTIBLE' in c_str and mapa['COMBUSTIBLE'] == -1: mapa['COMBUSTIBLE'] = idx
                        elif ('TIPO DE SURTIDO' in c_str or 'TIPO SURTIDO' in c_str) and mapa['TIPO_SURTIDO'] == -1: mapa['TIPO_SURTIDO'] = idx
                        elif ('SITIO' in c_str or 'LUGAR' in c_str or 'ESTACION' in c_str) and mapa['SITIO'] == -1: mapa['SITIO'] = idx
                        elif ('HORA' in c_str or 'TIEMPO' in c_str) and mapa['HORA'] == -1: mapa['HORA'] = idx
                        elif ('RUTA' in c_str or 'ZONA' in c_str) and mapa['RUTA'] == -1: mapa['RUTA'] = idx
                    
                    if i > 0:
                        fila_arriba = " ".join([str(x).upper() for x in df_raw.iloc[i-1].values])
                        if 'EXTRAC' in fila_arriba or 'RETIRO' in fila_arriba: current_grupo = 'EXTRACCION'
                        elif 'CENTRO' in fila_arriba: current_grupo = 'RUTA CENTRO'
                    continue

                if mapa and mapa['LITROS'] != -1 and mapa['UNIDAD'] != -1:
                    if mapa['FECHA'] != -1:
                        fecha_val = fila_valores[mapa['FECHA']]
                        if pd.isna(fecha_val) or str(fecha_val).strip() == '' or str(fecha_val).strip() == 'nan': continue
                        try:
                            if hasattr(fecha_val, 'date'): f_date = fecha_val.date()
                            else:
                                dt_parsed = pd.to_datetime(str(fecha_val), dayfirst=True, errors='coerce')
                                if pd.isna(dt_parsed): continue
                                f_date = dt_parsed.date()
                            if f_date != t_date: continue
                        except: continue
                    else: continue 
                        
                    unidad = str(fila_valores[mapa['UNIDAD']]).upper().strip()
                    if unidad in ['NAN', 'NONE', '', 'TOTAL'] or 'TOTAL' in unidad: continue
                    
                    chofer = str(fila_valores[mapa['CHOFER']]).title().strip() if mapa['CHOFER'] != -1 else "-"
                    if 'TRANSBORDO' in chofer.upper(): continue
                    
                    try: 
                        val_lts = str(fila_valores[mapa['LITROS']]).upper().replace(',', '').replace(' ', '').replace('LTS', '').replace('L', '').strip()
                        lts = float(val_lts)
                    except: continue
                    if lts <= 0: continue
                    
                    ruta = str(fila_valores[mapa['RUTA']]).title().strip() if mapa['RUTA'] != -1 else "-"
                    sitio = str(fila_valores[mapa['SITIO']]).title().strip() if mapa['SITIO'] != -1 else "-"
                    
                    hora_raw = str(fila_valores[mapa['HORA']]).strip() if mapa['HORA'] != -1 else "-"
                    hora_limpia = hora_raw.split(' ')[1][:5] if ' ' in hora_raw and ':' in hora_raw else hora_raw
                    if hora_limpia in ['nan', 'NaN', 'NaT']: hora_limpia = "-"
                    
                    combustible = str(fila_valores[mapa['COMBUSTIBLE']]).upper().strip() if mapa['COMBUSTIBLE'] != -1 else "GASOIL"
                    tipo_surt = str(fila_valores[mapa['TIPO_SURTIDO']]).upper().strip() if mapa['TIPO_SURTIDO'] != -1 else ""
                    
                    grupo = current_grupo
                    if 'EXTRAC' in tipo_surt or 'RETIRO' in tipo_surt or 'PLANTA' in tipo_surt:
                        grupo = 'EXTRACCION'
                    elif 'EXTRAC' in str(sheet).upper():
                        grupo = 'EXTRACCION'
                        
                    datos_totales.append({
                        "GRUPO": grupo,
                        "UNIDAD": unidad,
                        "CHOFER": chofer,
                        "RUTA": ruta,
                        "HORA": hora_limpia,
                        "SITIO": sitio,
                        "TIPO_SURTIDO": tipo_surt,
                        "COMBUSTIBLE": combustible,
                        "LITROS": lts
                    })
                    
        if not datos_totales: return pd.DataFrame(), "OK"
        return pd.DataFrame(datos_totales), "OK"
    except Exception as e:
        return None, f"Error: {str(e)}"

# ==========================================
# CREADOR HTML PIZARRAS COMBUSTIBLE (1, 2 Y 3) (SÚPER MODULAR)
# ==========================================
def html_pizarras_combustible_completas(df, fecha_str):
    color_header = "#0d47a1" 
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 60px; max-width: 180px; object-fit: contain; display: block;">' if logo else '<span style="font-size:30px; color:white;">⛽ DROTACA</span>'
    
    if df.empty:
        return f"<div style='padding: 50px; text-align: center; color: #555; font-size: 20px;'>No se registraron movimientos de combustible para la fecha: {fecha_str}</div>"

    def std_tipo(t):
        t = str(t).upper()
        if 'ESTACION' in t or 'E/S' in t or 'E / S' in t: return 'ESTACION DE SERVICIO'
        if 'BIDON' in t: return 'BIDON'
        if 'RESERVA' in t or 'TANQUE' in t or 'BASE' in t or 'PLANTA' in t: return 'TANQUE RESERVA'
        return 'OTROS'
    
    df['TIPO_STD'] = df['TIPO_SURTIDO'].apply(std_tipo)

    # 1. Filtros por Pizarra
    df_p1 = df[df['GRUPO'].isin(['RUTA CORTA', 'EXTRACCION'])]
    df_p2 = df[df['GRUPO'].str.contains('CENTRO|LARGA|OCCIDENTE', na=False)]
    
    # Totales Pizarra 1
    df_surt_p1 = df_p1[df_p1['GRUPO'] == 'RUTA CORTA']
    t_gasoil_p1 = df_surt_p1[df_surt_p1['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
    t_gasol_p1 = df_surt_p1[df_surt_p1['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()
    t_extrac_p1 = df_p1[df_p1['GRUPO'] == 'EXTRACCION']['LITROS'].sum()
    
    # Totales Pizarra 2
    t_gasoil_p2 = df_p2[df_p2['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
    t_gasol_p2 = df_p2[df_p2['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()
    t_surtido_p2 = df_p2['LITROS'].sum()
    
    # Totales Pizarra 3 (Global)
    df_surtido = df[df['GRUPO'] != 'EXTRACCION']
    t_gasoil = df_surtido[df_surtido['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
    t_gasolina = df_surtido[df_surtido['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()
    t_extraccion = df[df['GRUPO'] == 'EXTRACCION']['LITROS'].sum()
    t_surtido_total = df_surtido['LITROS'].sum()
    
    def generar_filas(df_filtro):
        df_filtro = df_filtro.reset_index(drop=True)
        html = ""
        for i, r in df_filtro.iterrows():
            item_num = i + 1
            bg = "#f9f9f9" if i % 2 != 0 else "#ffffff"
            color_lts = "#0d47a1" if 'GASOIL' in r['COMBUSTIBLE'] else "#e65100"
            html += f"""
            <tr style="background-color: {bg}; text-align: center; font-size: 13px; color: #000;">
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold;">{item_num}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['UNIDAD']}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['CHOFER']}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['RUTA']}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['HORA']}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['SITIO']}</td>
                <td style="border: 1px solid #000; padding: 8px;">{r['TIPO_SURTIDO']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; color: {color_lts};">{r['COMBUSTIBLE']}</td>
                <td style="border: 1px solid #000; padding: 8px; font-weight: bold; font-size: 15px;">{r['LITROS']:,.0f}</td>
            </tr>
            """
        return html

    def construir_tabla(titulo, filas, total_lts):
        if not filas: return ""
        return f"""
        <div style="margin-top: 25px;">
            <div style="background-color: {color_header}; color: white; padding: 10px; font-weight: bold; font-size: 16px; text-align: center; border: 2px solid #000; border-bottom: none;">{titulo}</div>
            <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
                <thead>
                    <tr style="background-color: #e0e0e0; font-size: 12px; text-align: center; color: #000;">
                        <th style="border: 1px solid #000; padding: 8px;">ITEM</th>
                        <th style="border: 1px solid #000; padding: 8px;">UNIDAD</th>
                        <th style="border: 1px solid #000; padding: 8px;">CHOFER</th>
                        <th style="border: 1px solid #000; padding: 8px;">RUTA</th>
                        <th style="border: 1px solid #000; padding: 8px;">HORA</th>
                        <th style="border: 1px solid #000; padding: 8px;">SITIO</th>
                        <th style="border: 1px solid #000; padding: 8px;">TIPO SURTIDO</th>
                        <th style="border: 1px solid #000; padding: 8px;">COMBUSTIBLE</th>
                        <th style="border: 1px solid #000; padding: 8px;">LITROS</th>
                    </tr>
                </thead>
                <tbody>
                    {filas}
                    <tr style="background-color: #ffffff; font-size: 15px; color: #000;">
                        <td colspan="8" style="border: 1px solid #000; padding: 12px; text-align: right; font-weight: bold;">TOTAL {titulo}:</td>
                        <td style="border: 1px solid #000; padding: 12px; text-align: center; font-weight: bold; color: {color_header};">{total_lts:,.0f} Lts</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    # Generación de Tablas Internas
    html_corta = construir_tabla("RUTA CORTA", generar_filas(df_p1[df_p1['GRUPO'] == 'RUTA CORTA']), df_surt_p1['LITROS'].sum())
    html_extrac = construir_tabla("EXTRACCIONES (DROTACA 2.0 / CIUDAD DROTACA)", generar_filas(df_p1[df_p1['GRUPO'] == 'EXTRACCION']), t_extrac_p1)
    
    html_centro = construir_tabla("RUTA CENTRO OCCIDENTE", generar_filas(df_p2[df_p2['GRUPO'] == 'RUTA CENTRO']), df_p2[df_p2['GRUPO'] == 'RUTA CENTRO']['LITROS'].sum())
    html_larga = construir_tabla("RUTA OCCIDENTE", generar_filas(df_p2[df_p2['GRUPO'].str.contains('OCCIDENTE|LARGA')]), df_p2[df_p2['GRUPO'].str.contains('OCCIDENTE|LARGA')]['LITROS'].sum())

    # Generación Pizarra 3 (Resumen por Tipo)
    html_resumen_tipos = ""
    if t_surtido_total > 0:
        filas_tipos = ""
        for tipo in ['ESTACION DE SERVICIO', 'BIDON', 'TANQUE RESERVA', 'OTROS']:
            df_tipo = df_surtido[df_surtido['TIPO_STD'] == tipo]
            if not df_tipo.empty:
                t_gasoil_tipo = df_tipo[df_tipo['COMBUSTIBLE'].str.contains('GASOIL|DIESEL')]['LITROS'].sum()
                t_gasol_tipo = df_tipo[df_tipo['COMBUSTIBLE'].str.contains('GASOLINA')]['LITROS'].sum()
                t_tipo = df_tipo['LITROS'].sum()
                pct_tipo = (t_tipo / t_surtido_total) * 100
                
                nombre_mostrar = "TANQUE RESERVA CIUDAD DROTACA" if tipo == 'TANQUE RESERVA' else tipo.title()
                
                filas_tipos += f"""
                <tr style="text-align: center; font-size: 16px; color: #000; border-bottom: 1px solid #ccc;">
                    <td style="border: 1px solid #000; padding: 15px; font-weight: bold; text-align: left; padding-left: 20px;">{nombre_mostrar}</td>
                    <td style="border: 1px solid #000; padding: 15px; color: #0d47a1; font-weight: bold;">{t_gasoil_tipo:,.0f} Lts</td>
                    <td style="border: 1px solid #000; padding: 15px; color: #e65100; font-weight: bold;">{t_gasol_tipo:,.0f} Lts</td>
                    <td style="border: 1px solid #000; padding: 15px; font-weight: bold; font-size: 18px;">{t_tipo:,.0f} Lts</td>
                    <td style="border: 1px solid #000; padding: 15px; font-weight: bold; color: #2e7d32;">{pct_tipo:.1f}%</td>
                </tr>
                """
        
        html_resumen_tipos = f"""
        <div style="margin-top: 30px;">
            <table style="width: 100%; border-collapse: collapse; border: 2px solid #000;">
                <thead>
                    <tr style="background-color: #e0e0e0; font-size: 14px; text-align: center; color: #000;">
                        <th style="border: 1px solid #000; padding: 12px; width: 30%;">TIPO DE SURTIDO</th>
                        <th style="border: 1px solid #000; padding: 12px; width: 20%;">GASOIL (Lts)</th>
                        <th style="border: 1px solid #000; padding: 12px; width: 20%;">GASOLINA (Lts)</th>
                        <th style="border: 1px solid #000; padding: 12px; width: 20%;">TOTAL (Lts)</th>
                        <th style="border: 1px solid #000; padding: 12px; width: 10%;">%</th>
                    </tr>
                </thead>
                <tbody>
                    {filas_tipos}
                    <tr style="background-color: #f2f2f2; font-size: 18px; color: #000;">
                        <td style="border: 1px solid #000; padding: 15px; text-align: right; font-weight: bold;">GRAN TOTAL:</td>
                        <td style="border: 1px solid #000; padding: 15px; text-align: center; font-weight: bold; color: #0d47a1;">{t_gasoil:,.0f} Lts</td>
                        <td style="border: 1px solid #000; padding: 15px; text-align: center; font-weight: bold; color: #e65100;">{t_gasolina:,.0f} Lts</td>
                        <td style="border: 1px solid #000; padding: 15px; text-align: center; font-weight: bold; font-size: 20px;">{t_surtido_total:,.0f} Lts</td>
                        <td style="border: 1px solid #000; padding: 15px; text-align: center; font-weight: bold;">100%</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    # Módulo Creador de Bloques Estadísticos (Adaptativo)
    def crear_bloque_estadistica(gasoil, gasolina, tercero, label_tercero="TOTAL EXTRACCIONES"):
        return f"""
        <div style="display: flex; justify-content: space-around; background-color: #f8f9fa; border: 2px solid #000; padding: 15px; margin-bottom: 20px;">
            <div style="text-align: center; border-right: 2px solid #ccc; flex: 1;">
                <span style="font-size: 13px; font-weight: bold; color: #666;">TOTAL GASOIL SURTIDO</span><br>
                <span style="font-size: 28px; font-weight: bold; color: #0d47a1;">{gasoil:,.0f} Lts</span>
            </div>
            <div style="text-align: center; border-right: 2px solid #ccc; flex: 1;">
                <span style="font-size: 13px; font-weight: bold; color: #666;">TOTAL GASOLINA SURTIDA</span><br>
                <span style="font-size: 28px; font-weight: bold; color: #e65100;">{gasolina:,.0f} Lts</span>
            </div>
            <div style="text-align: center; flex: 1;">
                <span style="font-size: 13px; font-weight: bold; color: #666;">{label_tercero}</span><br>
                <span style="font-size: 28px; font-weight: bold; color: #2e7d32;">{tercero:,.0f} Lts</span>
            </div>
        </div>
        """

    bloque_p1 = crear_bloque_estadistica(t_gasoil_p1, t_gasol_p1, t_extrac_p1, "TOTAL EXTRACCIONES")
    bloque_p2 = crear_bloque_estadistica(t_gasoil_p2, t_gasol_p2, t_surtido_p2, "TOTAL SURTIDO (LTS)")
    bloque_p3 = crear_bloque_estadistica(t_gasoil, t_gasolina, t_extraccion, "TOTAL EXTRACCIONES")

    def crear_cabecera_html(num_pizarra, titulo):
        return f"""
        <table style="width: 100%; border-collapse: collapse; background-color: {color_header}; color: white; margin-bottom: 20px;">
            <tr>
                <td style="padding: 20px; width: 25%; text-align: center; vertical-align: middle;">{area_logo}</td>
                <td style="padding: 20px; font-size: 26px; font-weight: bold; text-align: center; width: 50%; text-transform: uppercase; letter-spacing: 2px;">REPORTE DE SURTIDO ({num_pizarra}/3)<br><span style="font-size:16px; font-weight:normal; letter-spacing: 0px;">{titulo}</span></td>
                <td style="padding: 20px; font-size: 18px; text-align: center; width: 25%; font-weight: bold; vertical-align: middle;">Fecha: {fecha_str}</td>
            </tr>
        </table>
        """

    html_final = f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    
    <div style="text-align:right; margin-bottom:15px; display: flex; justify-content: flex-end; gap: 15px;">
        <button onclick="descargar_p1()" style="background:#1565c0;color:#fff;border:none;padding:12px 15px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:13px;box-shadow:0 4px 6px rgba(0,0,0,0.2);">⬇️ 1: Corta/Extr</button>
        <button onclick="descargar_p2()" style="background:#1565c0;color:#fff;border:none;padding:12px 15px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:13px;box-shadow:0 4px 6px rgba(0,0,0,0.2);">⬇️ 2: Centro/Occ</button>
        <button onclick="descargar_p3()" style="background:#e65100;color:#fff;border:none;padding:12px 15px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:13px;box-shadow:0 4px 6px rgba(0,0,0,0.2);">⬇️ 3: Resumen Tipos</button>
    </div>
    
    <div id="pizarra-surtido-1" style="font-family: Arial, sans-serif; width: 1100px; margin: 0 auto 50px auto; background-color: #fff; border: 3px solid #000; box-sizing: border-box;">
        {crear_cabecera_html("1", "Ruta Corta y Extracciones")}
        <div style="padding: 0 20px;">
            {bloque_p1}
            {html_corta}
            {html_extrac}
        </div>
        <div style="margin-top: 20px; text-align: center; font-size: 12px; color: #777; border-top: 1px solid #ccc; padding: 10px;">Departamento de Flota y Logística</div>
    </div>

    <div id="pizarra-surtido-2" style="font-family: Arial, sans-serif; width: 1100px; margin: 0 auto 50px auto; background-color: #fff; border: 3px solid #000; box-sizing: border-box;">
        {crear_cabecera_html("2", "Ruta Centro Occidente y Occidente")}
        <div style="padding: 0 20px;">
            {bloque_p2}
            {html_centro}
            {html_larga}
        </div>
        <div style="margin-top: 20px; text-align: center; font-size: 12px; color: #777; border-top: 1px solid #ccc; padding: 10px;">Departamento de Flota y Logística</div>
    </div>
    
    <div id="pizarra-surtido-3" style="font-family: Arial, sans-serif; width: 1100px; margin: 0 auto 50px auto; background-color: #fff; border: 3px solid #000; box-sizing: border-box;">
        {crear_cabecera_html("3", "Resumen Ejecutivo por Tipo de Surtido")}
        <div style="padding: 0 20px;">
            {bloque_p3}
            <h3 style="text-align: center; color: {color_header}; font-size: 22px; text-transform: uppercase; margin-top: 40px;">Distribución Estratégica del Combustible</h3>
            {html_resumen_tipos}
        </div>
        <div style="margin-top: 40px; text-align: center; font-size: 12px; color: #777; border-top: 1px solid #ccc; padding: 10px;">Departamento de Flota y Logística</div>
    </div>
    
    <script>
    function descargar_p1() {{ 
        html2canvas(document.getElementById('pizarra-surtido-1'), {{scale: 2}}).then(canvas => {{ 
            var link = document.createElement('a'); 
            link.download = 'Pizarra_Combustible_1_{fecha_str.replace("/","-")}.png'; 
            link.href = canvas.toDataURL(); link.click(); 
        }}); 
    }}
    function descargar_p2() {{ 
        html2canvas(document.getElementById('pizarra-surtido-2'), {{scale: 2}}).then(canvas => {{ 
            var link = document.createElement('a'); 
            link.download = 'Pizarra_Combustible_2_{fecha_str.replace("/","-")}.png'; 
            link.href = canvas.toDataURL(); link.click(); 
        }}); 
    }}
    function descargar_p3() {{ 
        html2canvas(document.getElementById('pizarra-surtido-3'), {{scale: 2}}).then(canvas => {{ 
            var link = document.createElement('a'); 
            link.download = 'Pizarra_Combustible_3_{fecha_str.replace("/","-")}.png'; 
            link.href = canvas.toDataURL(); link.click(); 
        }}); 
    }}
    </script>
    """
    return html_final

# ==========================================
# INTERFAZ PRINCIPAL - TABS
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["📍 Carga de Despachos", "📈 Rep. Monitoreo", "⛽ Surtidos de Combustible", "📄 Informes PDF"])

# ---------------------------------------------------------
# TAB 1: CARGA DE DESPACHOS
# ---------------------------------------------------------
with tab1:
    st.title("📍 PCD - Monitoreo y Kilometraje")
    
    modo_despacho = st.radio("⚙️ Modo de Operación (Despachos):", ["📥 Cargar Nuevos Excel (Procesar)", "🔍 Consultar Pizarra Histórica (Sheets)"], horizontal=True)
    
    if modo_despacho == "📥 Cargar Nuevos Excel (Procesar)":
        col_d, col_k = st.columns(2)
        with col_d:
            fecha_reporte = st.date_input("📅 Selecciona la fecha del reporte:", datetime.now())
            fecha_str = fecha_reporte.strftime("%d/%m/%Y")
        with col_k:
            f_km = st.file_uploader("🗺️ Subir Excel KILOMETRAJE (Opcional, para cruzar kms)", type=["xlsx", "xls", "xlsm"])

        st.markdown("---")
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("### 🟠 ORIENTE")
            st.caption("Copia la tabla desde Excel (desde ITEM hasta BULTOS) y pégala abajo.")
            txt_oriente = st.text_area("📋 Pegar datos de Oriente:", height=150)
        with col2:
            st.markdown("### 🟢 OCCIDENTE")
            f_occidente = st.file_uploader("📂 Subir Excel OCCIDENTE", type=["xlsx", "xls", "xlsm"], key="occ")
        with col3:
            st.markdown("### 🔴 CENTRO OCCIDENTE")
            f_centro = st.file_uploader("📂 Subir Excel CENTRO OCCIDENTE", type=["xlsx", "xls", "xlsm"], key="cen")

        if st.button("1️⃣ Extraer Datos y Cruzar KMs", type="primary", use_container_width=True):
            with st.spinner("Extrayendo datos de los Excel..."):
                if 'final_dfs_despachos' in st.session_state: del st.session_state['final_dfs_despachos']
                st.session_state['log_cruces_km'] = []
                dict_dfs = {}
                dict_km = {}
                if f_km:
                    dict_km, msj_km = procesar_excel_km(f_km, fecha_reporte)
                    if msj_km != "OK": st.warning(msj_km)
                    else: st.success("✅ Kilometraje cargado exitosamente en memoria.")
                
                if txt_oriente:
                    df_oriente, msj_or = procesar_texto_oriente(txt_oriente)
                    if df_oriente is not None:
                        df_oriente = aplicar_kilometraje(df_oriente, dict_km, "Oriente")
                        df_oriente['Region'] = "ORIENTE"
                        dict_dfs["ORIENTE"] = df_oriente
                    else: st.error(f"Oriente: {msj_or}")
                    
                if f_occidente:
                    df_occidente, msj_oc = procesar_excel_region(f_occidente, ["PIZZARRA", "OCCIDENTE"], "Occidente")
                    if df_occidente is not None:
                        df_occidente = aplicar_kilometraje(df_occidente, dict_km, "Occidente")
                        df_occidente['Region'] = "OCCIDENTE"
                        dict_dfs["OCCIDENTE"] = df_occidente
                    else: st.error(f"Occidente: {msj_oc}")
                    
                if f_centro:
                    df_centro, msj_ce = procesar_excel_region(f_centro, ["TABLA", "CENTRO"], "Centro")
                    if df_centro is not None:
                        df_centro = aplicar_kilometraje(df_centro, dict_km, "Centro")
                        df_centro['Region'] = "CENTRO"
                        dict_dfs["CENTRO"] = df_centro
                    else: st.error(f"Centro: {msj_ce}")

                if dict_dfs:
                    st.session_state['pre_dfs_despachos'] = dict_dfs
                    st.session_state['fecha_str_despachos'] = fecha_str
                    st.success("✅ ¡Datos extraídos! Por favor revisa y edita los números en el paso 2 abajo.")
                else:
                    st.error("No se extrajeron datos para procesar.")

        if 'pre_dfs_despachos' in st.session_state:
            st.markdown("---")
            st.markdown("### ✏️ Paso 2: Revisión y Edición Manual")
            edited_dict_dfs = {}
            for reg, df in st.session_state['pre_dfs_despachos'].items():
                st.write(f"**Ruta {reg}**")
                edited_df = st.data_editor(df, key=f"editor_desp_{reg}", use_container_width=True, hide_index=True)
                edited_dict_dfs[reg] = edited_df
                
            if st.button("2️⃣ Confirmar Edición y Generar Pizarra Final", type="primary", use_container_width=True):
                st.session_state['final_dfs_despachos'] = edited_dict_dfs

        if 'final_dfs_despachos' in st.session_state:
            st.markdown("---")
            st.success("✅ ¡Mega-Pizarra Nacional Generada Exitosamente!")
            dict_dfs = st.session_state['final_dfs_despachos']
            fecha_s = st.session_state['fecha_str_despachos']
            components.html(html_pizarra_nacional(dict_dfs, fecha_s), height=1400, scrolling=True)
            
            df_global = pd.concat(list(dict_dfs.values()), ignore_index=True)
            st.session_state['datos_guardar_monitoreo'] = df_global
            
            st.markdown("---")
            st.subheader("📱 Mensaje de WhatsApp (Resumen para Gerencia)")
            g_rutas = len(df_global); g_cubiertos = df_global['CUBIERTOS'].sum(); g_bultos = df_global['BULTOS'].sum(); g_kms = df_global['KILOMETROS'].sum()
            st.code(generar_ws_nacional(dict_dfs, g_rutas, g_cubiertos, g_bultos, g_kms, fecha_s), language="markdown")

        if st.session_state.get('datos_guardar_monitoreo') is not None:
            if st.button("💾 Guardar Datos en la Bóveda (Sheets)", use_container_width=True):
                with st.spinner("Enviando datos a Google Sheets..."):
                    df_g = st.session_state['datos_guardar_monitoreo']
                    df_to_save = df_g[['UNIDAD', 'CHOFER', 'AYUDANTE', 'DESPACHOS', 'CUBRIR', 'CUBIERTOS', 'PENDIENTES', 'BULTOS', 'KILOMETROS', 'Region']].copy()
                    try:
                        dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                        fecha_rep_obj = datetime.strptime(st.session_state['fecha_str_despachos'], "%d/%m/%Y")
                        dia_str = dias_semana[fecha_rep_obj.weekday()]
                        df_to_save.insert(0, 'Fecha', st.session_state['fecha_str_despachos'])
                        df_to_save.insert(1, 'Dia_Semana', dia_str)
                    except: pass 
                    
                    if guardar_en_google_sheets(df_to_save, "MONITOREO_DESPACHOS"):
                        st.success("✅ ¡Datos detallados de Monitoreo y Kilometraje guardados correctamente!")
                        st.session_state['datos_guardar_monitoreo'] = None
                        if 'final_dfs_despachos' in st.session_state: del st.session_state['final_dfs_despachos']
                        if 'pre_dfs_despachos' in st.session_state: del st.session_state['pre_dfs_despachos']
    
    else:
        st.markdown("### 🔍 Máquina del Tiempo: Recuperar Pizarra de Despachos")
        fecha_consulta = st.date_input("📅 Selecciona la fecha a consultar:", datetime.now(), key="fecha_hist_desp")
        f_str_cons = fecha_consulta.strftime("%d/%m/%Y")
        if st.button("🔍 Buscar en Base de Datos", type="primary"):
            with st.spinner("Buscando en la Bóveda (Google Sheets)..."):
                df_db = extraer_datos_sheets("MONITOREO_DESPACHOS")
                if not df_db.empty:
                    df_db['Fecha_DT'] = pd.to_datetime(df_db['Fecha'], format='%d/%m/%Y', errors='coerce')
                    df_dia = df_db[df_db['Fecha_DT'].dt.date == fecha_consulta].copy()
                    if not df_dia.empty:
                        dict_dfs = {}
                        for reg in df_dia['Region'].unique():
                            df_reg = df_dia[df_dia['Region'] == reg].copy()
                            df_reg = df_reg.reset_index(drop=True)
                            df_reg['ITEM'] = range(1, len(df_reg) + 1)
                            for col in ['CUBRIR', 'CUBIERTOS', 'PENDIENTES', 'BULTOS', 'KILOMETROS']:
                                df_reg[col] = pd.to_numeric(df_reg[col], errors='coerce').fillna(0)
                            dict_dfs[reg] = df_reg
                        st.success(f"✅ ¡Pizarra recuperada! ({len(df_dia)} rutas encontradas para el {f_str_cons})")
                        components.html(html_pizarra_nacional(dict_dfs, f_str_cons), height=1400, scrolling=True)
                        st.markdown("---")
                        st.subheader("📱 Mensaje de WhatsApp Recuperado")
                        st.code(generar_ws_nacional(dict_dfs, len(df_dia), df_dia['CUBIERTOS'].sum(), df_dia['BULTOS'].sum(), df_dia['KILOMETROS'].sum(), f_str_cons), language="markdown")
                    else: st.warning(f"⚠️ No hay registros de Despachos guardados para la fecha {f_str_cons}.")
                else: st.error("No se pudo conectar con la Base de Datos o la hoja está vacía.")

# ---------------------------------------------------------
# TAB 2: REPORTE GERENCIAL
# ---------------------------------------------------------
with tab2:
    st.header("📉 Reporte Gerencial de Despachos y KMs")
    if st.button("🔄 Sincronizar Base de Datos", type="primary", key="btn_sync_mon"):
        with st.spinner("Conectando con la Bóveda..."):
            st.cache_data.clear()
            st.session_state['db_monitoreo'] = extraer_datos_sheets("MONITOREO_DESPACHOS")
            st.success("¡Base de datos cargada!")
            
    if 'db_monitoreo' in st.session_state and not st.session_state['db_monitoreo'].empty:
        db_m = st.session_state['db_monitoreo']
        c_cubiertos = 'CUBIERTOS' if 'CUBIERTOS' in db_m.columns else ('Clientes_Cubiertos' if 'Clientes_Cubiertos' in db_m.columns else None)
        c_kms = 'KILOMETROS' if 'KILOMETROS' in db_m.columns else ('Kilometros' if 'Kilometros' in db_m.columns else None)
        
        def to_float_seguro(val):
            try:
                if isinstance(val, (int, float)): return float(val)
                v_clean = str(val).replace(',', '').strip()
                if v_clean == '': return 0.0
                return float(v_clean)
            except: return 0.0

        for col in [c_cubiertos, c_kms]:
            if col and col in db_m.columns: db_m[col] = db_m[col].apply(to_float_seguro)
                
        db_m['Fecha_DT'] = pd.to_datetime(db_m['Fecha'], format='%d/%m/%Y', errors='coerce')
        db_m = db_m.sort_values('Fecha_DT')
        
        st.markdown("---")
        c1, c2 = st.columns(2)
        rango_m = c1.date_input("📅 Rango de Fechas a Evaluar:", [])
        filtro_region = c2.selectbox("🌎 Filtrar por Región:", ["Todas"] + list(db_m['Region'].unique()))
        
        if len(rango_m) == 2:
            f_ini, f_fin = rango_m
            db_m_filt = db_m[(db_m['Fecha_DT'].dt.date >= f_ini) & (db_m['Fecha_DT'].dt.date <= f_fin)].copy()
            if filtro_region != "Todas": db_m_filt = db_m_filt[db_m_filt['Region'] == filtro_region]
            
            if db_m_filt.empty: st.warning("⚠️ No hay registros en ese rango/región.")
            else:
                t_farm = db_m_filt[c_cubiertos].sum() if c_cubiertos else 0
                t_kms = db_m_filt[c_kms].sum() if c_kms else 0
                dias_unicos = db_m_filt['Fecha'].nunique()
                promedio_pedidos = t_farm / dias_unicos if dias_unicos > 0 else 0
                
                col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
                col_kpi1.metric("Promedio Diario de Pedidos", f"{promedio_pedidos:,.0f}")
                col_kpi2.metric("Total Kms Recorridos", f"{t_kms:,.2f} Km")
                col_kpi3.metric("Total de Pedidos Entregados", f"{t_farm:,.0f}")
                
                st.markdown("---")
                df_resumen_fecha = db_m_filt.groupby(['Fecha', 'Dia_Semana']).agg({c_cubiertos: 'sum', c_kms: 'sum', 'Fecha_DT': 'first'}).reset_index().sort_values('Fecha_DT')
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_resumen_fecha['Fecha'], y=df_resumen_fecha[c_cubiertos], name="Pedidos Entregados", marker_color='#1a4685'))
                fig.add_trace(go.Scatter(x=df_resumen_fecha['Fecha'], y=df_resumen_fecha[c_kms], name="Kilometraje Total", mode='lines+markers', yaxis='y2', line=dict(color='#d32f2f', width=3)))
                fig.update_layout(title="Pedidos Entregados vs Kilometraje por Día", yaxis=dict(title="Pedidos", side='left'), yaxis2=dict(title="Kilómetros", overlaying='y', side='right'), legend=dict(x=0, y=1.1, orientation="h"), height=450)
                st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------
# TAB 3: SURTIDOS DE COMBUSTIBLE
# ---------------------------------------------------------
with tab3:
    st.title("⛽ Control de Surtidos y Extracciones")
    modo_comb = st.radio("⚙️ Modo de Operación (Combustible):", ["📥 Cargar Nuevos Excel (Procesar)", "🔍 Consultar Pizarra Histórica (Sheets)"], horizontal=True)

    if modo_comb == "📥 Cargar Nuevos Excel (Procesar)":
        st.info("Carga los 3 archivos Excel. El sistema identificará automáticamente los surtidos y las extracciones.")
        fecha_comb = st.date_input("📅 Fecha del Surtido:", datetime.now(), key="fecha_comb_ui")
        f_str_comb = fecha_comb.strftime("%d/%m/%Y")
        
        cc1, cc2, cc3 = st.columns(3)
        with cc1: f_corta = st.file_uploader("📂 Excel: Ruta Corta", type=["xlsx", "xls", "xlsm"], key="fcorta")
        with cc2: f_centro_ext = st.file_uploader("📂 Excel: Centro / Extracción", type=["xlsx", "xls", "xlsm"], key="fcentro")
        with cc3: f_larga = st.file_uploader("📂 Excel: Ruta Larga", type=["xlsx", "xls", "xlsm"], key="flarga")
            
        if st.button("⚡ Procesar Surtidos y Generar Pizarras", type="primary", use_container_width=True):
            with st.spinner("Procesando litros de combustible y construyendo Pizarras Gemelas..."):
                datos_combustible = []
                if f_corta:
                    df_corta, msg = procesar_excel_surtidos(f_corta, "RUTA CORTA", fecha_comb)
                    if df_corta is not None and not df_corta.empty: datos_combustible.append(df_corta)
                    elif msg != "OK": st.warning(f"Aviso Corta: {msg}")
                if f_centro_ext:
                    df_ce, msg = procesar_excel_surtidos(f_centro_ext, "RUTA CENTRO", fecha_comb)
                    if df_ce is not None and not df_ce.empty: datos_combustible.append(df_ce)
                    elif msg != "OK": st.warning(f"Aviso Centro/Extracción: {msg}")
                if f_larga:
                    df_larga, msg = procesar_excel_surtidos(f_larga, "RUTA OCCIDENTE", fecha_comb)
                    if df_larga is not None and not df_larga.empty: datos_combustible.append(df_larga)
                    elif msg != "OK": st.warning(f"Aviso Occidente: {msg}")
                    
                if datos_combustible:
                    st.success("✅ ¡Datos procesados! Renderizando Pizarras de Combustible...")
                    df_comb_final = pd.concat(datos_combustible, ignore_index=True)
                    
                    components.html(html_pizarras_combustible_completas(df_comb_final, f_str_comb), height=2500, scrolling=True)
                    
                    st.markdown("---")
                    st.subheader("📱 Mensajes de WhatsApp")
                    cw1, cw2, cw3 = st.columns(3)
                    with cw1:
                        st.markdown("**📱 Pizarra 1: Corta y Extracción**")
                        st.code(generar_ws_surtido_p1(df_comb_final, f_str_comb), language="markdown")
                    with cw2:
                        st.markdown("**📱 Pizarra 2: Centro y Occidente**")
                        st.code(generar_ws_surtido_p2(df_comb_final, f_str_comb), language="markdown")
                    with cw3:
                        st.markdown("**📱 Pizarra 3: Resumen Gerencial**")
                        st.code(generar_ws_surtido_p3(df_comb_final, f_str_comb), language="markdown")
                        
                    st.session_state['datos_guardar_combustible'] = df_comb_final
                else: st.warning("No se detectaron datos válidos para procesar en la fecha seleccionada.")

        if st.session_state.get('datos_guardar_combustible') is not None:
            if st.button("💾 Guardar Combustible en Sheets", use_container_width=True):
                with st.spinner("Guardando en base de datos..."):
                    df_g = st.session_state['datos_guardar_combustible']
                    df_to_save = df_g[['GRUPO', 'UNIDAD', 'CHOFER', 'RUTA', 'HORA', 'SITIO', 'TIPO_SURTIDO', 'COMBUSTIBLE', 'LITROS']].copy()
                    df_to_save.insert(0, 'Fecha', f_str_comb)
                    if guardar_en_google_sheets(df_to_save, "SURTIDO_COMBUSTIBLE"):
                        st.success("✅ ¡Datos detallados de Surtido guardados correctamente!")
                        st.session_state['datos_guardar_combustible'] = None

    else:
        st.markdown("### 🔍 Máquina del Tiempo: Recuperar Pizarra de Surtidos")
        fecha_consulta_comb = st.date_input("📅 Selecciona la fecha a consultar:", datetime.now(), key="fecha_hist_comb")
        f_str_cons_comb = fecha_consulta_comb.strftime("%d/%m/%Y")
        if st.button("🔍 Buscar en Base de Datos", type="primary"):
            with st.spinner("Buscando en la Bóveda (Google Sheets)..."):
                df_db = extraer_datos_sheets("SURTIDO_COMBUSTIBLE")
                if not df_db.empty:
                    df_db['Fecha_DT'] = pd.to_datetime(df_db['Fecha'], format='%d/%m/%Y', errors='coerce')
                    df_dia = df_db[df_db['Fecha_DT'].dt.date == fecha_consulta_comb].copy()
                    if not df_dia.empty:
                        df_dia['LITROS'] = pd.to_numeric(df_dia['LITROS'], errors='coerce').fillna(0)
                        st.success(f"✅ ¡Pizarra recuperada! ({len(df_dia)} registros encontrados para el {f_str_cons_comb})")
                        components.html(html_pizarras_combustible_completas(df_dia, f_str_cons_comb), height=2500, scrolling=True)
                        
                        st.markdown("---")
                        st.subheader("📱 Mensajes de WhatsApp Recuperados")
                        cw1, cw2, cw3 = st.columns(3)
                        with cw1:
                            st.markdown("**📱 Pizarra 1: Corta y Extracción**")
                            st.code(generar_ws_surtido_p1(df_dia, f_str_cons_comb), language="markdown")
                        with cw2:
                            st.markdown("**📱 Pizarra 2: Centro y Occidente**")
                            st.code(generar_ws_surtido_p2(df_dia, f_str_cons_comb), language="markdown")
                        with cw3:
                            st.markdown("**📱 Pizarra 3: Resumen Gerencial**")
                            st.code(generar_ws_surtido_p3(df_dia, f_str_cons_comb), language="markdown")
                    else: st.warning(f"⚠️ No hay registros de Surtido guardados para la fecha {f_str_cons_comb}.")
                else: st.error("No se pudo conectar con la Base de Datos o la hoja está vacía.")

# ---------------------------------------------------------
# TAB 4: GENERADOR DE INFORMES PDF
# ---------------------------------------------------------
with tab4:
    st.title("📄 Generador de Informes Gerenciales (PDF)")
    st.info("Genera un reporte gerencial profesional (A4) consolidado por Semanas y Total Mensual, listo para imprimir o guardar como PDF.")
    
    meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
    
    col_f1, col_f2 = st.columns(2)
    with col_f1: mes_sel = st.selectbox("Selecciona el Mes:", list(meses_dict.values()), index=datetime.now().month-1)
    with col_f2: anio_sel = st.number_input("Selecciona el Año:", min_value=2023, max_value=2030, value=datetime.now().year)
    
    if st.button("📊 Generar Informe PDF", type="primary", use_container_width=True):
        with st.spinner("Compilando bases de datos y calculando estadísticas semanales..."):
            num_mes = list(meses_dict.keys())[list(meses_dict.values()).index(mes_sel)]
            df_comb = extraer_datos_sheets("SURTIDO_COMBUSTIBLE")
            df_comb_mes = pd.DataFrame()
            if not df_comb.empty:
                df_comb['Fecha_DT'] = pd.to_datetime(df_comb['Fecha'], format='%d/%m/%Y', errors='coerce')
                df_comb_mes = df_comb[(df_comb['Fecha_DT'].dt.month == num_mes) & (df_comb['Fecha_DT'].dt.year == anio_sel)].copy()
            
            if df_comb_mes.empty:
                st.warning(f"No hay datos registrados en Google Sheets para el mes de {mes_sel} {anio_sel}.")
            else:
                df_comb_mes['LITROS'] = pd.to_numeric(df_comb_mes['LITROS'], errors='coerce').fillna(0)
                df_comb_mes['Dia_del_Mes'] = df_comb_mes['Fecha_DT'].dt.day
                df_comb_mes['Semana_Mes'] = df_comb_mes['Dia_del_Mes'].apply(lambda x: f"Semana {(x - 1) // 7 + 1}")
                
                resumen_semanal = df_comb_mes.groupby('Semana_Mes').agg(
                    Total_Gasolina=('LITROS', lambda x: df_comb_mes.loc[x.index][df_comb_mes.loc[x.index, 'COMBUSTIBLE'].str.contains('GASOLINA', na=False)]['LITROS'].sum()),
                    Total_Gasoil=('LITROS', lambda x: df_comb_mes.loc[x.index][df_comb_mes.loc[x.index, 'COMBUSTIBLE'].str.contains('GASOIL|DIESEL', na=False)]['LITROS'].sum()),
                    Total_Extracciones=('LITROS', lambda x: df_comb_mes.loc[x.index][df_comb_mes.loc[x.index, 'GRUPO'] == 'EXTRACCION']['LITROS'].sum()),
                    Gran_Total=('LITROS', 'sum')
                ).reset_index()
                
                color_header = "#1a4685"
                logo_b64 = obtener_logo_base64()
                logo_img = f'<img src="{logo_b64}" style="max-height: 50px;">' if logo_b64 else '<h3>DROTACA</h3>'
                
                filas_html = ""
                for _, row in resumen_semanal.iterrows():
                    filas_html += f"""
                    <tr style="text-align: center; border-bottom: 1px solid #ddd;">
                        <td style="padding: 12px; font-weight: bold;">{row['Semana_Mes']}</td>
                        <td style="padding: 12px; color: #e65100; font-weight: bold;">{row['Total_Gasolina']:,.0f} Lts</td>
                        <td style="padding: 12px; color: #0d47a1; font-weight: bold;">{row['Total_Gasoil']:,.0f} Lts</td>
                        <td style="padding: 12px; color: #2e7d32; font-weight: bold;">{row['Total_Extracciones']:,.0f} Lts</td>
                        <td style="padding: 12px; font-weight: bold; background-color: #f1f1f1;">{row['Gran_Total']:,.0f} Lts</td>
                    </tr>
                    """
                
                total_mes = resumen_semanal.sum(numeric_only=True)
                filas_html += f"""
                <tr style="text-align: center; background-color: {color_header}; color: white; font-weight: bold; font-size: 16px;">
                    <td style="padding: 15px;">TOTAL MENSUAL</td>
                    <td style="padding: 15px;">{total_mes['Total_Gasolina']:,.0f} Lts</td>
                    <td style="padding: 15px;">{total_mes['Total_Gasoil']:,.0f} Lts</td>
                    <td style="padding: 15px;">{total_mes['Total_Extracciones']:,.0f} Lts</td>
                    <td style="padding: 15px;">{total_mes['Gran_Total']:,.0f} Lts</td>
                </tr>
                """
                
                html_pdf = f"""
                <style>
                    @media print {{
                        .no-print {{ display: none !important; }}
                        body {{ margin: 0; padding: 0; background-color: white; }}
                        #page-a4 {{ border: none !important; width: 100% !important; padding: 0 !important; box-shadow: none !important; }}
                    }}
                    #page-a4 {{
                        background: white; width: 210mm; min-height: 297mm; margin: auto; padding: 20mm;
                        box-shadow: 0 0 10px rgba(0,0,0,0.1); font-family: 'Arial', sans-serif; box-sizing: border-box;
                    }}
                    .titulo {{ color: {color_header}; text-align: center; margin-bottom: 5px; text-transform: uppercase; }}
                    .subtitulo {{ text-align: center; color: #555; margin-top: 0; margin-bottom: 30px; font-size: 18px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                    th {{ background-color: #f2f2f2; color: #333; padding: 12px; border-bottom: 2px solid #ccc; }}
                </style>
                
                <div class="no-print" style="text-align: center; margin-bottom: 20px;">
                    <button onclick="window.print()" style="background-color: #d32f2f; color: white; border: none; padding: 15px 30px; font-size: 18px; font-weight: bold; cursor: pointer; border-radius: 5px; box-shadow: 0 4px 6px rgba(0,0,0,0.2);">
                        🖨️ IMPRIMIR / GUARDAR COMO PDF
                    </button>
                    <p style="color: #666; font-size: 14px;">*En la ventana de impresión, selecciona 'Destino: Guardar como PDF'.</p>
                </div>

                <div id="page-a4">
                    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid {color_header}; padding-bottom: 10px; margin-bottom: 30px;">
                        <div>{logo_img}</div>
                        <div style="text-align: right; color: #666; font-size: 14px;">
                            <b>Fecha de Emisión:</b> {datetime.now().strftime("%d/%m/%Y")}<br>
                            <b>Departamento:</b> Flota y Logística
                        </div>
                    </div>
                    
                    <h1 class="titulo">REPORTE GERENCIAL DE COMBUSTIBLE</h1>
                    <h3 class="subtitulo">Periodo: {mes_sel.upper()} {anio_sel}</h3>
                    
                    <p style="font-size: 14px; color: #333; line-height: 1.6; text-align: justify;">
                        El presente documento expone el resumen consolidado de las operaciones de surtido de combustible a nivel nacional, abarcando las rutas de Oriente, Centro Occidente y Occidente, así como las extracciones realizadas en base, agrupadas por semanas operativas correspondientes al mes indicado.
                    </p>
                    
                    <table>
                        <thead>
                            <tr>
                                <th>SEMANA OPERATIVA</th>
                                <th>GASOLINA</th>
                                <th>GASOIL</th>
                                <th>EXTRACCIONES</th>
                                <th>TOTAL CONSOLIDADO</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filas_html}
                        </tbody>
                    </table>
                    
                    <div style="margin-top: 50px; padding: 15px; border: 1px solid #ccc; background-color: #fafafa; border-radius: 5px;">
                        <h4 style="margin: 0 0 10px 0; color: {color_header};">📌 Nota Operativa:</h4>
                        <p style="margin: 0; font-size: 13px; color: #555;">
                            - Los datos reflejados provienen de las pizarras logísticas diarias cerradas y auditadas.<br>
                            - Semana 1 corresponde del 1 al 7 del mes; Semana 2 del 8 al 14, y así sucesivamente.
                        </p>
                    </div>
                    
                    <div style="margin-top: 100px; display: flex; justify-content: space-around; text-align: center;">
                        <div style="width: 250px;">
                            <hr style="border-top: 1px solid #000;">
                            <b>Firma y Sello</b><br>
                            <span style="font-size: 12px; color: #666;">Gerencia de Operaciones</span>
                        </div>
                        <div style="width: 250px;">
                            <hr style="border-top: 1px solid #000;">
                            <b>Firma y Sello</b><br>
                            <span style="font-size: 12px; color: #666;">Dirección General</span>
                        </div>
                    </div>
                </div>
                """
                components.html(html_pdf, height=1200, scrolling=True)