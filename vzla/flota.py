# ==========================================
# Archivo: flota.py (Módulo de Mantenimiento y Gastos)
# ==========================================
import streamlit as st
import pandas as pd
import re
import base64
from pathlib import Path
from datetime import datetime, timedelta
import streamlit.components.v1 as components
import gspread
import textwrap
import traceback
from google.oauth2.service_account import Credentials

# ==========================================
# DICCIONARIOS Y DATA MAESTRA DE FLOTA
# ==========================================
FLOTA_MASTER_DATA = [
    ("FORD", "CARGO", "A02CS4V"), ("CHEVROLET", "BUS", "A0378AK"), ("CHEVROLET", "BUS", "A0424AK"),
    ("FOTON", "FOTON", "A04AV4T"), ("FORD", "CARGO", "A07CZ7A"),
    ("MITSUBISHI", "L200", "A12BK8R"), 
    ("MITSUBISHI", "L300", "A13CS4G"), ("MITSUBISHI", "L300", "A15AT9B"), 
    ("MITSUBISHI", "L300", "A20CD8V"), ("MITSUBISHI", "L200", "A20EF9G"), ("MITSUBISHI", "CANTER", "A21BC1D"),
    ("FOTON", "FOTON", "A24AS1R"), ("MITSUBISHI", "CANTER", "A28AR0B"), ("MITSUBISHI", "CANTER", "A28BS4D"),
    ("CHEVROLET", "C 3500", "A29AI9V"), 
    ("CHEVROLET", "C 3500", "A30AS1A"), ("MITSUBISHI", "CANTER", "A31BS3D"),
    ("PEUGEOT", "PEUGEOT", "A33B16D"), ("MITSUBISHI", "CANTER", "A35BK2D"), ("MITSUBISHI", "CANTER", "A38BK6D"),
    ("MITSUBISHI", "CANTER", "A39BK8D"), ("MITSUBISHI", "L300", "A40AK1N"), ("DONG FENG", "DONG FENG", "A41AV9T"),
    ("TOYOTA", "HIACE", "A41BD1F"), ("FORD", "SUPER DUTY", "A42CV9G"), ("FORD", "SUPER DUTY", "A43CV4G"),
    ("CHEVROLET", "C 3500", "A50AU2B"),
    ("CHEVROLET", "FSR", "A53CG1D"),
    ("CHEVROLET", "C 3500", "A54DL8G"), ("CHEVROLET", "C 3500", "A57AU7B"), ("MITSUBISHI", "CANTER", "A60AV4F"),
    ("FORD", "GRUA", "A62AZ9D"), ("HINO MOTORS", "HINO MOTORS", "A64EH14"), ("CHEVROLET", "C 3500", "A65AT5R"),
    ("CHEVROLET", "C 3500", "A65AT9R"), ("CHEVROLET", "C 3500", "A65CY8V"), ("HINO MOTORS", "HINO MOTORS", "A65EH1A"),
    ("HINO MOTORS", "HINO MOTORS", "A65EH2A"), ("HINO MOTORS", "HINO MOTORS", "A65EH3A"), ("FORD", "SUPER DUTY", "A66BW8M"),
    ("HINO MOTORS", "HINO MOTORS", "A66EH5A"), ("MITSUBISHI", "L300", "A70AJ0O"), ("MITSUBISHI", "CANTER", "A70CJ6A"),
    ("CHEVROLET", "C 3500", "A70DL0G"), ("MITSUBISHI", "CANTER", "A71CJ9A"), ("MITSUBISHI", "L300", "A73CS4A"),
    ("DONG FENG", "DONG FENG", "A75EB7P"), ("CHEVROLET", "C 3500", "A80CL8D"),
    ("CHEVROLET", "C 3500", "A82AKQN"), ("CHEVROLET", "C 3500", "A91BF5F"), ("CHEVROLET", "C 3500", "A80CL2D"),
    ("CHEVROLET", "C 3500", "A91DB8M"), ("MITSUBISHI", "L300", "A94CS9A"), ("MITSUBISHI", "CANTER", "A95AS3F"),
    ("IVECO", "EUROCARGO", "A97CL6D"), ("MITSUBISHI", "TOURING", "AF793NA"),
    ("FORD", "FIESTA", "AF925DD"), ("FORD", "FIESTA", "AI532MG"),
    ("CHEVROLET", "GRAND VITARA", "LAY65J"), ("HONDA", "HONDA", "MDS43M")
]

VEHICULOS_VALIDOS = [f"{p} {m}" for ma, m, p in FLOTA_MASTER_DATA]

def normalizar_unidad(texto):
    texto = str(texto).upper().strip()
    if not texto: return ""
    for vehiculo in VEHICULOS_VALIDOS:
        placa = vehiculo.split(" ")[0]
        if placa in texto: return vehiculo
    return texto 

def obtener_dia_semana(fecha_str):
    try:
        obj = datetime.strptime(fecha_str, "%d/%m/%Y")
        dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
        return dias[obj.weekday()]
    except:
        return ""

# ==========================================
# CREADORES DE MENSAJES WHATSAPP (ESTILO ORACIÓN)
# ==========================================
def generar_ws_planificadas(df):
    if df.empty: return ""
    fecha = df['Fecha'].iloc[0]
    dia = obtener_dia_semana(fecha)
    total = len(df)
    return f"*Mantenimiento vehicular planificado* 🛠️\n📅 Fecha: {fecha} ({dia})\n\n*Resumen de actividades:*\nTotal mantenimientos planificados: {total}\n\n✅ Enviaremos los resultados al final de la jornada."

def generar_ws_realizadas(df):
    if df.empty: return ""
    fecha = df['Fecha'].iloc[0]
    dia = obtener_dia_semana(fecha)
    total = len(df)
    condiciones = df['Condición'].fillna('').astype(str).str.upper()
    operativos = len(condiciones[condiciones.str.contains('OPERATIVO')])
    en_proceso = total - operativos
    porcentaje = round((operativos / total) * 100, 1) if total > 0 else 0
    msg = f"*Mantenimiento vehicular realizado* 🚛\n📅 Fecha: {fecha} ({dia})\n\n"
    msg += f"*Resultados de la jornada:*\nTotal de mantenimientos: {total}\n🟢 Operativos: {operativos}\n🔴 Pendientes / en proceso: {en_proceso}\n📊 Avance de tareas: {porcentaje}%\n\n"
    return msg

def generar_ws_gastos(df, fecha_evaluada, tipo_categoria="Todo Incluido"):
    if df.empty: return ""
    total_usd = df['TOTAL $'].sum()
    total_items = len(df)
    resumen_tipo = df.groupby('TIPO')['TOTAL $'].sum().reset_index()
    msg = f"*Resumen diario de gastos ({tipo_categoria.title()})* 💰\n📅 Fecha: {fecha_evaluada}\n\n*Datos generales:*\n🧾 Gastos procesados: {total_items}\n💵 Inversión total: *${total_usd:,.2f}*\n\n*Distribución del gasto:*\n"
    for _, row in resumen_tipo.iterrows():
        msg += f"▪️ {row['TIPO'].title()}: ${row['TOTAL $']:,.2f}\n"
    msg += "\n✅ *Pizarra de desglose financiero adjunta.*"
    return msg

def generar_ws_combustible(df):
    if df.empty: return ""
    r = df.iloc[0]
    t1 = float(r['Tanque_1_50K'])
    t2 = float(r['Tanque_2_12K'])
    t3 = float(r['Tanque_3_7K'])
    total_tanques = t1 + t2 + t3
    dias_restantes = int(total_tanques / 1500) if total_tanques > 0 else 0
    de_str = str(r['De']).title()
    para_str = str(r['Para']).title()
    msg = f"*Reporte de combustible Drotaca* ⛽\n📅 Fecha: {r['Fecha']}\nDe: {de_str}\nPara: {para_str}\n\n"
    msg += f"*Reserva en base (bidones):*\n• Gasolina: {r['Gasolina_Bidones']:,.0f} Lts\n• Gasoil: {r['Gasoil_Bidones']:,.0f} Lts\n\n"
    msg += f"*Nivel de tanques (Ciudad Drotaca):*\n"
    msg += f"• Tanque principal (50.000 L): {t1:,.0f} Lts\n"
    msg += f"• Tanque 2 (12.000 L): {t2:,.0f} Lts\n"
    msg += f"• Tanque 3 (7.850 L): {t3:,.0f} Lts\n\n"
    msg += f"📊 *Total disponible:* {total_tanques:,.0f} Lts de gasoil\n"
    msg += f"⏳ *Proyección:* Tenemos gasoil para *{dias_restantes} días* en condiciones normales (1.500 Lts de consumo diario estimado).\n\n"
    msg += "✅ *Aval fotográfico e infografía adjunta.*"
    return msg

def generar_ws_estatus_dinamico(tipo_reporte, fecha, total, t_activas, t_inactivas, t_con_gps, t_sin_gps, inactivas_df, sin_gps_df):
    dia = obtener_dia_semana(fecha)
    
    if "Solo Flota" in tipo_reporte:
        msg = f"*Estatus Diario Operativo (Flota)* 🚚\n📅 Fecha: {fecha} ({dia})\n\n"
        msg += f"• Flota Total: {total} unidades\n• Operativas: {t_activas}\n• En Taller: {t_inactivas}\n\n"
        if not inactivas_df.empty:
            msg += "*Detalle Inoperativas:*\n"
            for _, row in inactivas_df.iterrows():
                msg += f"🔻 {row['Placa']} ({str(row['Modelo']).title()}) - {str(row['Motivo']).title()}\n"
            msg += "\n"
        msg += "✅ *Pizarra visual operativa adjunta.*"
        return msg
        
    elif "Solo GPS" in tipo_reporte:
        msg = f"*Estatus Diario Satelital (GPS)* 🛰️\n📅 Fecha: {fecha} ({dia})\n\n"
        msg += f"• Flota Total: {total} unidades\n• Transmitiendo: {t_con_gps}\n• Sin Transmisión: {t_sin_gps}\n\n"
        if not sin_gps_df.empty:
            msg += "*Detalle Sin Transmisión:*\n"
            for _, row in sin_gps_df.iterrows():
                msg += f"🔻 {row['Placa']} ({str(row['Modelo']).title()}) - {str(row['Motivo']).title()}\n"
            msg += "\n"
        msg += "✅ *Pizarra visual satelital adjunta.*"
        return msg
        
    else:
        msg = f"*Estatus Diario Operativo y Satelital (GPS)* 🚚🛰️\n📅 Fecha: {fecha} ({dia})\n\n"
        msg += f"*RESUMEN MECÁNICO (FLOTA)* 🔧\n• Flota Total: {total} unidades\n• Operativas: {t_activas}\n• En Taller: {t_inactivas}\n\n"
        if not inactivas_df.empty:
            msg += "*Detalle Inoperativas:*\n"
            for _, row in inactivas_df.iterrows():
                msg += f"🔻 {row['Placa']} ({str(row['Modelo']).title()}) - {str(row['Motivo']).title()}\n"
            msg += "\n"
            
        msg += f"*RESUMEN SATELITAL (GPS)* 📡\n• Transmitiendo: {t_con_gps}\n• Sin Transmisión: {t_sin_gps}\n\n"
        if not sin_gps_df.empty:
            msg += "*Detalle Sin Transmisión:*\n"
            for _, row in sin_gps_df.iterrows():
                msg += f"🔻 {row['Placa']} ({str(row['Modelo']).title()}) - {str(row['Motivo']).title()}\n"
            msg += "\n"
            
        msg += "✅ *Pizarra visual de estatus general adjunta.*"
        return msg

# ==========================================
# FUNCIÓN PARA INYECTAR LOGO
# ==========================================
def obtener_logo_base64():
    try:
        ruta_logo = Path(__file__).parent / "logo.png"
        if ruta_logo.exists():
            with open(ruta_logo, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode()
                return f"data:image/png;base64,{encoded_string}"
        return None
    except Exception:
        return None

# ==========================================
# PUENTE A GOOGLE SHEETS (MODERNIZADO)
# ==========================================
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

# RECONSTRUCTOR BLINDADO DE LLAVE
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"

CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def guardar_en_google_sheets(df_para_guardar, nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        
        # Conexión directa a prueba de balas usando tu ID
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        
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
                pass
            else:
                raise error_interno
                
        st.cache_data.clear() 
        return True
        
    except Exception as e:
        st.error(f"🛑 Falla en la conexión: {e}")
        st.code(traceback.format_exc(), language="python")
        return False

@st.cache_data(ttl=60)
def extraer_datos_sheets(nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        
        # Conexión directa a prueba de balas usando tu ID
        doc = cliente.open_by_key("1wCM3tcfQJtIQ4gDB0gLe9gJ4_ON7Vl6U4cBGuxXTKZ0")
        
        hoja = doc.worksheet(nombre_hoja)
        data = hoja.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()


# ==========================================
# EXTRACTORES DE TEXTO INTELIGENTES (NUEVOS)
# ==========================================
def procesar_texto_planificadas(texto):
    lineas = texto.split('\n')
    datos = []
    fecha_def = datetime.now().strftime("%d/%m/%Y")
    
    for linea_raw in lineas:
        linea = linea_raw.strip()
        if not linea or "UNIDAD" in linea.upper() or "ACTIVIDAD" in linea.upper(): continue
        
        # Detectar el separador REAL: TAB (pegado desde Excel) o pipe '|'. NO colapsar tabs.
        cols = None
        if '\t' in linea_raw:
            cols = [c.strip() for c in linea_raw.split('\t')]
        elif '|' in linea:
            cols = [c.strip() for c in linea.split('|')]
        
        if cols is not None:
            # Quitar columnas totalmente vacías en los extremos
            while cols and cols[0] == '': cols.pop(0)
            while cols and cols[-1] == '': cols.pop()
            if len(cols) < 2: continue
            
            # Fecha: si la primera columna tiene fecha, se usa; si no, la del día
            f_match = re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', cols[0])
            if f_match:
                f_val = f_match.group(0)
                fila = cols[1:]
            else:
                f_val = fecha_def
                fila = cols
            
            # fila = [UNIDAD, ACTIVIDAD..., MECÁNICO]
            if len(fila) >= 3:
                unidad = fila[0]
                mec = fila[-1]
                act = " - ".join([x for x in fila[1:-1] if x])
            elif len(fila) == 2:
                unidad, act, mec = fila[0], fila[1], "Sin asignar"
            else:
                unidad, act, mec = fila[0], "", "Sin asignar"
            
            unidad = normalizar_unidad(unidad) or unidad
            mec = mec.title() if mec and mec != "Sin asignar" else "Sin asignar"
            datos.append({"Fecha": f_val, "Unidad": unidad, "Actividad": act, "Mecánico": mec})
        else:
            # Respaldo: texto sin tabs ni '|' (heurística por espacios)
            f_val = fecha_def
            m_fecha = re.search(r'^(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+', linea)
            if m_fecha:
                f_val = m_fecha.group(1)
                linea = linea[m_fecha.end():].strip()
            
            linea = re.sub(r'\s+', ' ', linea)
            partes = linea.split()
            if len(partes) < 2: continue
            
            posible_unidad = f"{partes[0]} {partes[1]}"
            unidad = normalizar_unidad(posible_unidad)
            if not unidad or unidad == posible_unidad: 
                unidad = posible_unidad
            
            linea_resto = linea.replace(partes[0], "", 1).replace(partes[1], "", 1).strip()
            p_resto = linea_resto.split()
            if len(p_resto) >= 3:
                mec = f"{p_resto[-2]} {p_resto[-1]}"
                act = " ".join(p_resto[:-2])
            else:
                mec = "Sin asignar"
                act = linea_resto
            
            datos.append({"Fecha": f_val, "Unidad": unidad, "Actividad": act, "Mecánico": mec.title()})
            
    if not datos: return pd.DataFrame([{"Fecha": fecha_def, "Unidad": "", "Actividad": "", "Mecánico": ""}])
    return pd.DataFrame(datos)

def normalizar_condicion(cond):
    """Reduce cualquier texto de condición a solo 3 valores:
    OPERATIVO, PENDIENTE o EN PROCESO. Todo lo demás (REALIZADO, NO OPERATIVO,
    vacío o cualquier otra escritura) se convierte en EN PROCESO."""
    c = str(cond).upper().strip()
    if 'PENDIENTE' in c:
        return "PENDIENTE"
    if 'OPERATIVO' in c and 'NO OPERATIVO' not in c and 'NO-OPERATIVO' not in c and 'INOPERATIVO' not in c:
        return "OPERATIVO"
    return "EN PROCESO"

def _orden_condicion(cond):
    # Operativo primero; el resto abajo
    return {"OPERATIVO": 0, "EN PROCESO": 1, "PENDIENTE": 2}.get(str(cond).upper().strip(), 3)

def procesar_texto_realizadas(texto):
    lineas = texto.split('\n')
    datos = []
    fecha_def = datetime.now().strftime("%d/%m/%Y")
    
    for linea_raw in lineas:
        linea = linea_raw.strip()
        if not linea or "UNIDAD" in linea.upper() or "CONDICIÓN" in linea.upper() or "CONDICION" in linea.upper(): continue
        
        # Detectar separador REAL: TAB (pegado desde Excel) o pipe '|'. NO colapsar tabs.
        cols = None
        if '\t' in linea_raw:
            cols = [c.strip() for c in linea_raw.split('\t')]
        elif '|' in linea:
            cols = [c.strip() for c in linea.split('|')]
        
        if cols is not None:
            while cols and cols[0] == '': cols.pop(0)
            while cols and cols[-1] == '': cols.pop()
            if len(cols) < 2: continue
            
            f_match = re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', cols[0])
            if f_match:
                f_val = f_match.group(0)
                fila = cols[1:]
            else:
                f_val = fecha_def
                fila = cols
            
            # fila = [UNIDAD, RESUMEN..., CONDICIÓN]
            if len(fila) >= 3:
                unidad = fila[0]
                cond = fila[-1]
                act = " - ".join([x for x in fila[1:-1] if x])
            elif len(fila) == 2:
                unidad, act, cond = fila[0], fila[1], "OPERATIVO"
            else:
                unidad, act, cond = fila[0], "", "OPERATIVO"
            
            unidad = normalizar_unidad(unidad) or unidad
            datos.append({"Fecha": f_val, "Unidad": unidad, "Resumen Actividad": act, "Condición": normalizar_condicion(cond)})
        else:
            # Respaldo: texto sin tabs ni '|' (heurística por espacios)
            f_val = fecha_def
            m_fecha = re.search(r'^(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+', linea)
            if m_fecha:
                f_val = m_fecha.group(1)
                linea = linea[m_fecha.end():].strip()
            
            linea = re.sub(r'\s+', ' ', linea)
            partes = linea.split()
            if len(partes) < 2: continue
            
            posible_unidad = f"{partes[0]} {partes[1]}"
            unidad = normalizar_unidad(posible_unidad)
            if not unidad or unidad == posible_unidad: 
                unidad = posible_unidad
            
            linea_resto = linea.replace(partes[0], "", 1).replace(partes[1], "", 1).strip()
            linea_up = linea_resto.upper()
            cond = ""
            for kw in ["OPERATIVO", "REALIZADO", "EN PROCESO", "PENDIENTE"]:
                if linea_up.endswith(kw):
                    cond = kw
                    linea_resto = linea_resto[:len(linea_resto)-len(kw)].strip()
                    break
            if linea_resto.endswith("-") or linea_resto.endswith("|"):
                linea_resto = linea_resto[:-1].strip()
                
            datos.append({"Fecha": f_val, "Unidad": unidad, "Resumen Actividad": linea_resto, "Condición": normalizar_condicion(cond)})
            
    if not datos: return pd.DataFrame([{"Fecha": fecha_def, "Unidad": "", "Resumen Actividad": "", "Condición": ""}])
    df = pd.DataFrame(datos)
    # Ordenar: OPERATIVO primero, lo demás abajo
    df['_ord'] = df['Condición'].apply(_orden_condicion)
    df = df.sort_values(by='_ord', kind='stable').drop(columns=['_ord']).reset_index(drop=True)
    return df

# ==========================================
# LÓGICA DE EXCEL DE GASTOS
# ==========================================
def procesar_excel_gastos(df_raw):
    col_str = " ".join([str(c).upper() for c in df_raw.columns])
    if "FECHA" not in col_str and "UNIDAD" not in col_str:
        for i in range(min(15, len(df_raw))):
            fila_str = " ".join([str(x).upper() for x in df_raw.iloc[i].values])
            if "FECHA" in fila_str and ("UNIDAD" in fila_str or "PLACA" in fila_str):
                df_raw.columns = df_raw.iloc[i]
                df_raw = df_raw.iloc[i+1:].reset_index(drop=True)
                break
                
    mapa = {'FECHA': None, 'TIPO': None, 'UNIDAD': None, 'ITEM': None, 'CTD': None, 'P_UNIT': None, 'TOTAL': None}
    
    for col in df_raw.columns:
        upper_col = str(col).upper().strip()
        if 'UNIT' in upper_col and '$' in upper_col: mapa['P_UNIT'] = col
        if 'TOTAL' in upper_col and '$' in upper_col: mapa['TOTAL'] = col

    for col in df_raw.columns:
        upper_col = str(col).upper().strip()
        if 'FECHA' in upper_col and not mapa['FECHA']: mapa['FECHA'] = col
        elif 'TIPO' in upper_col and not mapa['TIPO']: mapa['TIPO'] = col
        elif ('UNIDAD' in upper_col or 'PLACA' in upper_col) and not mapa['UNIDAD']: mapa['UNIDAD'] = col
        elif ('ITEM' in upper_col or 'DESCRIP' in upper_col or 'MOTIVO' in upper_col) and not mapa['ITEM']: mapa['ITEM'] = col
        elif ('CTD' in upper_col or 'CANT' in upper_col) and not mapa['CTD']: mapa['CTD'] = col
        elif 'UNIT' in upper_col and not mapa['P_UNIT']: mapa['P_UNIT'] = col
        elif 'TOTAL' in upper_col and not mapa['TOTAL']: mapa['TOTAL'] = col
    
    df_clean = pd.DataFrame()
    
    if mapa['FECHA']:
        serie_fechas = df_raw[mapa['FECHA']]
        fechas_dt = pd.to_datetime(serie_fechas, dayfirst=True, errors='coerce')
        fechas_dt = fechas_dt.fillna(pd.to_datetime(serie_fechas, errors='coerce'))
        df_clean['Fecha_Registro'] = fechas_dt.dt.strftime('%d/%m/%Y').fillna('Sin Fecha')
    else:
        df_clean['Fecha_Registro'] = datetime.now().strftime('%d/%m/%Y')
        
    df_clean['TIPO'] = df_raw[mapa['TIPO']].fillna('NO DEFINIDO') if mapa['TIPO'] else 'NO DEFINIDO'
    df_clean['UNIDAD'] = df_raw[mapa['UNIDAD']].fillna('NO DEFINIDO').apply(normalizar_unidad) if mapa['UNIDAD'] else 'NO DEFINIDO'
    df_clean['ITEM'] = df_raw[mapa['ITEM']].fillna('-') if mapa['ITEM'] else '-'
    df_clean['CTD'] = pd.to_numeric(df_raw[mapa['CTD']], errors='coerce').fillna(1) if mapa['CTD'] else 1
    
    def limpiar_moneda(val):
        if pd.isna(val): return 0.0
        val_str = str(val).replace('$', '').replace('Bs.S', '').replace('Bs.', '').replace(' ', '').strip()
        if ',' in val_str and '.' in val_str:
            if val_str.rfind(',') > val_str.rfind('.'):
                val_str = val_str.replace('.', '').replace(',', '.')
            else:
                val_str = val_str.replace(',', '')
        elif ',' in val_str:
            val_str = val_str.replace(',', '.')
        try: return float(val_str)
        except: return 0.0

    df_clean['P. UNIT $'] = df_raw[mapa['P_UNIT']].apply(limpiar_moneda) if mapa['P_UNIT'] else 0.0
    df_clean['TOTAL $'] = df_raw[mapa['TOTAL']].apply(limpiar_moneda) if mapa['TOTAL'] else (df_clean['CTD'] * df_clean['P. UNIT $'])
    
    df_clean = df_clean[df_clean['TOTAL $'] > 0]
    return df_clean

def procesar_texto_combustible(texto):
    datos = {
        "Fecha": datetime.now().strftime("%d/%m/%Y"),
        "De": "", "Para": "", 
        "Gasolina_Bidones": 0.0, "Gasoil_Bidones": 0.0, 
        "Tanque_1_50K": 0.0, "Tanque_2_12K": 0.0, "Tanque_3_7K": 0.0
    }
    
    m_fecha = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
    if m_fecha: datos["Fecha"] = m_fecha.group(1)
    
    m_de = re.search(r'DE:\s*(.*)', texto, re.IGNORECASE)
    if m_de: datos["De"] = m_de.group(1).strip()
    
    m_para = re.search(r'PARA:\s*(.*)', texto, re.IGNORECASE)
    if m_para: datos["Para"] = m_para.group(1).strip()
    
    m_gas = re.search(r'gasolina[^\n]*\n.*?(\d+[\.,]?\d*)\s*lts', texto, re.IGNORECASE)
    if m_gas: datos["Gasolina_Bidones"] = float(m_gas.group(1).replace('.','').replace(',','.'))
    
    m_gasoil_b = re.search(r'(\d+[\.,]?\d*)\s*lts.*?gasoil.*?bidones', texto, re.IGNORECASE | re.DOTALL)
    if not m_gasoil_b:
        m_gasoil_b = re.search(r'gasoil[^\n]*\n.*?(\d+[\.,]?\d*)\s*lts.*?bidones', texto, re.IGNORECASE)
    if m_gasoil_b: datos["Gasoil_Bidones"] = float(m_gasoil_b.group(1).replace('.','').replace(',','.'))
    
    numeros_tanque = re.findall(r'(\d+[\.,]\d+|\d{3,})\s*lts', texto, re.IGNORECASE)
    if m_gas: numeros_tanque = [n for n in numeros_tanque if n.replace('.','').replace(',','.') != m_gas.group(1).replace('.','').replace(',','.')]
    if m_gasoil_b: numeros_tanque = [n for n in numeros_tanque if n.replace('.','').replace(',','.') != m_gasoil_b.group(1).replace('.','').replace(',','.')]
    
    if len(numeros_tanque) >= 1: datos["Tanque_1_50K"] = float(numeros_tanque[0].replace('.','').replace(',','.'))
    if len(numeros_tanque) >= 2: datos["Tanque_2_12K"] = float(numeros_tanque[1].replace('.','').replace(',','.'))
    if len(numeros_tanque) >= 3: datos["Tanque_3_7K"] = float(numeros_tanque[2].replace('.','').replace(',','.'))
    
    return pd.DataFrame([datos])

# ==========================================
# PIZARRAS HTML GLOBALES
# ==========================================
def html_pizarra_flota(df, tipo="PLANIFICADAS"):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 60px; max-width: 150px; object-fit: contain; display: block;">' if logo else '<span style="font-size:30px;">🚚</span>'
    color_header = "#1a237e" 
    titulo_texto = "Mantenimiento vehicular planificado" if tipo == "PLANIFICADAS" else "Mantenimiento vehicular realizado"
    if tipo == "PLANIFICADAS":
        cols = f"<th style='padding: 12px; border-bottom: 2px solid {color_header}; text-align: left; width: 30%;'>Placa / Unidad</th><th style='padding: 12px; border-bottom: 2px solid {color_header}; text-align: left; width: 45%;'>Tipo de mantenimiento</th><th style='padding: 12px; border-bottom: 2px solid {color_header}; text-align: left; width: 25%;'>Mecánico asignado</th>"
        filas = "".join([f"<tr style='background-color: {'#ffffff' if i % 2 == 0 else '#f8f9fa'};'><td style='padding: 12px; border-bottom: 1px solid #eee; font-weight: bold; color: {color_header};'>{r.get('Unidad','')}</td><td style='padding: 12px; border-bottom: 1px solid #eee; color: #000000;'>{str(r.get('Actividad','')).capitalize()}</td><td style='padding: 12px; border-bottom: 1px solid #eee; font-weight: bold; color: #000000;'>{str(r.get('Mecánico','')).title()}</td></tr>" for i, r in df.iterrows()])
    else:
        cols = f"<th style='padding: 12px; border-bottom: 2px solid {color_header}; text-align: left; width: 25%;'>Placa / Unidad</th><th style='padding: 12px; border-bottom: 2px solid {color_header}; text-align: left; width: 55%;'>Tipo de mantenimiento</th><th style='padding: 12px; border-bottom: 2px solid {color_header}; text-align: center; width: 20%;'>Condición</th>"
        filas = ""
        for i, r in df.iterrows():
            cond = str(r.get('Condición', '')).upper()
            if "OPERATIVO" in cond: color_cond = "#2e7d32"
            elif "REALIZADO" in cond: color_cond = "#e65100"
            else: color_cond = "#b71c1c"
            bg = "#ffffff" if i % 2 == 0 else "#f8f9fa"
            filas += f"<tr style='background-color: {bg};'><td style='padding: 12px; border-bottom: 1px solid #eee; font-weight: bold; color: {color_header};'>{r.get('Unidad','')}</td><td style='padding: 12px; border-bottom: 1px solid #eee; color: #000000;'>{str(r.get('Resumen Actividad','')).capitalize()}</td><td style='padding: 12px; border-bottom: 1px solid #eee; font-weight: bold; color: {color_cond}; text-align: center;'>{cond.title()}</td></tr>"
    fecha_doc = df['Fecha'].iloc[0] if not df.empty else datetime.now().strftime("%d/%m/%Y")
    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;">
        <button onclick="descargarPizarraFlota()" style="background:{color_header};color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Pizarra Flota</button>
    </div>
    <div id="pizarra-flota" style="font-family: Arial, sans-serif; width: 900px; margin: auto; background-color: #fff; border: 2px solid {color_header}; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
        <div style="background-color: {color_header}; color: white; padding: 20px 30px; display: flex; align-items: center; justify-content: space-between; border-bottom: 4px solid #b71c1c;">
            <div style="flex: 0 0 150px; text-align: left;">{area_logo}</div>
            <div style="flex: 1; text-align: right;">
                <h2 style="margin: 0; color: white; font-size: 22px; text-transform: uppercase;">{titulo_texto}</h2>
                <p style="margin: 5px 0 0; color: #e8eaf6; font-size: 14px;">Fecha de reporte: {fecha_doc}</p>
            </div>
        </div>
        <div style="padding: 25px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 14px; background-color: white;">
                <thead><tr style="background-color: #e8eaf6; color: {color_header}; font-size: 14px; font-weight: bold;">{cols}</tr></thead>
                <tbody>{filas}</tbody>
            </table>
            <div style="margin-top: 30px; text-align: center; font-size: 13px; font-weight: bold; color: #000000; border-top: 1px solid #eee; padding-top: 15px;">
                Departamento de Flota y Logística
            </div>
        </div>
    </div>
    <script>
    function descargarPizarraFlota() {{ html2canvas(document.getElementById('pizarra-flota'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Pizarra_Flota_{tipo}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}
    </script>
    """

def html_pizarra_gastos(df, fecha, tipo_categoria="TODO INCLUIDO"):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 50px; max-width: 150px; object-fit: contain; display: block;">' if logo else '<span style="font-size:30px;">💰</span>'
    color_header = "#1a237e" 
    
    filas_html = ""
    unidades = df['UNIDAD'].unique()
    for unidad in unidades:
        df_unidad = df[df['UNIDAD'] == unidad]
        total_unidad = df_unidad['TOTAL $'].sum()
        filas_html += f'<tr style="background-color: #e3f2fd; border-top: 2px solid #90caf9;"><td colspan="3" style="padding: 10px; font-weight: bold; color: {color_header}; font-size: 15px;">🚛 {unidad}</td><td style="padding: 10px; font-weight: bold; color: {color_header}; text-align: right; font-size: 15px;">${total_unidad:,.2f}</td></tr>'
        tipos = df_unidad['TIPO'].unique()
        for tipo in tipos:
            df_tipo = df_unidad[df_unidad['TIPO'] == tipo]
            total_tipo = df_tipo['TOTAL $'].sum()
            filas_html += f'<tr style="background-color: #f5f5f5;"><td style="padding: 8px 10px 8px 30px; font-weight: bold; color: #000000; font-size: 13px;" colspan="3">↳ {str(tipo).title()}</td><td style="padding: 8px 10px; font-weight: bold; color: #000000; text-align: right; font-size: 13px;">${total_tipo:,.2f}</td></tr>'
            for _, r in df_tipo.iterrows():
                item_completo = str(r['ITEM']).strip().capitalize()
                filas_html += f'<tr style="background-color: #ffffff; border-bottom: 1px solid #eee;"><td style="padding: 6px 10px 6px 50px; color: #000000; font-size: 12px; width: 50%; word-wrap: break-word;">{item_completo}</td><td style="padding: 6px 10px; color: #000000; font-size: 12px; text-align: center; width: 10%;">Cant: {r["CTD"]}</td><td style="padding: 6px 10px; color: #000000; font-size: 12px; text-align: right; width: 20%;">${r["P. UNIT $"]:,.2f} c/u</td><td style="padding: 6px 10px; color: #000000; font-weight: bold; font-size: 12px; text-align: right; width: 20%;">${r["TOTAL $"]:,.2f}</td></tr>'
    gran_total = df['TOTAL $'].sum()
    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;"><button onclick="descargarPizarraGastos()" style="background:{color_header};color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Pizarra de Gastos</button></div>
    <div id="pizarra-gastos" style="font-family: Arial, sans-serif; width: 900px; margin: auto; background-color: #fff; border: 2px solid {color_header}; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
        <div style="background-color: {color_header}; color: white; padding: 20px 30px; display: flex; align-items: center; justify-content: space-between; border-bottom: 4px solid #4caf50;">
            <div style="flex: 0 0 150px; text-align: left;">{area_logo}</div>
            <div style="flex: 1; text-align: right;"><h2 style="margin: 0; color: white; font-size: 20px; text-transform: uppercase;">Reporte de gastos: {str(tipo_categoria).title()}</h2><p style="margin: 5px 0 0; color: #e8eaf6; font-size: 14px;">Fecha de registro: {fecha}</p></div>
        </div>
        <div style="padding: 20px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 14px; table-layout: fixed;"><tbody>{filas_html}</tbody><tfoot><tr style="background-color: {color_header}; color: white;"><td colspan="3" style="padding: 15px; font-weight: bold; font-size: 16px; text-align: right;">GRAN TOTAL INVERTIDO:</td><td style="padding: 15px; font-weight: bold; font-size: 18px; text-align: right;">${gran_total:,.2f}</td></tr></tfoot></table>
            <div style="margin-top: 20px; text-align: center; font-size: 12px; color: #000000; border-top: 1px solid #eee; padding-top: 10px;">Departamento de Flota y Logística</div>
        </div>
    </div>
    <script>function descargarPizarraGastos() {{ html2canvas(document.getElementById('pizarra-gastos'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Pizarra_Gastos_{fecha}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>
    """

def html_pizarra_combustible_3t(df, img1, img2, img3):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 50px; max-width: 150px; object-fit: contain; display: block;">' if logo else '<span style="font-size:30px;">⛽</span>'
    r = df.iloc[0]
    t1, t2, t3 = float(r['Tanque_1_50K']), float(r['Tanque_2_12K']), float(r['Tanque_3_7K'])
    total_lts = t1 + t2 + t3
    dias_proy = int(total_lts / 1500) if total_lts > 0 else 0
    p1, p2, p3 = min(round((t1/50000)*100,1),100), min(round((t2/12000)*100,1),100), min(round((t3/7850)*100,1),100)

    def get_color(pct): return "#d32f2f" if pct <= 20 else ("#ff9800" if pct <= 45 else "#4caf50")
    c1, c2, c3 = get_color(p1), get_color(p2), get_color(p3)

    i1_html = f'<img src="data:image/jpeg;base64,{img1}" style="width:100%; height:180px; object-fit:cover; border-radius:5px; border:2px solid #ccc;">' if img1 else '<div style="background:#eee; height:180px; display:flex; align-items:center; justify-content:center; border-radius:5px; color:#888; font-size:12px;">Sin foto</div>'
    i2_html = f'<img src="data:image/jpeg;base64,{img2}" style="width:100%; height:180px; object-fit:cover; border-radius:5px; border:2px solid #ccc;">' if img2 else '<div style="background:#eee; height:180px; display:flex; align-items:center; justify-content:center; border-radius:5px; color:#888; font-size:12px;">Sin foto</div>'
    i3_html = f'<img src="data:image/jpeg;base64,{img3}" style="width:100%; height:180px; object-fit:cover; border-radius:5px; border:2px solid #ccc;">' if img3 else '<div style="background:#eee; height:180px; display:flex; align-items:center; justify-content:center; border-radius:5px; color:#888; font-size:12px;">Sin foto</div>'

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;"><button onclick="descargarPizarraComb()" style="background:#f44336;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Pizarra Combustible</button></div>
    <div id="pizarra-comb" style="font-family:Arial, sans-serif; width:900px; margin:auto; background-color:#fff; border:2px solid #1a237e; border-radius:10px; overflow:hidden; box-shadow:0 4px 8px rgba(0,0,0,0.1);">
        <div style="background-color:#1a237e; color:white; padding:20px 30px; display:flex; align-items:center; justify-content:space-between; border-bottom:4px solid #f44336;">
            <div style="flex:0 0 150px; text-align:left;">{area_logo}</div>
            <div style="flex:1; text-align:right;"><h2 style="margin:0; color:white; font-size:22px; text-transform:uppercase;">Control de reserva de combustible</h2><p style="margin:5px 0 0; color:#e8eaf6; font-size:14px;">Fecha: {r['Fecha']} | De: {str(r['De']).title()}</p></div>
        </div>
        <div style="padding:20px;">
            <div style="display:flex; gap:20px; margin-bottom:30px;">
                <div style="flex:1; background:#e8eaf6; padding:15px; border-radius:8px; border-left:5px solid #2196f3; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                    <p style="margin:0 0 5px 0; color:#555; font-size:13px; font-weight:bold; text-transform:uppercase;">Reserva bidones base</p>
                    <p style="margin:0; font-size:14px;">⛽ Gasolina: <b style="color:#d32f2f;">{r['Gasolina_Bidones']:,.0f} L</b></p>
                    <p style="margin:0; font-size:14px;">⛽ Gasoil: <b style="color:#d32f2f;">{r['Gasoil_Bidones']:,.0f} L</b></p>
                </div>
                <div style="flex:2; background:#e8f5e9; padding:15px; border-radius:8px; border-left:5px solid #4caf50; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                    <p style="margin:0 0 5px 0; color:#2e7d32; font-size:14px; font-weight:bold; text-transform:uppercase;">Total gasoil disponible (Ciudad Drotaca)</p>
                    <h2 style="margin:0; font-size:32px; color:#1a237e;">{total_lts:,.0f} <span style="font-size:16px;">Lts</span></h2>
                    <p style="margin:5px 0 0 0; font-size:15px; color:#e65100; font-weight:bold;">⏳ Tenemos gasoil para {dias_proy} días operativos en condiciones normales.</p>
                    <p style="margin:2px 0 0 0; font-size:12px; color:#555;">*(Cálculo en base a un consumo de 1.500 Lts por día)*</p>
                </div>
            </div>
            <div style="display:flex; gap:20px;">
                <div style="flex:1; background:#fafafa; border:1px solid #ddd; border-radius:8px; padding:15px; text-align:center;">
                    <h3 style="margin:0 0 5px 0; color:#333; font-size:15px;">Tanque Principal (Gasoil)</h3>
                    <p style="margin:0 0 5px 0; font-size:18px; font-weight:bold; color:#1a237e;">{t1:,.0f} L <span style="font-size:12px; color:#888;">de 50.000</span></p>
                    <p style="margin:0 0 15px 0; font-size:16px; font-weight:bold; color:{c1};">{p1}%</p>
                    
                    <div style="width:80px; height:180px; margin:0 auto 15px auto; border:4px solid #212121; border-radius:8px 8px 4px 4px; position:relative; background:#e0e0e0; box-shadow:inset 0 0 10px rgba(0,0,0,0.2);">
                        <div style="position:absolute; bottom:80%; left:0; width:100%; height:1px; background:#444; border-top:1px solid white; z-index:2;"></div><span style="position:absolute; bottom:80%; left:-35px; color:#333; font-size:10px; font-weight:bold; z-index:2;">40 K</span>
                        <div style="position:absolute; bottom:60%; left:0; width:100%; height:1px; background:#444; z-index:2;"></div><span style="position:absolute; bottom:60%; left:-35px; color:#333; font-size:10px; z-index:2;">30 K</span>
                        <div style="position:absolute; bottom:40%; left:0; width:100%; height:1px; background:#444; border-top:1px solid white; z-index:2;"></div><span style="position:absolute; bottom:40%; left:-35px; color:#333; font-size:10px; font-weight:bold; z-index:2;">20 K</span>
                        <div style="position:absolute; bottom:20%; left:0; width:100%; height:1px; background:#444; z-index:2;"></div><span style="position:absolute; bottom:20%; left:-35px; color:#333; font-size:10px; z-index:2;">10 K</span>
                        
                        <div style="position:absolute; bottom:0; left:0; width:100%; height:{p1}%; background-color:{c1}; transition:height 0.5s; z-index:1;"></div>
                    </div>
                    {i1_html}
                </div>
                
                <div style="flex:1; background:#fafafa; border:1px solid #ddd; border-radius:8px; padding:15px; text-align:center;">
                    <h3 style="margin:0 0 5px 0; color:#e65100; font-size:15px;">Tanque Aux. 1</h3>
                    <p style="margin:0 0 5px 0; font-size:18px; font-weight:bold; color:#1a237e;">{t2:,.0f} L <span style="font-size:12px; color:#888;">de 12.000</span></p>
                    <p style="margin:0 0 15px 0; font-size:16px; font-weight:bold; color:{c2};">{p2}%</p>
                    
                    <div style="width:80px; height:180px; margin:0 auto 15px auto; border:4px solid #e65100; border-radius:8px 8px 4px 4px; position:relative; background:#e0e0e0; box-shadow:inset 0 0 10px rgba(0,0,0,0.2);">
                        <div style="position:absolute; bottom:75%; left:0; width:100%; height:1px; background:#444; border-top:1px solid white; z-index:2;"></div><span style="position:absolute; bottom:75%; left:-30px; color:#333; font-size:10px; font-weight:bold; z-index:2;">9 K</span>
                        <div style="position:absolute; bottom:50%; left:0; width:100%; height:1px; background:#444; z-index:2;"></div><span style="position:absolute; bottom:50%; left:-30px; color:#333; font-size:10px; z-index:2;">6 K</span>
                        <div style="position:absolute; bottom:25%; left:0; width:100%; height:1px; background:#444; border-top:1px solid white; z-index:2;"></div><span style="position:absolute; bottom:25%; left:-30px; color:#333; font-size:10px; font-weight:bold; z-index:2;">3 K</span>
                        
                        <div style="position:absolute; bottom:0; left:0; width:100%; height:{p2}%; background-color:{c2}; transition:height 0.5s; z-index:1;"></div>
                    </div>
                    {i2_html}
                </div>
                
                <div style="flex:1; background:#fafafa; border:1px solid #ddd; border-radius:8px; padding:15px; text-align:center;">
                    <h3 style="margin:0 0 5px 0; color:#e65100; font-size:15px;">Tanque Aux. 2</h3>
                    <p style="margin:0 0 5px 0; font-size:18px; font-weight:bold; color:#1a237e;">{t3:,.0f} L <span style="font-size:12px; color:#888;">de 7.850</span></p>
                    <p style="margin:0 0 15px 0; font-size:16px; font-weight:bold; color:{c3};">{p3}%</p>
                    
                    <div style="width:80px; height:180px; margin:0 auto 15px auto; border:4px solid #e65100; border-radius:8px 8px 4px 4px; position:relative; background:#e0e0e0; box-shadow:inset 0 0 10px rgba(0,0,0,0.2);">
                        <div style="position:absolute; bottom:76.4%; left:0; width:100%; height:1px; background:#444; border-top:1px solid white; z-index:2;"></div><span style="position:absolute; bottom:76.4%; left:-30px; color:#333; font-size:10px; font-weight:bold; z-index:2;">6 K</span>
                        <div style="position:absolute; bottom:50.9%; left:0; width:100%; height:1px; background:#444; z-index:2;"></div><span style="position:absolute; bottom:50.9%; left:-30px; color:#333; font-size:10px; z-index:2;">4 K</span>
                        <div style="position:absolute; bottom:25.4%; left:0; width:100%; height:1px; background:#444; border-top:1px solid white; z-index:2;"></div><span style="position:absolute; bottom:25.4%; left:-30px; color:#333; font-size:10px; font-weight:bold; z-index:2;">2 K</span>
                        
                        <div style="position:absolute; bottom:0; left:0; width:100%; height:{p3}%; background-color:{c3}; transition:height 0.5s; z-index:1;"></div>
                    </div>
                    {i3_html}
                </div>
            </div>
        </div>
        <div style="background-color:#eee; text-align:center; padding:10px; font-size:12px; color:#555; border-top:1px solid #ccc;">Departamento de Flota y Logística</div>
    </div>
    <script>function descargarPizarraComb() {{ html2canvas(document.getElementById('pizarra-comb'), {{scale:2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Control_Combustible_{r['Fecha'].replace('/','-')}.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>
    """

# ==========================================
# NUEVA PIZARRA INTEGRADA (FLOTA + GPS) - DINAMICA
# ==========================================
def html_pizarra_estatus_dinamico(tipo_reporte, fecha, activas_resumen, inactivas_df, total, t_activas, t_inactivas, con_gps_resumen, sin_gps_df, t_con_gps, t_sin_gps):
    logo = obtener_logo_base64()
    
    if "Ambas" in tipo_reporte:
        titulo_principal = "Estatus Operativo y Satelital"
        logo_icon = "🚚🛰️"
        width_pizarra = "1000px"
    elif "Flota" in tipo_reporte:
        titulo_principal = "Estatus Operativo (Mecánico)"
        logo_icon = "🚚"
        width_pizarra = "750px"
    else:
        titulo_principal = "Estatus Satelital (GPS)"
        logo_icon = "🛰️"
        width_pizarra = "750px"
        
    area_logo = f'<img src="{logo}" style="max-height: 50px; max-width: 150px; object-fit: contain; display: block;">' if logo else f'<span style="font-size:30px;">{logo_icon}</span>'
    
    # 1. Filas Mecánicas (Operatividad)
    filas_op = ""
    for idx, (modelo, cant) in enumerate(activas_resumen.items()):
        bg = "#ffffff" if idx % 2 == 0 else "#e8f5e9"
        filas_op += f"<tr style='background-color: {bg}; border-bottom: 1px solid #c5e1a5;'><td style='padding: 8px 12px; font-weight: bold; color: #2e7d32; font-size: 13px;'>{str(modelo).title()}</td><td style='padding: 8px 12px; text-align: center; font-weight: bold; color: #333; font-size: 13px;'>{cant}</td></tr>"

    filas_inop = ""
    if not inactivas_df.empty:
        for idx, row in inactivas_df.iterrows():
            bg = "#ffffff" if idx % 2 == 0 else "#ffebee"
            filas_inop += f"<tr style='background-color: {bg}; border-bottom: 1px solid #ef9a9a;'><td style='padding: 8px 12px; font-weight: bold; color: #c62828; font-size: 13px;'>{row['Placa']}</td><td style='padding: 8px 12px; color: #555; font-size: 12px;'>{str(row['Motivo']).title()}</td></tr>"
    else:
        filas_inop = "<tr><td colspan='2' style='padding: 15px; text-align: center; font-weight: bold; color: #2e7d32; font-size: 13px;'>Flota 100% operativa</td></tr>"

    # 2. Filas Satelitales (GPS)
    filas_gps_op = ""
    for idx, (modelo, cant) in enumerate(con_gps_resumen.items()):
        bg = "#ffffff" if idx % 2 == 0 else "#e3f2fd"
        filas_gps_op += f"<tr style='background-color: {bg}; border-bottom: 1px solid #90caf9;'><td style='padding: 8px 12px; font-weight: bold; color: #1565c0; font-size: 13px;'>{str(modelo).title()}</td><td style='padding: 8px 12px; text-align: center; font-weight: bold; color: #333; font-size: 13px;'>{cant}</td></tr>"

    filas_gps_inop = ""
    if not sin_gps_df.empty:
        for idx, row in sin_gps_df.iterrows():
            bg = "#ffffff" if idx % 2 == 0 else "#ffebee"
            filas_gps_inop += f"<tr style='background-color: {bg}; border-bottom: 1px solid #ef9a9a;'><td style='padding: 8px 12px; font-weight: bold; color: #c62828; font-size: 13px;'>{row['Placa']}</td><td style='padding: 8px 12px; color: #555; font-size: 12px;'>{str(row['Motivo']).title()}</td></tr>"
    else:
        filas_gps_inop = "<tr><td colspan='2' style='padding: 15px; text-align: center; font-weight: bold; color: #1565c0; font-size: 13px;'>100% transmitiendo</td></tr>"

    col_mecanica = f"""
    <div style="flex: 1; border: 1px solid #eee; border-radius: 8px; padding: 15px; background: #fafafa;">
        <h3 style="margin: 0 0 15px 0; color: #2e7d32; font-size: 16px; border-bottom: 2px solid #2e7d32; padding-bottom: 5px;">🔧 ESTATUS MECÁNICO</h3>
        
        <h4 style="margin: 0 0 5px 0; font-size: 13px; color: #555;">🟢 Operativas (Agrupadas)</h4>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">
            <thead><tr style="background-color: #e8f5e9; color: #2e7d32;"><th style="padding: 6px 12px; text-align: left; font-size: 12px;">Modelo</th><th style="padding: 6px 12px; text-align: center; font-size: 12px;">Cantidad</th></tr></thead>
            <tbody>{filas_op}</tbody>
        </table>
        
        <h4 style="margin: 0 0 5px 0; font-size: 13px; color: #555;">🔴 En Taller / Inactivas</h4>
        <table style="width: 100%; border-collapse: collapse;">
            <thead><tr style="background-color: #ffebee; color: #c62828;"><th style="padding: 6px 12px; text-align: left; font-size: 12px;">Placa</th><th style="padding: 6px 12px; text-align: left; font-size: 12px;">Motivo</th></tr></thead>
            <tbody>{filas_inop}</tbody>
        </table>
    </div>
    """
    
    col_satelital = f"""
    <div style="flex: 1; border: 1px solid #eee; border-radius: 8px; padding: 15px; background: #fafafa;">
        <h3 style="margin: 0 0 15px 0; color: #1565c0; font-size: 16px; border-bottom: 2px solid #1565c0; padding-bottom: 5px;">📡 ESTATUS SATELITAL</h3>
        
        <h4 style="margin: 0 0 5px 0; font-size: 13px; color: #555;">🔵 Transmitiendo (Agrupadas)</h4>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">
            <thead><tr style="background-color: #e3f2fd; color: #1565c0;"><th style="padding: 6px 12px; text-align: left; font-size: 12px;">Modelo</th><th style="padding: 6px 12px; text-align: center; font-size: 12px;">Cantidad</th></tr></thead>
            <tbody>{filas_gps_op}</tbody>
        </table>
        
        <h4 style="margin: 0 0 5px 0; font-size: 13px; color: #555;">🔴 Sin Transmisión</h4>
        <table style="width: 100%; border-collapse: collapse;">
            <thead><tr style="background-color: #ffebee; color: #c62828;"><th style="padding: 6px 12px; text-align: left; font-size: 12px;">Placa</th><th style="padding: 6px 12px; text-align: left; font-size: 12px;">Motivo / Falla</th></tr></thead>
            <tbody>{filas_gps_inop}</tbody>
        </table>
    </div>
    """

    if "Ambas" in tipo_reporte:
        cuerpo = f'<div style="display: flex; padding: 20px; gap: 20px; align-items: flex-start;">{col_mecanica}{col_satelital}</div>'
        kpis_html = f"""
        <div style="flex: 1; padding: 15px; border-right: 1px solid #ccc; display: flex; flex-direction: column; justify-content: center;">
            <span style="font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">FLOTA TOTAL</span>
            <span style="font-size: 38px; color: #1a237e; font-weight: 900;">{total}</span>
        </div>
        <div style="flex: 2; padding: 15px; border-right: 1px solid #ccc; display: flex;">
            <div style="flex: 1; border-right: 1px dashed #ddd;">
                <span style="font-size: 12px; color: #2e7d32; font-weight: bold; text-transform: uppercase;">OPERATIVAS</span><br>
                <span style="font-size: 34px; color: #2e7d32; font-weight: 900;">{t_activas}</span>
            </div>
            <div style="flex: 1;">
                <span style="font-size: 12px; color: #d32f2f; font-weight: bold; text-transform: uppercase;">EN TALLER</span><br>
                <span style="font-size: 34px; color: #d32f2f; font-weight: 900;">{t_inactivas}</span>
            </div>
        </div>
        <div style="flex: 2; padding: 15px; display: flex;">
            <div style="flex: 1; border-right: 1px dashed #ddd;">
                <span style="font-size: 12px; color: #1565c0; font-weight: bold; text-transform: uppercase;">CON SEÑAL GPS</span><br>
                <span style="font-size: 34px; color: #1565c0; font-weight: 900;">{t_con_gps}</span>
            </div>
            <div style="flex: 1;">
                <span style="font-size: 12px; color: #d32f2f; font-weight: bold; text-transform: uppercase;">SIN TRANSMISIÓN</span><br>
                <span style="font-size: 34px; color: #d32f2f; font-weight: 900;">{t_sin_gps}</span>
            </div>
        </div>
        """
    elif "Flota" in tipo_reporte:
        cuerpo = f'<div style="display: flex; padding: 20px; gap: 20px; align-items: flex-start;">{col_mecanica}</div>'
        kpis_html = f"""
        <div style="flex: 1; padding: 15px; border-right: 1px solid #ccc; display: flex; flex-direction: column; justify-content: center;">
            <span style="font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">FLOTA TOTAL</span>
            <span style="font-size: 38px; color: #1a237e; font-weight: 900;">{total}</span>
        </div>
        <div style="flex: 2; padding: 15px; display: flex;">
            <div style="flex: 1; border-right: 1px dashed #ddd;">
                <span style="font-size: 12px; color: #2e7d32; font-weight: bold; text-transform: uppercase;">OPERATIVAS</span><br>
                <span style="font-size: 34px; color: #2e7d32; font-weight: 900;">{t_activas}</span>
            </div>
            <div style="flex: 1;">
                <span style="font-size: 12px; color: #d32f2f; font-weight: bold; text-transform: uppercase;">EN TALLER</span><br>
                <span style="font-size: 34px; color: #d32f2f; font-weight: 900;">{t_inactivas}</span>
            </div>
        </div>
        """
    else:
        cuerpo = f'<div style="display: flex; padding: 20px; gap: 20px; align-items: flex-start;">{col_satelital}</div>'
        kpis_html = f"""
        <div style="flex: 1; padding: 15px; border-right: 1px solid #ccc; display: flex; flex-direction: column; justify-content: center;">
            <span style="font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">FLOTA TOTAL</span>
            <span style="font-size: 38px; color: #1a237e; font-weight: 900;">{total}</span>
        </div>
        <div style="flex: 2; padding: 15px; display: flex;">
            <div style="flex: 1; border-right: 1px dashed #ddd;">
                <span style="font-size: 12px; color: #1565c0; font-weight: bold; text-transform: uppercase;">CON SEÑAL GPS</span><br>
                <span style="font-size: 34px; color: #1565c0; font-weight: 900;">{t_con_gps}</span>
            </div>
            <div style="flex: 1;">
                <span style="font-size: 12px; color: #d32f2f; font-weight: bold; text-transform: uppercase;">SIN TRANSMISIÓN</span><br>
                <span style="font-size: 34px; color: #d32f2f; font-weight: 900;">{t_sin_gps}</span>
            </div>
        </div>
        """

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;">
        <button onclick="descargarPizarraIntegrada()" style="background:#1a237e;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Pizarra Actual</button>
    </div>
    
    <div id="pizarra-integrada" style="font-family: Arial, sans-serif; width: {width_pizarra}; margin: auto; background-color: #fff; border: 2px solid #1a237e; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
        
        <div style="background-color: #1a237e; color: white; padding: 20px 30px; display: flex; align-items: center; justify-content: space-between; border-bottom: 4px solid #fbc02d;">
            <div style="flex: 0 0 150px; text-align: left;">{area_logo}</div>
            <div style="flex: 1; text-align: right;">
                <h2 style="margin: 0; color: white; font-size: 24px; text-transform: uppercase; letter-spacing: 1px;">{titulo_principal}</h2>
                <p style="margin: 5px 0 0; color: #e8eaf6; font-size: 15px;">Fecha de evaluación: {fecha}</p>
            </div>
        </div>
        
        <div style="display: flex; background-color: #f5f5f5; border-bottom: 2px solid #ccc; text-align: center;">
            {kpis_html}
        </div>
        
        {cuerpo}
        
        <div style="background-color: #eee; text-align: center; padding: 10px; font-size: 12px; color: #555; border-top: 1px solid #ccc;">
            Departamento de Flota y Logística - Operatividad y Monitoreo
        </div>
    </div>
    <script>
    function descargarPizarraIntegrada() {{ html2canvas(document.getElementById('pizarra-integrada'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Estatus_Dinámico.png'; link.href = canvas.toDataURL(); link.click(); }}); }}
    </script>
    """

# ==========================================
# INTERFAZ PRINCIPAL Y TABS (ACTUALIZADAS)
# ==========================================
st.title("🚛 PCD - Ecosistema de Flota y Logística")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📝 Planificadas", 
    "🔧 Realizadas", 
    "📊 Histórico", 
    "⛽ Combustible",
    "📈 Rep. Combustible",
    "🚚🛰️ Estatus y GPS",
    "📊 Rep. Estatus Integrado"
])

# ---------------------------------------------------------
# TAB 1: PLANIFICADAS
# ---------------------------------------------------------
with tab1:
    col1_p, col2_p = st.columns([1, 1.5])
    with col1_p:
        st.markdown("### 📥 1. Pegar Datos")
        st.caption("Pega el listado (con o sin '|'). El sistema detectará la fecha, placa, actividad y mecánico automáticamente.")
        txt_plan = st.text_area("Texto Planificadas:", height=200, key="txt_p")
        
        btn_plan = st.button("⚡ Procesar Planificadas", type="primary", use_container_width=True)
        
    if btn_plan and txt_plan: st.session_state['df_plan_edit'] = procesar_texto_planificadas(txt_plan)
    if 'df_plan_edit' in st.session_state:
        with col2_p:
            st.subheader("✏️ Revisión de Planificadas")
            df_edit_p = st.data_editor(st.session_state['df_plan_edit'], num_rows="dynamic", use_container_width=True)
            if st.button("💾 Guardar Planificadas en Sheets"):
                with st.spinner("Guardando..."):
                    df_to_save = df_edit_p.copy()
                    dias_semana, meses_ano = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    dias, semanas, meses = [], [], []
                    for f_str in df_to_save['Fecha']:
                        match_f = re.search(r'\d{2}/\d{2}/\d{4}', str(f_str))
                        if match_f:
                            try:
                                obj_fecha = datetime.strptime(match_f.group(), "%d/%m/%Y")
                                fecha_shift = obj_fecha + timedelta(days=2) 
                                dias.append(dias_semana[obj_fecha.weekday()])
                                semanas.append(f"Semana {fecha_shift.isocalendar()[1]}")
                                meses.append(meses_ano[obj_fecha.month - 1])
                            except:
                                dias.append(""); semanas.append(""); meses.append("")
                        else:
                            dias.append(""); semanas.append(""); meses.append("")
                    df_to_save['Dia'], df_to_save['Semana'], df_to_save['Mes'] = dias, semanas, meses
                    guardar_en_google_sheets(df_to_save, "FLOTA_PLANIFICADO")
                    st.success("✅ Guardado en FLOTA_PLANIFICADO.")
            st.markdown("---")
            st.subheader("📱 Mensaje WhatsApp")
            st.code(generar_ws_planificadas(df_edit_p), language="markdown")
            st.markdown("---")
            st.subheader("📸 Generar Imagen")
            components.html(html_pizarra_flota(df_edit_p, "PLANIFICADAS"), height=550, scrolling=True)

# ---------------------------------------------------------
# TAB 2: REALIZADAS
# ---------------------------------------------------------
with tab2:
    col1_r, col2_r = st.columns([1, 1.5])
    with col1_r:
        st.markdown("### 📥 1. Pegar Datos")
        st.caption("Pega el listado (con o sin '|'). El sistema detectará la fecha, placa, resumen y el estatus al final.")
        txt_real = st.text_area("Texto Realizadas:", height=200, key="txt_r")
        
        btn_real = st.button("⚡ Procesar Realizadas", type="primary", use_container_width=True)
        
    if btn_real and txt_real: st.session_state['df_real_edit'] = procesar_texto_realizadas(txt_real)
    if 'df_real_edit' in st.session_state:
        with col2_r:
            st.subheader("✏️ Revisión de Realizadas")
            df_edit_r = st.data_editor(st.session_state['df_real_edit'], num_rows="dynamic", use_container_width=True)
            if st.button("💾 Guardar Realizadas en Sheets"):
                with st.spinner("Guardando..."):
                    df_to_save = df_edit_r.copy()
                    dias_semana, meses_ano = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    dias, semanas, meses = [], [], []
                    for f_str in df_to_save['Fecha']:
                        match_f = re.search(r'\d{2}/\d{2}/\d{4}', str(f_str))
                        if match_f:
                            try:
                                obj_fecha = datetime.strptime(match_f.group(), "%d/%m/%Y")
                                fecha_shift = obj_fecha + timedelta(days=2) 
                                dias.append(dias_semana[obj_fecha.weekday()])
                                semanas.append(f"Semana {fecha_shift.isocalendar()[1]}")
                                meses.append(meses_ano[obj_fecha.month - 1])
                            except:
                                dias.append(""); semanas.append(""); meses.append("")
                        else:
                            dias.append(""); semanas.append(""); meses.append("")
                    df_to_save['Dia'], df_to_save['Semana'], df_to_save['Mes'] = dias, semanas, meses
                    guardar_en_google_sheets(df_to_save, "FLOTA_REALIZADO")
                    st.success("✅ Guardado en FLOTA_REALIZADO.")
            st.markdown("---")
            st.subheader("📱 Mensaje WhatsApp")
            
            mask_ocultar = df_edit_r['Condición'].astype(str).str.upper().str.contains('REALIZADO') & ~df_edit_r['Condición'].astype(str).str.upper().str.contains('OPERATIVO')
            df_visual_r = df_edit_r[~mask_ocultar].copy()
            
            st.code(generar_ws_realizadas(df_visual_r), language="markdown")
            st.markdown("---")
            st.subheader("📸 Generar Imagen")
            components.html(html_pizarra_flota(df_visual_r, "REALIZADAS"), height=650, scrolling=True)

# ---------------------------------------------------------
# TAB 3: AUDITORÍA MANTENIMIENTOS
# ---------------------------------------------------------
with tab3:
    if st.button("🔄 Descargar Data Histórica de Mantenimientos", type="primary"):
        with st.spinner("Conectando con Google Sheets..."):
            st.session_state['db_f_plan'] = extraer_datos_sheets("FLOTA_PLANIFICADO")
            st.session_state['db_f_real'] = extraer_datos_sheets("FLOTA_REALIZADO")
            st.success("¡Base de Datos Descargada!")
    if 'db_f_plan' in st.session_state and not st.session_state['db_f_plan'].empty:
        df_p = st.session_state['db_f_plan']
        df_r = st.session_state.get('db_f_real', pd.DataFrame())
        semanas_disp = sorted(list(set([str(x) for x in df_p.get('Semana', pd.Series()).dropna().unique() if str(x).strip()])))
        meses_disp = sorted(list(set([str(x) for x in df_p.get('Mes', pd.Series()).dropna().unique() if str(x).strip()])))
        c1, c2 = st.columns(2)
        filtro_mes = c1.selectbox("Filtrar por Mes:", ["Todos"] + meses_disp, key="fm_mtto")
        filtro_semana = c2.selectbox("Filtrar por Semana:", ["Todas"] + semanas_disp if semanas_disp else ["Sin datos"], key="fs_mtto")
        df_p_filt = df_p.copy()
        df_r_filt = df_r.copy() if not df_r.empty else pd.DataFrame()
        if filtro_mes != "Todos":
            if 'Mes' in df_p_filt.columns: df_p_filt = df_p_filt[df_p_filt['Mes'] == filtro_mes]
            if not df_r_filt.empty and 'Mes' in df_r_filt.columns: df_r_filt = df_r_filt[df_r_filt['Mes'] == filtro_mes]
        if filtro_semana != "Todas":
            if 'Semana' in df_p_filt.columns: df_p_filt = df_p_filt[df_p_filt['Semana'] == filtro_semana]
            if not df_r_filt.empty and 'Semana' in df_r_filt.columns: df_r_filt = df_r_filt[df_r_filt['Semana'] == filtro_semana]
        col1_k, col2_k, col3_k = st.columns(3)
        total_plan = len(df_p_filt)
        col1_k.metric("Mantenimientos planificados", total_plan)
        total_real = len(df_r_filt[df_r_filt['Condición'].isin(['OPERATIVO', 'REALIZADO'])]) if not df_r_filt.empty and 'Condición' in df_r_filt.columns else 0
        col2_k.metric("Mantenimientos completados", total_real)
        efectividad = round((total_real / total_plan) * 100, 1) if total_plan > 0 else 0
        col3_k.metric("Efectividad de taller", f"{efectividad}%")

# ---------------------------------------------------------
# TAB 4: ⛽ CONTROL DE COMBUSTIBLE (3 TANQUES)
# ---------------------------------------------------------
with tab4:
    st.header("⛽ Módulo de Reserva de Combustible")
    
    col_comb1, col_comb2 = st.columns([1, 1.5])
    
    with col_comb1:
        st.info("Pega el reporte de cambio de guardia y sube los avales fotográficos de los tanques.")
        txt_comb = st.text_area("Pega WhatsApp de Combustible:", height=200)
        
        st.write("📸 Sube las fotos de los tanques:")
        ci1, ci2, ci3 = st.columns(3)
        with ci1: img_comb1 = st.file_uploader("Tanque 1 (50K)", type=["png", "jpg", "jpeg"])
        with ci2: img_comb2 = st.file_uploader("Tanque 2 (12K)", type=["png", "jpg", "jpeg"])
        with ci3: img_comb3 = st.file_uploader("Tanque 3 (7.85K)", type=["png", "jpg", "jpeg"])
        
        btn_comb = st.button("⚡ Extraer y Calcular", type="primary", use_container_width=True)
        
    if btn_comb and txt_comb:
        st.session_state['df_combustible'] = procesar_texto_combustible(txt_comb)
        st.session_state['img_comb_1_b64'] = base64.b64encode(img_comb1.getvalue()).decode() if img_comb1 else None
        st.session_state['img_comb_2_b64'] = base64.b64encode(img_comb2.getvalue()).decode() if img_comb2 else None
        st.session_state['img_comb_3_b64'] = base64.b64encode(img_comb3.getvalue()).decode() if img_comb3 else None
            
    if 'df_combustible' in st.session_state:
        with col_comb2:
            st.subheader("✏️ Verificación de Datos")
            st.caption("Ajusta los litros de los tanques según lo que marque la foto.")
            df_edit_c = st.data_editor(st.session_state['df_combustible'], num_rows="dynamic", use_container_width=True, hide_index=True)
            
            if st.button("💾 Guardar en Sheets", use_container_width=True):
                with st.spinner("Guardando en la bóveda..."):
                    df_to_save = df_edit_c.copy()
                    
                    dias_semana, meses_ano = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    dias, semanas, meses, d_proy = [], [], [], []
                    
                    for i, r in df_to_save.iterrows():
                        total_t = float(r['Tanque_1_50K']) + float(r['Tanque_2_12K']) + float(r['Tanque_3_7K'])
                        d_proy.append(int(total_t / 1500) if total_t > 0 else 0)
                        try:
                            obj_fecha = datetime.strptime(str(r['Fecha']), "%d/%m/%Y")
                            fecha_shift = obj_fecha + timedelta(days=2)
                            dias.append(dias_semana[obj_fecha.weekday()])
                            semanas.append(f"Semana {fecha_shift.isocalendar()[1]}")
                            meses.append(meses_ano[obj_fecha.month - 1])
                        except:
                            dias.append(""); semanas.append(""); meses.append("")
                            
                    df_to_save['Total_Tanques'] = [float(r['Tanque_1_50K']) + float(r['Tanque_2_12K']) + float(r['Tanque_3_7K']) for _, r in df_to_save.iterrows()]
                    df_to_save['Dias_Proyectados'] = d_proy
                    df_to_save['Dia'] = dias
                    df_to_save['Semana'] = semanas
                    df_to_save['Mes'] = meses
                    
                    guardar_en_google_sheets(df_to_save, "FLOTA_COMBUSTIBLE")
                    st.success("✅ Reserva guardada con éxito en FLOTA_COMBUSTIBLE.")
            
            st.markdown("---")
            st.subheader("📱 Mensaje de Estatus")
            st.code(generar_ws_combustible(df_edit_c), language="markdown")
            
            st.markdown("---")
            st.subheader("📸 Pizarra de Monitoreo")
            components.html(html_pizarra_combustible_3t(df_edit_c, st.session_state.get('img_comb_1_b64'), st.session_state.get('img_comb_2_b64'), st.session_state.get('img_comb_3_b64')), height=850, scrolling=True)

# ---------------------------------------------------------
# TAB 5: 📈 REPORTE DE COMBUSTIBLE (PDF)
# ---------------------------------------------------------
with tab5:
    st.header("📈 Informe Gerencial de Combustible (PDF)")
    
    if st.button("🔄 Sincronizar Base de Datos de Combustible", type="primary", key="btn_sync_comb"):
        with st.spinner("Conectando..."):
            st.session_state['db_combustible'] = extraer_datos_sheets("FLOTA_COMBUSTIBLE")
            st.success("¡Base de datos cargada!")
            
    if 'db_combustible' in st.session_state and not st.session_state['db_combustible'].empty:
        db_c = st.session_state['db_combustible']
        db_c['Fecha_DT'] = pd.to_datetime(db_c['Fecha'], format='%d/%m/%Y', errors='coerce')
        db_c = db_c.sort_values('Fecha_DT')
        
        if 'Total_Tanques' not in db_c.columns:
            if 'Gasoil_Tanque' in db_c.columns:
                db_c['Total_Tanques'] = db_c['Gasoil_Tanque']
            else:
                db_c['Total_Tanques'] = 0
                
        def limpiar_num(x):
            try: return float(str(x).replace(',', ''))
            except: return 0.0
            
        db_c['Total_Tanques'] = db_c['Total_Tanques'].apply(limpiar_num)
        
        st.markdown("---")
        rango_c = st.date_input("📅 Rango de Fechas a Evaluar:", [], key="rango_comb")
        
        if len(rango_c) == 2:
            f_ini, f_fin = rango_c
            db_c_filt = db_c[(db_c['Fecha_DT'].dt.date >= f_ini) & (db_c['Fecha_DT'].dt.date <= f_fin)].copy()
            
            if db_c_filt.empty:
                st.warning("⚠️ No hay registros en ese rango de fechas.")
            else:
                consumo_total = 0
                recargas = 0
                lts_recargados = 0
                
                valores_tanque = db_c_filt['Total_Tanques'].tolist()
                for i in range(1, len(valores_tanque)):
                    diferencia = valores_tanque[i] - valores_tanque[i-1]
                    if diferencia < 0:
                        consumo_total += abs(diferencia)
                    elif diferencia > 500:
                        recargas += 1
                        lts_recargados += diferencia
                
                col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
                col_kpi1.metric("Gasoil consumido (Lts)", f"{consumo_total:,.0f}")
                col_kpi2.metric("Eventos de recarga", recargas)
                col_kpi3.metric("Lts inyectados a tanques", f"{lts_recargados:,.0f}")
                
                logo = obtener_logo_base64()
                area_logo = f'<div style="background-color: #1a237e; padding: 12px 20px; border-radius: 8px; display: flex; align-items: center; justify-content: center;"><img src="{logo}" style="max-height: 55px; max-width: 180px; object-fit: contain;"></div>' if logo else '<div style="background-color: #1a237e; padding: 12px 20px; border-radius: 8px; display: flex; align-items: center; justify-content: center;"><span style="font-size:24px; color: white; font-weight: bold;">⛽ Reporte</span></div>'
                
                filas_c = ""
                for _, r in db_c_filt.iterrows():
                    pct_val = (r['Total_Tanques']/69850)*100
                    color_pct = "#2e7d32" if pct_val > 50 else ("#f57c00" if pct_val > 30 else "#d32f2f")
                    filas_c += f"""
                    <tr style="border-bottom: 1px solid #ddd; background: #fff;">
                        <td style="padding: 10px; text-align: center;">{r['Fecha']}</td>
                        <td style="padding: 10px; font-weight: bold; color: #1a237e; text-align: right;">{float(r.get('Gasolina_Bidones', 0)):,.0f} L</td>
                        <td style="padding: 10px; font-weight: bold; color: #1a237e; text-align: right;">{float(r.get('Gasoil_Bidones', 0)):,.0f} L</td>
                        <td style="padding: 10px; font-weight: bold; color: #1a237e; text-align: right;">{r['Total_Tanques']:,.0f} L</td>
                        <td style="padding: 10px; font-weight: bold; color: {color_pct}; text-align: center;">{pct_val:.1f}%</td>
                    </tr>
                    """

                pdf_combustible = f"""
                <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                
                <div style="text-align: right; margin-bottom: 25px;">
                    <button onclick="descargarPDFComb()" style="background-color: #d32f2f; color: white; border: none; padding: 15px 30px; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold;">⬇️ DESCARGAR REPORTE LOGÍSTICO (PDF)</button>
                </div>
                
                <div id="pdf-comb" style="width: 800px; margin: auto; background: white; padding: 40px; box-sizing: border-box; box-shadow: 0 0 10px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
                    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #1a237e; padding-bottom: 15px; margin-bottom: 25px;">
                        <div>{area_logo}</div>
                        <div style="text-align: right;">
                            <h1 style="margin: 0; color: #1a237e; font-size: 22px; text-transform: uppercase;">Reporte de combustible</h1>
                            <p style="margin: 5px 0 0; color: #555; font-size: 14px;">Periodo: {f_ini.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}</p>
                        </div>
                    </div>
                    
                    <div style="display: flex; gap: 15px; margin-bottom: 30px;">
                        <div style="flex: 1; background-color: #f8f9fa; border-top: 4px solid #1a237e; padding: 15px; text-align: center; border-radius: 5px;">
                            <p style="margin: 0; font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">Consumo estimado</p>
                            <h2 style="margin: 5px 0 0; color: #1a237e; font-size: 26px;">{consumo_total:,.0f} <span style="font-size: 14px;">Lts</span></h2>
                        </div>
                        <div style="flex: 1; background-color: #f8f9fa; border-top: 4px solid #2e7d32; padding: 15px; text-align: center; border-radius: 5px;">
                            <p style="margin: 0; font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">Eventos de recarga</p>
                            <h2 style="margin: 5px 0 0; color: #2e7d32; font-size: 26px;">{recargas}</h2>
                        </div>
                        <div style="flex: 1; background-color: #f8f9fa; border-top: 4px solid #e65100; padding: 15px; text-align: center; border-radius: 5px;">
                            <p style="margin: 0; font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">Volumen inyectado</p>
                            <h2 style="margin: 5px 0 0; color: #e65100; font-size: 26px;">{lts_recargados:,.0f} <span style="font-size: 14px;">Lts</span></h2>
                        </div>
                    </div>
                    
                    <div style="margin-top: 20px; text-align: center; font-size: 11px; color: #aaa; border-top: 1px solid #eee; padding-top: 10px;">
                        Departamento de Flota y Logística - Documento Generado por PCD
                    </div>
                </div>
                
                <script>
                function descargarPDFComb() {{
                    const {{ jsPDF }} = window.jspdf;
                    html2canvas(document.getElementById('pdf-comb'), {{ scale: 2 }}).then(canvas => {{
                        const doc = new jsPDF('p', 'pt', 'a4');
                        const pdfWidth = doc.internal.pageSize.getWidth();
                        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
                        doc.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, pdfWidth, pdfHeight);
                        doc.save('Reporte_Combustible.pdf');
                    }});
                }}
                </script>
                """
                components.html(pdf_combustible, height=900, scrolling=True)

# ---------------------------------------------------------
# TAB 6: 🚚🛰️ ESTATUS Y GPS UNIFICADO (NUEVO DINAMICO)
# ---------------------------------------------------------
with tab6:
    st.header("🚚🛰️ Control Operativo y Satelital")
    
    tipo_reporte = st.radio("🎯 Selecciona el tipo de reporte a generar:", 
                            ["Ambas", "Solo Flota", "Solo GPS"], 
                            horizontal=True)
    
    st.info("El sistema asume 100% de operatividad por defecto. Selecciona únicamente las unidades que presenten fallas mecánicas o de transmisión.")
    
    col_e1, col_e2 = st.columns([1, 1])
    
    inactivas_seleccionadas = []
    sin_gps_seleccionadas = []
    df_inactivas_edit = pd.DataFrame()
    df_sin_gps_edit = pd.DataFrame()
    
    with col_e1:
        fecha_estatus = st.date_input("📅 Fecha de evaluación:", datetime.now(), key="fecha_unificada")
        fecha_str_estatus = fecha_estatus.strftime("%d/%m/%Y")
        opciones_vehiculos = [f"{m} ({p})" for ma, m, p in FLOTA_MASTER_DATA]
        
        if tipo_reporte in ["Ambas", "Solo Flota"]:
            st.markdown("### 🔧 1. Estatus Mecánico (Taller)")
            inactivas_seleccionadas = st.multiselect(
                "🔴 Unidades INOPERATIVAS:", 
                opciones_vehiculos,
                placeholder="Buscar por placa o modelo...",
                key="ms_inop"
            )
            
            if inactivas_seleccionadas:
                datos_inactivas = []
                for sel in inactivas_seleccionadas:
                    mod = sel.split(" (")[0]
                    pla = sel.split(" (")[1].replace(")", "")
                    datos_inactivas.append({"Modelo": mod, "Placa": pla, "Motivo": "Mantenimiento preventivo"})
                df_inactivas_edit = st.data_editor(pd.DataFrame(datos_inactivas), use_container_width=True, hide_index=True, key="ed_inop")
                
        if tipo_reporte in ["Ambas", "Solo GPS"]:
            st.markdown("### 📡 2. Estatus Satelital (GPS)")
            sin_gps_seleccionadas = st.multiselect(
                "🔴 Unidades SIN TRANSMISIÓN:", 
                opciones_vehiculos,
                placeholder="Buscar por placa o modelo...",
                key="ms_gps"
            )
            
            if sin_gps_seleccionadas:
                datos_sin_gps = []
                for sel in sin_gps_seleccionadas:
                    mod = sel.split(" (")[0]
                    pla = sel.split(" (")[1].replace(")", "")
                    datos_sin_gps.append({"Modelo": mod, "Placa": pla, "Motivo": "Revisión técnica"})
                df_sin_gps_edit = st.data_editor(pd.DataFrame(datos_sin_gps), use_container_width=True, hide_index=True, key="ed_gps")
            
    with col_e2:
        total_flota = len(FLOTA_MASTER_DATA)
        total_inactivas = len(inactivas_seleccionadas)
        total_activas = total_flota - total_inactivas
        total_sin_gps = len(sin_gps_seleccionadas)
        total_con_gps = total_flota - total_sin_gps
        
        st.markdown("### 📊 3. Resumen y Acciones")
        
        k1, k2, k3 = st.columns(3)
        k1.metric("Total Flota", total_flota)
        if tipo_reporte == "Solo Flota":
            k2.metric("Mecánicamente Activas", total_activas)
            k3.metric("En Taller", total_inactivas)
        elif tipo_reporte == "Solo GPS":
            k2.metric("GPS Transmitiendo", total_con_gps)
            k3.metric("Sin Transmisión", total_sin_gps)
        else:
            k2.metric("Mecánicamente Activas", total_activas)
            k3.metric("GPS Transmitiendo", total_con_gps)
        
        # Procesamiento de diccionarios para gráficas
        placas_inactivas_list = [p.split(" (")[1].replace(")", "") for p in inactivas_seleccionadas]
        activas_dict = {}
        for ma, mo, pla in FLOTA_MASTER_DATA:
            if pla not in placas_inactivas_list:
                activas_dict[mo] = activas_dict.get(mo, 0) + 1
        activas_ordenadas = dict(sorted(activas_dict.items(), key=lambda item: item[1], reverse=True))

        placas_sin_gps_list = [p.split(" (")[1].replace(")", "") for p in sin_gps_seleccionadas]
        con_gps_dict = {}
        for ma, mo, pla in FLOTA_MASTER_DATA:
            if pla not in placas_sin_gps_list:
                con_gps_dict[mo] = con_gps_dict.get(mo, 0) + 1
        con_gps_ordenadas = dict(sorted(con_gps_dict.items(), key=lambda item: item[1], reverse=True))

        st.markdown("---")
        
        c_btn_u1, c_btn_u2 = st.columns(2)
        with c_btn_u1:
            if st.button("📱 Generar WhatsApp", use_container_width=True):
                st.code(generar_ws_estatus_dinamico(tipo_reporte, fecha_str_estatus, total_flota, total_activas, total_inactivas, total_con_gps, total_sin_gps, df_inactivas_edit, df_sin_gps_edit), language="markdown")
                
        with c_btn_u2:
            if st.button("💾 Guardar Estatus Integrado", type="primary", use_container_width=True):
                with st.spinner("Guardando en base de datos..."):
                    dias_semana, meses_ano = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"], ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    dia_v = dias_semana[fecha_estatus.weekday()]
                    fecha_shift = fecha_estatus + timedelta(days=2)
                    sem_v = f"Semana {fecha_shift.isocalendar()[1]}"
                    mes_v = meses_ano[fecha_estatus.month - 1]
                    
                    det_inop = " | ".join([f"{r['Placa']} ({r['Motivo']})" for _, r in df_inactivas_edit.iterrows()]) if not df_inactivas_edit.empty else "100% Operativa"
                    det_gps = " | ".join([f"{r['Placa']} ({r['Motivo']})" for _, r in df_sin_gps_edit.iterrows()]) if not df_sin_gps_edit.empty else "100% Transmitiendo"
                    
                    df_save_int = pd.DataFrame([{
                        "Fecha": fecha_str_estatus,
                        "Total_Flota": total_flota,
                        "Activas": total_activas,
                        "Inactivas": total_inactivas,
                        "Detalle_Inactivas": det_inop,
                        "Con_GPS": total_con_gps,
                        "Sin_GPS": total_sin_gps,
                        "Detalle_Sin_GPS": det_gps,
                        "Dia": dia_v,
                        "Semana": sem_v,
                        "Mes": mes_v
                    }])
                    
                    guardar_en_google_sheets(df_save_int, "FLOTA_ESTATUS_INTEGRADO")
                    st.success("✅ Estatus general guardado en 'FLOTA_ESTATUS_INTEGRADO'.")

    st.markdown("---")
    st.subheader("📸 Pizarra Visual")
    components.html(html_pizarra_estatus_dinamico(tipo_reporte, fecha_str_estatus, activas_ordenadas, df_inactivas_edit, total_flota, total_activas, total_inactivas, con_gps_ordenadas, df_sin_gps_edit, total_con_gps, total_sin_gps), height=900, scrolling=True)


# ---------------------------------------------------------
# TAB 7: 📊 REPORTE DE ESTATUS INTEGRADO (PDF)
# ---------------------------------------------------------
with tab7:
    st.header("📊 Informe Gerencial Integrado (Mecánica + GPS)")
    
    if st.button("🔄 Sincronizar Base de Datos Integrada", type="primary", key="btn_sync_int"):
        with st.spinner("Conectando..."):
            st.session_state['db_integrada'] = extraer_datos_sheets("FLOTA_ESTATUS_INTEGRADO")
            st.success("¡Base de datos cargada!")
            
    if 'db_integrada' in st.session_state and not st.session_state['db_integrada'].empty:
        db_i = st.session_state['db_integrada']
        db_i['Fecha_DT'] = pd.to_datetime(db_i['Fecha'], format='%d/%m/%Y', errors='coerce')
        db_i = db_i.sort_values('Fecha_DT')
        
        st.markdown("---")
        rango_i = st.date_input("📅 Rango de Fechas a Evaluar:", [], key="rango_int")
        
        if len(rango_i) == 2:
            f_ini, f_fin = rango_i
            db_i_filt = db_i[(db_i['Fecha_DT'].dt.date >= f_ini) & (db_i['Fecha_DT'].dt.date <= f_fin)].copy()
            
            if db_i_filt.empty:
                st.warning("⚠️ No hay registros en ese rango de fechas.")
            else:
                total_flota_master = len(FLOTA_MASTER_DATA)
                
                promedio_activas = db_i_filt['Activas'].mean()
                disp_porcentaje = (promedio_activas / total_flota_master) * 100
                
                promedio_gps = db_i_filt['Con_GPS'].mean()
                gps_porcentaje = (promedio_gps / total_flota_master) * 100
                
                logo = obtener_logo_base64()
                area_logo = f'<div style="background-color: #1a237e; padding: 12px 20px; border-radius: 8px; display: flex; align-items: center; justify-content: center;"><img src="{logo}" style="max-height: 55px; max-width: 180px; object-fit: contain;"></div>' if logo else '<div style="background-color: #1a237e; padding: 12px 20px; border-radius: 8px; display: flex; align-items: center; justify-content: center;"><span style="font-size:24px; color: white; font-weight: bold;">📊 REPORTE</span></div>'
                
                filas_i = ""
                for _, r in db_i_filt.iterrows():
                    filas_i += f"""
                    <tr style="border-bottom: 1px solid #ddd; background: #fff;">
                        <td style="padding: 8px; text-align: center; color: #555;">{r['Fecha']}</td>
                        <td style="padding: 8px; font-weight: bold; color: #333; text-align: center;">{r['Total_Flota']}</td>
                        <td style="padding: 8px; font-weight: bold; color: #2e7d32; text-align: center;">{r['Activas']}</td>
                        <td style="padding: 8px; font-weight: bold; color: #d32f2f; text-align: center;">{r['Inactivas']}</td>
                        <td style="padding: 8px; font-weight: bold; color: #1565c0; text-align: center;">{r['Con_GPS']}</td>
                        <td style="padding: 8px; font-weight: bold; color: #d32f2f; text-align: center;">{r['Sin_GPS']}</td>
                    </tr>
                    """
                    
                filas_chart_i = ""
                max_chart = total_flota_master
                for _, r in db_i_filt.iterrows():
                    ancho_m = (r['Activas'] / max_chart) * 100
                    ancho_g = (r['Con_GPS'] / max_chart) * 100
                    color_m = "#2e7d32" if ancho_m >= 90 else ("#f57c00" if ancho_m >= 80 else "#d32f2f")
                    color_g = "#1565c0" if ancho_g >= 90 else ("#f57c00" if ancho_g >= 80 else "#d32f2f")
                    
                    filas_chart_i += f"""
                    <div style="margin-bottom: 15px; border-bottom: 1px dashed #ccc; padding-bottom: 10px;">
                        <div style="font-size: 13px; font-weight: bold; color: #555; margin-bottom: 5px;">{r['Fecha']} ({str(r['Dia']).title()})</div>
                        <div style="display: flex; gap: 10px;">
                            <div style="flex: 1;">
                                <div style="font-size: 11px; color: {color_m}; margin-bottom: 2px;">{r['Activas']} Operativas</div>
                                <div style="background-color: #e0e0e0; border-radius: 5px; width: 100%; height: 12px; overflow: hidden;">
                                    <div style="background-color: {color_m}; width: {ancho_m}%; height: 100%;"></div>
                                </div>
                            </div>
                            <div style="flex: 1;">
                                <div style="font-size: 11px; color: {color_g}; margin-bottom: 2px;">{r['Con_GPS']} Transmitiendo</div>
                                <div style="background-color: #e0e0e0; border-radius: 5px; width: 100%; height: 12px; overflow: hidden;">
                                    <div style="background-color: {color_g}; width: {ancho_g}%; height: 100%;"></div>
                                </div>
                            </div>
                        </div>
                    </div>
                    """

                pdf_estatus_int = f"""
                <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                
                <div style="text-align: right; margin-bottom: 25px;">
                    <button onclick="descargarPDFEstInt()" style="background-color: #1a237e; color: white; border: none; padding: 15px 30px; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold;">⬇️ DESCARGAR REPORTE INTEGRADO (PDF)</button>
                </div>
                
                <div id="pdf-est-int" style="width: 800px; margin: auto; background: white; padding: 40px; box-sizing: border-box; box-shadow: 0 0 10px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
                    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #1a237e; padding-bottom: 15px; margin-bottom: 25px;">
                        <div>{area_logo}</div>
                        <div style="text-align: right;">
                            <h1 style="margin: 0; color: #1a237e; font-size: 22px; text-transform: uppercase;">Reporte Integrado (Mecánica y GPS)</h1>
                            <p style="margin: 5px 0 0; color: #555; font-size: 14px;">Periodo: {f_ini.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}</p>
                        </div>
                    </div>
                    
                    <div style="display: flex; gap: 15px; margin-bottom: 30px;">
                        <div style="flex: 1; background-color: #f8f9fa; border-top: 4px solid #2e7d32; padding: 15px; text-align: center; border-radius: 5px;">
                            <p style="margin: 0; font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">Disponibilidad Mecánica</p>
                            <h2 style="margin: 5px 0 0; color: #2e7d32; font-size: 26px;">{disp_porcentaje:.1f}%</h2>
                            <p style="margin: 5px 0 0; font-size: 12px; color: #555;">Promedio: {promedio_activas:.1f} activas</p>
                        </div>
                        <div style="flex: 1; background-color: #f8f9fa; border-top: 4px solid #1565c0; padding: 15px; text-align: center; border-radius: 5px;">
                            <p style="margin: 0; font-size: 12px; color: #666; font-weight: bold; text-transform: uppercase;">Cobertura GPS</p>
                            <h2 style="margin: 5px 0 0; color: #1565c0; font-size: 26px;">{gps_porcentaje:.1f}%</h2>
                            <p style="margin: 5px 0 0; font-size: 12px; color: #555;">Promedio: {promedio_gps:.1f} on-line</p>
                        </div>
                    </div>
                    
                    <h3 style="color: #1a237e; border-bottom: 2px solid #1a237e; padding-bottom: 5px;">Evolución Operativa vs Satelital</h3>
                    <div style="padding: 20px; background-color: #f8f9fa; border-radius: 8px; margin-bottom: 30px; border: 1px solid #eee;">
                        {filas_chart_i}
                    </div>
                    
                    <h3 style="color: #1a237e; border-bottom: 2px solid #1a237e; padding-bottom: 5px;">Bitácora Diaria Integrada</h3>
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                        <thead>
                            <tr style="background-color: #1a237e; color: white;">
                                <th style="padding: 10px; text-align: center;">Fecha</th>
                                <th style="padding: 10px; text-align: center;">Flota Total</th>
                                <th style="padding: 10px; text-align: center; background-color: #2e7d32;">Activas</th>
                                <th style="padding: 10px; text-align: center; background-color: #d32f2f;">Taller</th>
                                <th style="padding: 10px; text-align: center; background-color: #1565c0;">Con GPS</th>
                                <th style="padding: 10px; text-align: center; background-color: #d32f2f;">Sin GPS</th>
                            </tr>
                        </thead>
                        <tbody>
                            {filas_i}
                        </tbody>
                    </table>
                    
                    <div style="margin-top: 30px; text-align: center; font-size: 11px; color: #aaa; border-top: 1px solid #eee; padding-top: 10px;">
                        Departamento de Flota y Logística - Documento Generado por PCD
                    </div>
                </div>
                
                <script>
                function descargarPDFEstInt() {{
                    const {{ jsPDF }} = window.jspdf;
                    html2canvas(document.getElementById('pdf-est-int'), {{ scale: 2 }}).then(canvas => {{
                        const doc = new jsPDF('p', 'pt', 'a4');
                        const pdfWidth = doc.internal.pageSize.getWidth();
                        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
                        doc.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, pdfWidth, pdfHeight);
                        doc.save('Reporte_Integrado_Flota_GPS.pdf');
                    }});
                }}
                </script>
                """
                components.html(pdf_estatus_int, height=1200, scrolling=True)
