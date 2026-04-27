# ==========================================
# Archivo: seguridad.py (Módulo de Prevención y Control)
# ==========================================
import streamlit as st
import pandas as pd
import re
import base64
from pathlib import Path
from datetime import datetime, timedelta
import streamlit.components.v1 as components
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# CONFIGURACIÓN INICIAL
st.set_page_config(page_title="PCD - Seguridad", layout="wide")

# ==========================================
# FUNCIÓN PARA INYECTAR LOGO (Base64)
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
# PUENTE A GOOGLE SHEETS
# ==========================================
def obtener_cliente_sheets():
    alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credenciales = ServiceAccountCredentials.from_json_keyfile_name('credenciales.json', alcance)
    return gspread.authorize(credenciales)

def guardar_en_google_sheets(df_para_guardar, nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        doc = cliente.open("PCD_BaseDatos")
        try:
            hoja = doc.worksheet(nombre_hoja)
        except Exception:
            hoja = doc.sheet1
            st.warning(f"⚠️ No se encontró la hoja '{nombre_hoja}'. Guardando en la primera pestaña.")
            
        df_string = df_para_guardar.astype(str)
        valores_a_insertar = df_string.values.tolist()
        hoja.append_rows(valores_a_insertar)
    except Exception as e:
        st.error(f"Error de conexión a Sheets: {e}")

def extraer_datos_sheets(nombre_hoja):
    try:
        cliente = obtener_cliente_sheets()
        doc = cliente.open("PCD_BaseDatos")
        hoja = doc.worksheet(nombre_hoja)
        data = hoja.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()

# ==========================================
# MOTOR MATEMÁTICO DE REPORTES (CON PARCHE DE MADRUGADA)
# ==========================================
def parsear_hora_para_orden(hora_str):
    try:
        limpio = str(hora_str).upper().replace(' ', '').replace('.', '')
        if 'M' in limpio and 'AM' not in limpio and 'PM' not in limpio:
            limpio = limpio.replace('M', 'PM')
        if 'AM' not in limpio and 'PM' not in limpio:
            limpio += 'PM' 
            
        match = re.search(r'(\d{1,2}):(\d{2})\s*([AP]M)', limpio)
        if match:
            h = int(match.group(1))
            m = int(match.group(2))
            ampm = match.group(3)
            
            if ampm == 'PM' and h != 12: h += 12
            if ampm == 'AM' and h == 12: h = 0
            
            # --- PARCHE GERENCIAL DE MADRUGADA ---
            # Si la hora es menor a las 6:00 AM, le sumamos 24 horas 
            # para que el sistema lo ponga al final del día anterior
            if h < 6:
                h += 24
            # ---------------------------------------
                
            return h * 100 + m
        return 9999
    except:
        return 9999

def hora_a_minutos(hora_str):
    h_int = parsear_hora_para_orden(hora_str)
    if h_int == 9999: return None
    h = h_int // 100
    m = h_int % 100
    return h * 60 + m

def minutos_a_hora(mins):
    if pd.isna(mins): return "N/R"
    mins = int(mins)
    h = mins // 60
    m = mins % 60
    
    # Revertimos el parche de madrugada para que se muestre bien en el reporte
    if h >= 24:
        h -= 24
        
    ampm = "AM"
    if h >= 12:
        ampm = "PM"
        if h > 12: h -= 12
    if h == 0: h = 12
    return f"{h:02d}:{m:02d} {ampm}"

# ==========================================
# MOTOR DE EXPANSIÓN DE FECHAS
# ==========================================
def expandir_df_rol(df):
    """Detecta rangos 'AL' en la columna Fecha y multiplica la fila para cada día de la semana."""
    if df is None or df.empty: return df
    filas_exp = []
    dias_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    
    for _, r in df.iterrows():
        fecha_str = str(r.get('Fecha', ''))
        fechas = re.findall(r'\d{2}/\d{2}/\d{4}', fecha_str)
        
        # Si encuentra un rango con "AL" (Hasta 60 días para cubrir todo Caracas)
        if len(fechas) == 2 and "AL" in fecha_str.upper():
            try:
                f_ini = datetime.strptime(fechas[0], "%d/%m/%Y")
                f_fin = datetime.strptime(fechas[1], "%d/%m/%Y")
                delta = f_fin - f_ini
                
                if 0 <= delta.days <= 60:
                    for i in range(delta.days + 1):
                        f_act = f_ini + timedelta(days=i)
                        r_new = r.copy()
                        r_new['Fecha'] = f"{dias_es[f_act.weekday()].upper()} {f_act.strftime('%d/%m/%Y')}"
                        filas_exp.append(r_new)
                    continue
            except: pass
            
        if len(fechas) >= 1 and "AL" not in fecha_str.upper():
            try:
                f_ini = datetime.strptime(fechas[0], "%d/%m/%Y")
                r_new = r.copy()
                r_new['Fecha'] = f"{dias_es[f_ini.weekday()].upper()} {fechas[0]}"
                filas_exp.append(r_new)
                continue
            except: pass
            
        filas_exp.append(r)
        
    return pd.DataFrame(filas_exp)

def formatear_fecha_rol(texto):
    m_rango = re.search(r'(\d{2}/\d{2}/\d{4})\s*AL\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if m_rango:
        f1, f2 = m_rango.groups()
        try:
            obj_f1 = datetime.strptime(f1, "%d/%m/%Y")
            dias_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            dia_str = dias_es[obj_f1.weekday()].upper()
            return f"{dia_str} {f1} AL {f2}"
        except:
            return f"{f1} AL {f2}"
            
    m_caracas = re.search(r'CARACAS\s+(\d{2}/\d{2}/\d{4}.*)', texto, re.IGNORECASE)
    if m_caracas:
        rango_match = re.search(r'(\d{2}/\d{2}/\d{4}\s*AL\s*\d{2}/\d{2}/\d{4})', m_caracas.group(1), re.IGNORECASE)
        if rango_match:
            return f"CARACAS - {rango_match.group(1).upper()}"
        limpio = m_caracas.group(1).split('NOMBRE-PERSONAL')[0].strip()
        return f"CARACAS - {limpio.upper()}"
        
    m_fecha = re.search(r'\d{2}/\d{2}/\d{4}', texto)
    if m_fecha:
        dias_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
        fecha_str = m_fecha.group()
        try:
            obj_fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
            dia_semana = dias_es[obj_fecha.weekday()]
            return f"{dia_semana.upper()} {fecha_str}"
        except:
            return fecha_str
    return None

# ==========================================
# MOTORES DE EXTRACCIÓN Y ORDENAMIENTO
# ==========================================
def extraer_rol_pdf(archivo_pdf):
    try:
        import pdfplumber
    except ImportError:
        return None, None, "FALTA_LIBRERIA"
        
    datos = []
    vacaciones_info = "N/A"
    
    def limpiar_personal(texto):
        if not texto: return ""
        items = []
        for item in re.split(r'[\n,]', str(texto)):
            i_clean = item.strip()
            if len(re.sub(r'[^a-zA-Z0-9]', '', i_clean)) == 0: continue
            if i_clean.upper() not in ["N/A", "0"]: items.append(i_clean)
        return "\n".join(items)

    try:
        with pdfplumber.open(archivo_pdf) as pdf:
            for page in pdf.pages:
                tablas = page.extract_tables()
                texto_pagina = page.extract_text() or ""
                
                fecha_actual = formatear_fecha_rol(texto_pagina)
                if not fecha_actual:
                    fecha_actual = datetime.now().strftime("%d/%m/%Y")

                for tabla in tablas:
                    for fila in tabla:
                        fila_limpia = [str(x).strip() if x is not None else "" for x in fila]
                        if not fila_limpia: continue
                        
                        texto_fila = " ".join(fila_limpia)

                        if "OFICIAL" in texto_fila.upper() and ("VACASIONES" in texto_pagina.upper() or "VACACIONES" in texto_pagina.upper()):
                            if len(fila_limpia) >= 3:
                                vac = fila_limpia[-1]
                                if vac: vacaciones_info = vac
                            continue

                        nueva_fecha = formatear_fecha_rol(texto_fila)
                        if nueva_fecha: fecha_actual = nueva_fecha

                        area_str = fila_limpia[0].upper()
                        areas_validas = ["FLOTA", "ALMACEN", "ALMACÉN", "CIUDAD DROTACA", "TORRE", "CARACAS", "SUPERVISOR", "MONITORES"]

                        if any(a in area_str for a in areas_validas) and "TOTAL" not in area_str:
                            area = fila_limpia[0].replace('\n', ' ').strip()
                            cantidad, diurno, nocturno = "", "", ""

                            if len(fila_limpia) >= 4 and (fila_limpia[1].isdigit() or fila_limpia[1] == ""):
                                cantidad, diurno, nocturno = fila_limpia[1], fila_limpia[2], fila_limpia[3]
                            elif len(fila_limpia) == 3 and fila_limpia[1].isdigit():
                                cantidad, diurno = fila_limpia[1], fila_limpia[2]
                            elif len(fila_limpia) >= 3:
                                diurno, nocturno = fila_limpia[1], fila_limpia[2]
                            elif len(fila_limpia) == 2:
                                diurno = fila_limpia[1]

                            datos.append({
                                "Fecha": fecha_actual, "Área": area, "Cantidad": cantidad,
                                "Diurno": limpiar_personal(diurno), "Nocturno": limpiar_personal(nocturno)
                            })
        return pd.DataFrame(datos), vacaciones_info, "OK"
    except Exception as e:
        return None, None, str(e)

def extraer_apertura(texto):
    datos = {"Fecha": datetime.now().strftime("%d/%m/%Y"), "Sede": "Drotaca 2.0", "Hora Apertura": "", "Hora Alarma": "", "Observaciones": ""}
    m_apertura = re.search(r'Hora:\s*(.*?(?:AM|PM|am|pm))', texto)
    if m_apertura: datos["Hora Apertura"] = m_apertura.group(1).strip()
    m_alarma = re.search(r'Alarmas de Seguridad:\s*(.*?(?:AM|PM|am|pm))', texto)
    if m_alarma: datos["Hora Alarma"] = m_alarma.group(1).strip()
    obs = re.search(r'Observaci[oó]n:\s*(.*)', texto, re.IGNORECASE)
    if obs: datos["Observaciones"] = obs.group(1).strip()
    elif "recorrido" in texto.lower(): datos["Observaciones"] = "Recorrido por pasillos y almacén. Todo en orden."
    return pd.DataFrame([datos])

def extraer_cierre_drotaca(texto):
    fecha = datetime.now().strftime("%d/%m/%Y")
    m_fecha = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})', texto)
    if m_fecha: fecha = m_fecha.group(1)
    departamentos = []
    for linea in texto.split('\n'):
        if '✅' in linea:
            partes = linea.split('✅')[-1].strip()
            m = re.search(r'(.*?)\s+(\d{1,2}:\d{2}\s*[APap]?[Mm]?)', partes)
            if m: departamentos.append({"Fecha": fecha, "Departamento": m.group(1).strip(), "Hora Salida": m.group(2).strip()})
    df = pd.DataFrame(departamentos)
    if not df.empty:
        df['Orden'] = df['Hora Salida'].apply(parsear_hora_para_orden)
        df = df.sort_values(by='Orden').drop(columns=['Orden']).reset_index(drop=True)
    return df

def extraer_cierre_juanita(texto):
    fecha = datetime.now().strftime("%d/%m/%Y")
    grupos = []
    for linea in texto.split('\n'):
        m = re.search(r'(Personal.*?)\s+Salida Hora:\s*(\d{1,2}:\d{2}\s*[a-zA-Z]{0,2})', linea, re.IGNORECASE)
        if m: grupos.append({"Fecha": fecha, "Grupo": m.group(1).strip(), "Hora Salida": m.group(2).strip()})
    df = pd.DataFrame(grupos)
    if not df.empty:
        df['Orden'] = df['Hora Salida'].apply(parsear_hora_para_orden)
        df = df.sort_values(by='Orden').drop(columns=['Orden']).reset_index(drop=True)
    m_cierre = re.search(r'cierre del mismo Hora:\s*(\d{1,2}:\d{2}\s*[a-zA-Z]{0,2})', texto, re.IGNORECASE)
    hora_cierre = m_cierre.group(1).strip() if m_cierre else "N/A"
    oficiales = []
    bloque_ofi = re.search(r'Oficiales de seguridad[^\n]*\n(.*?)(?:Todo sin [Nn]ovedad|Sin [Nn]ovedad)', texto, re.IGNORECASE | re.DOTALL)
    if bloque_ofi:
        for linea in bloque_ofi.group(1).split('\n'):
            limpio = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', linea).strip()
            if len(limpio) > 3: oficiales.append(limpio)
    return df, hora_cierre, oficiales

def extraer_personal_cierre(texto):
    fecha = datetime.now().strftime("%d/%m/%Y")
    data = []
    area_actual = "Área General"
    for linea in texto.split('\n'):
        linea = linea.strip()
        if not linea: continue
        m = re.match(r'^[\✓\✔]\s*(.*)', linea)
        if not m: m = re.match(r'^[vV]\s+(.*)', linea)
        if m:
            persona_raw = m.group(1).strip().replace('*', '')
            if (len(persona_raw.split()) == 1 or persona_raw.lower() in ['transbordo', 'oriente', 'occidente', 'centro', 'nacional', 'local']) and len(data) > 0 and data[-1]["Area Asignada"] == area_actual:
                data[-1]["Personal"] += f" ➔ {persona_raw.capitalize()}"
            else:
                data.append({"Fecha": fecha, "Area Asignada": area_actual, "Personal": " ".join([p.capitalize() for p in persona_raw.split()])})
        else:
            t = linea.replace('*', '').strip()
            if len(t) > 60 or "buenas noches" in t.lower() or "bendiciones" in t.lower() or "nocturno" in t.lower() or "distribución" in t.lower(): continue
            area_actual = t.capitalize()
    return pd.DataFrame(data)

# ==========================================
# GENERADORES DE TEXTO PARA WHATSAPP
# ==========================================
def generar_ws_apertura(df):
    if df.empty: return ""
    r = df.iloc[0]
    return f"*REPORTE DE APERTURA - DROTACA 2.0*\n📅 Fecha: {r['Fecha']}\n\nApertura de Sede: {r['Hora Apertura']}\nDesactivación Alarma: {r['Hora Alarma']}\nObservación: {r['Observaciones']}\n\n✅ Área segura e iniciamos jornada con la bendición de Dios."

def generar_ws_cierre_drotaca(df):
    if df.empty: return ""
    ultimo = df.iloc[-1]
    return f"*REPORTE DE CIERRE - DROTACA 2.0*\n📅 Fecha: {df['Fecha'].iloc[0]}\n\n*Métrica del día:*\nTotal Dptos despachados: {len(df)}\nÚltima salida registrada: {ultimo['Departamento']} a las {ultimo['Hora Salida']}\n\n✅ Sede asegurada, instalaciones operativas y sin novedades."

def generar_ws_cierre_juanita(df, h_cierre, oficiales):
    if df.empty: return ""
    ultimo = df.iloc[-1]
    ofi_str = ", ".join(oficiales) if oficiales else "No registrados"
    return f"*REPORTE DE CIERRE - ALMACÉN JUANITA*\n📅 Fecha: {df['Fecha'].iloc[0]}\nHora Cierre General: {h_cierre}\n\n*Métrica del día:*\nTotal grupos despachados: {len(df)}\nÚltimo grupo: {ultimo['Grupo']} a las {ultimo['Hora Salida']}\nOficiales activos: {ofi_str}\n\n✅ Todo sin novedad."

def generar_ws_personal_cierre(df):
    if df.empty: return ""
    msg = f"*ASIGNACIÓN DE PERSONAL - CIERRE DROTACA 2.0*\n📅 Fecha: {df['Fecha'].iloc[0]}\n\n"
    for area in df['Area Asignada'].unique():
        msg += f"*{area}*\n"
        for p in df[df['Area Asignada'] == area]['Personal']: msg += f"✓ {p}\n"
        msg += "\n"
    return msg.strip() + "\n\n✅ Personal notificado y en posición."

def generar_ws_rol_guardia(df, titulo, vacaciones="N/A"):
    msg = f"*ASIGNACIÓN DE GUARDIAS - {titulo.upper()}*\n"
    for fecha in df['Fecha'].unique():
        msg += f"\n📅 *OFICIALES DE SEGURIDAD {fecha}*\n"
        for _, r in df[df['Fecha'] == fecha].iterrows():
            msg += f"\n*{r['Área']}*\n"
            if r['Diurno'] and str(r['Diurno']).strip(): msg += f"Diurno: {', '.join([x.strip() for x in r['Diurno'].split(chr(10)) if x.strip()])}\n"
            if r['Nocturno'] and str(r['Nocturno']).strip(): msg += f"Nocturno: {', '.join([x.strip() for x in r['Nocturno'].split(chr(10)) if x.strip()])}\n"
    if str(vacaciones).strip() and vacaciones != "N/A": msg += f"\n*Personal de Vacaciones:* {vacaciones}\n"
    return msg.strip() + "\n\n✅ Todo el personal debe estar en sus áreas correspondientes."

# ==========================================
# CREADOR DEL REPORTE SEMANAL GERENCIAL
# ==========================================
def analizar_datos_semanales(df_ape, df_dro, df_jua, semana_filtro, mes_filtro):
    def filtrar(df):
        if df.empty: return df
        d = df.copy()
        if mes_filtro != "Todos" and 'Mes' in d.columns: d = d[d['Mes'] == mes_filtro]
        if semana_filtro != "Todas" and 'Semana' in d.columns: d = d[d['Semana'] == semana_filtro]
        return d

    f_ape, f_dro, f_jua = filtrar(df_ape), filtrar(df_dro), filtrar(df_jua)
    dias_orden = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    resumen = {d: {"fecha": "", "ape": "N/R", "jua": "N/R", "dro": "N/R", "m_dro": [], "m_jua": []} for d in dias_orden}
    
    if not f_ape.empty and 'Dia' in f_ape.columns:
        for _, r in f_ape.iterrows():
            dia = str(r['Dia']).capitalize()
            if dia in resumen:
                resumen[dia]['fecha'] = str(r.get('Fecha', ''))
                resumen[dia]['ape'] = str(r.get('Hora Apertura', 'N/R'))
                
    if not f_jua.empty and 'Dia' in f_jua.columns:
        col_cierre = 'Hora Cierre' if 'Hora Cierre' in f_jua.columns else 'Hora Cierre Total'
        for _, r in f_jua.iterrows():
            dia = str(r['Dia']).capitalize()
            if dia in resumen:
                resumen[dia]['fecha'] = str(r.get('Fecha', ''))
                h = str(r.get(col_cierre, 'N/R'))
                resumen[dia]['jua'] = h
                m = hora_a_minutos(h)
                if m: resumen[dia]['m_jua'].append(m)

    if not f_dro.empty and 'Dia' in f_dro.columns and 'Hora Salida' in f_dro.columns:
        for fecha, group in f_dro.groupby('Fecha'):
            dia = str(group.iloc[0].get('Dia', '')).capitalize()
            if dia in resumen:
                resumen[dia]['fecha'] = fecha
                g = group.copy()
                g['min'] = g['Hora Salida'].apply(hora_a_minutos)
                g = g.dropna(subset=['min']).sort_values('min')
                if not g.empty:
                    resumen[dia]['dro'] = str(g.iloc[-1]['Hora Salida'])
                    resumen[dia]['m_dro'].append(g.iloc[-1]['min'])
    return resumen

def generar_texto_semanal(resumen, semana_str):
    texto = f"*REPORTE DE GESTIÓN SEMANAL - {semana_str.upper()}* 📊\n\n"
    dro_mins, jua_mins = [], []
    hay_datos = False
    
    for dia, data in resumen.items():
        if data['fecha']:
            hay_datos = True
            texto += f"*{dia.upper()} ({data['fecha']})*\n🔓 Apertura: {data['ape']}\n📦 Cierre Juanita: {data['jua']}\n🏢 Cierre Drotaca: {data['dro']}\n\n"
            dro_mins.extend(data['m_dro'])
            jua_mins.extend(data['m_jua'])
            
    if not hay_datos:
        return "⚠️ No hay datos registrados para la semana y mes seleccionados."

    texto += "*(--- ANÁLISIS DE LA SEMANA ---)*\n"
    if dro_mins: texto += f"⏱️ Promedio Cierre Drotaca: {minutos_a_hora(sum(dro_mins)/len(dro_mins))}\n🚨 Cierre más tardío Drotaca: {minutos_a_hora(max(dro_mins))}\n"
    if jua_mins: texto += f"⏱️ Promedio Cierre Juanita: {minutos_a_hora(sum(jua_mins)/len(jua_mins))}\n🚨 Cierre más tardío Juanita: {minutos_a_hora(max(jua_mins))}\n"
    return texto + "\n✅ Semana analizada y bajo control operativo."

# ==========================================
# GENERADORES VISUALES (HTML/CANVAS/PDF)
# ==========================================

# ---------------------------------------------------------
# FUNCION MAESTRA PARA AGRUPAR DIAS (SÁBADO, DOMINGO, LUNES A VIERNES)
# ---------------------------------------------------------
def agrupar_por_dias(df_rol):
    """Toma el DF de roles expandido y lo agrupa para ahorrar hojas en el PDF y en la imagen."""
    def extraer_dt(texto):
        m = re.search(r'\d{2}/\d{2}/\d{4}', str(texto))
        if m: return datetime.strptime(m.group(), "%d/%m/%Y")
        return datetime.min
        
    df_rol['dt'] = df_rol['Fecha'].apply(extraer_dt)
    df_rol = df_rol.sort_values('dt')
    
    pages_data = []
    
    # Sábado
    df_sat = df_rol[df_rol['dt'].dt.weekday == 5]
    if not df_sat.empty:
        fecha_str = f"SÁBADO {df_sat['dt'].iloc[0].strftime('%d/%m/%Y')}"
        pages_data.append((fecha_str, df_sat.drop_duplicates(subset=['Área', 'Diurno', 'Nocturno'])))
        
    # Domingo
    df_sun = df_rol[df_rol['dt'].dt.weekday == 6]
    if not df_sun.empty:
        fecha_str = f"DOMINGO {df_sun['dt'].iloc[0].strftime('%d/%m/%Y')}"
        pages_data.append((fecha_str, df_sun.drop_duplicates(subset=['Área', 'Diurno', 'Nocturno'])))
        
    # Lunes a Viernes
    df_lv = df_rol[df_rol['dt'].dt.weekday.isin([0, 1, 2, 3, 4])]
    if not df_lv.empty:
        min_dt = df_lv['dt'].min()
        max_dt = df_lv['dt'].max()
        dias_es = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]
        n1 = dias_es[min_dt.weekday()]
        n2 = dias_es[max_dt.weekday()]
        
        if min_dt != max_dt:
            fecha_str = f"{n1} {min_dt.strftime('%d/%m/%Y')} AL {n2} {max_dt.strftime('%d/%m/%Y')}"
        else:
            fecha_str = f"{n1} {min_dt.strftime('%d/%m/%Y')}"
            
        pages_data.append((fecha_str, df_lv.drop_duplicates(subset=['Área', 'Diurno', 'Nocturno'])))
        
    return pages_data

def html_reporte_semanal_pizarra(resumen, titulo):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 50px;">' if logo else '🛡️'
    filas = ""
    dro_mins, jua_mins, count = [], [], 0
    
    for dia, data in resumen.items():
        if not data['fecha']: continue
        bg = "#ffffff" if count % 2 == 0 else "#f8f9fa"
        count += 1
        filas += f"<tr style='background-color: {bg}; border-bottom: 1px solid #ddd; text-align: center;'><td style='padding: 12px; font-weight: bold; color: #1a237e; border-right: 1px solid #eee;'>{dia.upper()}<br><span style='font-size:11px; color:#666;'>{data['fecha']}</span></td><td style='padding: 12px; font-weight: bold; color: #2e7d32; border-right: 1px solid #eee;'>{data['ape']}</td><td style='padding: 12px; font-weight: bold; color: #e65100; border-right: 1px solid #eee;'>{data['jua']}</td><td style='padding: 12px; font-weight: bold; color: #000000;'>{data['dro']}</td></tr>"
        dro_mins.extend(data['m_dro']); jua_mins.extend(data['m_jua'])

    if count == 0: filas = "<tr><td colspan='4' style='padding: 20px; text-align: center; color: red;'>No hay datos para los filtros seleccionados</td></tr>"

    kpi_dro = f"Promedio Cierre Drotaca: {minutos_a_hora(sum(dro_mins)/len(dro_mins)) if dro_mins else 'N/A'}"
    kpi_jua = f"Promedio Cierre Juanita: {minutos_a_hora(sum(jua_mins)/len(jua_mins)) if jua_mins else 'N/A'}"

    return f"""<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><div style="text-align:right; margin-bottom:15px;"><button onclick="descargarSemanal()" style="background:#1a237e;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;">⬇️ Descargar Reporte Semanal</button></div><div id="pizarra-semanal" style="font-family: Arial, sans-serif; width: 800px; margin: auto; background-color: #fff; border: 2px solid #1a237e; border-radius: 10px; overflow: hidden;"><div style="background-color: #1a237e; color: white; padding: 20px; display: flex; align-items: center; justify-content: space-between;">{area_logo}<div style="text-align: right;"><h2 style="margin: 0; font-size: 20px;">REPORTE SEMANAL DE GESTIÓN</h2><p style="margin: 5px 0 0; font-size: 14px; opacity: 0.9;">{titulo}</p></div></div><div style="padding: 20px;"><table style="width: 100%; border-collapse: collapse;"><thead><tr style="background-color: #e8eaf6; color: #1a237e; font-size: 13px;"><th style="padding: 10px; border-bottom: 2px solid #1a237e;">DÍA</th><th style="padding: 10px; border-bottom: 2px solid #1a237e;">APERTURA</th><th style="padding: 10px; border-bottom: 2px solid #1a237e;">CIERRE JUANITA</th><th style="padding: 10px; border-bottom: 2px solid #1a237e;">CIERRE DROTACA</th></tr></thead><tbody>{filas}</tbody></table><div style="margin-top: 20px; padding: 15px; background-color: #f1f8e9; border-radius: 5px; display: flex; justify-content: space-around; border: 1px solid #c5e1a5;"><div style="text-align: center; color: #2e7d32; font-weight: bold;">📊 {kpi_dro}</div><div style="text-align: center; color: #e65100; font-weight: bold;">📊 {kpi_jua}</div></div></div></div><script>function descargarSemanal() {{ html2canvas(document.getElementById('pizarra-semanal'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Reporte_Semanal_Cierres.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>"""

def html_reporte_rol_pdf(df_rol, titulo, vacaciones="N/A"):
    # FORZAR EXPANSIÓN
    df_rol = expandir_df_rol(df_rol)
    
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 60px; max-width: 150px; object-fit: contain; display: block;">' if logo else '<span style="font-size:30px;">🛡️</span>'
    
    content_html = ""
    
    # -------------------------------------------------------------
    # RECOLECCIÓN DE PERSONAL PARA ESTADÍSTICAS
    # -------------------------------------------------------------
    oficiales_unicos = set()
    for _, r in df_rol.iterrows():
        # Ignorar al supervisor en el conteo de oficiales
        if 'SUPERVISOR' in str(r.get('Área', '')).upper():
            continue
            
        d_list = str(r.get('Diurno', '')).split('\n')
        n_list = str(r.get('Nocturno', '')).split('\n')
        for linea in d_list + n_list:
            limpio = linea.strip()
            limpio = limpio.replace('✓', '').replace('☑', '').replace('✔', '').replace('v ', '').replace('V ', '').strip()
            limpio = re.split(r'[:\(]', limpio)[0].strip()
            if limpio and limpio.upper() not in ['N/A', '0', 'NO ASIGNADO', 'NINGUNO']:
                oficiales_unicos.add(limpio.title())
                
    total_activos = len(oficiales_unicos)
    
    vac_list = []
    if str(vacaciones).strip() and str(vacaciones).upper() not in ["N/A", "0", "NINGUNO", ""]:
        v_clean = re.sub(r'[✓☑✔\-]', '', str(vacaciones))
        vac_list = [v.strip().title() for v in re.split(r'[,\n]', v_clean) if v.strip()]
    total_vacaciones = len(vac_list)
    
    total_plantilla = total_activos + total_vacaciones
    porcentaje_activo = round((total_activos / total_plantilla) * 100, 1) if total_plantilla > 0 else 0
    # -------------------------------------------------------------

    # Agrupar los días para el PDF: Sábado (1 hoja), Domingo (1 hoja), Lunes a Viernes (1 hoja)
    grupos_paginas = agrupar_por_dias(df_rol)
            
    for idx, (fecha_titulo, df_f) in enumerate(grupos_paginas):
        
        sup_row = df_f[df_f['Área'].astype(str).str.upper().str.contains('SUPERVISOR')]
        sup_name = "NO ASIGNADO"
        if not sup_row.empty:
            d_sup = str(sup_row.iloc[0].get('Diurno', '')).strip()
            n_sup = str(sup_row.iloc[0].get('Nocturno', '')).strip()
            sup_name = d_sup if d_sup else n_sup
            sup_name = sup_name.replace('\n', ' / ').replace('✓', '').replace('☑', '').strip()
            df_f = df_f[~df_f['Área'].astype(str).str.upper().str.contains('SUPERVISOR')]

        filas = ""
        
        for i, r in df_f.iterrows():
            bg = "#ffffff" if i % 2 == 0 else "#f8f9fa"
            
            # EL LIMPIADOR ANTI-AMONTONAMIENTOS: Agrega espacios vitales
            def fix_text(t):
                t = str(t).replace(':', ': ').replace('(', ' (')
                return re.sub(r'\s+', 't').strip()
            
            diurno_list = [fix_text(l) for l in str(r.get('Diurno', '')).split('\n') if l.strip()]
            nocturno_list = [fix_text(l) for l in str(r.get('Nocturno', '')).split('\n') if l.strip()]
            
            # El uso de display:block básico soluciona los traslapes raros
            diurno_html = "".join([f"<div style='margin-bottom:6px; display:block;'>✓ {'<b>'+l+'</b>' if 'ingresa' in l.lower() else l}</div>" for l in diurno_list])
            nocturno_html = "".join([f"<div style='margin-bottom:6px; display:block;'>✓ {'<b>'+l+'</b>' if 'ingresa' in l.lower() else l}</div>" for l in nocturno_list])
            
            cantidad = r.get('Cantidad', '')
            if pd.isna(cantidad) or str(cantidad).strip() == '':
                cantidad = str(len(diurno_list) + len(nocturno_list))
            
            filas += f"<tr style='background-color: {bg}; page-break-inside: avoid;'><td style='padding: 12px; border: 1px solid #ddd; font-weight: bold; color: #1a237e; vertical-align: top;'>{r['Área']}</td><td style='padding: 12px; border: 1px solid #ddd; font-weight: bold; color: #333; vertical-align: top; text-align: center;'>{cantidad}</td><td style='padding: 12px; border: 1px solid #ddd; color: #333; vertical-align: top;'>{diurno_html}</td><td style='padding: 12px; border: 1px solid #ddd; color: #333; vertical-align: top;'>{nocturno_html}</td></tr>"
        
        # PÁGINAS INDIVIDUALES CLARAS:
        page_break = '<div style="page-break-before: always; clear: both;"></div>' if idx > 0 else ''
        
        content_html += f"""
        {page_break}
        <div style="padding: 0px; font-family: Arial, sans-serif; background: white; margin: 0; box-sizing: border-box;">
            <div style="display: flex; justify-content: space-between; align-items: center; background-color: #1a237e; color: white; padding: 20px 30px; border-bottom: 4px solid #b71c1c;">
                <div style="flex: 0 0 150px; text-align: left;">{area_logo}</div>
                <div style="flex: 1; text-align: right;">
                    <h2 style="margin: 0; color: white; font-size: 22px; text-transform: uppercase;">{titulo}</h2>
                    <p style="margin: 5px 0 0; color: #e8eaf6; font-size: 14px;">Semana Operativa (Sábado a Viernes)</p>
                </div>
            </div>
            
            <div style="padding: 25px 30px;">
                <div style="background-color: #e8eaf6; border: 2px solid #1a237e; border-radius: 8px; padding: 15px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;">
                    <div style="font-size: 18px; color: #1a237e; font-weight: bold;">
                        📅 {fecha_titulo}
                    </div>
                    <div style="font-size: 14px; color: #1a237e; font-weight: bold; background: #fff; padding: 6px 15px; border-radius: 4px; border: 1px solid #1a237e; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                        👮‍♂️ SUPERVISOR DE SEGURIDAD: {sup_name.upper()}
                    </div>
                </div>
                
                <table style="width: 100%; border-collapse: collapse; font-size: 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); background-color: white;">
                    <thead><tr style="background-color: #1a237e; color: white; font-size: 13px;"><th style="padding: 12px; border: 1px solid #1a237e; text-align: left; width: 25%;">ÁREA ASIGNADA</th><th style="padding: 12px; border: 1px solid #1a237e; text-align: center; width: 10%;">CANT.</th><th style="padding: 12px; border: 1px solid #1a237e; text-align: left; width: 32%;">TURNO DIURNO</th><th style="padding: 12px; border: 1px solid #1a237e; text-align: left; width: 33%;">TURNO NOCTURNO</th></tr></thead>
                    <tbody>{filas}</tbody>
                </table>
                
                <div style="margin-top: 30px; text-align: center; font-size: 12px; color: #666; border-top: 1px solid #eee; padding-top: 10px;">
                    Documento Oficial de Seguridad Drotaca - Página {idx + 1}
                </div>
            </div>
        </div>
        """

    # -------------------------------------------------------------
    # HOJA FINAL: ESTADÍSTICAS (CON CHECK EN LUGAR DE EMOJI POLICÍA)
    # -------------------------------------------------------------
    nombres_html = "".join([f"<div style='width: 33%; padding: 6px 0; color: #333;'>✓ {nombre}</div>" for nombre in sorted(oficiales_unicos)])
    vac_html = "".join([f"<div style='width: 33%; padding: 6px 0; color: #b71c1c;'>✓ {nombre}</div>" for nombre in sorted(vac_list)]) if vac_list else "<div style='color: #666; font-style: italic; width: 100%;'>No hay personal registrado en vacaciones.</div>"

    final_page = f"""
    <div style="page-break-before: always; clear: both;"></div>
    <div style="padding: 0px; font-family: Arial, sans-serif; background: white; margin: 0; box-sizing: border-box; min-height: 1040px; display: block;">
        <div style="display: flex; justify-content: space-between; align-items: center; background-color: #1a237e; color: white; padding: 20px 30px; border-bottom: 4px solid #b71c1c;">
            <div style="flex: 0 0 150px; text-align: left;">{area_logo}</div>
            <div style="flex: 1; text-align: right;">
                <h2 style="margin: 0; color: white; font-size: 22px; text-transform: uppercase;">RESUMEN Y ESTADÍSTICAS</h2>
                <p style="margin: 5px 0 0; color: #e8eaf6; font-size: 14px;">Semana Operativa</p>
            </div>
        </div>
        
        <div style="padding: 30px;">
            <div style="display: flex; justify-content: space-between; gap: 20px; margin-bottom: 30px;">
                <div style="flex: 1; background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; padding: 20px; text-align: center; border-top: 4px solid #1a237e; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    <h3 style="margin: 0; color: #666; font-size: 14px; text-transform: uppercase;">Total Oficiales</h3>
                    <p style="margin: 10px 0 0; font-size: 32px; font-weight: bold; color: #1a237e;">{total_plantilla}</p>
                </div>
                <div style="flex: 1; background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; padding: 20px; text-align: center; border-top: 4px solid #2e7d32; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    <h3 style="margin: 0; color: #666; font-size: 14px; text-transform: uppercase;">Personal Activo</h3>
                    <p style="margin: 10px 0 0; font-size: 32px; font-weight: bold; color: #2e7d32;">{total_activos} <span style="font-size: 16px; color: #888;">({porcentaje_activo}%)</span></p>
                </div>
                <div style="flex: 1; background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; padding: 20px; text-align: center; border-top: 4px solid #b71c1c; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    <h3 style="margin: 0; color: #666; font-size: 14px; text-transform: uppercase;">De Vacaciones</h3>
                    <p style="margin: 10px 0 0; font-size: 32px; font-weight: bold; color: #b71c1c;">{total_vacaciones}</p>
                </div>
            </div>

            <h3 style="color: #1a237e; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 15px; font-size: 16px;">🛡️ NÓMINA DE OFICIALES ACTIVOS</h3>
            <div style="display: flex; flex-wrap: wrap; font-size: 13px; font-weight: bold; margin-bottom: 30px; background: #fdfdfd; padding: 15px; border-radius: 5px; border: 1px solid #eee;">
                {nombres_html}
            </div>

            <h3 style="color: #b71c1c; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 15px; font-size: 16px;">PERSONAL DE VACACIONES</h3>
            <div style="display: flex; flex-wrap: wrap; font-size: 13px; font-weight: bold; background: #fffafb; padding: 15px; border-radius: 5px; border: 1px solid #ffebee;">
                {vac_html}
            </div>
            
            <div style="margin-top: 40px; text-align: center; font-size: 12px; color: #666; border-top: 1px solid #eee; padding-top: 15px;">
                Documento Oficial de Seguridad Drotaca - Estadísticas
            </div>
        </div>
    </div>
    """
    
    content_html += final_page

    pages_html = f"""
    <div style="padding: 0px; font-family: Arial, sans-serif; background: #fff;">
        {content_html}
    </div>
    """

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;"><button onclick="descargarPDFGerencial()" style="background:#b71c1c;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:16px;">📄 Descargar PDF Gerencial</button></div>
    <div id="pdf-gerencial" style="background: #fff; padding: 0; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">{pages_html}</div>
    <script>
    function descargarPDFGerencial() {{
        var element = document.getElementById('pdf-gerencial');
        var opt = {{ 
            margin: 0, 
            filename: 'Reporte_Guardias.pdf', 
            image: {{ type: 'jpeg', quality: 0.98 }}, 
            html2canvas: {{ scale: 2, useCORS: true, letterRendering: false, logging: false }}, 
            jsPDF: {{ unit: 'mm', format: 'a4', orientation: 'portrait' }},
            pagebreak: {{ mode: 'css', avoid: 'tr' }}
        }};
        html2pdf().set(opt).from(element).save();
    }}
    </script>
    """

def html_rol_guardia(df, titulo, vacaciones="N/A"):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 55px; width: auto;">' if logo else '🛡️'
    html_content = ""
    
    for fecha_titulo in df['Fecha'].unique():
        df_f = df[df['Fecha'] == fecha_titulo]
        
        html_content += f'<div style="background-color: #1a237e; color: white; padding: 10px; margin-top: 15px; font-weight: bold; text-align: center; border-radius: 5px; font-size: 15px;">OFICIALES DE SEGURIDAD {str(fecha_titulo).upper()}</div><table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 10px;"><thead><tr style="background-color: #e8eaf6; color: #1a237e; font-size: 13px;"><th style="padding: 10px; border: 1px solid #ddd; text-align: left; width: 25%;">ÁREA ASIGNADA</th><th style="padding: 10px; border: 1px solid #ddd; text-align: center; width: 10%;">CANT.</th><th style="padding: 10px; border: 1px solid #ddd; text-align: left; width: 32%;">TURNO DIURNO</th><th style="padding: 10px; border: 1px solid #ddd; text-align: left; width: 33%;">TURNO NOCTURNO</th></tr></thead><tbody>'
        for i, r in df_f.iterrows():
            bg = "#ffffff" if i % 2 == 0 else "#f8f9fa"
            
            diurno_list = [str(l).replace(':', ': ') for l in str(r.get('Diurno', '')).split('\n') if l.strip()]
            nocturno_list = [str(l).replace(':', ': ') for l in str(r.get('Nocturno', '')).split('\n') if l.strip()]
            
            diurno_html = "".join([f"<div style='margin-bottom:4px;'>✓ {'<b>'+l+'</b>' if 'ingresa' in l.lower() else l}</div>" for l in diurno_list])
            nocturno_html = "".join([f"<div style='margin-bottom:4px;'>✓ {'<b>'+l+'</b>' if 'ingresa' in l.lower() else l}</div>" for l in nocturno_list])
            
            cantidad = r.get('Cantidad', '')
            if pd.isna(cantidad) or str(cantidad).strip() == '':
                cantidad = str(len(diurno_list) + len(nocturno_list))
                
            html_content += f"<tr style='background-color: {bg};'><td style='padding: 10px; border: 1px solid #ddd; font-weight: bold; color: #1a237e; vertical-align: top;'>{r['Área']}</td><td style='padding: 10px; border: 1px solid #ddd; font-weight: bold; color: #333; vertical-align: top; text-align: center;'>{cantidad}</td><td style='padding: 10px; border: 1px solid #ddd; color: #333; vertical-align: top;'>{diurno_html}</td><td style='padding: 10px; border: 1px solid #ddd; color: #333; vertical-align: top;'>{nocturno_html}</td></tr>"
        html_content += "</tbody></table>"

    caja_vacaciones = f'<div style="margin-top: 15px; padding: 10px; background-color: #e8eaf6; border-radius: 5px; color: #1a237e; font-weight: bold; text-align: center;">PERSONAL DE VACACIONES: <span style="color: #333; font-weight: normal;">{vacaciones}</span></div>' if str(vacaciones).strip() and vacaciones != "N/A" else ""

    return f"""<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><div style="text-align:right; margin-bottom:15px;"><button onclick="descargarRol()" style="background:#1a237e;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;">⬇️ Descargar Rol de Guardia</button></div><div id="pizarra-rol" style="font-family: Arial, sans-serif; width: 800px; margin: auto; background-color: #fff; border: 2px solid #1a237e; border-radius: 12px; overflow: hidden; padding-bottom: 20px;"><div style="background-color: #1a237e; color: white; padding: 25px; display: flex; align-items: center; justify-content: space-between;"><div>{area_logo}</div><div style="text-align: right;"><h2 style="margin: 0; font-size: 22px;">{titulo.upper()}</h2><p style="margin: 5px 0 0; font-size: 14px;">Asignación Oficial de Seguridad</p></div></div><div style="padding: 0 20px;">{html_content}{caja_vacaciones}</div><p style="text-align: center; color: green; font-weight: bold; margin-top: 20px;">✅ Personal notificado y en posición.</p></div><script>function descargarRol() {{ html2canvas(document.getElementById('pizarra-rol'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Rol_Guardia.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>"""

def html_cierre_drotaca(df):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 50px;">' if logo else '🛡️'
    filas = "".join([f"<tr style='background-color: {'#ffffff' if i % 2 == 0 else '#f8f9fa'}; border-bottom: 1px solid #eee;'><td style='padding: 12px; font-weight: bold; color: #333;'>✅ {r['Departamento']}</td><td style='padding: 12px; text-align: right; color: #000000; font-weight: bold;'>{r['Hora Salida']}</td></tr>" for i, r in df.iterrows()])
    return f"""<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><div style="text-align:right; margin-bottom:15px;"><button onclick="descargarPizarraSeguridad()" style="background:#1a237e;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Cierre</button></div><div id="pizarra-seguridad" style="font-family: Arial, sans-serif; width: 600px; margin: auto; background-color: #fff; border: 2px solid #1a237e; border-radius: 10px; overflow: hidden;"><div style="background-color: #1a237e; color: white; padding: 20px; display: flex; align-items: center; justify-content: space-between;">{area_logo}<div style="text-align: right;"><h2 style="margin: 0; font-size: 22px;">CIERRE DROTACA 2.0</h2><p style="margin: 5px 0 0; font-size: 14px; opacity: 0.9;">Fecha: {df['Fecha'].iloc[0] if not df.empty else ''}</p></div></div><div style="padding: 20px;"><table style="width: 100%; border-collapse: collapse;"><thead><tr style="background-color: #e8eaf6; color: #1a237e; font-size: 14px; font-weight: bold;"><th style="padding: 10px; text-align: left;">DEPARTAMENTO</th><th style="padding: 10px; text-align: right;">HORA DE SALIDA</th></tr></thead><tbody>{filas}</tbody></table></div></div><script>function descargarPizarraSeguridad() {{ html2canvas(document.getElementById('pizarra-seguridad'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Cierre_Drotaca.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>"""

def html_cierre_juanita(df, hora_cierre, oficiales):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 45px;">' if logo else '🛡️'
    filas = "".join([f"<tr style='border-bottom: 1px solid #ddd;'><td style='padding: 10px;'>🚶 {r['Grupo']}</td><td style='padding: 10px; text-align: right; font-weight:bold;'>{r['Hora Salida']}</td></tr>" for i, r in df.iterrows()])
    lista_ofi = "<br>".join([f"👮 {o}" for o in oficiales]) if oficiales else "<i>No registrados</i>"
    return f"""<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><div style="text-align:right; margin-bottom:15px;"><button onclick="descargarJuanita()" style="background:#004d40;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Cierre</button></div><div id="pizarra-juanita" style="font-family: Arial, sans-serif; width: 550px; margin: auto; background-color: #fff; border: 2px solid #004d40; border-radius: 10px; overflow: hidden;"><div style="background-color: #004d40; color: white; padding: 20px; display: flex; align-items: center; justify-content: space-between;">{area_logo}<div style="text-align: right;"><h2 style="margin: 0; font-size: 20px;">CIERRE JUANITA</h2><p style="margin: 5px 0 0; font-size: 12px; font-weight: bold;">Hora Cierre: {hora_cierre}</p></div></div><div style="padding: 20px;"><table style="width: 100%; border-collapse: collapse; font-size: 14px;"><tbody>{filas}</tbody></table><div style="margin-top: 15px; padding: 15px; background-color: #e0f2f1; border-radius: 5px;"><h4 style="margin: 0 0 10px 0; color: #004d40;">Oficiales de Guardia:</h4><div style="font-size: 13px; color: #333;">{lista_ofi}</div></div><p style="text-align: center; color: green; font-weight: bold; margin-top: 15px;">✅ Todo sin novedad</p></div></div><script>function descargarJuanita() {{ html2canvas(document.getElementById('pizarra-juanita'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Cierre_Juanita.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>"""

def html_apertura(df):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 55px;">' if logo else '🛡️'
    r = df.iloc[0]
    return f"""<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><div style="text-align:right; margin-bottom:15px;"><button onclick="descargarApertura()" style="background:#1a237e;color:#fff;border:none;padding:10px 20px;border-radius:5px;font-weight:bold;cursor:pointer;">⬇️ Descargar Apertura</button></div><div id="pizarra-apertura" style="font-family: Arial, sans-serif; width: 600px; margin: auto; background-color: #fff; border: 2px solid #1a237e; border-radius: 12px; overflow: hidden;"><div style="background-color: #1a237e; color: white; padding: 25px; display: flex; align-items: center; justify-content: space-between;">{area_logo}<div style="text-align: right;"><h2 style="margin: 0; font-size: 24px;">APERTURA DROTACA 2.0</h2><p style="margin: 5px 0 0; font-size: 15px; font-weight: bold;">Fecha: {r['Fecha']}</p></div></div><div style="padding: 30px; text-align: center;"><h3 style="color: #444; margin-bottom: 5px; font-size: 18px;">Hora de Apertura</h3><p style="font-size: 36px; font-weight: bold; color: #1a237e; margin-top: 0; margin-bottom: 20px;">{r['Hora Apertura']}</p><h3 style="color: #444; margin-bottom: 5px; font-size: 18px;">Desactivación de Alarmas</h3><p style="font-size: 36px; font-weight: bold; color: #1a237e; margin-top: 0;">{r['Hora Alarma']}</p><div style="margin-top: 35px; padding: 20px; background-color: #f0f2f6; border-radius: 10px; text-align: left; border-left: 6px solid #1a237e;"><b style="color: #1a237e; font-size: 16px;">Novedades/Observaciones:</b><br><p style="margin: 8px 0 0; font-size: 15px; color: #000; line-height: 1.4;">{r['Observaciones']}</p></div></div></div><script>function descargarApertura() {{ html2canvas(document.getElementById('pizarra-apertura'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Apertura_Drotaca.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>"""

def html_personal_cierre(df):
    logo = obtener_logo_base64()
    area_logo = f'<img src="{logo}" style="max-height: 50px;">' if logo else '🛡️'
    filas = ""
    for i, area in enumerate(df['Area Asignada'].unique()):
        personal = df[df['Area Asignada'] == area]['Personal'].tolist()
        lista_personal = "<br>".join([f"✓ {p}" for p in personal])
        filas += f"<tr style='background-color: {'#ffffff' if i % 2 == 0 else '#f8f9fa'}; border-bottom: 1px solid #eee;'><td style='padding: 12px; font-weight: bold; color: #1a237e; vertical-align: top; width: 40%;'>{area}</td><td style='padding: 12px; color: #333; font-weight: bold; font-size: 15px;'>{lista_personal}</td></tr>"
    return f"""<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script><div style="text-align:right; margin-bottom:15px;"><button onclick="descargarPersonal()" style="background:#1a237e;color:#fff;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:14px;">⬇️ Descargar Asignación</button></div><div id="pizarra-personal" style="font-family: Arial, sans-serif; width: 600px; margin: auto; background-color: #fff; border: 2px solid #1a237e; border-radius: 10px; overflow: hidden;"><div style="background-color: #1a237e; color: white; padding: 20px; display: flex; align-items: center; justify-content: space-between;">{area_logo}<div style="text-align: right;"><h2 style="margin: 0; font-size: 20px;">ASIGNACIÓN DE CIERRE DROTACA</h2><p style="margin: 5px 0 0; font-size: 14px; opacity: 0.9;">Fecha: {df['Fecha'].iloc[0] if not df.empty else ''}</p></div></div><div style="padding: 20px;"><table style="width: 100%; border-collapse: collapse;"><thead><tr style="background-color: #e8eaf6; color: #1a237e; font-size: 14px; font-weight: bold;"><th style="padding: 10px; text-align: left;">ÁREA ASIGNADA</th><th style="padding: 10px; text-align: left;">PERSONAL DE GUARDIA</th></tr></thead><tbody>{filas}</tbody></table></div></div><script>function descargarPersonal() {{ html2canvas(document.getElementById('pizarra-personal'), {{scale: 2}}).then(canvas => {{ var link = document.createElement('a'); link.download = 'Personal_Cierre.png'; link.href = canvas.toDataURL(); link.click(); }}); }}</script>"""

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("🛡️ PCD - Control y Seguridad Integral")

tab1, tab2 = st.tabs(["📝 Módulo Operativo (Diario)", "📈 Módulo Gerencial (Reportes)"])

with tab1:
    tipo_reporte = st.radio("Reporte a procesar:", ["Rol Guardia Fin de Semana", "Rol Guardia Lunes", "Apertura Drotaca 2.0", "Personal Cierre Drotaca", "Cierre Drotaca 2.0", "Cierre Juanita"], horizontal=True)

    col1, col2 = st.columns([1, 1.5])
    with col1:
        if tipo_reporte in ["Rol Guardia Fin de Semana", "Rol Guardia Lunes"]:
            archivo_rol = st.file_uploader("📂 Sube el PDF del Rol de Guardia", type=["pdf"])
            procesar_seg = st.button("⚡ Procesar PDF", type="primary", use_container_width=True)
            texto_seguridad = None
        else:
            texto_seguridad = st.text_area("Pega WhatsApp aquí:", height=250)
            procesar_seg = st.button("⚡ Procesar Datos", type="primary", use_container_width=True)
            archivo_rol = None

    if procesar_seg:
        with st.spinner("Extrayendo información..."):
            if tipo_reporte in ["Rol Guardia Fin de Semana", "Rol Guardia Lunes"] and archivo_rol:
                df_seg, vac_info, status = extraer_rol_pdf(archivo_rol)
                if status == "FALTA_LIBRERIA": st.error("🚨 Falta la librería para leer PDFs. Ve a tu terminal y ejecuta:\n\n`pip install pdfplumber`")
                elif df_seg is not None and not df_seg.empty:
                    st.session_state['seg_tipo'], st.session_state['seg_data'] = "ROL_GUARDIA", df_seg
                    st.session_state['seg_titulo'], st.session_state['seg_meta'] = tipo_reporte, vac_info
                    st.session_state['ws_msg'] = generar_ws_rol_guardia(df_seg, tipo_reporte, vac_info)
                else: st.error(f"❌ Error leyendo el PDF. Detalle: {status}")
            elif texto_seguridad:
                if tipo_reporte == "Cierre Drotaca 2.0":
                    df_seg = extraer_cierre_drotaca(texto_seguridad)
                    st.session_state['seg_tipo'], st.session_state['seg_data'], st.session_state['ws_msg'] = "CIERRE_DROTACA", df_seg, generar_ws_cierre_drotaca(df_seg)
                elif tipo_reporte == "Cierre Juanita":
                    df_seg, h_cierre, oficiales = extraer_cierre_juanita(texto_seguridad)
                    st.session_state['seg_tipo'], st.session_state['seg_data'] = "CIERRE_JUANITA", df_seg
                    st.session_state['seg_meta'] = {"hora": h_cierre, "ofi": oficiales}
                    st.session_state['ws_msg'] = generar_ws_cierre_juanita(df_seg, h_cierre, oficiales)
                elif tipo_reporte == "Apertura Drotaca 2.0":
                    df_seg = extraer_apertura(texto_seguridad)
                    st.session_state['seg_tipo'], st.session_state['seg_data'], st.session_state['ws_msg'] = "APERTURA", df_seg, generar_ws_apertura(df_seg)
                elif tipo_reporte == "Personal Cierre Drotaca":
                    df_seg = extraer_personal_cierre(texto_seguridad)
                    st.session_state['seg_tipo'], st.session_state['seg_data'], st.session_state['ws_msg'] = "PERSONAL_CIERRE", df_seg, generar_ws_personal_cierre(df_seg)

    if st.session_state.get('seg_data') is not None:
        with col2:
            st.subheader("✏️ Revisión de Datos")
            df_editado = st.data_editor(st.session_state['seg_data'], num_rows="dynamic", use_container_width=True, hide_index=True)
            
            if st.button("💾 Guardar en Bitácora (Sheets)", use_container_width=True):
                with st.spinner("Guardando en Google Sheets..."):
                    
                    df_to_save = expandir_df_rol(df_editado.copy()) if st.session_state['seg_tipo'] == "ROL_GUARDIA" else df_editado.copy()
                        
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

                    if st.session_state['seg_tipo'] == "CIERRE_DROTACA":
                        c_esp = ['Fecha', 'Departamento', 'Hora Salida', 'Dia', 'Semana', 'Mes']
                        for c in c_esp:
                            if c not in df_to_save.columns: df_to_save[c] = ""
                        guardar_en_google_sheets(df_to_save[c_esp], "SEG_CIERRE_DROTACA")
                    elif st.session_state['seg_tipo'] == "CIERRE_JUANITA":
                        df_to_save['Oficiales'], df_to_save['Hora Cierre'] = ", ".join(st.session_state['seg_meta']['ofi']), st.session_state['seg_meta']['hora']
                        c_esp = ['Fecha', 'Grupo', 'Hora Salida', 'Oficiales', 'Hora Cierre', 'Dia', 'Semana', 'Mes']
                        for c in c_esp:
                            if c not in df_to_save.columns: df_to_save[c] = ""
                        guardar_en_google_sheets(df_to_save[c_esp], "SEG_CIERRE_JUANITA")
                    elif st.session_state['seg_tipo'] == "APERTURA":
                        c_esp = ['Fecha', 'Sede', 'Hora Apertura', 'Hora Alarma', 'Observaciones', 'Dia', 'Semana', 'Mes']
                        for c in c_esp:
                            if c not in df_to_save.columns: df_to_save[c] = ""
                        guardar_en_google_sheets(df_to_save[c_esp], "SEG_APERTURA")
                    elif st.session_state['seg_tipo'] == "PERSONAL_CIERRE":
                        c_esp = ['Fecha', 'Area Asignada', 'Personal', 'Dia', 'Semana', 'Mes']
                        for c in c_esp:
                            if c not in df_to_save.columns: df_to_save[c] = ""
                        guardar_en_google_sheets(df_to_save[c_esp], "SEG_PERSONAL_CIERRE")
                    elif st.session_state['seg_tipo'] == "ROL_GUARDIA":
                        c_esp = ['Fecha', 'Área', 'Diurno', 'Nocturno', 'Dia', 'Semana', 'Mes']
                        for c in c_esp:
                            if c not in df_to_save.columns: df_to_save[c] = ""
                        guardar_en_google_sheets(df_to_save[c_esp], "SEG_ROL_GUARDIA")
                    st.success("✅ Guardado exitoso.")
            
            st.markdown("---")
            st.subheader("📱 Mensaje WhatsApp")
            st.text_area("Copiar texto:", st.session_state.get('ws_msg', ''), height=180)

            st.markdown("---")
            st.subheader("📸 Generar Imagen")
            if st.session_state['seg_tipo'] == "CIERRE_DROTACA": components.html(html_cierre_drotaca(df_editado), height=300 + (len(df_editado) * 50), scrolling=True)
            elif st.session_state['seg_tipo'] == "CIERRE_JUANITA": components.html(html_cierre_juanita(df_editado, st.session_state['seg_meta']['hora'], st.session_state['seg_meta']['ofi']), height=350 + (len(df_editado) * 45), scrolling=True)
            elif st.session_state['seg_tipo'] == "APERTURA": components.html(html_apertura(df_editado), height=550, scrolling=True)
            elif st.session_state['seg_tipo'] == "PERSONAL_CIERRE": components.html(html_personal_cierre(df_editado), height=250 + (len(df_editado['Area Asignada'].unique()) * 60) + (len(df_editado) * 20), scrolling=True)
            elif st.session_state['seg_tipo'] == "ROL_GUARDIA": 
                fechas_unicas = len(df_editado['Fecha'].unique())
                alt_rol = 300 + (fechas_unicas * 100) + (len(df_editado) * 60)
                components.html(html_rol_guardia(df_editado, st.session_state['seg_titulo'], st.session_state.get('seg_meta', 'N/A')), height=alt_rol, scrolling=True)

# ---------------------------------------------------------
# PESTAÑA 2: MÓDULO GERENCIAL
# ---------------------------------------------------------
with tab2:
    st.header("📊 Inteligencia de Seguridad")
    if st.button("🔄 Sincronizar Base de Datos", type="primary"):
        with st.spinner("Conectando con Google Sheets y descargando datos..."):
            st.session_state['db_ape'] = extraer_datos_sheets("SEG_APERTURA")
            st.session_state['db_dro'] = extraer_datos_sheets("SEG_CIERRE_DROTACA")
            st.session_state['db_jua'] = extraer_datos_sheets("SEG_CIERRE_JUANITA")
            st.session_state['db_rol'] = extraer_datos_sheets("SEG_ROL_GUARDIA")
            st.session_state['db_sincronizada'] = True
            st.success("¡Base de Datos Sincronizada Exitosamente!")

    if st.session_state.get('db_sincronizada'):
        df_ape = st.session_state.get('db_ape', pd.DataFrame())
        df_dro = st.session_state.get('db_dro', pd.DataFrame())
        df_jua = st.session_state.get('db_jua', pd.DataFrame())
        df_rol = st.session_state.get('db_rol', pd.DataFrame())

        st.markdown("---")
        tipo_informe = st.radio("Seleccione el enfoque del reporte:", [
            "Reporte Semanal Detallado (Cierres)", 
            "Resumen Histórico Global",
            "Reporte Semanal de Guardias (PDF)"
        ], horizontal=True)
        
        if tipo_informe == "Reporte Semanal Detallado (Cierres)":
            st.subheader("📅 Generador de Reporte Semanal")
            semanas_disp, meses_disp = set(), set()
            for d in [df_ape, df_dro, df_jua]:
                if not d.empty:
                    if 'Semana' in d.columns: semanas_disp.update([str(x) for x in d['Semana'].dropna().unique() if str(x).strip()])
                    if 'Mes' in d.columns: meses_disp.update([str(x) for x in d['Mes'].dropna().unique() if str(x).strip()])
            
            c1, c2 = st.columns(2)
            filtro_mes = c1.selectbox("Filtrar por Mes:", ["Todos"] + sorted(list(meses_disp)))
            filtro_semana = c2.selectbox("Filtrar por Semana:", ["Todas"] + sorted(list(semanas_disp)) if semanas_disp else ["Sin datos"])
            
            if st.button("Generar Reporte de esta Semana"):
                with st.spinner("Analizando la semana seleccionada..."):
                    resumen_semanal = analizar_datos_semanales(df_ape, df_dro, df_jua, filtro_semana, filtro_mes)
                    col_res_txt, col_res_img = st.columns([1, 1.5])
                    with col_res_txt: st.text_area("📋 Texto para WhatsApp:", generar_texto_semanal(resumen_semanal, filtro_semana), height=500)
                    with col_res_img: components.html(html_reporte_semanal_pizarra(resumen_semanal, f"{filtro_semana.upper()}" if filtro_semana != "Todas" else "REPORTE GENERAL"), height=650, scrolling=True)

        elif tipo_informe == "Resumen Histórico Global":
            st.subheader("📈 Métricas Históricas Acumuladas")
            if not df_dro.empty and 'Hora Salida' in df_dro.columns:
                df_dro['Minutos'] = df_dro['Hora Salida'].apply(hora_a_minutos)
                df_d_valid = df_dro.dropna(subset=['Minutos'])
                if not df_d_valid.empty:
                    st.markdown("#### 🏢 Sede Principal Drotaca")
                    promedio_general = df_d_valid['Minutos'].mean()
                    top_tarde = df_d_valid.groupby('Departamento')['Minutos'].mean().reset_index().sort_values('Minutos', ascending=False).head(3)
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Cierre Promedio", minutos_a_hora(promedio_general)); m2.metric("Más Demorado", top_tarde.iloc[0]['Departamento']); m3.metric("Días Registrados", len(df_dro['Fecha'].unique()))
                    cuerpo_informe_dro = f"*INFORME GERENCIAL GLOBAL - DROTACA 2.0* 🏢\n\n⏱️ *Cierre Promedio Histórico:* {minutos_a_hora(promedio_general)}.\n\n🚨 *Departamentos con salidas más tardías:*\n"
                    for i, r_inf in enumerate(top_tarde.itertuples(), 1): cuerpo_informe_dro += f"{i}. {r_inf.Departamento}: {minutos_a_hora(r_inf.Minutos)}\n"
                    st.text_area("📋 Resumen Drotaca:", cuerpo_informe_dro, height=150)
            
            st.markdown("---")
            col_cierre = 'Hora Cierre' if 'Hora Cierre' in df_jua.columns else 'Hora Cierre Total'
            if not df_jua.empty and col_cierre in df_jua.columns:
                df_jua['Minutos_C'] = df_jua[col_cierre].apply(hora_a_minutos)
                df_j_valid = df_jua.dropna(subset=['Minutos_C'])
                if not df_j_valid.empty:
                    st.markdown("#### 📦 Almacén Juanita")
                    col_j1, col_j2 = st.columns(2)
                    col_j1.metric("Promedio Cierre Total", minutos_a_hora(df_j_valid['Minutos_C'].mean())); col_j2.metric("Días Analizados", len(df_jua['Fecha'].unique()))
                    st.text_area("📋 Resumen Juanita:", f"*INFORME GERENCIAL GLOBAL - ALMACÉN JUANITA* 📦\n\n⏱️ *Cierre Promedio Histórico:* {minutos_a_hora(df_j_valid['Minutos_C'].mean())}.\n\n", height=100)

        elif tipo_informe == "Reporte Semanal de Guardias (PDF)":
            st.subheader("📄 Generador de PDF Gerencial (Rol de Guardia)")
            if df_rol.empty:
                st.warning("No hay datos de guardias en la base de datos.")
            else:
                # EXPANSIÓN Y RE-CÁLCULO DINÁMICO PARA FILTROS PERFECTOS (CARACAS Y SÁBADO/DOMINGO)
                df_rol_exp = expandir_df_rol(df_rol.copy())
                
                dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                meses_ano = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                semanas_calc, meses_calc = [], []
                
                for f_str in df_rol_exp['Fecha']:
                    m = re.search(r'\d{2}/\d{2}/\d{4}', str(f_str))
                    if m:
                        try:
                            obj = datetime.strptime(m.group(), "%d/%m/%Y")
                            f_shift = obj + timedelta(days=2)
                            semanas_calc.append(f"Semana {f_shift.isocalendar()[1]}")
                            meses_calc.append(meses_ano[obj.month - 1])
                        except:
                            semanas_calc.append("")
                            meses_calc.append("")
                    else:
                        semanas_calc.append("")
                        meses_calc.append("")
                        
                df_rol_exp['Semana'] = semanas_calc
                df_rol_exp['Mes'] = meses_calc
                
                semanas_rol = sorted(list(set([str(x) for x in df_rol_exp['Semana'].dropna().unique() if str(x).strip()])))
                meses_rol = sorted(list(set([str(x) for x in df_rol_exp['Mes'].dropna().unique() if str(x).strip()])))
                
                c1, c2 = st.columns(2)
                filtro_mes_rol = c1.selectbox("Filtrar por Mes:", ["Todos"] + meses_rol, key='mes_rol')
                filtro_sem_rol = c2.selectbox("Filtrar por Semana:", ["Todas"] + semanas_rol if semanas_rol else ["Sin datos"], key='sem_rol')
                
                if st.button("Generar PDF de Guardias"):
                    with st.spinner("Armando el documento PDF continuo..."):
                        df_f = df_rol_exp.copy()
                        if filtro_mes_rol != "Todos": df_f = df_f[df_f['Mes'] == filtro_mes_rol]
                        if filtro_sem_rol != "Todas": df_f = df_f[df_f['Semana'] == filtro_sem_rol]
                        
                        if df_f.empty:
                            st.warning("No hay datos para los filtros seleccionados.")
                        else:
                            def sort_date(text):
                                m = re.search(r'\d{2}/\d{2}/\d{4}', str(text))
                                if m: return datetime.strptime(m.group(), "%d/%m/%Y")
                                return datetime.min
                            
                            df_f['s_key'] = df_f['Fecha'].apply(sort_date)
                            df_f = df_f.sort_values('s_key').drop(columns=['s_key'])
                            
                            titulo_pdf = f"Reporte {filtro_sem_rol} - {filtro_mes_rol}" if filtro_sem_rol != "Todas" else "Reporte General"
                            
                            st.info("💡 Haz clic en el botón rojo 'Descargar PDF Gerencial' para guardar el documento.")
                            components.html(html_reporte_rol_pdf(df_f, titulo_pdf, "N/A"), height=800, scrolling=True)