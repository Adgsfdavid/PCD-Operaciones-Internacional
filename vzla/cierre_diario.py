# ==========================================
# Archivo: cierre_diario.py (Reporte Master de Operaciones)
# ==========================================
import streamlit as st
import pandas as pd
import base64
from datetime import datetime, timedelta
from pathlib import Path
import streamlit.components.v1 as components
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Reporte Diario Master", layout="wide")

# ==========================================
# FUNCIONES GLOBALES Y DE GOOGLE SHEETS
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

def procesar_imagen_subida(uploaded_file):
    if uploaded_file is not None:
        uploaded_file.seek(0)
        return base64.b64encode(uploaded_file.read()).decode()
    return None

# ==========================================
# CREDENCIALES DIRECTAS DE GOOGLE
# ==========================================
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])

ddef guardar_en_sheets(nombre_hoja, df):
    try:
        alcance = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # --- AQUÍ ESTÁ LA CORRECCIÓN CLAVE ---
        credenciales = ServiceAccountCredentials.from_json_keyfile_dict(CREDENCIALES_GOOGLE, alcance)
        cliente = gspread.authorize(credenciales)
        
        # --- LÓGICA DE BASE DE DATOS ---
        user_data = st.session_state.get("user_data", {})
        nombre_bd = user_data.get("sheet_name", "PCD_BaseDatos")
        doc = cliente.open(nombre_bd)
        
        try:
            hoja = doc.worksheet(nombre_hoja)
        except gspread.exceptions.WorksheetNotFound:
            return False, f"La hoja '{nombre_hoja}' no existe en tu archivo {nombre_bd}. ¡Créala primero!"
            
        df_guardar = df.copy().astype(str)
        
        # Opcional: Agregar marca de tiempo si no existe en el DataFrame original
        if "Fecha Sistema" not in df_guardar.columns:
            from datetime import datetime
            df_guardar.insert(0, "Fecha Sistema", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        
        hoja.append_rows(df_guardar.values.tolist())
        return True, "Datos guardados en la nube exitosamente."
    except Exception as e:
        return False, str(e)

# ==========================================
# INTERFAZ DE USUARIO: PESTAÑAS PRINCIPALES
# ==========================================
st.title("📚 Reporte Master de Operaciones Diarias")
st.markdown("Consolida las pizarras de tu guardia, gestión de comensales y guardias semanales en un único PDF gerencial.")

col_f, col_d = st.columns(2)
with col_f:
    fecha_cierre = st.date_input("📅 Fecha de la Bitácora:", datetime.now())
    fecha_str = fecha_cierre.strftime("%d/%m/%Y")
    
    # Cálculo automático del rango semanal (Lunes a Domingo)
    fecha_fin = fecha_cierre + timedelta(days=6)
    rango_semana_str = f"Semana del {fecha_str} al {fecha_fin.strftime('%d/%m/%Y')}"
    
with col_d:
    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    dia_actual = dias_semana[fecha_cierre.weekday()]
    es_lunes = (fecha_cierre.weekday() == 0)
    
    if es_lunes:
        st.info(f"**Día Operativo:** {dia_actual} (¡Día de Cargar Guardias Semanales!)")
    else:
        st.info(f"**Día Operativo:** {dia_actual}")

st.markdown("---")

tab_cierre, tab_comensales, tab_flota, tab_monitoreo = st.tabs([
    "📋 Control de Cierre Diario", 
    "🍽️ Pizarra de Comensales",
    "🚛 Guardia de Flota",
    "🖥️ Guardia de Monitoreo"
])

# ==========================================
# MÓDULO 2: PIZARRA DE COMENSALES
# ==========================================
with tab_comensales:
    st.header("🍽️ Gestión y Generador de Pizarra: Comensales")
    
    if 'df_comensales' not in st.session_state:
        st.session_state['df_comensales'] = pd.DataFrame({
            "Departamento": ["ADMINISTRACION", "OPERACIONES", "EXTERNOS"],
            "Desayuno": [0, 0, 0], "Almuerzo": [0, 0, 0], "Cena": [0, 0, 0]
        })
    
    col_p1, col_p2 = st.columns([1, 2])
    with col_p1:
        st.markdown("### 📥 1. Pegar Datos")
        texto_pegado = st.text_area("Pega aquí el listado de Excel:", height=200, key="txt_com")
        if st.button("🔄 Procesar Datos", use_container_width=True, key="btn_p_com"):
            if texto_pegado:
                lineas = texto_pegado.strip().split('\n')
                datos_limpios = []
                leyendo = False
                for linea in lineas:
                    if not linea.strip(): continue
                    if "DEPARTAMENTO" in linea.upper() or "DESAYUNO" in linea.upper():
                        leyendo = True
                        continue
                    if "TOTAL" in linea.upper() or "TOTAL COMIDAS" in linea.upper(): break
                    if leyendo:
                        partes = linea.split('\t')
                        dep = partes[0].strip() if len(partes) > 0 else ""
                        if dep and dep.upper() != "ADICIONALES":
                            def s_int(v):
                                try: return int(v.strip())
                                except: return 0
                            datos_limpios.append({"Departamento": dep, "Desayuno": s_int(partes[1]) if len(partes)>1 else 0, "Almuerzo": s_int(partes[2]) if len(partes)>2 else 0, "Cena": s_int(partes[3]) if len(partes)>3 else 0})
                if datos_limpios:
                    st.session_state['df_comensales'] = pd.DataFrame(datos_limpios)
                    st.rerun()
                    
    with col_p2:
        st.markdown("### ✏️ 2. Pizarra Digital Editable")
        fecha_pizarra = st.text_input("📅 Fecha", value=fecha_str, key="f_com")
        df_editado_com = st.data_editor(st.session_state['df_comensales'], num_rows="dynamic", use_container_width=True, hide_index=True, key="edit_com")
        
        t_des = pd.to_numeric(df_editado_com["Desayuno"], errors='coerce').fillna(0).sum()
        t_alm = pd.to_numeric(df_editado_com["Almuerzo"], errors='coerce').fillna(0).sum()
        t_cen = pd.to_numeric(df_editado_com["Cena"], errors='coerce').fillna(0).sum()
        t_gen = t_des + t_alm + t_cen
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("☕ Desayunos", int(t_des)); c2.metric("🍲 Almuerzos", int(t_alm)); c3.metric("🥪 Cenas", int(t_cen)); c4.metric("🔥 TOTAL", int(t_gen))

        c_btn1, c_btn2, c_btn3 = st.columns(3)
        with c_btn1:
            if st.button("📱 WhatsApp", use_container_width=True, key="w_com"):
                if t_gen > 0:
                    msg = f"📊 *PIZARRA DE COMENSALES*\n📅 *Fecha:* {fecha_pizarra}\n\n"
                    for _, r in df_editado_com.iterrows():
                        if (pd.to_numeric(r['Desayuno'], errors='coerce').fillna(0) + pd.to_numeric(r['Almuerzo'], errors='coerce').fillna(0) + pd.to_numeric(r['Cena'], errors='coerce').fillna(0)) > 0:
                            msg += f"🏢 *{r['Departamento']}*\n"
                            if r['Desayuno'] > 0: msg += f"   ☕ Desayuno: {int(r['Desayuno'])}\n"
                            if r['Almuerzo'] > 0: msg += f"   🍲 Almuerzo: {int(r['Almuerzo'])}\n"
                            if r['Cena'] > 0: msg += f"   🥪 Cena: {int(r['Cena'])}\n"
                    msg += f"\n🔥 *TOTAL GENERAL:* {int(t_gen)} platos"
                    st.code(msg, language="markdown")
        with c_btn2:
            if st.button("💾 Integrar a PDF", type="primary", use_container_width=True, key="i_com"):
                if t_gen > 0:
                    st.session_state['comensales_guardados'] = df_editado_com
                    st.session_state['total_comensales'] = t_gen
                    st.success("✅ Listo para el PDF.")
        with c_btn3:
            if st.button("☁️ Guardar en Sheets", use_container_width=True):
                if not df_editado_com.empty:
                    with st.spinner("Guardando..."):
                        exito, msj = guardar_en_sheets("PIZARRA_COMENSALES", df_editado_com)
                        if exito: st.success(msj)
                        else: st.error(msj)

    # --- Pizarra Gráfica Comensales ---
    st.markdown("---")
    filas_html_com = ""
    for _, row in df_editado_com.iterrows():
        d = int(pd.to_numeric(row['Desayuno'], errors='coerce')) if pd.notna(row['Desayuno']) else 0
        a = int(pd.to_numeric(row['Almuerzo'], errors='coerce')) if pd.notna(row['Almuerzo']) else 0
        c = int(pd.to_numeric(row['Cena'], errors='coerce')) if pd.notna(row['Cena']) else 0
        if (d+a+c) > 0:
            filas_html_com += f"<tr><td style='padding:10px; border:1px solid #000; font-weight:bold; font-size:15px;'>{str(row['Departamento']).upper()}</td><td style='padding:10px; border:1px solid #000; text-align:center; font-size:18px; font-weight:900;'>{str(d) if d>0 else ''}</td><td style='padding:10px; border:1px solid #000; text-align:center; font-size:18px; font-weight:900;'>{str(a) if a>0 else ''}</td><td style='padding:10px; border:1px solid #000; text-align:center; font-size:18px; font-weight:900;'>{str(c) if c>0 else ''}</td></tr>"

    logo_b64 = obtener_logo_base64()
    img_logo = f'<img src="{logo_b64}" alt="Logo" style="height: 50px; object-fit: contain;">' if logo_b64 else '<h2 style="margin: 0; color: white;">DROTACA</h2>'
    
    html_pizarra_com = f"""
    <div id="piz-com" style="background:white; border:2px solid #ccc; border-radius:12px; width:750px; margin:auto; box-shadow:0px 10px 30px rgba(0,0,0,0.3); font-family:sans-serif;">
        <div style="background-color:#0d47a1; padding:20px 40px; color:white;">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(255,255,255,0.3); padding-bottom:10px; margin-bottom:10px;">{img_logo}<span style="color:#ffeb3b; font-weight:bold; font-style:italic;">¡Conectamos con la salud!</span></div>
            <h1 style="text-align:center; margin:0; font-size:28px;">PIZARRA DE COMENSALES</h1>
            <div style="text-align:center; font-size:16px; color:#ffeb3b; font-weight:bold;">FECHA: {fecha_pizarra}</div>
        </div>
        <div style="padding:20px 40px;">
            <table style="width:100%; border-collapse:collapse; margin-bottom:20px;">
                <thead><tr><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:left;">DEPARTAMENTOS</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white;">DESAYUNO</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white;">ALMUERZO</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white;">CENA</th></tr></thead>
                <tbody>{filas_html_com}</tbody>
            </table>
            <div style="border-top:3px solid #0d47a1; padding-top:15px; display:flex; justify-content:space-around; font-weight:bold; color:#333;">
                <span>DESAYUNOS: <span style="color:#0d47a1; font-size:20px;">{int(t_des)}</span></span>
                <span>ALMUERZOS: <span style="color:#0d47a1; font-size:20px;">{int(t_alm)}</span></span>
                <span>CENAS: <span style="color:#0d47a1; font-size:20px;">{int(t_cen)}</span></span>
            </div>
            <h2 style="text-align:center; margin:15px 0 0 0; font-size:28px; background-color:#f8f9fa; padding:10px; border:2px solid #0d47a1; border-radius:8px;">TOTAL COMIDAS: <span style="color:#e65100;">{int(t_gen)}</span></h2>
            <div style="margin-top:20px; text-align:center; font-size:12px; color:#666; font-weight:bold;">DEPARTAMENTO DE SERVICIOS GENERALES</div>
        </div>
    </div>
    """
    
    if t_gen > 0:
        st.components.v1.html(f"""
        <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head>
        <body style="margin:0; padding:0;">
            <div style="text-align:center; margin-bottom:15px;"><button onclick="d()" style="background:#28a745; color:white; border:none; padding:10px 20px; font-weight:bold; cursor:pointer; border-radius:5px;">📸 DESCARGAR PIZARRA</button></div>
            {html_pizarra_com}
            <script>function d(){{ html2canvas(document.getElementById('piz-com'), {{scale:2}}).then(c=>{{let a=document.createElement('a'); a.download='Comensales.png'; a.href=c.toDataURL('image/png'); a.click();}}); }}</script>
        </body></html>
        """, height=900, scrolling=True)


# ==========================================
# MÓDULO 3: GUARDIA DE FLOTA (Lógica Inteligente)
# ==========================================
with tab_flota:
    st.header("🚛 Guardia de Flota")
    st.write("Pega el texto de WhatsApp. El sistema asignará automáticamente los días según el cargo.")
    
    if 'df_flota' not in st.session_state:
        st.session_state['df_flota'] = pd.DataFrame({"Nombre": [], "Cargo y Turno": [], "Días": [], "Horario": []})
        
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        txt_flota = st.text_area("Pega listado de Flota:", height=150, placeholder="Javier Hidalgo, Supervisor Diurno (5AM / 5PM)")
        if st.button("🔄 Procesar Flota", use_container_width=True):
            if txt_flota:
                datos_f = []
                for linea in txt_flota.strip().split('\n'):
                    if not linea.strip(): continue
                    partes = linea.split(',')
                    if len(partes) >= 2:
                        nombre = partes[0].strip()
                        resto = partes[1].split('(')
                        cargo = resto[0].strip()
                        horario = resto[1].replace(')','').strip() if len(resto)>1 else "No definido"
                        
                        # --- LÓGICA INTELIGENTE DE DÍAS ---
                        cargo_up = cargo.upper()
                        dias = "No definido"
                        
                        if "NOCTURNO" in cargo_up or "RECORRIDO" in cargo_up:
                            dias = "Lunes a Viernes"
                        elif "DIURNO" in cargo_up:
                            dias = "Lunes a Domingo"
                        elif "PATIO" in cargo_up:
                            dias = "Lunes a Sábado"
                            # Ajuste de horario inteligente para Patio
                            horario = f"L-V: {horario.split('/')[0].strip()} / {horario.split('/')[1].strip()} | Sáb: 8AM / 12PM" if '/' in horario else "L-V: 8AM/6PM | Sáb: 8AM/12PM"
                            
                        datos_f.append({"Nombre": nombre, "Cargo y Turno": cargo, "Días": dias, "Horario": horario})
                    else:
                        datos_f.append({"Nombre": linea.strip(), "Cargo y Turno": "", "Días": "", "Horario": ""})
                st.session_state['df_flota'] = pd.DataFrame(datos_f)
                st.rerun()

    with col_f2:
        st.markdown("### ✏️ Pizarra Flota Editable")
        fecha_flota = st.text_input("📅 Rango de Semana:", value=rango_semana_str)
        df_edit_flota = st.data_editor(st.session_state['df_flota'], num_rows="dynamic", use_container_width=True, hide_index=True)
        
        cf1, cf2 = st.columns(2)
        with cf1:
            if st.button("📱 WhatsApp Flota", use_container_width=True):
                msg = f"🚛 *GUARDIA DE FLOTA*\n📅 *{fecha_flota}*\n\n"
                for _, r in df_edit_flota.iterrows():
                    msg += f"👤 *{r['Nombre']}*\n🔹 Rol: {r['Cargo y Turno']}\n🗓️ Días: {r['Días']}\n⏰ Horario: {r['Horario']}\n\n"
                st.code(msg, language="markdown")
        with cf2:
            if st.button("☁️ Guardar en Sheets", use_container_width=True, key="gs_flota"):
                if not df_edit_flota.empty:
                    with st.spinner("Guardando..."):
                        exito, msj = guardar_en_sheets("GUARDIA_FLOTA", df_edit_flota)
                        if exito: st.success(msj)
                        else: st.error(msj)

    # --- Pizarra Gráfica Flota ---
    st.markdown("---")
    filas_html_f = ""
    for _, row in df_edit_flota.iterrows():
        # COLOR NEGRO (#000000) APLICADO EN EL HORARIO
        filas_html_f += f"<tr><td style='padding:12px; border:1px solid #000; font-weight:bold; font-size:15px;'>{row['Nombre']}</td><td style='padding:12px; border:1px solid #000; font-size:14px;'>{row['Cargo y Turno']}</td><td style='padding:12px; border:1px solid #000; text-align:center; font-size:14px; font-weight:bold; color:#0d47a1;'>{row['Días']}</td><td style='padding:12px; border:1px solid #000; text-align:center; font-size:14px; font-weight:bold; color:#000000;'>{row['Horario']}</td></tr>"

    html_pizarra_flota = f"""
    <div id="piz-flota" style="background:white; border:2px solid #ccc; border-radius:12px; width:800px; margin:auto; box-shadow:0px 10px 30px rgba(0,0,0,0.3); font-family:sans-serif;">
        <div style="background-color:#0d47a1; padding:20px 40px; color:white;">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(255,255,255,0.3); padding-bottom:10px; margin-bottom:10px;">{img_logo}<span style="color:#ffeb3b; font-weight:bold; font-style:italic;">¡Conectamos con la salud!</span></div>
            <h1 style="text-align:center; margin:0; font-size:28px;">GUARDIA DE FLOTA</h1>
            <div style="text-align:center; font-size:16px; color:#ffeb3b; font-weight:bold;">{fecha_flota}</div>
        </div>
        <div style="padding:20px 40px;">
            <table style="width:100%; border-collapse:collapse; margin-bottom:20px;">
                <thead><tr><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:left;">PERSONAL</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:left;">CARGO Y TURNO</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:center;">DÍAS</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:center;">HORARIO</th></tr></thead>
                <tbody>{filas_html_f}</tbody>
            </table>
            <div style="margin-top:20px; text-align:center; font-size:12px; color:#666; font-weight:bold;">DEPARTAMENTO DE FLOTA Y LOGÍSTICA</div>
        </div>
    </div>
    """
    
    if not df_edit_flota.empty:
        st.components.v1.html(f"""
        <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head>
        <body style="margin:0; padding:0;">
            <div style="text-align:center; margin-bottom:15px;"><button onclick="d()" style="background:#28a745; color:white; border:none; padding:10px 20px; font-weight:bold; cursor:pointer; border-radius:5px;">📸 DESCARGAR PIZARRA</button></div>
            {html_pizarra_flota}
            <script>function d(){{ html2canvas(document.getElementById('piz-flota'), {{scale:2}}).then(c=>{{let a=document.createElement('a'); a.download='Flota.png'; a.href=c.toDataURL('image/png'); a.click();}}); }}</script>
        </body></html>
        """, height=700, scrolling=True)


# ==========================================
# MÓDULO 4: GUARDIA DE MONITOREO 
# ==========================================
with tab_monitoreo:
    st.header("🖥️ Guardia de Monitoreo")
    st.write("Estructura el personal de monitoreo de la semana.")
    
    if 'df_mon' not in st.session_state:
        st.session_state['df_mon'] = pd.DataFrame({"Nombre y Turno": ["Pedro Molina (Madrugada)", "Gabriel Vera (Sábado)", "Jesus Brito (Domingo)"], "Horario y Días": ["05:45 a.m. – 08:00 a.m. (L a V)", "05:45 a.m. – 06:00 p.m.", "05:45 a.m. – 07:00 p.m."]})
    
    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        st.markdown("### 📥 1. Pegar Datos (Opcional)")
        txt_mon = st.text_area("Texto rápido:", height=150)
        st.caption("Para monitoreo suele ser más fácil editar la tabla de la derecha directamente.")
        
    with col_m2:
        st.markdown("### ✏️ Pizarra Monitoreo")
        fecha_mon = st.text_input("📅 Semana Monitoreo:", value=rango_semana_str)
        df_edit_mon = st.data_editor(st.session_state['df_mon'], num_rows="dynamic", use_container_width=True, hide_index=True)
        unidades_ramon = st.text_input("🚛 Unidades del Sr. Ramón (Responsable):", value="Carlos Carneiro")
        
        cm1, cm2 = st.columns(2)
        with cm1:
            if st.button("📱 WhatsApp Monitoreo", use_container_width=True):
                msg = f"🖥️ *GUARDIA DE MONITOREO*\n📅 *{fecha_mon}*\n\n🕒 *Guardias Programadas:*\n\n"
                for _, r in df_edit_mon.iterrows():
                    msg += f"👤 *{r['Nombre y Turno']}*\n⏰ {r['Horario y Días']}\n\n"
                msg += f"🚛 *Unidades del Sr. Ramón*\nResponsable: {unidades_ramon}"
                st.code(msg, language="markdown")
        with cm2:
            if st.button("☁️ Guardar en Sheets", use_container_width=True, key="gs_mon"):
                if not df_edit_mon.empty:
                    df_guardar = df_edit_mon.copy()
                    df_guardar['Unidades Sr Ramon'] = unidades_ramon # Agregamos el responsable extra a la hoja
                    with st.spinner("Guardando..."):
                        exito, msj = guardar_en_sheets("GUARDIA_MONITOREO", df_guardar)
                        if exito: st.success(msj)
                        else: st.error(msj)

    # --- Pizarra Gráfica Monitoreo ---
    st.markdown("---")
    filas_html_m = ""
    for _, row in df_edit_mon.iterrows():
        filas_html_m += f"<tr><td style='padding:12px; border:1px solid #000; font-weight:bold; font-size:15px; color:#0d47a1;'>{row['Nombre y Turno']}</td><td style='padding:12px; border:1px solid #000; font-size:14px; font-weight:bold;'>{row['Horario y Días']}</td></tr>"

    html_pizarra_mon = f"""
    <div id="piz-mon" style="background:white; border:2px solid #ccc; border-radius:12px; width:750px; margin:auto; box-shadow:0px 10px 30px rgba(0,0,0,0.3); font-family:sans-serif;">
        <div style="background-color:#0d47a1; padding:20px 40px; color:white;">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(255,255,255,0.3); padding-bottom:10px; margin-bottom:10px;">{img_logo}<span style="color:#ffeb3b; font-weight:bold; font-style:italic;">¡Conectamos con la salud!</span></div>
            <h1 style="text-align:center; margin:0; font-size:28px;">GUARDIA DE MONITOREO</h1>
            <div style="text-align:center; font-size:16px; color:#ffeb3b; font-weight:bold;">{fecha_mon}</div>
        </div>
        <div style="padding:20px 40px;">
            <table style="width:100%; border-collapse:collapse; margin-bottom:20px;">
                <thead><tr><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:left;">PERSONAL (TURNO)</th><th style="border:2px solid #000; padding:10px; background-color:#0d47a1; color:white; text-align:left;">HORARIO Y DÍAS</th></tr></thead>
                <tbody>{filas_html_m}</tbody>
            </table>
            
            <div style="background-color:#f8f9fa; padding:15px; border-left:5px solid #e65100; margin-top:20px; border-radius:5px;">
                <h4 style="margin:0 0 5px 0; color:#333;">🚛 Unidades del Sr. Ramón</h4>
                <span style="font-size:15px; font-weight:bold; color:#0d47a1;">Responsable: {unidades_ramon}</span>
            </div>
            
            <div style="margin-top:20px; text-align:center; font-size:12px; color:#666; font-weight:bold;">COORDINACIÓN DE MONITOREO Y FLOTA</div>
        </div>
    </div>
    """
    
    if not df_edit_mon.empty:
        st.components.v1.html(f"""
        <html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script></head>
        <body style="margin:0; padding:0;">
            <div style="text-align:center; margin-bottom:15px;"><button onclick="d()" style="background:#28a745; color:white; border:none; padding:10px 20px; font-weight:bold; cursor:pointer; border-radius:5px;">📸 DESCARGAR PIZARRA</button></div>
            {html_pizarra_mon}
            <script>function d(){{ html2canvas(document.getElementById('piz-mon'), {{scale:2}}).then(c=>{{let a=document.createElement('a'); a.download='Monitoreo.png'; a.href=c.toDataURL('image/png'); a.click();}}); }}</script>
        </body></html>
        """, height=700, scrolling=True)


# ==========================================
# MÓDULO 1: CONTROL DE CIERRE DIARIO (EL PDF MASTER)
# ==========================================
with tab_cierre:
    
    # --- LISTAS DINÁMICAS (FUSIÓN FLOTA/GPS) ---
    procesos_lunes = []
    if es_lunes:
        procesos_lunes = [
            "L1. Guardia De Seguridad",
            "L2. Guardia De Flota",
            "L3. Guardia De Monitoreo",
            "L4. Pizarra De Menú Semanal"
        ]

    # NOTA: Ahora usamos el nombre Dinámico
    procesos_manana = [
        "2. Apertura Drotaca 2.0",
        "3. Mantenimiento Preventivo Vehicular a Realizar",
        "4. Pizarra Estatus (Dinámica)",
        "5. Pizarra de Despachos de Rutas del Día Anterior",
        "6. Pizarra Salida de Rutas",
        "7. Pizarra Tráfico (Transbordo + Oriente)"
    ]

    procesos_tarde = [
        "8. Reporte de Personal de Guardia Drotaca 2.0",
        "9. Cierre Juanita",
        "10. Cierre Drotaca 2.0",
        "11. Resultado de Mantenimiento Preventivo Realizado",
        "12. Pizarra de Control de Reserva de Combustible (El Tigre)",
        "13. Pizarra de Resultado de Surtido / Extracción / Ruta Larga y Corta"
    ]

    imagenes_b64 = {}
    total_procesos = len(procesos_lunes) + len(procesos_manana) + len(procesos_tarde)
    procesos_cargados = 0

    # --- SECCIÓN ESPECIAL LUNES ---
    if es_lunes:
        st.markdown("<div style='background-color:#fff3e0; padding:15px; border-left:5px solid #e65100; border-radius:5px;'>", unsafe_allow_html=True)
        st.subheader("📅 GUARDIAS SEMANALES (Solo Lunes)")
        st.write("Sube las imágenes de las guardias generadas en las pestañas contiguas y el menú semanal.")
        cols_l = st.columns(2)
        for i, proceso in enumerate(procesos_lunes):
            with cols_l[i % 2]:
                archivo = st.file_uploader(f"📎 {proceso}", type=['png', 'jpg', 'jpeg'], key=f"l_{i}")
                if archivo:
                    imagenes_b64[proceso] = procesar_imagen_subida(archivo)
                    procesos_cargados += 1
                    st.success("✅ Cargado")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    st.subheader("☀️ TURNO DE MAÑANA (Apertura y Control)")
    cols_m = st.columns(2)
    for i, proceso in enumerate(procesos_manana):
        with cols_m[i % 2]:
            # AQUI: Lógica Dinámica para la Pizarra #4
            if proceso == "4. Pizarra Estatus (Dinámica)":
                tipo_piz_4 = st.radio("Selecciona qué contiene la Pizarra #4:", ["Ambas", "Solo Flota", "Solo GPS"], horizontal=True, key="rad_piz_4")
                nombre_label = f"📎 4. Estatus ({tipo_piz_4})"
                archivo = st.file_uploader(nombre_label, type=['png', 'jpg', 'jpeg'], key=f"m_{i}")
                if archivo:
                    imagenes_b64[proceso] = procesar_imagen_subida(archivo)
                    procesos_cargados += 1
                    st.success("✅ Cargado")
            else:
                archivo = st.file_uploader(f"📎 {proceso}", type=['png', 'jpg', 'jpeg'], key=f"m_{i}")
                if archivo:
                    imagenes_b64[proceso] = procesar_imagen_subida(archivo)
                    procesos_cargados += 1
                    st.success("✅ Cargado")

    st.markdown("---")

    st.subheader("🌙 TURNO DE TARDE (Cierre y Resultados)")
    cols_t = st.columns(2)
    for i, proceso in enumerate(procesos_tarde):
        with cols_t[i % 2]:
            archivo = st.file_uploader(f"📎 {proceso}", type=['png', 'jpg', 'jpeg'], key=f"t_{i}")
            if archivo:
                imagenes_b64[proceso] = procesar_imagen_subida(archivo)
                procesos_cargados += 1
                st.success("✅ Cargado")

    # --- SECCIÓN: SMART FORM - RESUMEN EJECUTIVO FINAL ---
    st.markdown("---")
    st.subheader("📝 Datos para el Resumen Ejecutivo Final")

    with st.expander("Desplegar Formulario de Resumen", expanded=True):
        c_h1, c_h2, c_h3 = st.columns(3)
        hora_apertura = c_h1.text_input("Apertura Drotaca", "07:00 AM")
        hora_cierre_dro = c_h2.text_input("Cierre Drotaca", "11:00 PM")
        hora_cierre_jua = c_h3.text_input("Cierre Juanita", "08:00 PM")

        c_m1, c_m2 = st.columns(2)
        mant_plan = c_m1.number_input("Mantenimientos Planificados", min_value=0, value=3)
        mant_real = c_m2.number_input("Mantenimientos Realizados", min_value=0, value=3)

        c_d1, c_d2, c_d3 = st.columns(3)
        farmacias = c_d1.number_input("Total Farmacias Entregadas", min_value=0, value=455)
        bultos = c_d2.number_input("Total Bultos Entregados", min_value=0, value=800)
        kms = c_d3.number_input("Kms Recorridos (Larga y Corta)", min_value=0, value=5585)

        st.markdown("**Combustible (Ciudad Drotaca)**")
        c_c1, c_c2 = st.columns(2)
        gasoil_lts = c_c1.number_input("Total Gasoil Disponible (Lts)", min_value=0, value=38500)
        gasoil_dias = c_c2.number_input("Autonomía (Días Operativos)", min_value=0, value=25)
        
        # MOSTRAR INTEGRACIÓN DE COMENSALES
        total_com_vinc = st.session_state.get('total_comensales', 0)
        if total_com_vinc > 0:
            st.success(f"🍽️ **Dato Vinculado Activo:** Se integrarán {int(total_com_vinc)} comensales en las métricas del PDF.")

    # --- SECCIÓN: BARRA DE PROGRESO ---
    st.markdown("---")
    st.subheader("📊 Estatus del Cierre Diario")
    progreso = procesos_cargados / total_procesos if total_procesos > 0 else 0
    st.progress(progreso)

    if procesos_cargados == total_procesos:
        st.success(f"🎉 ¡Excelente! Has cargado {procesos_cargados} de {total_procesos} procesos (100%).")
    else:
        st.warning(f"⏳ Llevas {procesos_cargados} de {total_procesos} procesos. (Las pizarras vacías serán omitidas en el PDF).")

    # ==========================================
    # MOTOR GENERADOR DEL PDF MASTER (CON PASADA ÚNICA ACTUALIZADA)
    # ==========================================
    if st.button("🖨️ GENERAR REPORTE MASTER PDF", type="primary", use_container_width=True):
        if procesos_cargados == 0 and 'comensales_guardados' not in st.session_state:
            st.error("Debes cargar al menos 1 pizarra o proceso para poder generar el reporte.")
        else:
            with st.spinner("Ensamblando Reporte Master y calculando paginación..."):
                logo_b64 = obtener_logo_base64()
                logo_html = f'<img src="{logo_b64}" style="max-height: 120px; filter: drop-shadow(0px 0px 10px rgba(0,0,0,0.5));">' if logo_b64 else '<h1 style="color:white; filter: drop-shadow(0px 0px 5px rgba(0,0,0,0.5));">DROTACA</h1>'
                img_portada = "https://images.unsplash.com/photo-1601584115197-04ecc0da31d7?q=80&w=2070&auto=format&fit=crop"
                
                color_primario = "#0d47a1" 
                
                procesos_activos = [p for p in (procesos_lunes + procesos_manana + procesos_tarde) if p in imagenes_b64]
                
                if 'comensales_guardados' in st.session_state:
                    procesos_activos.append("14. Pizarra de Comensales")
                
                html_indice_items = ""
                html_paginas_contenido = ""
                current_page = 3 # Pág 1: Portada, Pág 2: Índice
                ya_procesado = set()
                
                lista_completa = procesos_lunes + procesos_manana + procesos_tarde
                if 'comensales_guardados' in st.session_state:
                    lista_completa.append("14. Pizarra de Comensales")
                    
                for proc in lista_completa:
                    if proc not in procesos_activos or proc in ya_procesado:
                        continue
                        
                    # CASO 1: Mantenimientos Fusionados
                    if "3. Mantenimiento" in proc or "11. Resultado" in proc:
                        p3 = "3. Mantenimiento Preventivo Vehicular a Realizar"
                        p11 = "11. Resultado de Mantenimiento Preventivo Realizado"
                        
                        if p3 in procesos_activos:
                            html_indice_items += f'<tr><td style="padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 13px; color: #333;">{p3}</td><td style="padding: 8px 10px; font-size: 13px; font-weight: bold; color: {color_primario}; text-align: right;">Pág. {current_page}</td></tr>'
                            ya_procesado.add(p3)
                        if p11 in procesos_activos:
                            html_indice_items += f'<tr><td style="padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 13px; color: #333;">{p11}</td><td style="padding: 8px 10px; font-size: 13px; font-weight: bold; color: {color_primario}; text-align: right;">Pág. {current_page}</td></tr>'
                            ya_procesado.add(p11)
                            
                        img3 = imagenes_b64.get(p3)
                        img11 = imagenes_b64.get(p11)
                        b_img3 = f'<div style="flex:1; padding: 5px; border: 2px solid {color_primario}; text-align: center; display: flex; flex-direction: column; justify-content: center;"><h3 style="margin: 0 0 5px 0; font-size: 14px; color: {color_primario};">PLANIFICADO</h3><img src="data:image/png;base64,{img3}" style="max-height: 330px; max-width: 100%; object-fit: contain;"></div>' if img3 else ""
                        b_img11 = f'<div style="flex:1; padding: 5px; border: 2px solid #2e7d32; text-align: center; display: flex; flex-direction: column; justify-content: center; margin-top: 10px;"><h3 style="margin: 0 0 5px 0; font-size: 14px; color: #2e7d32;">REALIZADO</h3><img src="data:image/png;base64,{img11}" style="max-height: 330px; max-width: 100%; object-fit: contain;"></div>' if img11 else ""
                        
                        html_paginas_contenido += f"""
                        <div class="page">
                            <div class="header-band">
                                <span style="font-weight: bold;">Reporte de Operaciones Diarias</span>
                                <span>Fecha: {fecha_str}</span>
                            </div>
                            <h2 style="color: {color_primario}; text-align: center; font-size: 20px; margin-top: 20px; text-transform: uppercase;">CONTROL DE MANTENIMIENTO PREVENTIVO</h2>
                            <div style="padding: 0 40px; height: 230mm; display: flex; flex-direction: column;">
                                {b_img3}
                                {b_img11}
                            </div>
                            <div class="footer-band">
                                <span style="float: left; padding-left: 20px;">Departamento de Flota y Logística</span>
                                <span style="float: right; padding-right: 20px;">Página {current_page}</span>
                                <div style="clear: both;"></div>
                            </div>
                        </div>
                        """
                        current_page += 1
                        
                    # CASO 2: Comensales Integrados
                    elif proc == "14. Pizarra de Comensales":
                        html_indice_items += f'<tr><td style="padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 13px; color: #333;">{proc}</td><td style="padding: 8px 10px; font-size: 13px; font-weight: bold; color: {color_primario}; text-align: right;">Pág. {current_page}</td></tr>'
                        ya_procesado.add(proc)
                        
                        df_com = st.session_state['comensales_guardados']
                        tabla_pdf_html = f"""<table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px;"><thead><tr style="background-color: {color_primario}; color: white;"><th style="padding: 12px; border: 1px solid #ccc; text-align: left;">Departamento / Área</th><th style="padding: 12px; border: 1px solid #ccc; text-align: center;">☕ Desayuno</th><th style="padding: 12px; border: 1px solid #ccc; text-align: center;">🍲 Almuerzo</th><th style="padding: 12px; border: 1px solid #ccc; text-align: center;">🥪 Cena</th><th style="padding: 12px; border: 1px solid #ccc; text-align: center; background-color: #002171;">TOTAL</th></tr></thead><tbody>"""
                        t_des, t_alm, t_cen = 0, 0, 0
                        
                        for _, row in df_com.iterrows():
                            d = int(pd.to_numeric(row['Desayuno'], errors='coerce')) if pd.notna(row['Desayuno']) else 0
                            a = int(pd.to_numeric(row['Almuerzo'], errors='coerce')) if pd.notna(row['Almuerzo']) else 0
                            c = int(pd.to_numeric(row['Cena'], errors='coerce')) if pd.notna(row['Cena']) else 0
                            t_fila = d + a + c
                            if t_fila > 0:
                                t_des += d; t_alm += a; t_cen += c
                                tabla_pdf_html += f'<tr><td style="padding: 10px; border: 1px solid #ddd; font-weight: bold; color: #333;">{row["Departamento"]}</td><td style="padding: 10px; border: 1px solid #ddd; text-align: center;">{d if d > 0 else "-"}</td><td style="padding: 10px; border: 1px solid #ddd; text-align: center;">{a if a > 0 else "-"}</td><td style="padding: 10px; border: 1px solid #ddd; text-align: center;">{c if c > 0 else "-"}</td><td style="padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold; background-color: #f8f9fa;">{t_fila}</td></tr>'
                                
                        tabla_pdf_html += f'<tr style="background-color: #e3f2fd; font-weight: 900; color: {color_primario}; font-size: 16px;"><td style="padding: 12px; border: 1px solid #ccc; text-align: right;">TOTAL GENERAL:</td><td style="padding: 12px; border: 1px solid #ccc; text-align: center;">{t_des}</td><td style="padding: 12px; border: 1px solid #ccc; text-align: center;">{t_alm}</td><td style="padding: 12px; border: 1px solid #ccc; text-align: center;">{t_cen}</td><td style="padding: 12px; border: 1px solid #ccc; text-align: center; font-size: 18px; color: #e65100;">{t_des + t_alm + t_cen}</td></tr></tbody></table>'
                        
                        html_paginas_contenido += f"""
                        <div class="page">
                            <div class="header-band">
                                <span style="font-weight: bold;">Reporte de Operaciones Diarias</span>
                                <span>Fecha: {fecha_str}</span>
                            </div>
                            <div style="padding: 0 40px; height: calc(100% - 100px); display: flex; flex-direction: column;">
                                <h2 style="color: {color_primario}; text-align: center; font-size: 20px; margin-top: 20px; text-transform: uppercase; flex-shrink: 0;">14. PIZARRA DE COMENSALES</h2>
                                {tabla_pdf_html}
                            </div>
                            <div class="footer-band">
                                <span style="float: left; padding-left: 20px;">Departamento de Flota y Logística</span>
                                <span style="float: right; padding-right: 20px;">Página {current_page}</span>
                                <div style="clear: both;"></div>
                            </div>
                        </div>
                        """
                        current_page += 1
                        
                    # CASO 3: Normales y Dinámicos (Imágenes)
                    else:
                        titulo_indice = proc
                        titulo_pdf = proc
                        
                        # --- LÓGICA PARA EL TÍTULO DINÁMICO EN EL PDF ---
                        if proc == "4. Pizarra Estatus (Dinámica)":
                            tipo_selec = st.session_state.get("rad_piz_4", "Ambas")
                            if tipo_selec == "Ambas":
                                titulo_pdf = "4. ESTATUS OPERATIVO Y SATELITAL"
                            elif tipo_selec == "Solo Flota":
                                titulo_pdf = "4. ESTATUS MECÁNICO (FLOTA)"
                            else:
                                titulo_pdf = "4. ESTATUS SATELITAL (GPS)"
                            titulo_indice = titulo_pdf
                        
                        elif proc in ["L1. Guardia De Seguridad", "L2. Guardia De Flota", "L3. Guardia De Monitoreo"]:
                            titulo_pdf = "GUARDIAS SEMANALES"
                        elif proc == "L4. Pizarra De Menú Semanal":
                            titulo_pdf = "MENÚ SEMANAL"
                            
                        html_indice_items += f'<tr><td style="padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 13px; color: #333;">{titulo_indice}</td><td style="padding: 8px 10px; font-size: 13px; font-weight: bold; color: {color_primario}; text-align: right;">Pág. {current_page}</td></tr>'
                        ya_procesado.add(proc)
                        
                        html_paginas_contenido += f"""
                        <div class="page">
                            <div class="header-band">
                                <span style="font-weight: bold;">Reporte de Operaciones Diarias</span>
                                <span>Fecha: {fecha_str}</span>
                            </div>
                            <div style="padding: 0 40px; height: calc(100% - 100px); display: flex; flex-direction: column;">
                                <h2 style="color: {color_primario}; text-align: center; font-size: 20px; margin-top: 20px; text-transform: uppercase; flex-shrink: 0;">{titulo_pdf}</h2>
                                <div style="flex-grow: 1; display: flex; justify-content: center; align-items: center; overflow: hidden; padding-bottom: 20px;">
                                    <img src="data:image/png;base64,{imagenes_b64[proc]}" style="max-width: 100%; max-height: 100%; object-fit: contain; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); border: 1px solid #ccc;">
                                </div>
                            </div>
                            <div class="footer-band">
                                <span style="float: left; padding-left: 20px;">Departamento de Flota y Logística</span>
                                <span style="float: right; padding-right: 20px;">Página {current_page}</span>
                                <div style="clear: both;"></div>
                            </div>
                        </div>
                        """
                        current_page += 1

                total_comensales_str = f"{st.session_state.get('total_comensales', 0):,.0f}"

                html_master = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700;900&display=swap');
                        body {{ margin: 0; padding: 0; font-family: 'Montserrat', sans-serif; background-color: #f0f0f0; -webkit-print-color-adjust: exact; }}
                        @media print {{
                            @page {{ size: A4; margin: 0mm; }} 
                            body {{ background-color: white; margin: 0; padding: 0; }}
                            .page {{ margin: 0 !important; box-shadow: none !important; border: none !important; width: 210mm !important; height: 295mm !important; max-height: 295mm !important; overflow: hidden !important; page-break-after: always !important; page-break-inside: avoid !important; }}
                            .page:last-of-type {{ page-break-after: auto !important; }}
                            .no-print {{ display: none !important; }}
                        }}
                        .page {{ width: 210mm; height: 295mm; padding: 0; margin: 10mm auto; background: white; box-shadow: 0 0 10px rgba(0,0,0,0.2); position: relative; overflow: hidden; box-sizing: border-box; }}
                        .cover-page {{ display: flex; flex-direction: column; justify-content: space-between; background: linear-gradient(180deg, rgba(0,0,0,0.4) 0%, rgba(0,0,0,0) 25%, rgba(0,0,0,0) 70%, rgba(0,0,0,0.6) 100%), url('{img_portada}') center/cover; color: white; }}
                        .cover-top {{ padding: 50px; text-align: left; }}
                        .cover-middle {{ position: absolute; top: 58%; left: 50%; transform: translate(-50%, -50%); text-align: center; width: 90%; text-shadow: 0px 4px 15px rgba(0,0,0,0.9); }}
                        .cover-title {{ font-size: 50px; font-weight: 900; margin: 0; letter-spacing: 2px; text-transform: uppercase; color: white; }}
                        .cover-subtitle {{ font-size: 24px; font-weight: 400; color: #ffffff; margin-top: 15px; letter-spacing: 5px; text-transform: uppercase; }}
                        .cover-bottom {{ position: absolute; bottom: 0; left: 0; width: 100%; background-color: white; color: #333; text-shadow: none; padding: 25px 50px; display: flex; justify-content: space-between; align-items: center; box-sizing: border-box; border-top: 5px solid {color_primario}; }}
                        .inner-padding {{ padding: 40px 50px; height: calc(100% - 80px); overflow: hidden; }}
                        .header-band {{ background-color: {color_primario}; color: white; padding: 15px 30px; display: flex; justify-content: space-between; font-size: 14px; font-weight: bold; }}
                        .footer-band {{ position: absolute; bottom: 0; width: 100%; background-color: #333; color: white; padding: 12px 0; font-size: 12px; font-weight: bold; }}
                        .section-title {{ color: {color_primario}; border-bottom: 3px solid #e65100; padding-bottom: 10px; font-weight: 900; font-size: 18px; text-transform: uppercase; margin-bottom: 20px; margin-top: 15px; }}
                        .metric-box {{ background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; padding: 12px 15px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; }}
                        .metric-label {{ font-size: 13px; font-weight: bold; color: #555; }}
                        .metric-value {{ font-size: 16px; font-weight: 900; color: {color_primario}; }}
                    </style>
                </head>
                <body>
                    <div class="no-print" style="text-align: center; padding: 20px; background-color: #fff; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                        <button onclick="window.print()" style="background-color: #e65100; color: white; border: none; padding: 15px 30px; font-size: 18px; font-weight: bold; cursor: pointer; border-radius: 5px;">
                            🖨️ IMPRIMIR / GUARDAR COMO PDF
                        </button>
                        <p style="color: #666; font-size: 14px; margin-top: 10px;">*En la ventana que se abrirá, asegúrate de seleccionar 'Destino: Guardar como PDF'.</p>
                    </div>

                    <div class="page cover-page">
                        <div class="cover-top">{logo_html}</div>
                        <div class="cover-middle">
                            <h1 class="cover-title">Reporte Master</h1>
                            <h2 class="cover-title" style="font-size: 38px;">de Operaciones</h2>
                            <h3 class="cover-subtitle">Cierre de Día y Control Logístico</h3>
                        </div>
                        <div class="cover-bottom">
                            <div>
                                <p style="margin: 0; font-size: 14px; color: #777;">REALIZADO POR:</p>
                                <p style="margin: 5px 0 0 0; font-size: 20px; font-weight: 900; color: {color_primario};">Departamento de Flota y Logística</p>
                            </div>
                            <div style="text-align: right;">
                                <p style="margin: 0; font-size: 14px; color: #777;">FECHA OPERATIVA:</p>
                                <p style="margin: 5px 0 0 0; font-size: 24px; font-weight: 900; color: #333;">{dia_actual.upper()} {fecha_str}</p>
                            </div>
                        </div>
                    </div>

                    <div class="page">
                        <div class="header-band">
                            <span style="font-weight: bold;">Reporte de Operaciones Diarias</span>
                            <span>Fecha: {fecha_str}</span>
                        </div>
                        <div class="inner-padding">
                            <div style="display: flex; gap: 20px; height: 100%;">
                                <div style="flex: 1;">
                                    <h2 class="section-title">Métricas Finales del Día</h2>
                                    <div class="metric-box"><span class="metric-label">🕘 Apertura Drotaca:</span><span class="metric-value">{hora_apertura}</span></div>
                                    <div class="metric-box"><span class="metric-label">🕙 Cierre Drotaca:</span><span class="metric-value">{hora_cierre_dro}</span></div>
                                    <div class="metric-box"><span class="metric-label">🕗 Cierre Juanita:</span><span class="metric-value">{hora_cierre_jua}</span></div>
                                    <div class="metric-box" style="margin-top: 15px;"><span class="metric-label">Mantenimientos Planificados:</span><span class="metric-value" style="color: #ff9800;">{mant_plan}</span></div>
                                    <div class="metric-box"><span class="metric-label">Mantenimientos Realizados:</span><span class="metric-value" style="color: #2e7d32;">{mant_real}</span></div>
                                    <div class="metric-box" style="margin-top: 15px;"><span class="metric-label">Total Farmacias Entregadas:</span><span class="metric-value">{farmacias:,.0f}</span></div>
                                    <div class="metric-box"><span class="metric-label">Total Bultos Entregados:</span><span class="metric-value">{bultos:,.0f}</span></div>
                                    <div class="metric-box"><span class="metric-label">Total KMs Recorridos:</span><span class="metric-value">{kms:,.0f} Kms</span></div>
                                    <div class="metric-box" style="margin-top: 15px;"><span class="metric-label">Total Comensales del Día:</span><span class="metric-value">{total_comensales_str} Personas</span></div>
                                    <div style="background-color: #e3f2fd; border-left: 5px solid {color_primario}; padding: 10px; margin-top: 15px;">
                                        <span class="metric-label" style="display: block; margin-bottom: 2px;">Gasoil Disp. (Ciudad Drotaca):</span>
                                        <span class="metric-value" style="display: block; font-size: 18px;">{gasoil_lts:,.0f} Lts</span>
                                        <span style="font-size: 11px; color: #555; font-weight: bold;">Autonomía para {gasoil_dias} días operativos.</span>
                                    </div>
                                </div>
                                <div style="flex: 1.2;">
                                    <h2 class="section-title">Índice Operativo</h2>
                                    <table style="width: 100%; border-collapse: collapse; margin-top: 5px;">
                                        <thead>
                                            <tr style="background-color: {color_primario}; color: white;">
                                                <th style="padding: 8px; text-align: left; font-size: 12px;">Proceso Operativo</th>
                                                <th style="padding: 8px; text-align: right; font-size: 12px;">Página</th>
                                            </tr>
                                        </thead>
                                        <tbody>{html_indice_items}</tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                        <div class="footer-band">
                            <span style="float: left; padding-left: 20px;">Departamento de Flota y Logística</span>
                            <span style="float: right; padding-right: 20px;">Página 2</span>
                            <div style="clear: both;"></div>
                        </div>
                    </div>

                    {html_paginas_contenido}

                </body>
                </html>
                """

                components.html(html_master, height=1200, scrolling=True)
                st.toast("✅ Reporte Master ensamblado correctamente. Desplázate hacia abajo para imprimirlo.")