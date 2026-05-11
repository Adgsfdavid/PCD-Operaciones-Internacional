import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import openpyxl
import io
import os

# --- 1. CONFIGURACIÓN DE CONEXIÓN ---
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"
CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def conectar_bd_compras():
    # ID: 1M9VQOHU6LHniSBu3A_o2qDFlZITrDRm_f7qYGU5iOX8
    return obtener_cliente_sheets().open_by_key("1M9VQOHU6LHniSBu3A_o2qDFlZITrDRm_f7qYGU5iOX8")

# --- 2. CARGA INICIAL DE DATOS ---
try:
    doc = conectar_bd_compras()
    hoja_sol = doc.worksheet("SOLICITUDES")
    data_raw = hoja_sol.get_all_records()
    df_global = pd.DataFrame(data_raw)
    
    # Limpieza de nombres de columnas para evitar el desplazamiento
    df_global.columns = [c.strip() for c in df_global.columns]
except Exception as e:
    df_global = pd.DataFrame()

st.title("🛒 Gestión y Cronometría de Compras")
tab_nueva, tab_control, tab_metricas = st.tabs(["📝 Crear Solicitud", "✅ Check de Entrega", "📊 Historial y Auditoría"])

# ---------------------------------------------------------
# PESTAÑA 1: GENERAR SOLICITUD
# ---------------------------------------------------------
with tab_nueva:
    with st.container(border=True):
        col_id, col_cat = st.columns(2)
        with col_id:
            try:
                hoja_cnt = doc.worksheet("CONTADOR")
                proximo_id = int(hoja_cnt.acell('A2').value)
            except: proximo_id = 1
            st.subheader(f"Planilla N° {proximo_id}")
            fecha_sol = st.date_input("Fecha Solicitud:", datetime.now())
        with col_cat:
            categoria = st.selectbox("Tipo de Solicitud:", [
                "Solicitud para stock (almacén de flota)", 
                "Solicitud para vehículos"
            ])

    if 'filas_compras' not in st.session_state:
        st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)

    df_editado = st.data_editor(st.session_state['filas_compras'], num_rows="dynamic", use_container_width=True, hide_index=True)
    nota_p = st.text_area("Notas para el Excel (B22):")

    if st.button("💾 GUARDAR SOLICITUD", type="primary", use_container_width=True):
        df_valido = df_editado[df_editado['Descripción'].str.strip() != ""].copy()
        if not df_valido.empty:
            with st.spinner("Sincronizando con Google Sheets..."):
                registros = []
                for _, r in df_valido.iterrows():
                    # ORDEN ESTRICTO DE 10 COLUMNAS PARA EVITAR DESPLAZAMIENTO
                    registros.append([
                        proximo_id, 
                        fecha_sol.strftime("%d/%m/%Y"), 
                        categoria, 
                        r['Cantidad'], 
                        r['Descripción'].upper(), 
                        "PENDIENTE", 
                        "", # Fecha_Entrega
                        "", # Dias_Resolucion
                        st.session_state.get('usuario','admin_vzla'), 
                        nota_p
                    ])
                
                hoja_sol.append_rows(registros)
                doc.worksheet("CONTADOR").update_acell('A2', proximo_id + 1)
                
                # Manejo de Excel
                ruta_excel = os.path.join(os.path.dirname(__file__), "Planilla de Solicitud de compra.xlsx")
                if os.path.exists(ruta_excel):
                    wb = openpyxl.load_workbook(ruta_excel)
                    ws = wb.active
                    ws["C6"] = proximo_id
                    ws["B22"] = f"NOTA: {nota_p}"
                    for i, (idx, row) in enumerate(df_valido.iterrows(), start=8):
                        if i <= 20:
                            ws[f"C{i}"], ws[f"D{i}"] = row['Cantidad'], row['Descripción'].upper()
                    out = io.BytesIO()
                    wb.save(out)
                    st.download_button("⬇️ Descargar Planilla Excel", out.getvalue(), f"Solicitud_{proximo_id}.xlsx")
                
                st.success("Solicitud registrada correctamente.")
                st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)
                st.rerun()

# ---------------------------------------------------------
# PESTAÑA 2: CHECK DE ENTREGA
# ---------------------------------------------------------
with tab_control:
    st.header("📋 Pendientes de Compra")
    if not df_global.empty and 'Estatus' in df_global.columns:
        df_p = df_global[df_global['Estatus'].str.strip() == 'PENDIENTE'].copy()
        
        if df_p.empty:
            st.success("✅ No hay ítems pendientes.")
        else:
            st.dataframe(df_p[['ID_Planilla', 'Fecha_Solicitud', 'Descripcion', 'Cantidad']], use_container_width=True, hide_index=True)
            
            st.markdown("---")
            opciones = df_p.apply(lambda r: f"P-{r['ID_Planilla']} | {r['Cantidad']}x {r['Descripcion']}", axis=1).tolist()
            seleccion = st.selectbox("Seleccione el ítem recibido:", opciones)
            fecha_rec = st.date_input("Fecha de Recepción:", datetime.now())
            
            if st.button("🏁 CONFIRMAR ENTREGA"):
                id_sel = int(seleccion.split("|")[0].replace("P-", "").strip())
                desc_sel = seleccion.split("|")[1].split("x ", 1)[1].strip()
                
                data_actual = hoja_sol.get_all_records()
                for idx, row in enumerate(data_actual):
                    if int(row['ID_Planilla']) == id_sel and str(row['Descripcion']).strip() == desc_sel:
                        fila_sheet = idx + 2
                        f_ini = datetime.strptime(str(row['Fecha_Solicitud']), "%d/%m/%Y")
                        f_fin = datetime.combine(fecha_rec, datetime.min.time())
                        dias = (f_fin - f_ini).days
                        
                        # Actualización de columnas F, G, H
                        hoja_sol.update(f"F{fila_sheet}:H{fila_sheet}", [["COMPRADO", fecha_rec.strftime("%d/%m/%Y"), dias]])
                        st.success(f"Entregado en {dias} días.")
                        st.rerun()
                        break

# ---------------------------------------------------------
# PESTAÑA 3: AUDITORÍA
# ---------------------------------------------------------
with tab_metricas:
    st.header("📊 Auditoría de Tiempos")
    if not df_global.empty and 'Estatus' in df_global.columns:
        df_comp = df_global[df_global['Estatus'].str.strip() == 'COMPRADO'].copy()
        df_pend = df_global[df_global['Estatus'].str.strip() == 'PENDIENTE'].copy()
        
        if len(df_pend) >= 15:
            st.error(f"🚨 ALERTA CRÍTICA: {len(df_pend)} ítems pendientes.")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Solicitudes", len(df_global))
        m2.metric("Pendientes", len(df_pend))
        
        # FIX PARA EL KEYERROR: Verificar si la columna existe y tiene datos
        if not df_comp.empty and 'Dias_Resolucion' in df_comp.columns:
            promedio = pd.to_numeric(df_comp['Dias_Resolucion'], errors='coerce').mean()
            m3.metric("Tiempo Promedio", f"{promedio:.1f} días")
            st.dataframe(df_comp, use_container_width=True, hide_index=True)
        else:
            m3.metric("Tiempo Promedio", "0.0 días")
            st.info("No hay historial de compras completadas.")
