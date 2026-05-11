import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import openpyxl
import io
import os
from fpdf import FPDF
import streamlit.components.v1 as components

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
    return obtener_cliente_sheets().open_by_key("1M9VQOHU6LHniSBu3A_o2qDFlZITrDRm_f7qYGU5iOX8")

# --- 2. CARGA Y LIMPIEZA DE DATOS ---
try:
    doc = conectar_bd_compras()
    hoja_sol = doc.worksheet("SOLICITUDES")
    headers = [str(h).strip() for h in hoja_sol.row_values(1)]
    data_raw = hoja_sol.get_all_records()
    df_global = pd.DataFrame(data_raw)
    
    if not df_global.empty:
        df_global.columns = [str(c).strip() for c in df_global.columns]
        # Normalización de nombres para evitar KeyErrors
        rename_map = {'Fecha': 'Fecha_Solicitud', 'Tipo_Solicitud': 'Categoria'}
        df_global.rename(columns=rename_map, inplace=True)
        df_global['Fecha_Solicitud'] = pd.to_datetime(df_global['Fecha_Solicitud'], format='%d/%m/%Y', errors='coerce')
        for col in ['Estatus', 'Dias_Resolucion', 'Descripcion', 'Cantidad', 'ID_Planilla']:
            if col not in df_global.columns: df_global[col] = ""
except:
    df_global = pd.DataFrame()
    headers = []

st.title("🛒 Gestión y Cronometría de Compras")
tab_nueva, tab_control, tab_metricas = st.tabs(["📝 Crear Solicitud", "✅ Check de Entrega", "📊 Historial y Auditoría"])

# ---------------------------------------------------------
# PESTAÑA 1: CREAR SOLICITUD
# ---------------------------------------------------------
with tab_nueva:
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            try:
                hoja_cnt = doc.worksheet("CONTADOR")
                proximo_id = int(hoja_cnt.acell('A2').value)
            except: proximo_id = 1
            st.subheader(f"Planilla N° {proximo_id}")
            fecha_sol = st.date_input("Fecha Solicitud:", datetime.now())
        with c2:
            categoria = st.selectbox("Tipo de Solicitud:", ["Solicitud para stock (almacén de flota)", "Solicitud para vehículos"])

    if 'filas_compras' not in st.session_state:
        st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)
    
    df_editado = st.data_editor(st.session_state['filas_compras'], num_rows="dynamic", use_container_width=True, hide_index=True)
    nota_p = st.text_area("Notas (B22):")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 LIMPIAR", use_container_width=True):
            st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)
            st.rerun()
    with col2:
        if st.button("💾 GUARDAR SOLICITUD", type="primary", use_container_width=True):
            df_v = df_editado[df_editado['Descripción'].str.strip() != ""].copy()
            if not df_v.empty:
                with st.spinner("Guardando..."):
                    registros = []
                    for i, (_, r) in enumerate(df_v.iterrows(), start=1):
                        fila_dict = {"ID_Planilla": proximo_id, "Fecha": fecha_sol.strftime("%d/%m/%Y"), "Fecha_Solicitud": fecha_sol.strftime("%d/%m/%Y"), "Tipo_Solicitud": categoria, "Item_No": i, "Cantidad": r['Cantidad'], "Descripcion": r['Descripción'].upper(), "Estatus": "PENDIENTE", "Usuario": st.session_state.get('usuario','admin_vzla'), "Nota": nota_p}
                        registros.append([fila_dict.get(h, "") for h in headers])
                    
                    hoja_sol.append_rows(registros)
                    doc.worksheet("CONTADOR").update_acell('A2', proximo_id + 1)
                    st.success("✅ Guardado con éxito.")
                    st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)
                    st.rerun()

# ---------------------------------------------------------
# PESTAÑA 2: CHECK DE ENTREGA (CAMBIO DE ESTATUS)
# ---------------------------------------------------------
with tab_control:
    st.header("🏁 Marcar Ítems como Recibidos")
    st.info("Al marcar aquí, el sistema cambia el estatus a COMPRADO y calcula los días de demora.")
    
    if not df_global.empty:
        df_p = df_global[df_global['Estatus'].astype(str).str.strip() == 'PENDIENTE'].copy()
        if df_p.empty:
            st.success("🎉 No hay pendientes.")
        else:
            st.dataframe(df_p[['ID_Planilla', 'Fecha_Solicitud', 'Descripcion', 'Cantidad']], use_container_width=True, hide_index=True)
            opciones = df_p.apply(lambda r: f"P-{r['ID_Planilla']} | {r['Descripcion']}", axis=1).tolist()
            seleccion = st.selectbox("Seleccione el ítem que llegó:", opciones)
            fecha_rec = st.date_input("Fecha de Recepción:", datetime.now())
            
            if st.button("🏁 CONFIRMAR ENTREGA", type="primary"):
                id_sel = int(seleccion.split("|")[0].replace("P-", "").strip())
                desc_sel = seleccion.split("|")[1].strip()
                
                # Actualización en Sheets
                for idx, row in enumerate(hoja_sol.get_all_records()):
                    if int(row['ID_Planilla']) == id_sel and str(row['Descripcion']).strip() == desc_sel:
                        fila = idx + 2
                        f_ini = datetime.strptime(str(row.get('Fecha', row.get('Fecha_Solicitud'))), "%d/%m/%Y")
                        dias = (datetime.combine(fecha_rec, datetime.min.time()) - f_ini).days
                        
                        # Actualizamos columnas de Estatus, Fecha Entrega y Días
                        hoja_sol.update_cell(fila, headers.index("Estatus")+1, "COMPRADO")
                        hoja_sol.update_cell(fila, headers.index("Fecha_Entrega")+1, fecha_rec.strftime("%d/%m/%Y"))
                        hoja_sol.update_cell(fila, headers.index("Dias_Resolucion")+1, dias)
                        st.success(f"Entregado en {dias} días.")
                        st.rerun()
                        break

# ---------------------------------------------------------
# PESTAÑA 3: HISTORIAL Y AUDITORÍA (NUEVOS FILTROS)
# ---------------------------------------------------------
with tab_metricas:
    st.header("📊 Auditoría de Eficiencia")
    
    if not df_global.empty:
        # --- FILTROS DISCRETOS ---
        c1, c2, c3 = st.columns(3)
        with c1:
            periodo = st.selectbox("Periodo:", ["Mes Actual", "Semana Actual", "Día de Hoy", "Todo el Historial"])
        with c2:
            f_tipo = st.multiselect("Categoría:", ["Solicitud para stock (almacén de flota)", "Solicitud para vehículos"], default=["Solicitud para stock (almacén de flota)", "Solicitud para vehículos"])
        with c3:
            f_estatus = st.multiselect("Estatus:", ["PENDIENTE", "COMPRADO"], default=["PENDIENTE", "COMPRADO"])

        # Aplicar Lógica de Filtros
        df_f = df_global.copy()
        hoy = datetime.now()
        
        if periodo == "Mes Actual":
            df_f = df_f[df_f['Fecha_Solicitud'].dt.month == hoy.month]
        elif periodo == "Semana Actual":
            lunes = hoy - timedelta(days=hoy.weekday())
            df_f = df_f[df_f['Fecha_Solicitud'].dt.date >= lunes.date()]
        elif periodo == "Día de Hoy":
            df_f = df_f[df_f['Fecha_Solicitud'].dt.date == hoy.date()]
            
        df_f = df_f[df_f['Categoria'].isin(f_tipo)]
        df_f = df_f[df_f['Estatus'].str.strip().isin(f_estatus)]

        # --- MÉTRICAS ---
        st.markdown("---")
        k1, k2, k3 = st.columns(3)
        k1.metric("Solicitudes", len(df_f))
        k2.metric("Pendientes", len(df_f[df_f['Estatus']=='PENDIENTE']))
        promedio = pd.to_numeric(df_f[df_f['Estatus']=='COMPRADO']['Dias_Resolucion'], errors='coerce').mean()
        k3.metric("Promedio Entrega", f"{promedio:.1f} días" if not pd.isna(promedio) else "0 d")

        # --- TABLA HTML ---
        filas_html = ""
        for _, r in df_f.sort_values('Fecha_Solicitud', ascending=False).iterrows():
            badge = f"<span style='background:{'#198754' if r['Estatus']=='COMPRADO' else '#ffc107'}; color:{'white' if r['Estatus']=='COMPRADO' else 'black'}; padding:3px 8px; border-radius:4px;'>{r['Estatus']}</span>"
            filas_html += f"<tr><td>{r['ID_Planilla']}</td><td>{r['Fecha_Solicitud'].strftime('%d/%m/%Y')}</td><td style='text-align:left;'>{r['Descripcion']}</td><td>{r['Cantidad']}</td><td>{badge}</td></tr>"

        html_table = f"""
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
        <style>body{{background:#0e1117; color:white; font-size:12px;}} .dataTables_wrapper{{color:white;}} table.dataTable tbody tr{{background:#212529; color:white;}}</style>
        <table id="audit" class="display" style="width:100%">
            <thead><tr><th>ID</th><th>FECHA</th><th>ÍTEM</th><th>CANT</th><th>ESTADO</th></tr></thead>
            <tbody>{filas_html}</tbody>
        </table>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        <script>$(document).ready(function(){{$('#audit').DataTable({{"order": [[0, "desc"]], "language": {{"search": "Buscador Rápido:"}}}});}});</script>
        """
        components.html(html_table, height=500, scrolling=True)
