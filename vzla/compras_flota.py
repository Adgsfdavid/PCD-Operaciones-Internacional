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

# --- 2. FUNCIÓN PDF PROFESIONAL CON ESTADÍSTICAS ---
def crear_pdf_auditoria_pro(df_filtrado, rango_texto):
    pdf = FPDF()
    pdf.add_page()
    azul_d = (13, 71, 161)
    
    # Encabezado Azul
    pdf.set_fill_color(*azul_d)
    pdf.rect(0, 0, 210, 35, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 8, "DROGUERIA DROTACA - AUDITORIA DE COMPRAS", ln=True, align='C')
    pdf.set_font("Arial", '', 9)
    pdf.cell(0, 5, f"REPORTE DE GESTION LOGISTICA | PERIODO: {rango_texto.upper()}", ln=True, align='C')
    
    # Línea divisoria Negra
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.5)
    pdf.line(10, 35, 200, 35)
    
    pdf.ln(20)
    pdf.set_text_color(0, 0, 0)
    
    # Tabla de Datos
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(240, 240, 240)
    cols = [("ID", 12), ("FECHA", 22), ("DESCRIPCION", 75), ("CANT", 12), ("SOLICITANTE", 35), ("ESTATUS", 34)]
    for col_name, width in cols:
        pdf.cell(width, 8, col_name, border=1, align='C', fill=True)
    pdf.ln()
    
    pdf.set_font("Arial", '', 7)
    for _, row in df_filtrado.iterrows():
        fecha_s = row['Fecha_Solicitud'].strftime('%d/%m/%Y') if pd.notnull(row['Fecha_Solicitud']) else "N/A"
        pdf.cell(12, 7, str(row['ID_Planilla']), border=1, align='C')
        pdf.cell(22, 7, fecha_s, border=1, align='C')
        pdf.cell(75, 7, str(row['Descripcion'])[:50], border=1, align='L')
        pdf.cell(12, 7, str(row['Cantidad']), border=1, align='C')
        pdf.cell(35, 7, str(row['Usuario'])[:20], border=1, align='C')
        pdf.cell(34, 7, str(row['Estatus']), border=1, align='C')
        pdf.ln()

    # --- SECCIÓN DE ESTADÍSTICAS DE INTERÉS ---
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(0,0,0)
    pdf.set_text_color(255,255,255)
    pdf.cell(0, 8, " ESTADISTICAS DE INTERES Y RENDIMIENTO", ln=True, fill=True)
    pdf.set_text_color(0,0,0)
    pdf.set_font("Arial", '', 9)
    pdf.ln(2)

    # Cálculo de métricas
    df_filtrado['Dia_Semana'] = df_filtrado['Fecha_Solicitud'].dt.day_name()
    dia_pico = df_filtrado['Dia_Semana'].mode()[0] if not df_filtrado.empty else "N/A"
    total = len(df_filtrado)
    listos = len(df_filtrado[df_filtrado['Estatus'] == 'COMPRADO'])
    efectividad = (listos / total * 100) if total > 0 else 0

    pdf.cell(0, 6, f"* Dia con mayor volumen de solicitudes: {dia_pico}", ln=True)
    pdf.cell(0, 6, f"* Efectividad de cumplimiento en este periodo: {efectividad:.1f}%", ln=True)
    pdf.cell(0, 6, f"* Solicitudes totales procesadas: {total} items.", ln=True)
    
    return pdf.output(dest='S').encode('latin-1', 'replace')

# --- 3. CARGA DE DATOS ---
try:
    doc = conectar_bd_compras()
    hoja_sol = doc.worksheet("SOLICITUDES")
    data_raw = hoja_sol.get_all_records()
    df_global = pd.DataFrame(data_raw)
    if not df_global.empty:
        df_global.columns = [str(c).strip() for c in df_global.columns]
        rename_map = {'Fecha': 'Fecha_Solicitud', 'Tipo_Solicitud': 'Categoria'}
        df_global.rename(columns=rename_map, inplace=True)
        df_global['Fecha_Solicitud'] = pd.to_datetime(df_global['Fecha_Solicitud'], format='%d/%m/%Y', errors='coerce')
except:
    df_global = pd.DataFrame()

st.title("🛒 Control de Compras PCD")
t1, t2, t3 = st.tabs(["📝 Solicitud", "✅ Check", "📊 Auditoría"])

# [Omitiendo código de T1 y T2 por brevedad, se mantienen igual]
# ... (Código de Pestaña 1 y 2 que ya tienes funcionando) ...

# ---------------------------------------------------------
# PESTAÑA 3: HISTORIAL Y AUDITORÍA (MODO SIMPLE)
# ---------------------------------------------------------
with t3:
    if not df_global.empty:
        st.subheader("📊 Resumen General de Gestión")
        
        # Filtros (Por defecto en 'Todo')
        c1, c2, c3 = st.columns(3)
        with c1:
            f_periodo = st.selectbox("Filtrar Periodo:", ["Todo el Historial", "Mes Actual", "Semana Actual", "Hoy"])
        with c2:
            f_cat = st.multiselect("Categoría:", df_global['Categoria'].unique(), default=df_global['Categoria'].unique())
        with c3:
            f_est = st.multiselect("Estatus:", ["PENDIENTE", "COMPRADO"], default=["PENDIENTE", "COMPRADO"])

        # Aplicar Filtros
        df_f = df_global.copy()
        hoy = datetime.now()
        if f_periodo == "Mes Actual": df_f = df_f[df_f['Fecha_Solicitud'].dt.month == hoy.month]
        elif f_periodo == "Semana Actual":
            lunes = hoy - timedelta(days=hoy.weekday())
            df_f = df_f[df_f['Fecha_Solicitud'].dt.date >= lunes.date()]
        elif f_periodo == "Hoy": df_f = df_f[df_f['Fecha_Solicitud'].dt.date == hoy.date()]

        df_f = df_f[df_f['Categoria'].isin(f_cat)]
        df_f = df_f[df_f['Estatus'].isin(f_est)]

        # Métricas Rápidas
        k1, k2, k3 = st.columns(3)
        k1.metric("Items", len(df_f))
        k2.metric("Pendientes", len(df_f[df_f['Estatus']=='PENDIENTE']))
        k3.metric("Efectividad", f"{(len(df_f[df_f['Estatus']=='COMPRADO'])/len(df_f)*100 if len(df_f)>0 else 0):.1f}%")

        # Tabla HTML (Buscador activo)
        filas_html = ""
        for _, r in df_f.sort_values('Fecha_Solicitud', ascending=False).iterrows():
            f_str = r['Fecha_Solicitud'].strftime('%d/%m/%Y') if pd.notnull(r['Fecha_Solicitud']) else "-"
            color = "#198754" if r['Estatus'] == 'COMPRADO' else "#ffc107"
            txt_c = "white" if r['Estatus'] == 'COMPRADO' else "black"
            filas_html += f"<tr><td>{r['ID_Planilla']}</td><td>{f_str}</td><td>{r['Descripcion']}</td><td>{r['Cantidad']}</td><td><span style='background:{color}; color:{txt_c}; padding:2px 5px; border-radius:3px;'>{r['Estatus']}</span></td></tr>"

        html_table = f"""
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
        <style>body{{background:#0e1117; color:white; font-size:12px;}} .dataTables_wrapper{{color:white;}} table.dataTable tbody tr{{background:#212529; color:white;}}</style>
        <table id="audit" class="display" style="width:100%">
            <thead><tr><th>ID</th><th>FECHA</th><th>DESCRIPCION</th><th>CANT</th><th>ESTATUS</th></tr></thead>
            <tbody>{filas_html}</tbody>
        </table>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        <script>$(document).ready(function(){{$('#audit').DataTable({{"order": [[0, "desc"]]}});}});</script>
        """
        components.html(html_table, height=500, scrolling=True)

        # Botón PDF
        st.markdown("---")
        if st.button("📄 GENERAR REPORTE PDF PROFESIONAL"):
            pdf_bytes = crear_pdf_auditoria_pro(df_f, f_periodo)
            st.download_button("⬇️ DESCARGAR PDF", pdf_bytes, f"Auditoria_{f_periodo}.pdf", "application/pdf")
    else:
        st.info("Sin datos registrados.")
