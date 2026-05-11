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

# --- 2. FUNCIONES DE APOYO ---
def crear_pdf_auditoria(df_filtrado, rango_texto):
    pdf = FPDF()
    pdf.add_page()
    
    # Colores Drotaca
    azul_drotaca = (13, 71, 161)
    dorado_drotaca = (212, 175, 55)
    
    # Encabezado
    pdf.set_fill_color(*azul_drotaca)
    pdf.rect(0, 0, 210, 40, 'F')
    
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "REPORTE DE AUDITORIA DE COMPRAS - FLOTA", ln=True, align='C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 5, f"Periodo: {rango_texto}", ln=True, align='C')
    pdf.cell(0, 5, f"Generado el: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    # Línea Dorada
    pdf.set_draw_color(*dorado_drotaca)
    pdf.set_line_width(1)
    pdf.line(0, 40, 210, 40)
    
    pdf.ln(20)
    pdf.set_text_color(0, 0, 0)
    
    # Tabla
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(230, 230, 230)
    
    cols = [("ID", 15), ("FECHA", 25), ("DESCRIPCION", 70), ("CANT", 15), ("SOLICITADO", 35), ("ESTADO", 30)]
    for col_name, width in cols:
        pdf.cell(width, 10, col_name, border=1, align='C', fill=True)
    pdf.ln()
    
    pdf.set_font("Arial", '', 7)
    for _, row in df_filtrado.iterrows():
        pdf.cell(15, 8, str(row['ID_Planilla']), border=1, align='C')
        pdf.cell(25, 8, row['Fecha_Solicitud'].strftime('%d/%m/%Y'), border=1, align='C')
        pdf.cell(70, 8, str(row['Descripcion'])[:45], border=1, align='L')
        pdf.cell(15, 8, str(row['Cantidad']), border=1, align='C')
        pdf.cell(35, 8, str(row['Usuario'])[:20], border=1, align='C')
        pdf.cell(30, 8, str(row['Estatus']), border=1, align='C')
        pdf.ln()
        
    return pdf.output(dest='S').encode('latin-1', 'replace')

# --- 3. CARGA DE DATOS ---
try:
    doc = conectar_bd_compras()
    hoja_sol = doc.worksheet("SOLICITUDES")
    data_raw = hoja_sol.get_all_records()
    df_global = pd.DataFrame(data_raw)
    df_global.columns = [str(c).strip() for c in df_global.columns]
    if 'Fecha' in df_global.columns:
        df_global.rename(columns={'Fecha': 'Fecha_Solicitud'}, inplace=True)
    df_global['Fecha_Solicitud'] = pd.to_datetime(df_global['Fecha_Solicitud'], format='%d/%m/%Y', errors='coerce')
except:
    df_global = pd.DataFrame()

st.title("🛒 Gestión y Cronometría de Compras")
tab_nueva, tab_control, tab_metricas = st.tabs(["📝 Crear Solicitud", "✅ Check de Entrega", "📊 Historial y Auditoría"])

# [LAS PESTAÑAS 1 Y 2 SE MANTIENEN IGUAL QUE TU CÓDIGO ANTERIOR]

# ---------------------------------------------------------
# PESTAÑA 3: HISTORIAL Y AUDITORÍA (CON FILTROS Y PDF)
# ---------------------------------------------------------
with tab_metricas:
    st.subheader("🔍 Filtros de Auditoría")
    
    if not df_global.empty:
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            tipo_filtro = st.radio("Filtrar por:", ["Mes Actual", "Última Semana", "Rango Personalizado"], horizontal=True)
        
        with col_f2:
            hoy = datetime.now()
            if tipo_filtro == "Rango Personalizado":
                fecha_inicio = st.date_input("Desde:", hoy - timedelta(days=30))
                fecha_fin = st.date_input("Hasta:", hoy)
                df_filtrado = df_global[(df_global['Fecha_Solicitud'].dt.date >= fecha_inicio) & (df_global['Fecha_Solicitud'].dt.date <= fecha_fin)].copy()
                rango_txt = f"{fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}"
            elif tipo_filtro == "Mes Actual":
                df_filtrado = df_global[df_global['Fecha_Solicitud'].dt.month == hoy.month].copy()
                rango_txt = hoy.strftime('%B %Y')
            else:
                lunes = hoy - timedelta(days=hoy.weekday())
                df_filtrado = df_global[df_global['Fecha_Solicitud'].dt.date >= lunes.date()].copy()
                rango_txt = "Esta Semana"

        # --- KPIS ---
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        total_f = len(df_filtrado)
        comp_f = len(df_filtrado[df_filtrado['Estatus'] == 'COMPRADO'])
        pend_f = len(df_filtrado[df_filtrado['Estatus'] == 'PENDIENTE'])
        prom_f = pd.to_numeric(df_filtrado[df_filtrado['Estatus'] == 'COMP_f']['Dias_Resolucion'], errors='coerce').mean() if comp_f > 0 else 0

        m1.metric("Solicitudes en Periodo", total_f)
        m2.metric("Pendientes", pend_f, delta=f"{pend_f}", delta_color="inverse")
        m3.metric("Entregados", comp_f)
        m4.metric("Promedio Entrega", f"{prom_f:.1f} d")

        # --- TABLA HTML DATATABLES ---
        filas_html = ""
        for _, row in df_filtrado.sort_values('Fecha_Solicitud', ascending=False).iterrows():
            badge = f"<span style='background: {'#198754' if row['Estatus']=='COMPRADO' else '#ffc107'}; color: {'white' if row['Estatus']=='COMPRADO' else 'black'}; padding: 4px 8px; border-radius: 4px;'>{row['Estatus']}</span>"
            filas_html += f"<tr><td>{row['ID_Planilla']}</td><td>{row['Fecha_Solicitud'].strftime('%d/%m/%Y')}</td><td style='text-align:left;'>{row['Descripcion']}</td><td>{row['Cantidad']}</td><td>{row['Usuario']}</td><td>{badge}</td></tr>"

        html_table = f"""
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
        <style>body {{ background: #0e1117; color: white; }} .dataTables_wrapper {{ color: white; }} table.dataTable tbody tr {{ background: #212529; }}</style>
        <table id="audit" class="display" style="width:100%">
            <thead><tr><th>ID</th><th>FECHA</th><th>ITEM</th><th>CANT</th><th>USUARIO</th><th>ESTATUS</th></tr></thead>
            <tbody>{filas_html}</tbody>
        </table>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        <script>$(document).ready(function() {{ $('#audit').DataTable({{"order": [[0, "desc"]], "language": {{"search": "Buscar:", "lengthMenu": "Mostrar _MENU_ items"}} }}); }});</script>
        """
        components.html(html_table, height=500, scrolling=True)

        # --- BOTÓN DE EXPORTACIÓN PDF ---
        st.markdown("### 🖨️ Exportar Resultados")
        if not df_filtrado.empty:
            pdf_bytes = crear_pdf_auditoria(df_filtrado, rango_txt)
            st.download_button(
                label="📄 DESCARGAR AUDITORÍA EN PDF PROFESIONAL",
                data=pdf_bytes,
                file_name=f"Auditoria_Compras_{datetime.now().strftime('%d%m%y')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary"
            )
    else:
        st.info("No hay datos para filtrar.")
