import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import openpyxl
import io
import os

# CONFIGURACIÓN DE CONEXIÓN
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

st.title("🛒 Gestión y Cronometría de Compras")
tab_nueva, tab_control, tab_metricas = st.tabs(["📝 Generar Solicitud", "⏳ Control de Entregas", "📊 Auditoría de Tiempos"])

with tab_nueva:
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            try:
                hoja_cnt = conectar_bd_compras().worksheet("CONTADOR")
                proximo_id = int(hoja_cnt.acell('A2').value)
            except: proximo_id = 1
            st.subheader(f"Planilla N° {proximo_id}")
            fecha_sol = st.date_input("Fecha de Solicitud:", datetime.now())
        with c2:
            # CATEGORÍAS ACTUALIZADAS
            categoria = st.selectbox("Tipo de Solicitud:", [
                "Solicitud para stock (almacén de flota)", 
                "Solicitud para vehículos"
            ])

    if 'filas_compras' not in st.session_state:
        st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 3)

    st.markdown("### 📋 Detalle de Ítems")
    # CONFIGURACIÓN DE TABLA CORREGIDA PARA EVITAR TYPEERROR
    df_editado = st.data_editor(
        st.session_state['filas_compras'], 
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True
    )
    nota_planilla = st.text_area("Nota (B22):", max_chars=200)

    if st.button("💾 GUARDAR Y GENERAR EXCEL", type="primary", use_container_width=True):
        df_valido = df_editado[df_editado['Descripción'] != ""].copy()
        if df_valido.empty:
            st.error("Debe ingresar al menos una descripción.")
        else:
            with st.spinner("Procesando..."):
                try:
                    doc = conectar_bd_compras()
                    hoja_sol = doc.worksheet("SOLICITUDES")
                    registros = [[proximo_id, fecha_sol.strftime("%d/%m/%Y"), categoria, r['Cantidad'], r['Descripción'], "PENDIENTE", "", "", st.session_state.get('usuario','admin'), nota_planilla] for _, r in df_valido.iterrows()]
                    hoja_sol.append_rows(registros)
                    doc.worksheet("CONTADOR").update_acell('A2', proximo_id + 1)
                    
                    # Generación de Excel
                    ruta_excel = os.path.join(os.path.dirname(__file__), "Planilla de Solicitud de compra.xlsx")
                    if os.path.exists(ruta_excel):
                        wb = openpyxl.load_workbook(ruta_excel)
                        ws = wb.active
                        ws["C6"] = proximo_id
                        ws["B22"] = f"NOTA: {nota_planilla}"
                        for i, row in enumerate(df_valido.iterrows(), start=8):
                            if i <= 20:
                                ws[f"C{i}"] = row[1]['Cantidad']
                                ws[f"D{i}"] = row[1]['Descripción']
                        output = io.BytesIO()
                        wb.save(output)
                        st.download_button("⬇️ Descargar Excel", output.getvalue(), f"Planilla_{proximo_id}.xlsx", "application/vnd.ms-excel")
                        st.success("Guardado correctamente.")
                    else: st.error("Archivo base no encontrado.")
                except Exception as e: st.error(f"Error: {e}")
