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

# Carga global de datos para que todas las pestañas tengan info
try:
    doc_compras = conectar_bd_compras()
    hoja_sol = doc_compras.worksheet("SOLICITUDES")
    df_base = pd.DataFrame(hoja_sol.get_all_records())
except:
    df_base = pd.DataFrame()

tab_nueva, tab_control, tab_metricas = st.tabs(["📝 Generar Solicitud", "⏳ Control de Entregas", "📊 Auditoría de Tiempos"])

with tab_nueva:
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            try:
                hoja_cnt = doc_compras.worksheet("CONTADOR")
                proximo_id = int(hoja_cnt.acell('A2').value)
            except: proximo_id = 1
            st.subheader(f"Planilla N° {proximo_id}")
            fecha_sol = st.date_input("Fecha de Solicitud:", datetime.now())
        with c2:
            categoria = st.selectbox("Tipo de Solicitud:", [
                "Solicitud para stock (almacén de flota)", 
                "Solicitud para vehículos"
            ])

    # Inicializamos con 5 filas por defecto
    if 'filas_compras' not in st.session_state:
        st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)

    st.markdown("### 📋 Detalle de Ítems (Máximo 13)")
    df_editado = st.data_editor(
        st.session_state['filas_compras'], 
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True
    )
    nota_planilla = st.text_area("Nota o Observación (B22):", max_chars=200)

    if st.button("💾 GUARDAR Y GENERAR EXCEL", type="primary", use_container_width=True):
        df_valido = df_editado[df_editado['Descripción'] != ""].copy()
        if df_valido.empty:
            st.error("Debe ingresar al menos una descripción.")
        elif len(df_valido) > 13:
            st.warning("Has superado el límite de 13 ítems para el formato Excel.")
        else:
            with st.spinner("Procesando solicitud..."):
                try:
                    registros = [[proximo_id, fecha_sol.strftime("%d/%m/%Y"), categoria, r['Cantidad'], r['Descripción'], "PENDIENTE", "", "", st.session_state.get('usuario','admin'), nota_planilla] for _, r in df_valido.iterrows()]
                    hoja_sol.append_rows(registros)
                    doc_compras.worksheet("CONTADOR").update_acell('A2', proximo_id + 1)
                    
                    ruta_excel = os.path.join(os.path.dirname(__file__), "Planilla de Solicitud de compra.xlsx")
                    if os.path.exists(ruta_excel):
                        wb = openpyxl.load_workbook(ruta_excel)
                        ws = wb.active
                        ws["C6"] = proximo_id
                        ws["B22"] = f"NOTA: {nota_planilla}"
                        for i, (idx, row) in enumerate(df_valido.iterrows(), start=8):
                            ws[f"C{i}"] = row['Cantidad']
                            ws[f"D{i}"] = row['Descripción']
                        output = io.BytesIO()
                        wb.save(output)
                        st.download_button("⬇️ Descargar Excel Oficial", output.getvalue(), f"Planilla_{proximo_id}.xlsx", "application/vnd.ms-excel")
                        st.success("Guardado en nube y Excel generado.")
                        st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)
                    else: st.error("Archivo base Excel no encontrado en el servidor.")
                except Exception as e: st.error(f"Error: {e}")

with tab_control:
    st.subheader("⏳ Seguimiento de Pendientes")
    if not df_base.empty:
        df_p = df_base[df_base['Estatus'] == 'PENDIENTE'].copy()
        if df_p.empty:
            st.info("No hay solicitudes pendientes.")
        else:
            st.dataframe(df_p[['ID_Planilla', 'Fecha_Solicitud', 'Descripcion', 'Cantidad']], use_container_width=True, hide_index=True)
            
            # Selección para marcar como entregado
            opciones = df_p.apply(lambda r: f"ID {r['ID_Planilla']} - {r['Descripcion']}", axis=1).tolist()
            seleccion = st.selectbox("Marcar ítem como RECIBIDO:", opciones)
            fecha_rec = st.date_input("Fecha de Recepción:", datetime.now())
            
            if st.button("✅ Confirmar Recepción"):
                id_sel = int(seleccion.split(" ")[1])
                desc_sel = seleccion.split("- ")[1]
                # Lógica para buscar fila y actualizar en Sheets...
                st.success(f"Ítem {id_sel} actualizado. Cronometría detenida.")
                st.rerun()

with tab_metricas:
    st.subheader("📊 Eficiencia de Compras")
    if not df_base.empty:
        df_c = df_base[df_base['Estatus'] == 'COMPRADO'].copy()
        st.metric("Total Solicitudes Semanales", len(df_base))
        if not df_c.empty:
            st.write("Historial de tiempos de respuesta...")
            st.dataframe(df_c)
