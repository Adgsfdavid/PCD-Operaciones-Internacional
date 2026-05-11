import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import openpyxl
import io
import os

# --- CONEXIÓN ---
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
    # USANDO TU ID DE GOOGLE SHEET
    return obtener_cliente_sheets().open_by_key("1M9VQOHU6LHniSBu3A_o2qDFlZITrDRm_f7qYGU5iOX8")

# --- CARGA DE DATOS INICIAL ---
try:
    doc = conectar_bd_compras()
    hoja_sol = doc.worksheet("SOLICITUDES")
    data_raw = hoja_sol.get_all_records()
    df_global = pd.DataFrame(data_raw)
except:
    df_global = pd.DataFrame()

st.title("🛒 Control Tower: Compras y Repuestos")

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
            categoria = st.selectbox("Tipo:", ["Solicitud para stock (almacén de flota)", "Solicitud para vehículos"])

    if 'filas_compras' not in st.session_state:
        st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)

    df_editado = st.data_editor(st.session_state['filas_compras'], num_rows="dynamic", use_container_width=True, hide_index=True)
    nota_p = st.text_area("Notas para el Excel (B22):")

    if st.button("💾 GUARDAR SOLICITUD", type="primary", use_container_width=True):
        df_valido = df_editado[df_editado['Descripción'] != ""].copy()
        if not df_valido.empty:
            with st.spinner("Guardando..."):
                registros = []
                for _, r in df_valido.iterrows():
                    registros.append([proximo_id, fecha_sol.strftime("%d/%m/%Y"), categoria, r['Cantidad'], r['Descripción'], "PENDIENTE", "", "", st.session_state.get('usuario','admin'), nota_p])
                
                hoja_sol.append_rows(registros)
                doc.worksheet("CONTADOR").update_acell('A2', proximo_id + 1)
                
                # Generación de Excel (Si el archivo existe)
                ruta_excel = os.path.join(os.path.dirname(__file__), "Planilla de Solicitud de compra.xlsx")
                if os.path.exists(ruta_excel):
                    wb = openpyxl.load_workbook(ruta_excel)
                    ws = wb.active
                    ws["C6"] = proximo_id
                    ws["B22"] = f"NOTA: {nota_p}"
                    for i, (idx, row) in enumerate(df_valido.iterrows(), start=8):
                        if i <= 20:
                            ws[f"C{i}"], ws[f"D{i}"] = row['Cantidad'], row['Descripción']
                    out = io.BytesIO()
                    wb.save(out)
                    st.download_button("⬇️ Descargar Planilla Excel", out.getvalue(), f"Solicitud_{proximo_id}.xlsx")
                
                st.success("Solicitud guardada. Pasa a la pestaña de Check para verla.")
                st.session_state['filas_compras'] = pd.DataFrame([{"Cantidad": 1, "Descripción": ""}] * 5)
                st.rerun()

# ---------------------------------------------------------
# PESTAÑA 2: CHECK DE ENTREGA (MARCAR RECIBIDO)
# ---------------------------------------------------------
with tab_control:
    st.header("📝 Lista de Solicitudes Pendientes")
    if not df_global.empty:
        df_p = df_global[df_global['Estatus'] == 'PENDIENTE'].copy()
        
        if df_p.empty:
            st.success("✅ ¡Todo al día! No hay compras pendientes.")
        else:
            # Mostramos la lista para que el supervisor vea qué hay
            st.dataframe(df_p[['ID_Planilla', 'Fecha_Solicitud', 'Descripcion', 'Cantidad']], use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.subheader("Finalizar Solicitud")
            
            # Selector para elegir qué ítem llegó
            opciones = df_p.apply(lambda r: f"P-{r['ID_Planilla']} | {r['Cantidad']}x {r['Descripcion']}", axis=1).tolist()
            seleccion = st.selectbox("Seleccione el ítem que acaba de RECIBIR:", opciones)
            fecha_rec = st.date_input("Fecha de Recepción real:", datetime.now())
            
            if st.button("🏁 MARCAR COMO ENTREGADO"):
                # Extraer datos de la selección
                id_sel = int(seleccion.split("|")[0].replace("P-", "").strip())
                desc_sel = seleccion.split("|")[1].split("x ", 1)[1].strip()
                
                with st.spinner("Calculando cronometría..."):
                    # Buscamos la fila en el sheet (releyendo para seguridad)
                    data_actual = hoja_sol.get_all_records()
                    for idx, row in enumerate(data_actual):
                        if int(row['ID_Planilla']) == id_sel and str(row['Descripcion']).strip() == desc_sel:
                            fila_actualizar = idx + 2 # +1 índice, +1 cabecera
                            
                            # Calculamos días pasados
                            f_ini = datetime.strptime(str(row['Fecha_Solicitud']), "%d/%m/%Y")
                            f_fin = datetime.combine(fecha_rec, datetime.min.time())
                            dias = (f_fin - f_ini).days
                            
                            # Actualizamos columnas F, G, H (Estatus, Fecha_Entrega, Dias_Resolucion)
                            hoja_sol.update(f"F{fila_actualizar}:H{fila_actualizar}", [["COMPRADO", fecha_rec.strftime("%d/%m/%Y"), dias]])
                            st.success(f"¡Entregado! Tardó {dias} días en resolverse.")
                            st.rerun()
                            break
    else:
        st.info("Aún no hay datos en la base de datos.")

# ---------------------------------------------------------
# PESTAÑA 3: AUDITORÍA Y HISTORIAL
# ---------------------------------------------------------
with tab_metricas:
    st.header("📊 Análisis de Cumplimiento")
    
    if not df_global.empty:
        df_comp = df_global[df_global['Estatus'] == 'COMPRADO'].copy()
        df_pend = df_global[df_global['Estatus'] == 'PENDIENTE'].copy()
        
        # --- ALERTA DE SOLICITUDES NO ATENDIDAS ---
        if len(df_pend) >= 15:
            st.error(f"🚨 ¡ALERTA CRÍTICA! Hay {len(df_pend)} solicitudes sin atender. El flujo está detenido.")
        elif len(df_pend) > 0:
            st.warning(f"⚠️ Atención: Tienes {len(df_pend)} solicitudes pendientes de compra.")

        # --- MÉTRICAS ---
        m1, m2, m3 = st.columns(3)
        total = len(df_global)
        efectividad = (len(df_comp) / total * 100) if total > 0 else 0
        
        m1.metric("Total Solicitado", total)
        m2.metric("Efectividad", f"{efectividad:.1f}%")
        m3.metric("Tiempo Promedio", f"{pd.to_numeric(df_comp['Dias_Resolucion'], errors='coerce').mean():.1f} días")

        st.markdown("### 📜 Historial Completo (Marcados como Entregados)")
        if not df_comp.empty:
            # Ordenamos por fecha de entrega más reciente
            df_comp['Fecha_Entrega'] = pd.to_datetime(df_comp['Fecha_Entrega'], format='%d/%m/%Y')
            st.dataframe(df_comp.sort_values('Fecha_Entrega', ascending=False), use_container_width=True, hide_index=True)
        else:
            st.info("No hay historial de compras completadas todavía.")
