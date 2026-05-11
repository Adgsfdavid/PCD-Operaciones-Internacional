# ==========================================
# Archivo: compras_flota.py (Control y Cronometría de Compras)
# ==========================================
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import openpyxl
import io
import os

# ==========================================
# CONEXIÓN A GOOGLE SHEETS (NUEVA BASE DE DATOS)
# ==========================================
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
    # CONEXIÓN DIRECTA A TU GOOGLE SHEET DE COMPRAS
    return obtener_cliente_sheets().open_by_key("1M9VQOHU6LHniSBu3A_o2qDFlZITrDRm_f7qYGU5iOX8")

# ==========================================
# INTERFAZ Y PESTAÑAS
# ==========================================
st.title("🛒 Gestión y Cronometría de Compras")
st.markdown("Generación de planillas oficiales y auditoría de tiempos de entrega.")

tab_nueva, tab_control, tab_metricas = st.tabs(["📝 Generar Solicitud", "⏳ Control de Entregas", "📊 Auditoría de Tiempos"])

# ---------------------------------------------------------
# PESTAÑA 1: NUEVA SOLICITUD Y GENERACIÓN DE EXCEL
# ---------------------------------------------------------
with tab_nueva:
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            try:
                hoja_cnt = conectar_bd_compras().worksheet("CONTADOR")
                proximo_id = int(hoja_cnt.acell('A2').value)
            except:
                proximo_id = 1
                
            st.subheader(f"Planilla N° {proximo_id}")
            fecha_sol = st.date_input("Fecha de Solicitud:", datetime.now())
        with c2:
            categoria = st.selectbox("Categoría:", ["Repuestos Mecánicos", "Consumibles", "Herramientas", "Oficina/Limpieza", "Otros"])

    if 'filas_compras' not in st.session_state:
        st.session_state['filas_compras'] = pd.DataFrame([
            {"Cantidad": 0, "Descripción": ""}, {"Cantidad": 0, "Descripción": ""}, {"Cantidad": 0, "Descripción": ""}
        ])

    st.markdown("### 📋 Ítems a solicitar (Máx. 13 ítems por planilla)")
    df_editado = st.data_editor(
        st.session_state['filas_compras'], num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Cantidad": st.column_config.NumberColumn(min_value=1, step=1),
            "Descripción": st.column_config.TextColumn(placeholder="Ej: Tornillo 3/8 para Unidad Cargo 815")
        }
    )
    nota_planilla = st.text_area("Nota o justificación (Aparecerá en la B22 del Excel):", max_chars=150)

    if st.button("💾 GUARDAR SOLICITUD Y GENERAR EXCEL", type="primary", use_container_width=True):
        df_valido = df_editado[(df_editado['Descripción'] != "") & (df_editado['Cantidad'] > 0)].copy()
        
        if df_valido.empty or len(df_valido) > 13:
            st.error("Debes agregar entre 1 y 13 ítems válidos para que encajen en el Excel (C8 a C20).")
        else:
            with st.spinner("Guardando en la nube y ensamblando Excel..."):
                try:
                    doc = conectar_bd_compras()
                    try:
                        hoja_sol = doc.worksheet("SOLICITUDES")
                    except:
                        hoja_sol = doc.add_worksheet("SOLICITUDES", 1000, 15)
                        hoja_sol.append_row(["ID_Planilla", "Fecha_Solicitud", "Categoria", "Cantidad", "Descripcion", "Estatus", "Fecha_Entrega", "Dias_Resolucion", "Usuario", "Nota"])
                    
                    # 1. Guardar cada ítem en Google Sheets
                    registros = []
                    for _, fila in df_valido.iterrows():
                        registros.append([
                            proximo_id, fecha_sol.strftime("%d/%m/%Y"), categoria, 
                            fila['Cantidad'], fila['Descripción'], "PENDIENTE", "", "", 
                            st.session_state.get('usuario', 'admin_vzla'), nota_planilla
                        ])
                    hoja_sol.append_rows(registros)
                    
                    try:
                        doc.worksheet("CONTADOR").update_acell('A2', proximo_id + 1)
                    except:
                        doc.add_worksheet("CONTADOR", 100, 2).update('A1:A2', [['Ultimo_ID'], [proximo_id + 1]])

                    st.success(f"✅ Registros guardados en Google Sheets.")

                    # 2. Rellenar Excel Original
                    ruta_excel = os.path.join(os.path.dirname(__file__), "Planilla de Solicitud de compra.xlsx")
                    if os.path.exists(ruta_excel):
                        wb = openpyxl.load_workbook(ruta_excel)
                        ws = wb.active

                        # Rellenar N° de Planilla (En la C6 según tu formato original)
                        ws["C6"] = proximo_id
                        ws["D6"] = categoria

                        # Rellenar Ítems (C8:C20 y D8:D20)
                        fila_excel = 8
                        for _, row in df_valido.iterrows():
                            ws[f"C{fila_excel}"] = row['Cantidad']
                            ws[f"D{fila_excel}"] = row['Descripción']
                            fila_excel += 1

                        # Rellenar Nota en B22
                        if nota_planilla:
                            ws["B22"] = f"NOTA: {nota_planilla}"

                        # Preparar descarga
                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)
                        
                        st.download_button(
                            label="⬇️ DESCARGAR PLANILLA EXCEL OFICIAL",
                            data=output,
                            file_name=f"Planilla_N{proximo_id}_{fecha_sol.strftime('%d%m%y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="secondary",
                            use_container_width=True
                        )
                        del st.session_state['filas_compras']
                    else:
                        st.warning("⚠️ Datos guardados, pero no se pudo generar el Excel porque falta subir el archivo 'Planilla de Solicitud de compra.xlsx' a GitHub en la carpeta 'vzla'.")

                except Exception as e:
                    st.error(f"Error en el proceso: {e}")

# ---------------------------------------------------------
# PESTAÑA 2: CONTROL Y CRONOMETRÍA (MARCAR ENTREGADOS)
# ---------------------------------------------------------
with tab_control:
    st.markdown("### ⏳ Seguimiento de Ítems Pendientes")
    st.write("Marca los ítems que ya fueron comprados y entregados para detener el cronómetro.")
    
    try:
        hoja_sol = conectar_bd_compras().worksheet("SOLICITUDES")
        data = hoja_sol.get_all_records()
        df_sol = pd.DataFrame(data)
    except:
        df_sol = pd.DataFrame()

    if not df_sol.empty:
        df_pendientes = df_sol[df_sol['Estatus'] == 'PENDIENTE'].copy()
        
        if df_pendientes.empty:
            st.info("🎉 ¡Excelente! No hay ítems pendientes de compra.")
        else:
            # Calcular días transcurridos actuales
            df_pendientes['Fecha_Solicitud'] = pd.to_datetime(df_pendientes['Fecha_Solicitud'], format='%d/%m/%Y', errors='coerce')
            df_pendientes['Días_En_Espera'] = (datetime.now() - df_pendientes['Fecha_Solicitud']).dt.days

            # Interfaz para cerrar el ticket
            item_a_cerrar = st.selectbox("Selecciona el ítem recibido:", 
                df_pendientes.apply(lambda r: f"Planilla {r['ID_Planilla']} | {r['Cantidad']}x {r['Descripcion']} ({r['Días_En_Espera']} días esperando)", axis=1))
            
            fecha_entrega = st.date_input("Fecha en que se recibió:", datetime.now(), key="f_entrega")
            
            if st.button("✅ Marcar como ENTREGADO / COMPRADO", type="primary"):
                plan_id = int(item_a_cerrar.split("|")[0].replace("Planilla", "").strip())
                desc_item = item_a_cerrar.split("x ", 1)[1].split("(")[0].strip()
                
                with st.spinner("Actualizando cronometría..."):
                    # Buscar la fila exacta en Google Sheets
                    for i, row in enumerate(data):
                        if int(row.get('ID_Planilla', 0)) == plan_id and str(row.get('Descripcion', '')).strip() == desc_item and row.get('Estatus') == 'PENDIENTE':
                            fila_sheet = i + 2 # +1 por índice 0, +1 por encabezado
                            
                            # Calcular días de resolución
                            f_sol = datetime.strptime(str(row['Fecha_Solicitud']), '%d/%m/%Y')
                            f_ent = datetime.strptime(fecha_entrega.strftime('%d/%m/%Y'), '%d/%m/%Y')
                            dias_resolucion = (f_ent - f_sol).days
                            
                            # Actualizar hoja: Estatus(F), Fecha_Entrega(G), Dias_Resolucion(H)
                            hoja_sol.update(f"F{fila_sheet}:H{fila_sheet}", [["COMPRADO", fecha_entrega.strftime("%d/%m/%Y"), dias_resolucion]])
                            st.success(f"Ticket cerrado. Tiempo de resolución cronometrado: {dias_resolucion} días.")
                            st.rerun()
                            break

# ---------------------------------------------------------
# PESTAÑA 3: AUDITORÍA DE TIEMPOS (MÉTRICAS)
# ---------------------------------------------------------
with tab_metricas:
    st.markdown("### 📊 Eficiencia del Departamento de Compras")
    if not df_sol.empty:
        df_cerrados = df_sol[df_sol['Estatus'] == 'COMPRADO'].copy()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Solicitudes", len(df_sol))
        c2.metric("Entregados", len(df_cerrados))
        c3.metric("Pendientes", len(df_sol[df_sol['Estatus'] == 'PENDIENTE']))
        
        if not df_cerrados.empty:
            df_cerrados['Dias_Resolucion'] = pd.to_numeric(df_cerrados['Dias_Resolucion'], errors='coerce')
            promedio_dias = df_cerrados['Dias_Resolucion'].mean()
            
            st.markdown("---")
            st.subheader(f"⏱️ Tiempo Promedio de Respuesta: {promedio_dias:.1f} días")
            
            st.dataframe(df_cerrados[['ID_Planilla', 'Fecha_Solicitud', 'Descripcion', 'Fecha_Entrega', 'Dias_Resolucion']], 
                         use_container_width=True, hide_index=True)
        else:
            st.info("No hay métricas de resolución aún.")