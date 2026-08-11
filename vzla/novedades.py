# ==========================================
# Archivo: novedades.py
# Módulo: Novedades y Status Diario (Sheet nuevo, separado)
# PARTE 1: Novedades de Despacho / Encomiendas / Novedades por Ruta
# ==========================================
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import pandas as pd
from datetime import datetime

# ==========================================
# CONFIGURACIÓN
# ==========================================
# ID del Google Sheet NUEVO (separado del principal)
SHEET_ID = "1D7w0ABnnatGd83TpJHFeVxxLOKRqoBBFbM9FyYYdFEg"

# Definición de las 5 hojas y sus encabezados exactos.
# (Status_Dia e Historial_Status se crean ya para tener todo listo; se usan en la Parte 2 y 3.)
HOJAS = {
    "Novedades_Despacho": ["ID", "FECHA RECLAMO", "FECHA ATENCIÓN", "CLIENTE/FARMACIA", "CÓDIGO",
                            "RUTA", "MOLÉCULA", "CANTIDAD", "NOVEDAD", "CONTEXTO", "ESTADO", "REGISTRADO POR"],
    "Encomiendas": ["ID", "FECHA", "DÍA", "MES", "MOVIMIENTO", "RUTA", "UNIDAD", "CHOFER",
                    "AYUDANTE", "TIPO DE ENCOMIENDA", "DETALLE", "REGISTRADO POR"],
    "Novedades_Ruta": ["ID", "FECHA", "DÍA", "MES", "RUTA", "UNIDAD", "CHOFER", "AYUDANTE",
                       "TIPO DE NOVEDAD", "DESCRIPCIÓN", "REGISTRADO POR"],
    "Status_Dia": ["ZONA", "UNIDAD", "CHOFER", "AYUDANTE", "RUTA/DESPACHO", "HORA",
                   "UBICACIÓN ACTUAL", "STATUS"],
    "Historial_Status": ["FECHA", "HORA DE CORTE", "ZONA", "UNIDAD", "CHOFER", "AYUDANTE",
                         "RUTA/DESPACHO", "HORA", "UBICACIÓN", "STATUS"],
}

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
MESES_ANO = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# ==========================================
# CONEXIÓN A GOOGLE SHEETS (mismo patrón que el resto del proyecto)
# ==========================================
@st.cache_resource
def obtener_cliente_sheets():
    cred = dict(st.secrets["gcp_service_account"])
    llave_sucia = cred["private_key"]
    llave_limpia = (llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "")
                                .replace("-----END PRIVATE KEY-----", "")
                                .replace("\\n", "").replace("\n", "").replace(" ", ""))
    cred["private_key"] = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(cred, scopes=alcance)
    return gspread.authorize(credenciales)

def abrir_libro():
    return obtener_cliente_sheets().open_by_key(SHEET_ID)

def asegurar_estructura(libro):
    """Crea las hojas que falten con sus encabezados. Solo la primera vez."""
    existentes = {w.title: w for w in libro.worksheets()}
    for nombre, cols in HOJAS.items():
        if nombre not in existentes:
            ws = libro.add_worksheet(title=nombre, rows=500, cols=max(12, len(cols) + 1))
            ws.update([cols])
        else:
            ws = existentes[nombre]
            if not ws.row_values(1):
                ws.update([cols])
    # Eliminar la hoja por defecto vacía ("Hoja 1"/"Sheet1") si sigue ahí
    for basura in ["Hoja 1", "Hoja1", "Sheet1"]:
        if basura in [w.title for w in libro.worksheets()] and basura not in HOJAS:
            try:
                libro.del_worksheet(libro.worksheet(basura))
            except Exception:
                pass

def nuevo_id():
    return datetime.now().strftime("%Y%m%d-%H%M%S")

def dia_y_mes(fecha_obj):
    return DIAS_SEMANA[fecha_obj.weekday()], MESES_ANO[fecha_obj.month - 1]

def guardar_fila(nombre_hoja, fila_lista):
    ws = abrir_libro().worksheet(nombre_hoja)
    ws.append_row(fila_lista, value_input_option="USER_ENTERED")

def leer_hoja_df(nombre_hoja):
    try:
        registros = abrir_libro().worksheet(nombre_hoja).get_all_records()
        return pd.DataFrame(registros)
    except Exception:
        return pd.DataFrame()

# ==========================================
# INTERFAZ
# ==========================================
st.title("📝 Novedades y Status Diario")

usuario = str(st.session_state.get("usuario", "-")).upper()
st.caption(f"Registrando como: **{usuario}**")

# Preparar el libro y la estructura (una sola vez por sesión)
try:
    libro = abrir_libro()
    if not st.session_state.get("nov_estructura_ok", False):
        asegurar_estructura(libro)
        st.session_state["nov_estructura_ok"] = True
except Exception as e:
    st.error("No se pudo conectar con el Google Sheet nuevo.")
    st.info("Verifica que el Sheet esté compartido como **Editor** con el bot "
            "`bot-pizarra@pcd-drotaca.iam.gserviceaccount.com`.")
    st.exception(e)
    st.stop()

tab1, tab2, tab3 = st.tabs(["🚨 Novedades de Despacho", "📦 Encomiendas", "🛞 Novedades por Ruta"])

# ------------------------------------------
# TAB 1: NOVEDADES DE DESPACHO
# ------------------------------------------
with tab1:
    st.subheader("🚨 Registrar novedad de despacho")
    with st.form("form_desp", clear_on_submit=True):
        c1, c2 = st.columns(2)
        f_reclamo = c1.date_input("Fecha de reclamo", format="DD/MM/YYYY")
        ya_atendida = c2.checkbox("¿Ya fue atendida?")
        f_atencion = c2.date_input("Fecha de atención", format="DD/MM/YYYY", disabled=not ya_atendida)

        cliente = st.text_input("Cliente / Farmacia")
        c3, c4 = st.columns(2)
        codigo = c3.text_input("Código")
        ruta = c4.text_input("Ruta")
        molecula = st.text_input("Molécula / Producto")
        cantidad = st.text_input("Cantidad")
        novedad = st.text_input("Novedad (ej: Producto en mal estado)")
        contexto = st.text_area("Contexto / Detalle")

        if st.form_submit_button("💾 Guardar novedad de despacho", type="primary", use_container_width=True):
            estado = "Atendida" if ya_atendida else "Pendiente"
            f_at_str = f_atencion.strftime("%d/%m/%Y") if ya_atendida else ""
            fila = [nuevo_id(), f_reclamo.strftime("%d/%m/%Y"), f_at_str, cliente, codigo,
                    ruta, molecula, cantidad, novedad, contexto, estado, usuario]
            try:
                guardar_fila("Novedades_Despacho", fila)
                st.success("✅ Novedad de despacho guardada.")
            except Exception as e:
                st.error(f"No se pudo guardar: {e}")

    with st.expander("📄 Ver registros de despacho"):
        df = leer_hoja_df("Novedades_Despacho")
        if df.empty:
            st.info("Aún no hay registros.")
        else:
            st.dataframe(df.iloc[::-1], use_container_width=True, hide_index=True)

# ------------------------------------------
# TAB 2: ENCOMIENDAS (RETIRO / ENTREGA)
# ------------------------------------------
with tab2:
    st.subheader("📦 Registrar retiro / entrega de encomienda")
    with st.form("form_enc", clear_on_submit=True):
        c1, c2 = st.columns(2)
        fecha = c1.date_input("Fecha", format="DD/MM/YYYY")
        movimiento = c2.selectbox("Movimiento", ["Retiro", "Entrega"])
        ruta = st.text_input("Ruta")
        c3, c4 = st.columns(2)
        unidad = c3.text_input("Unidad")
        chofer = c4.text_input("Chofer")
        ayudante = st.text_input("Ayudante")
        tipo_enc = st.text_input("Tipo de encomienda")
        detalle = st.text_area("Detalle (opcional)")

        if st.form_submit_button("💾 Guardar encomienda", type="primary", use_container_width=True):
            dia, mes = dia_y_mes(fecha)
            fila = [nuevo_id(), fecha.strftime("%d/%m/%Y"), dia, mes, movimiento, ruta,
                    unidad, chofer, ayudante, tipo_enc, detalle, usuario]
            try:
                guardar_fila("Encomiendas", fila)
                st.success("✅ Encomienda guardada.")
            except Exception as e:
                st.error(f"No se pudo guardar: {e}")

    with st.expander("📄 Ver registros de encomiendas"):
        df = leer_hoja_df("Encomiendas")
        if df.empty:
            st.info("Aún no hay registros.")
        else:
            st.dataframe(df.iloc[::-1], use_container_width=True, hide_index=True)

# ------------------------------------------
# TAB 3: NOVEDADES GENERALES POR RUTA
# ------------------------------------------
with tab3:
    st.subheader("🛞 Registrar novedad general por ruta")
    with st.form("form_ruta", clear_on_submit=True):
        c1, c2 = st.columns(2)
        fecha = c1.date_input("Fecha", format="DD/MM/YYYY", key="f_ruta")
        ruta = c2.text_input("Ruta")
        c3, c4 = st.columns(2)
        unidad = c3.text_input("Unidad")
        chofer = c4.text_input("Chofer")
        ayudante = st.text_input("Ayudante")
        tipo_nov = st.text_input("Tipo de novedad")
        descripcion = st.text_area("Descripción")

        if st.form_submit_button("💾 Guardar novedad por ruta", type="primary", use_container_width=True):
            dia, mes = dia_y_mes(fecha)
            fila = [nuevo_id(), fecha.strftime("%d/%m/%Y"), dia, mes, ruta, unidad,
                    chofer, ayudante, tipo_nov, descripcion, usuario]
            try:
                guardar_fila("Novedades_Ruta", fila)
                st.success("✅ Novedad por ruta guardada.")
            except Exception as e:
                st.error(f"No se pudo guardar: {e}")

    with st.expander("📄 Ver registros por ruta"):
        df = leer_hoja_df("Novedades_Ruta")
        if df.empty:
            st.info("Aún no hay registros.")
        else:
            st.dataframe(df.iloc[::-1], use_container_width=True, hide_index=True)