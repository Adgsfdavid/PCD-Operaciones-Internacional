# ==========================================
# Archivo: novedades.py
# Modulo: Novedades y Status Diario (Sheet nuevo, separado)
# VISOR DE SOLO LECTURA - el personal escribe en el Google Sheet;
# aqui solo se ve el resultado consolidado.
# ==========================================
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import pandas as pd
from datetime import datetime, date

# ==========================================
# CONFIGURACION
# ==========================================
SHEET_ID = "1D7w0ABnnatGd83TpJHFeVxxLOKRqoBBFbM9FyYYdFEg"

HOJAS = {
    "Novedades_Despacho": ["FECHA RECLAMO", "FECHA ATENCIÓN", "CLIENTE/FARMACIA", "CÓDIGO",
                            "RUTA", "MOLÉCULA", "CANTIDAD", "NOVEDAD", "CONTEXTO", "ESTADO", "REGISTRADO POR"],
    "Encomiendas": ["FECHA", "MOVIMIENTO", "RUTA", "UNIDAD", "CHOFER",
                    "AYUDANTE", "TIPO DE ENCOMIENDA", "DETALLE", "REGISTRADO POR"],
    "Novedades_Ruta": ["FECHA", "RUTA", "UNIDAD", "CHOFER", "AYUDANTE",
                       "TIPO DE NOVEDAD", "DESCRIPCIÓN", "REGISTRADO POR"],
    "Status_Dia": ["ZONA", "UNIDAD", "CHOFER", "AYUDANTE", "RUTA/DESPACHO", "HORA",
                   "UBICACIÓN ACTUAL", "STATUS"],
    "Historial_Status": ["FECHA", "HORA DE CORTE", "ZONA", "UNIDAD", "CHOFER", "AYUDANTE",
                         "RUTA/DESPACHO", "HORA", "UBICACIÓN", "STATUS"],
}

COL_FECHA = {
    "Novedades_Despacho": "FECHA RECLAMO",
    "Encomiendas": "FECHA",
    "Novedades_Ruta": "FECHA",
}

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
MESES_ANO = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# ==========================================
# CONEXION A GOOGLE SHEETS
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
    existentes = {w.title: w for w in libro.worksheets()}
    for nombre, cols in HOJAS.items():
        if nombre not in existentes:
            ws = libro.add_worksheet(title=nombre, rows=500, cols=max(12, len(cols) + 1))
            ws.update([cols])
        else:
            ws = existentes[nombre]
            if not ws.row_values(1):
                ws.update([cols])
    for basura in ["Hoja 1", "Hoja1", "Sheet1"]:
        if basura in [w.title for w in libro.worksheets()] and basura not in HOJAS:
            try:
                libro.del_worksheet(libro.worksheet(basura))
            except Exception:
                pass

@st.cache_data(ttl=60, show_spinner=False)
def leer_hoja_df(nombre_hoja):
    try:
        registros = abrir_libro().worksheet(nombre_hoja).get_all_records()
        return pd.DataFrame(registros)
    except Exception:
        return pd.DataFrame()

# ==========================================
# UTILIDADES
# ==========================================
def agregar_dia_mes(df, col_fecha):
    if df.empty or col_fecha not in df.columns:
        return df
    def _dia(v):
        try:
            d = pd.to_datetime(str(v), dayfirst=True)
            return DIAS_SEMANA[d.weekday()]
        except Exception:
            return ""
    def _mes(v):
        try:
            d = pd.to_datetime(str(v), dayfirst=True)
            return MESES_ANO[d.month - 1]
        except Exception:
            return ""
    df = df.copy()
    df.insert(1, "DÍA", df[col_fecha].apply(_dia))
    df.insert(2, "MES", df[col_fecha].apply(_mes))
    return df

def filtrar_por_dia(df, col_fecha, dia_sel):
    if df.empty or col_fecha not in df.columns:
        return df
    obj = pd.to_datetime(df[col_fecha].astype(str), dayfirst=True, errors="coerce")
    return df[obj.dt.date == dia_sel]

# ==========================================
# INTERFAZ (VISOR)
# ==========================================
st.title("📝 Novedades del Día")
st.caption("Vista de solo lectura. El personal registra directo en el Google Sheet; aquí ves el consolidado.")

try:
    libro = abrir_libro()
    if not st.session_state.get("nov_estructura_ok", False):
        asegurar_estructura(libro)
        st.session_state["nov_estructura_ok"] = True
except Exception as e:
    st.error("No se pudo conectar con el Google Sheet de novedades.")
    st.info("Verifica que el Sheet esté compartido como **Editor** con el bot "
            "`bot-pizarra@pcd-drotaca.iam.gserviceaccount.com`.")
    st.exception(e)
    st.stop()

c1, c2, c3 = st.columns([2, 2, 1])
dia_sel = c1.date_input("📅 Ver día", value=date.today(), format="DD/MM/YYYY")
ver_todo = c2.checkbox("Ver todo el histórico (ignorar día)")
if c3.button("🔄 Actualizar", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

df_desp = leer_hoja_df("Novedades_Despacho")
df_enc = leer_hoja_df("Encomiendas")
df_ruta = leer_hoja_df("Novedades_Ruta")

def preparar(df, hoja):
    df = agregar_dia_mes(df, COL_FECHA[hoja])
    if not ver_todo:
        df = filtrar_por_dia(df, COL_FECHA[hoja], dia_sel)
    return df

vd = preparar(df_desp, "Novedades_Despacho")
ve = preparar(df_enc, "Encomiendas")
vr = preparar(df_ruta, "Novedades_Ruta")

m1, m2, m3 = st.columns(3)
etiqueta = "histórico" if ver_todo else dia_sel.strftime("%d/%m/%Y")
m1.metric(f"🚨 Novedades despacho ({etiqueta})", len(vd))
m2.metric(f"📦 Encomiendas ({etiqueta})", len(ve))
m3.metric(f"🛞 Novedades por ruta ({etiqueta})", len(vr))

st.markdown("---")

tab1, tab2, tab3 = st.tabs(["🚨 Novedades de Despacho", "📦 Encomiendas", "🛞 Novedades por Ruta"])

with tab1:
    if vd.empty:
        st.info("Sin novedades de despacho para el filtro seleccionado.")
    else:
        st.dataframe(vd.iloc[::-1], use_container_width=True, hide_index=True)

with tab2:
    if ve.empty:
        st.info("Sin encomiendas para el filtro seleccionado.")
    else:
        st.dataframe(ve.iloc[::-1], use_container_width=True, hide_index=True)

with tab3:
    if vr.empty:
        st.info("Sin novedades por ruta para el filtro seleccionado.")
    else:
        st.dataframe(vr.iloc[::-1], use_container_width=True, hide_index=True)
