# ==========================================
# Archivo: control_solicitudes.py (Control de Solicitudes — Encomiendas y Retiros)
# ==========================================
import streamlit as st
import pandas as pd
import gspread
import textwrap
import uuid
from datetime import datetime, date, timedelta
from google.oauth2.service_account import Credentials
import io

# ==========================================
# CONFIGURACIÓN DE CONEXIÓN A GOOGLE SHEETS
# ==========================================
CREDENCIALES_GOOGLE = dict(st.secrets["gcp_service_account"])
llave_sucia = CREDENCIALES_GOOGLE["private_key"]
llave_limpia = llave_sucia.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace("\\n", "").replace("\n", "").replace(" ", "")
llave_perfecta = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(llave_limpia, 64)) + "\n-----END PRIVATE KEY-----\n"
CREDENCIALES_GOOGLE["private_key"] = llave_perfecta

# Google Sheet dedicado a este módulo (separado del de GPS Chinitas).
GOOGLE_SHEET_KEY_SOLICITUDES = "1qKHyrJS-0qRcg5R7mGmlbFlKEdkPlAdbfEDA8trikdw"

HOJA_SOLICITUDES = "Solicitudes"
HOJA_KPIS = "KPIs"

# Encabezado exacto de la hoja maestra "Solicitudes". Si el Sheet está vacío o
# recién creado, la app lo arma sola la primera vez que se abre — no hace
# falta prepararlo a mano.
COLUMNAS_SOLICITUDES = [
    "ID", "Fecha", "Día", "Semana", "Mes", "Solicitante", "Tipo de Retiro",
    "Ruta / Destino", "Chofer Asignado", "Estado",
    "Avisado (Fecha y Hora)", "Confirmado (Fecha y Hora)", "Días para Completarse",
]

TIPOS_RETIRO = ["Encomienda", "Retiro de Mercancía"]

ESTADO_PENDIENTE = "Pendiente"
ESTADO_AVISADO = "Avisado"
ESTADO_COMPLETADA = "Completada"

DIAS_ES = {"Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles", "Thursday": "Jueves",
           "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"}
MESES_ES = {"January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril", "May": "Mayo",
            "June": "Junio", "July": "Julio", "August": "Agosto", "September": "Septiembre",
            "October": "Octubre", "November": "Noviembre", "December": "Diciembre"}

# ==========================================
# CONEXIÓN Y AUTO-CONFIGURACIÓN DEL SHEET
# ==========================================
def obtener_cliente_sheets():
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credenciales = Credentials.from_service_account_info(CREDENCIALES_GOOGLE, scopes=alcance)
    return gspread.authorize(credenciales)

def _configurar_formulas_kpi(ws_kpi):
    """
    Escribe fórmulas de Google Sheets (no valores fijos) para que el resumen
    de KPIs se recalcule solo cada vez que cambie algo en "Solicitudes",
    incluso si nadie abrió la app — sirve para que el jefe lo revise
    directo desde Sheets sin depender de Streamlit.
    """
    filas = [
        ["Indicador", "Valor"],
        ["Total de solicitudes", "=COUNTA(Solicitudes!A2:A)"],
        ["Pendientes", '=COUNTIF(Solicitudes!J2:J,"Pendiente")'],
        ["Avisadas (esperando retiro)", '=COUNTIF(Solicitudes!J2:J,"Avisado")'],
        ["Completadas", '=COUNTIF(Solicitudes!J2:J,"Completada")'],
        ["% Completadas", '=IFERROR(ROUND(B5/B2*100,1),0)&"%"'],
        ["Promedio de días para completarse", "=IFERROR(ROUND(AVERAGE(Solicitudes!M2:M),1),0)"],
        ["Solicitudes de hoy", '=COUNTIF(Solicitudes!B2:B,TEXT(TODAY(),"dd/mm/yyyy"))'],
        ["Completadas hoy", '=COUNTIFS(Solicitudes!J2:J,"Completada",Solicitudes!L2:L,">="&TEXT(TODAY(),"dd/mm/yyyy"))'],
    ]
    ws_kpi.update("A1", filas, value_input_option="USER_ENTERED")
    try:
        ws_kpi.format("A1:B1", {"textFormat": {"bold": True}})
        ws_kpi.freeze(rows=1)
    except Exception:
        pass

@st.cache_resource(show_spinner=False)
def asegurar_estructura_sheet():
    """
    Se ejecuta una sola vez por sesión: conecta al Google Sheet y crea las
    hojas "Solicitudes" y "KPIs" con su encabezado/fórmulas si todavía no
    existen (por ejemplo, si el Sheet está recién creado y vacío). Si ya
    existen, no las toca — solo se asegura de que el encabezado de
    "Solicitudes" esté completo.
    """
    cliente = obtener_cliente_sheets()
    doc = cliente.open_by_key(GOOGLE_SHEET_KEY_SOLICITUDES)
    nombres = [ws.title for ws in doc.worksheets()]

    # --- Hoja "Solicitudes" ---
    if HOJA_SOLICITUDES in nombres:
        ws_sol = doc.worksheet(HOJA_SOLICITUDES)
    else:
        # Si hay una hoja por defecto vacía (p. ej. "Sheet1" / "Hoja 1"), la
        # reusamos renombrándola en vez de dejarla huérfana sin usar.
        ws_sol = None
        for ws in doc.worksheets():
            if ws.title in ("Sheet1", "Hoja 1", "Hoja1") and not ws.get_all_values():
                ws.update_title(HOJA_SOLICITUDES)
                ws_sol = ws
                break
        if ws_sol is None:
            ws_sol = doc.add_worksheet(title=HOJA_SOLICITUDES, rows=1000, cols=len(COLUMNAS_SOLICITUDES))

    primera_fila = ws_sol.row_values(1)
    if primera_fila != COLUMNAS_SOLICITUDES:
        ws_sol.update("A1", [COLUMNAS_SOLICITUDES])
        try:
            ws_sol.format("A1:M1", {"textFormat": {"bold": True}})
            ws_sol.freeze(rows=1)
        except Exception:
            pass

    # --- Hoja "KPIs" ---
    nombres = [ws.title for ws in doc.worksheets()]
    if HOJA_KPIS in nombres:
        ws_kpi = doc.worksheet(HOJA_KPIS)
    else:
        ws_kpi = doc.add_worksheet(title=HOJA_KPIS, rows=20, cols=3)
        _configurar_formulas_kpi(ws_kpi)

    return ws_sol, ws_kpi

# ==========================================
# LECTURA / ESCRITURA DE DATOS
# ==========================================
def leer_solicitudes(ws_sol):
    registros = ws_sol.get_all_records()
    df = pd.DataFrame(registros)
    if df.empty:
        df = pd.DataFrame(columns=COLUMNAS_SOLICITUDES)
    return df

def crear_solicitud(ws_sol, solicitante, tipo_retiro, ruta, chofer, fecha_solicitud):
    """
    fecha_solicitud: objeto date de la solicitud (por defecto hoy, pero
    editable desde el formulario — por ejemplo para cargar algo que pidieron
    el lunes aunque se registre en el sistema más tarde).
    """
    id_nuevo = uuid.uuid4().hex[:8].upper()
    dia_nombre = DIAS_ES.get(fecha_solicitud.strftime("%A"), fecha_solicitud.strftime("%A"))
    mes_nombre = MESES_ES.get(fecha_solicitud.strftime("%B"), fecha_solicitud.strftime("%B"))
    semana = fecha_solicitud.strftime("%W")
    fila = [
        id_nuevo, fecha_solicitud.strftime("%d/%m/%Y"), dia_nombre, semana, mes_nombre,
        solicitante.strip().upper(), tipo_retiro, ruta.strip().upper(), chofer.strip().upper(),
        ESTADO_PENDIENTE, "", "", "",
    ]
    ws_sol.append_row(fila, value_input_option="USER_ENTERED")
    return id_nuevo

def _buscar_fila(ws_sol, id_solicitud):
    celda = ws_sol.find(id_solicitud, in_column=1)
    if not celda:
        raise ValueError(f"No encontré la solicitud {id_solicitud} en el Sheet (¿la borraron o cambiaron el ID?).")
    return celda.row

def marcar_avisado(ws_sol, id_solicitud):
    fila = _buscar_fila(ws_sol, id_solicitud)
    ahora = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    ws_sol.update(f"J{fila}:K{fila}", [[ESTADO_AVISADO, ahora]], value_input_option="USER_ENTERED")

def marcar_completada(ws_sol, id_solicitud, fecha_solicitud_str):
    fila = _buscar_fila(ws_sol, id_solicitud)
    ahora = datetime.now()
    try:
        fecha_solicitud = datetime.strptime(fecha_solicitud_str, "%d/%m/%Y").date()
        dias = (ahora.date() - fecha_solicitud).days
    except Exception:
        dias = ""
    ws_sol.update(
        f"J{fila}:M{fila}",
        [[ESTADO_COMPLETADA, ws_sol.acell(f"K{fila}").value or "", ahora.strftime("%d/%m/%Y %I:%M %p"), dias]],
        value_input_option="USER_ENTERED",
    )

# ==========================================
# INTERFAZ STREAMLIT
# ==========================================
st.title("📋 Control de Solicitudes")
st.caption("Encomiendas y retiros de mercancía — pedido → avisado al chofer → retiro confirmado.")

try:
    ws_sol, ws_kpi = asegurar_estructura_sheet()
except Exception as e:
    st.error(f"No pude conectar con el Google Sheet de Control de Solicitudes: {e}")
    st.info("Verifica que el Sheet esté compartido (como Editor) con el correo de la cuenta de servicio "
            "que usa esta app (el `client_email` que está en tus Secrets, dentro de `gcp_service_account`).")
    st.stop()

if "recargar_solicitudes" not in st.session_state:
    st.session_state["recargar_solicitudes"] = 0

col_recarga, _ = st.columns([1, 4])
if col_recarga.button("🔄 Actualizar datos"):
    st.session_state["recargar_solicitudes"] += 1

df = leer_solicitudes(ws_sol)

# ---------------------------------------------------------
# ALERTAS
# ---------------------------------------------------------
if not df.empty:
    ahora = datetime.now()

    def _horas_desde(valor_fecha_hora):
        try:
            dt = datetime.strptime(valor_fecha_hora, "%d/%m/%Y %I:%M %p")
            return (ahora - dt).total_seconds() / 3600
        except Exception:
            return None

    pendientes_viejas = df[(df["Estado"] == ESTADO_PENDIENTE)]
    avisadas_sin_retiro = df[df["Estado"] == ESTADO_AVISADO].copy()
    if not avisadas_sin_retiro.empty:
        avisadas_sin_retiro["_horas"] = avisadas_sin_retiro["Avisado (Fecha y Hora)"].apply(_horas_desde)
        avisadas_demoradas = avisadas_sin_retiro[avisadas_sin_retiro["_horas"] >= 24]
    else:
        avisadas_demoradas = avisadas_sin_retiro

    if len(pendientes_viejas):
        st.warning(f"🔴 {len(pendientes_viejas)} solicitud(es) todavía sin avisarle al chofer.")
    if len(avisadas_demoradas):
        placas_demoradas = ", ".join(avisadas_demoradas["ID"].tolist())
        st.error(f"⚠️ {len(avisadas_demoradas)} solicitud(es) llevan más de 24 horas avisadas al chofer y "
                 f"todavía no se confirma el retiro (ID: {placas_demoradas}). Revísalas.")

st.markdown("---")

# ---------------------------------------------------------
# 1. CREAR NUEVA SOLICITUD
# ---------------------------------------------------------
st.subheader("➕ Nueva solicitud")
with st.form("form_nueva_solicitud", clear_on_submit=True):
    fc1, fc2 = st.columns(2)
    solicitante = fc1.text_input("Solicitante (Supervisor):")
    tipo_retiro = fc2.selectbox("Tipo de Retiro:", TIPOS_RETIRO)
    fc3, fc4 = st.columns(2)
    ruta = fc3.text_input("Ruta / Destino:")
    chofer = fc4.text_input("Chofer Asignado:")
    fecha_solicitud = st.date_input(
        "Fecha de la solicitud:", value=date.today(),
        help="Por defecto es hoy, pero la puedes cambiar — por ejemplo si estás cargando algo que pidieron el lunes."
    )
    enviado = st.form_submit_button("➕ Crear Solicitud", type="primary", use_container_width=True)

    if enviado:
        if not solicitante or not ruta or not chofer:
            st.error("Completa Solicitante, Ruta/Destino y Chofer Asignado.")
        else:
            nuevo_id = crear_solicitud(ws_sol, solicitante, tipo_retiro, ruta, chofer, fecha_solicitud)
            st.success(f"✅ Solicitud {nuevo_id} creada como Pendiente ({fecha_solicitud.strftime('%d/%m/%Y')}).")
            st.session_state["recargar_solicitudes"] += 1
            st.rerun()

st.markdown("---")

# ---------------------------------------------------------
# 2. PESTAÑAS: PENDIENTES / AVISADAS / COMPLETADAS / KPIs / INFORME
# ---------------------------------------------------------
tab_pend, tab_avis, tab_comp, tab_kpi, tab_informe = st.tabs(
    ["🔴 Pendientes", "🟡 Avisadas", "🟢 Completadas", "📊 KPIs", "🧾 Informe"]
)

with tab_pend:
    df_pend = df[df["Estado"] == ESTADO_PENDIENTE] if not df.empty else df
    if df_pend.empty:
        st.info("No hay solicitudes pendientes por avisar. 🎉")
    for _, r in df_pend.iterrows():
        with st.container(border=True):
            cA, cB = st.columns([4, 1])
            cA.markdown(f"**{r['ID']}** — {r['Tipo de Retiro']} · {r['Ruta / Destino']}  \n"
                        f"Solicitó: {r['Solicitante']} ({r['Fecha']}) · Chofer: {r['Chofer Asignado']}")
            if cB.button("📣 Ya avisé al chofer", key=f"avisar_{r['ID']}", use_container_width=True):
                marcar_avisado(ws_sol, r["ID"])
                st.session_state["recargar_solicitudes"] += 1
                st.rerun()

with tab_avis:
    df_avis = df[df["Estado"] == ESTADO_AVISADO] if not df.empty else df
    if df_avis.empty:
        st.info("No hay solicitudes esperando confirmación de retiro.")
    for _, r in df_avis.iterrows():
        with st.container(border=True):
            cA, cB = st.columns([4, 1])
            cA.markdown(f"**{r['ID']}** — {r['Tipo de Retiro']} · {r['Ruta / Destino']}  \n"
                        f"Chofer: {r['Chofer Asignado']} · Avisado: {r['Avisado (Fecha y Hora)']}")
            if cB.button("✅ Chofer confirmó el retiro", key=f"confirmar_{r['ID']}", use_container_width=True):
                marcar_completada(ws_sol, r["ID"], r["Fecha"])
                st.session_state["recargar_solicitudes"] += 1
                st.rerun()

with tab_comp:
    df_comp = df[df["Estado"] == ESTADO_COMPLETADA] if not df.empty else df
    if df_comp.empty:
        st.info("Todavía no hay solicitudes completadas.")
    else:
        st.dataframe(
            df_comp[["ID", "Fecha", "Solicitante", "Tipo de Retiro", "Ruta / Destino", "Chofer Asignado",
                     "Confirmado (Fecha y Hora)", "Días para Completarse"]],
            use_container_width=True, hide_index=True,
        )

with tab_kpi:
    total = len(df)
    n_pend = int((df["Estado"] == ESTADO_PENDIENTE).sum()) if total else 0
    n_avis = int((df["Estado"] == ESTADO_AVISADO).sum()) if total else 0
    n_comp = int((df["Estado"] == ESTADO_COMPLETADA).sum()) if total else 0
    dias_prom = pd.to_numeric(df["Días para Completarse"], errors="coerce").dropna().mean() if total else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total", total)
    k2.metric("🔴 Pendientes", n_pend)
    k3.metric("🟡 Avisadas", n_avis)
    k4.metric("🟢 Completadas", n_comp)
    st.metric("Promedio de días para completarse", f"{dias_prom:,.1f} días" if dias_prom else "—")

    st.caption("Este mismo resumen también está disponible con fórmulas en vivo dentro de la hoja "
               "\"KPIs\" del Google Sheet, para verlo sin abrir la app.")

    if total and "Chofer Asignado" in df.columns:
        st.markdown("**Por chofer:**")
        resumen_chofer = df.groupby("Chofer Asignado")["Estado"].value_counts().unstack(fill_value=0)
        st.dataframe(resumen_chofer, use_container_width=True)

with tab_informe:
    st.markdown("**Descargar informe (Excel)**")
    rango = st.radio("Rango:", ["Hoy", "Última semana", "Último mes", "Todo"], horizontal=True)

    if not df.empty:
        df_fecha = df.copy()
        df_fecha["_fecha_dt"] = pd.to_datetime(df_fecha["Fecha"], format="%d/%m/%Y", errors="coerce")
        hoy = pd.Timestamp(date.today())
        if rango == "Hoy":
            df_informe = df_fecha[df_fecha["_fecha_dt"] == hoy]
        elif rango == "Última semana":
            df_informe = df_fecha[df_fecha["_fecha_dt"] >= hoy - pd.Timedelta(days=7)]
        elif rango == "Último mes":
            df_informe = df_fecha[df_fecha["_fecha_dt"] >= hoy - pd.Timedelta(days=30)]
        else:
            df_informe = df_fecha
        df_informe = df_informe.drop(columns=["_fecha_dt"])
    else:
        df_informe = df

    st.dataframe(df_informe, use_container_width=True, hide_index=True)

    if not df_informe.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_informe.to_excel(writer, index=False, sheet_name="Solicitudes")
        st.download_button(
            "⬇️ Descargar Excel",
            data=buffer.getvalue(),
            file_name=f"informe_solicitudes_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.caption("No hay solicitudes en ese rango para exportar.")
