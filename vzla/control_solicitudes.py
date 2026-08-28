# ==========================================
# Archivo: control_solicitudes.py (Control de Solicitudes — Encomiendas y Retiros)
# ==========================================
import streamlit as st
import pandas as pd
import gspread
import textwrap
import uuid
from datetime import datetime, date, time, timedelta, timezone
from google.oauth2.service_account import Credentials
import io

# Streamlit Cloud corre sus servidores en UTC, no en hora de Venezuela — sin
# esto, todo lo que use ahora_vzla() queda ~4-5 horas adelantado. Esta
# función se usa en TODO el archivo en vez de datetime.now() directo.
# Venezuela no tiene horario de verano, así que un offset fijo de -4 es
# siempre correcto (se intenta primero con la base de datos de zonas
# horarias del sistema; si el servidor no la tiene instalada, se cae al
# offset fijo para que esto nunca truene).
try:
    from zoneinfo import ZoneInfo
    VENEZUELA_TZ = ZoneInfo("America/Caracas")
except Exception:
    VENEZUELA_TZ = timezone(timedelta(hours=-4))

def ahora_vzla():
    try:
        return datetime.now(VENEZUELA_TZ).replace(tzinfo=None)
    except Exception:
        return datetime.now(timezone(timedelta(hours=-4))).replace(tzinfo=None)

try:
    from fpdf import FPDF
    FPDF_DISPONIBLE = True
except ImportError:
    FPDF_DISPONIBLE = False

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
    "Detalle", "Duración del Retiro",
]

# El tipo de retiro ya no se elige — siempre es "Retiro de Encomienda". Se deja
# como constante (en vez de un selector) para no complicar el formulario.
TIPO_RETIRO_FIJO = "Retiro de Encomienda"
# Se conserva por compatibilidad con solicitudes viejas que puedan tener otro
# valor guardado en el Sheet (no se tocan al leerlas, solo las nuevas usan el fijo).
TIPOS_RETIRO = ["Encomienda", "Retiro de Mercancía", TIPO_RETIRO_FIJO]

SOLICITANTES_FIJOS = ["JOSE SUAREZ", "PROCURA", "PROMOCION COMERCIAL", "PROMOCION MEDICA", "OTROS"]

# "Solicitante" es el ÁREA/DEPARTAMENTO que pide el retiro (ej. "PROMOCION
# COMERCIAL"), NO la persona. El supervisor encargado de gestionar todo esto
# desde el sistema es una persona real, y esa es la que debe salir como
# "Supervisor encargado" en los mensajes — se toma del usuario que inició
# sesión, no del campo Solicitante.
NOMBRES_SUPERVISORES = {
    "faisal": "FAISAL YORDI",
    "admin_vzla": "ADMINISTRADOR",
}

def nombre_supervisor_actual():
    usuario = st.session_state.get("usuario", "")
    return NOMBRES_SUPERVISORES.get(usuario, usuario.upper() if usuario else "SUPERVISOR")

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

def _letra_columna(n):
    """1 -> A, 2 -> B, ... 14 -> N (para armar rangos tipo 'A1:N1' dinámicamente)."""
    letras = ""
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letras = chr(65 + resto) + letras
    return letras

@st.cache_resource(show_spinner=False)
def obtener_documento():
    """Cachea SOLO la conexión (lo caro: autenticar y abrir el Sheet)."""
    cliente = obtener_cliente_sheets()
    return cliente.open_by_key(GOOGLE_SHEET_KEY_SOLICITUDES)

def asegurar_estructura_sheet():
    """
    Se ejecuta en CADA carga de la página (a propósito, sin @st.cache_resource):
    crea las hojas "Solicitudes" y "KPIs" con su encabezado/fórmulas si
    todavía no existen, y si el código agrega una columna nueva más adelante
    (como pasó con "Detalle"), la agrega sola sin necesitar reiniciar la app
    — antes esto estaba cacheado y por eso una columna nueva se quedaba sin
    escribir hasta reiniciar el servidor. La revisión en sí es barata (un
    par de llamadas), así que no hay problema en repetirla siempre.
    """
    doc = obtener_documento()
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
            ultima_letra = _letra_columna(len(COLUMNAS_SOLICITUDES))
            ws_sol.format(f"A1:{ultima_letra}1", {"textFormat": {"bold": True}})
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

def crear_solicitud(ws_sol, solicitante, detalle, ruta, chofer, fecha_solicitud):
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
        solicitante.strip().upper(), TIPO_RETIRO_FIJO, ruta.strip().upper(), chofer.strip().upper(),
        ESTADO_PENDIENTE, "", "", "", detalle.strip().upper(), "",
    ]
    ws_sol.append_row(fila, value_input_option="USER_ENTERED")
    return id_nuevo

def generar_mensaje_chofer(r):
    """
    Arma el mensaje ya redactado, listo para copiar y mandarle al chofer por
    WhatsApp (o el medio que sea) con toda la info de la solicitud.
    """
    ahora = ahora_vzla()
    detalle = (r.get("Detalle") or "").strip()
    lineas = [
        "📦 *SOLICITUD DE RETIRO*",
        "",
        f"🗓️ {ahora.strftime('%d/%m/%Y')} - {ahora.strftime('%I:%M %p')}",
        f"👤 Supervisor encargado: {nombre_supervisor_actual()}",
        f"🏢 Área solicitante: {r['Solicitante']}",
        "",
        f"*Tipo:* {r['Tipo de Retiro']}",
    ]
    if detalle:
        lineas.append(f"*Detalle:* {detalle}")
    lineas += [
        f"*Ruta:* {r['Ruta / Destino']}",
        f"*Chofer:* {r['Chofer Asignado']}",
        "",
        "Por favor confirmar este mensaje ✅",
    ]
    return "\n".join(lineas)

def _buscar_fila(ws_sol, id_solicitud):
    celda = ws_sol.find(id_solicitud, in_column=1)
    if not celda:
        raise ValueError(f"No encontré la solicitud {id_solicitud} en el Sheet (¿la borraron o cambiaron el ID?).")
    return celda.row

def marcar_avisado(ws_sol, id_solicitud):
    fila = _buscar_fila(ws_sol, id_solicitud)
    ahora = ahora_vzla().strftime("%d/%m/%Y %I:%M %p")
    ws_sol.update(f"J{fila}:K{fila}", [[ESTADO_AVISADO, ahora]], value_input_option="USER_ENTERED")

def marcar_completada_detallada(ws_sol, id_solicitud, fecha_solicitud_str, fecha_retiro, hora_retiro, chofer_retiro):
    """
    Confirma la entrega con fecha, hora y chofer específicos (indicados a
    mano en el mini-panel "Confirmar Entrega"), calcula cuánto se demoró en
    días y horas desde que se creó la solicitud, y lo deja todo escrito en
    el Sheet (Estado, Confirmado, Días para Completarse y Duración).
    Devuelve (duracion_texto, confirmado_dt) para armar el mensaje de WhatsApp.
    """
    fila = _buscar_fila(ws_sol, id_solicitud)
    fecha_solicitud = _parsear_fecha(fecha_solicitud_str)
    confirmado_dt = datetime.combine(fecha_retiro, hora_retiro)
    inicio_dt = datetime.combine(fecha_solicitud, time.min)
    delta = confirmado_dt - inicio_dt
    if delta.total_seconds() < 0:
        delta = timedelta(0)
    dias_totales = delta.days
    horas_resto = delta.seconds // 3600
    if dias_totales == 0 and horas_resto == 0:
        duracion_texto = "menos de 1 hora"
    elif dias_totales == 0:
        duracion_texto = f"{horas_resto}h"
    else:
        duracion_texto = f"{dias_totales}d {horas_resto}h"

    ws_sol.update(
        f"J{fila}:M{fila}",
        [[ESTADO_COMPLETADA, ws_sol.acell(f"K{fila}").value or "", confirmado_dt.strftime("%d/%m/%Y %I:%M %p"), dias_totales]],
        value_input_option="USER_ENTERED",
    )
    ws_sol.update(f"O{fila}", [[duracion_texto]], value_input_option="USER_ENTERED")

    chofer_final = chofer_retiro.strip().upper() if chofer_retiro and chofer_retiro.strip() else ""
    if chofer_final:
        ws_sol.update(f"I{fila}", [[chofer_final]], value_input_option="USER_ENTERED")

    return duracion_texto, confirmado_dt

def generar_mensaje_confirmacion(r, confirmado_dt, chofer_retiro, duracion_texto):
    """Mensaje listo para copiar y mandar por WhatsApp confirmando que el retiro se realizó."""
    detalle = (r.get("Detalle") or "").strip()
    lineas = [
        "✅ *CONFIRMACIÓN DE RETIRO*",
        "",
        f"🗓️ Solicitado: {r['Fecha']}",
        f"📦 Retirado: {confirmado_dt.strftime('%d/%m/%Y')} - {confirmado_dt.strftime('%I:%M %p')}",
        f"👤 Supervisor encargado: {nombre_supervisor_actual()}",
        f"🏢 Área solicitante: {r['Solicitante']}",
        f"🚚 Chofer: {chofer_retiro or r['Chofer Asignado']}",
    ]
    if detalle:
        lineas.append(f"*Detalle:* {detalle}")
    lineas += [
        f"*Ruta:* {r['Ruta / Destino']}",
        "",
        f"⏱️ Tiempo de gestión: {duracion_texto}",
    ]
    return "\n".join(lineas)

def actualizar_solicitud(ws_sol, id_solicitud, solicitante, detalle, ruta, chofer, fecha_solicitud):
    """Edita los datos base de una solicitud ya creada (no toca Estado/Avisado/Confirmado)."""
    fila = _buscar_fila(ws_sol, id_solicitud)
    dia_nombre = DIAS_ES.get(fecha_solicitud.strftime("%A"), fecha_solicitud.strftime("%A"))
    mes_nombre = MESES_ES.get(fecha_solicitud.strftime("%B"), fecha_solicitud.strftime("%B"))
    semana = fecha_solicitud.strftime("%W")
    ws_sol.update(
        f"B{fila}:I{fila}",
        [[fecha_solicitud.strftime("%d/%m/%Y"), dia_nombre, semana, mes_nombre,
          solicitante.strip().upper(), TIPO_RETIRO_FIJO, ruta.strip().upper(), chofer.strip().upper()]],
        value_input_option="USER_ENTERED",
    )
    ws_sol.update(f"N{fila}", [[detalle.strip().upper()]], value_input_option="USER_ENTERED")

def borrar_solicitud(ws_sol, id_solicitud):
    fila = _buscar_fila(ws_sol, id_solicitud)
    ws_sol.delete_rows(fila)

def _texto_pdf(valor):
    """FPDF (fuente Helvetica) no soporta bien tildes/ñ como UTF-8 directo;
    esto las reemplaza por su versión sin tilde para que el PDF no reviente."""
    texto = str(valor) if valor is not None else ""
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "ñ": "n", "Ñ": "N",
    }
    for viejo, nuevo in reemplazos.items():
        texto = texto.replace(viejo, nuevo)
    return texto

def generar_pdf_informe(df_informe, titulo_rango):
    """Genera el informe en PDF (apaisado) con el resumen y la tabla de solicitudes del rango elegido."""
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, _texto_pdf("Informe de Control de Solicitudes"), ln=True)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, _texto_pdf(titulo_rango), ln=True)
    pdf.cell(0, 7, _texto_pdf(f"Generado: {ahora_vzla().strftime('%d/%m/%Y %I:%M %p')}"), ln=True)
    pdf.ln(3)

    total = len(df_informe)
    n_pend = int((df_informe["Estado"] == ESTADO_PENDIENTE).sum())
    n_avis = int((df_informe["Estado"] == ESTADO_AVISADO).sum())
    n_comp = int((df_informe["Estado"] == ESTADO_COMPLETADA).sum())
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, _texto_pdf(
        f"Total: {total}   |   Pendientes: {n_pend}   |   Avisadas: {n_avis}   |   Completadas: {n_comp}"
    ), ln=True)
    pdf.ln(4)

    columnas = ["ID", "Fecha", "Solicitante", "Tipo de Retiro", "Detalle", "Ruta / Destino",
                "Chofer Asignado", "Estado", "Confirmado (Fecha y Hora)", "Días para Completarse"]
    anchos = [18, 20, 30, 26, 35, 30, 28, 22, 32, 20]

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(13, 71, 161)
    pdf.set_text_color(255, 255, 255)
    for col, ancho in zip(columnas, anchos):
        pdf.cell(ancho, 8, _texto_pdf(col), border=1, fill=True, align="C")
    pdf.ln()

    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(0, 0, 0)
    for _, r in df_informe.iterrows():
        for col, ancho in zip(columnas, anchos):
            valor = _texto_pdf(r.get(col, ""))
            if len(valor) > int(ancho * 1.8):
                valor = valor[: int(ancho * 1.8) - 1] + "…"
            pdf.cell(ancho, 7, valor, border=1)
        pdf.ln()

    return bytes(pdf.output())

def _parsear_fecha(fecha_str):
    try:
        return datetime.strptime(fecha_str, "%d/%m/%Y").date()
    except Exception:
        return date.today()

def render_tarjeta(r, ws_sol, subinfo, accion_label=None, accion_fn=None, confirmar_entrega=False):
    """
    Dibuja una tarjeta de solicitud con: info + botón de acción principal
    (opcional, ej. "Ya avisé al chofer") + mensaje para el chofer + Editar/Borrar.
    Se usa igual en Pendientes, Avisadas y Completadas.

    confirmar_entrega=True (solo en Avisadas) reemplaza el botón simple por un
    mini-panel "📥 Confirmar Entrega" donde se valida fecha, hora y chofer real
    del retiro, y al confirmar genera el mensaje de confirmación para WhatsApp.
    """
    id_sol = r["ID"]
    with st.container(border=True):
        detalle_linea = f" — {r['Detalle']}" if r.get("Detalle") else ""
        cA, cB = st.columns([4, 1])
        cA.markdown(f"**{id_sol}** — {r['Tipo de Retiro']}{detalle_linea} · {r['Ruta / Destino']}  \n{subinfo}")
        if accion_label and accion_fn and not confirmar_entrega and cB.button(accion_label, key=f"accion_{id_sol}", use_container_width=True):
            accion_fn()
            st.session_state["recargar_solicitudes"] += 1
            st.rerun()

        if confirmar_entrega:
            msg_key = f"msg_confirmado_{id_sol}"
            if st.session_state.get(msg_key):
                st.success("✅ Entrega confirmada.")
                st.code(st.session_state[msg_key], language=None)
                if st.button("👍 Listo, continuar", key=f"cerrar_{id_sol}", use_container_width=True):
                    del st.session_state[msg_key]
                    st.session_state["recargar_solicitudes"] += 1
                    st.rerun()
            else:
                with st.expander("📥 Confirmar Entrega", expanded=False):
                    with st.form(f"form_confirmar_{id_sol}"):
                        cf1, cf2 = st.columns(2)
                        fecha_retiro = cf1.date_input("Fecha del retiro:", value=date.today(), key=f"fret_{id_sol}")
                        hora_retiro = cf2.time_input("Hora del retiro:", value=ahora_vzla().time(), key=f"hret_{id_sol}")
                        chofer_retiro = st.text_input(
                            "Chofer que retiró:", value=r["Chofer Asignado"], key=f"chret_{id_sol}",
                            help="Por defecto es el chofer asignado a la solicitud, pero puedes cambiarlo si fue otro chofer.",
                        )
                        confirmar = st.form_submit_button("✅ Confirmar Entrega", type="primary", use_container_width=True)
                        if confirmar:
                            duracion_texto, confirmado_dt = marcar_completada_detallada(
                                ws_sol, id_sol, r["Fecha"], fecha_retiro, hora_retiro, chofer_retiro
                            )
                            chofer_final = chofer_retiro.strip().upper() if chofer_retiro.strip() else r["Chofer Asignado"]
                            st.session_state[msg_key] = generar_mensaje_confirmacion(r, confirmado_dt, chofer_final, duracion_texto)
                            st.rerun()

        with st.expander("📋 Mensaje para el chofer"):
            st.code(generar_mensaje_chofer(r), language=None)

        with st.expander("✏️ Editar / 🗑️ Borrar"):
            with st.form(f"form_editar_{id_sol}"):
                ec1, ec2 = st.columns(2)
                sol_idx = SOLICITANTES_FIJOS.index(r["Solicitante"]) if r["Solicitante"] in SOLICITANTES_FIJOS else len(SOLICITANTES_FIJOS) - 1
                e_solicitante_sel = ec1.selectbox("Solicitante:", SOLICITANTES_FIJOS, index=sol_idx, key=f"esol_{id_sol}")
                e_solicitante_otro = ""
                if e_solicitante_sel == "OTROS":
                    valor_previo = r["Solicitante"] if r["Solicitante"] not in SOLICITANTES_FIJOS else ""
                    e_solicitante_otro = ec2.text_input("Especifique quién:", value=valor_previo, key=f"eotro_{id_sol}")
                e_detalle = st.text_input("Detalle:", value=r.get("Detalle", ""), key=f"edet_{id_sol}")
                ec3, ec4 = st.columns(2)
                e_ruta = ec3.text_input("Ruta / Destino:", value=r["Ruta / Destino"], key=f"eruta_{id_sol}")
                e_chofer = ec4.text_input("Chofer Asignado:", value=r["Chofer Asignado"], key=f"echofer_{id_sol}")
                e_fecha = st.date_input("Fecha:", value=_parsear_fecha(r["Fecha"]), key=f"efecha_{id_sol}")

                g1, g2 = st.columns(2)
                guardar = g1.form_submit_button("💾 Guardar Cambios", use_container_width=True)
                confirmar_borrado = g2.checkbox("Confirmar borrado", key=f"chkborrar_{id_sol}")
                borrar = g2.form_submit_button("🗑️ Borrar Solicitud", use_container_width=True)

                if guardar:
                    sol_final = e_solicitante_otro.strip() if e_solicitante_sel == "OTROS" else e_solicitante_sel
                    if not sol_final or not e_ruta or not e_chofer:
                        st.error("Completa Solicitante, Ruta/Destino y Chofer Asignado.")
                    else:
                        actualizar_solicitud(ws_sol, id_sol, sol_final, e_detalle, e_ruta, e_chofer, e_fecha)
                        st.success("✅ Solicitud actualizada.")
                        st.session_state["recargar_solicitudes"] += 1
                        st.rerun()

                if borrar:
                    if not confirmar_borrado:
                        st.error("Marca \"Confirmar borrado\" antes de borrar — es permanente.")
                    else:
                        borrar_solicitud(ws_sol, id_sol)
                        st.success(f"🗑️ Solicitud {id_sol} borrada.")
                        st.session_state["recargar_solicitudes"] += 1
                        st.rerun()

# ==========================================
# INTERFAZ STREAMLIT
# ==========================================
st.title("📋 Control de Solicitudes")

try:
    ws_sol, ws_kpi = asegurar_estructura_sheet()
except Exception as e:
    st.error(f"No pude conectar con el Google Sheet de Control de Solicitudes: {e}")
    st.info("Verifica que el Sheet esté compartido (como Editor) con el correo de la cuenta de servicio "
            "que usa esta app (el `client_email` que está en tus Secrets, dentro de `gcp_service_account`).")
    st.stop()

if "recargar_solicitudes" not in st.session_state:
    st.session_state["recargar_solicitudes"] = 0

df = leer_solicitudes(ws_sol)

# ---------------------------------------------------------
# ALERTAS
# ---------------------------------------------------------
if not df.empty:
    ahora = ahora_vzla()

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
# 0. RESUMEN / KPIs — visible siempre en la pantalla principal,
#    no solo dentro de la pestaña "KPIs".
# ---------------------------------------------------------
st.subheader("📊 Resumen General")
total_kpi = len(df)
n_pend_kpi = int((df["Estado"] == ESTADO_PENDIENTE).sum()) if total_kpi else 0
n_avis_kpi = int((df["Estado"] == ESTADO_AVISADO).sum()) if total_kpi else 0
n_comp_kpi = int((df["Estado"] == ESTADO_COMPLETADA).sum()) if total_kpi else 0
dias_prom_kpi = pd.to_numeric(df["Días para Completarse"], errors="coerce").dropna().mean() if total_kpi else 0

rk1, rk2, rk3, rk4, rk5 = st.columns(5)
rk1.metric("Total", total_kpi)
rk2.metric("🔴 Pendientes", n_pend_kpi)
rk3.metric("🟡 Avisadas", n_avis_kpi)
rk4.metric("🟢 Completadas", n_comp_kpi)
rk5.metric("Prom. días para completarse", f"{dias_prom_kpi:,.1f}" if dias_prom_kpi else "—")

st.markdown("---")

# ---------------------------------------------------------
# 1. CREAR NUEVA SOLICITUD
# ---------------------------------------------------------
st.subheader("➕ Nueva solicitud")

# El selector de "Solicitante" va FUERA del formulario porque necesita
# reaccionar al instante cuando se elige "OTROS" (los widgets dentro de un
# st.form solo se procesan al enviarlo, no muestran/ocultan nada al vuelo).
sc1, sc2 = st.columns(2)
solicitante_sel = sc1.selectbox("Solicitante (Supervisor):", SOLICITANTES_FIJOS, key="solicitante_sel_nueva")
solicitante_otro = ""
if solicitante_sel == "OTROS":
    solicitante_otro = sc2.text_input("Especifique quién solicita:", key="solicitante_otro_nueva")

with st.form("form_nueva_solicitud", clear_on_submit=True):
    detalle = st.text_input("¿Qué se solicita? (detalle):", placeholder="Ej: 2 BOTELLAS DE AGUA")
    fc3, fc4 = st.columns(2)
    ruta = fc3.text_input("Ruta / Destino:")
    chofer = fc4.text_input("Chofer Asignado:")
    fecha_solicitud = st.date_input(
        "Fecha de la solicitud:", value=date.today(),
        help="Por defecto es hoy, pero la puedes cambiar — por ejemplo si estás cargando algo que pidieron el lunes."
    )
    enviado = st.form_submit_button("➕ Crear Solicitud", type="primary", use_container_width=True)

    if enviado:
        solicitante_final = solicitante_otro.strip() if solicitante_sel == "OTROS" else solicitante_sel
        if not solicitante_final or not ruta or not chofer:
            st.error("Completa Solicitante (si elegiste 'OTROS', escribe quién), Ruta/Destino y Chofer Asignado.")
        else:
            nuevo_id = crear_solicitud(ws_sol, solicitante_final, detalle, ruta, chofer, fecha_solicitud)
            st.success(f"✅ Solicitud {nuevo_id} creada como Pendiente ({fecha_solicitud.strftime('%d/%m/%Y')}).")
            # El selector de "Solicitante" y el campo de "OTROS" viven FUERA del
            # st.form (ver comentario arriba), así que clear_on_submit no los
            # limpia solo — hay que borrarlos a mano para que el formulario
            # quede completamente en blanco y listo para cargar otra solicitud.
            st.session_state.pop("solicitante_sel_nueva", None)
            st.session_state.pop("solicitante_otro_nueva", None)
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
        render_tarjeta(
            r, ws_sol,
            subinfo=f"Solicitó: {r['Solicitante']} ({r['Fecha']}) · Chofer: {r['Chofer Asignado']}",
            accion_label="📣 Ya avisé al chofer",
            accion_fn=lambda r=r: marcar_avisado(ws_sol, r["ID"]),
        )

with tab_avis:
    df_avis = df[df["Estado"] == ESTADO_AVISADO] if not df.empty else df
    if df_avis.empty:
        st.info("No hay solicitudes esperando confirmación de retiro.")
    for _, r in df_avis.iterrows():
        render_tarjeta(
            r, ws_sol,
            subinfo=f"Chofer: {r['Chofer Asignado']} · Avisado: {r['Avisado (Fecha y Hora)']}",
            confirmar_entrega=True,
        )

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
        if st.checkbox("✏️ Editar o borrar una solicitud completada"):
            for _, r in df_comp.iterrows():
                render_tarjeta(
                    r, ws_sol,
                    subinfo=f"Chofer: {r['Chofer Asignado']} · Confirmado: {r['Confirmado (Fecha y Hora)']}",
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
        st.markdown("**Resumen por chofer (conteo):**")
        resumen_chofer = df.groupby("Chofer Asignado")["Estado"].value_counts().unstack(fill_value=0)
        st.dataframe(resumen_chofer, use_container_width=True)

        st.markdown("**Detalle por chofer (fecha, ruta y estado):**")
        columnas_detalle = ["Chofer Asignado", "Fecha", "Ruta / Destino", "Solicitante", "Estado",
                             "Avisado (Fecha y Hora)", "Confirmado (Fecha y Hora)"]
        columnas_detalle = [c for c in columnas_detalle if c in df.columns]
        st.dataframe(
            df[columnas_detalle].sort_values(["Chofer Asignado", "Fecha"]),
            use_container_width=True, hide_index=True,
        )

with tab_informe:
    st.markdown("**Informe descargable (Excel y PDF)**")
    rango = st.radio("Rango:", ["Día", "Semana", "Mes", "Todo"], horizontal=True)

    df_informe = df.copy()
    titulo_rango = "Todas las solicitudes"

    if not df.empty:
        df_informe["_fecha_dt"] = pd.to_datetime(df_informe["Fecha"], format="%d/%m/%Y", errors="coerce")

        if rango == "Día":
            fecha_sel = st.date_input("Elige el día:", value=date.today(), key="informe_dia")
            df_informe = df_informe[df_informe["_fecha_dt"].dt.date == fecha_sel]
            titulo_rango = f"Día: {fecha_sel.strftime('%d/%m/%Y')}"

        elif rango == "Semana":
            fecha_ref = st.date_input(
                "Elige cualquier día DENTRO de la semana que quieres ver:", value=date.today(), key="informe_semana"
            )
            iso_ref = fecha_ref.isocalendar()  # (año ISO, semana ISO, día)
            iso_datos = df_informe["_fecha_dt"].dt.isocalendar()
            df_informe = df_informe[(iso_datos["year"] == iso_ref[0]) & (iso_datos["week"] == iso_ref[1])]
            inicio_semana = fecha_ref - timedelta(days=fecha_ref.weekday())
            fin_semana = inicio_semana + timedelta(days=6)
            titulo_rango = f"Semana del {inicio_semana.strftime('%d/%m/%Y')} al {fin_semana.strftime('%d/%m/%Y')}"

        elif rango == "Mes":
            fecha_ref = st.date_input(
                "Elige cualquier día DENTRO del mes que quieres ver:", value=date.today(), key="informe_mes"
            )
            df_informe = df_informe[
                (df_informe["_fecha_dt"].dt.month == fecha_ref.month)
                & (df_informe["_fecha_dt"].dt.year == fecha_ref.year)
            ]
            titulo_rango = f"Mes: {MESES_ES.get(fecha_ref.strftime('%B'), fecha_ref.strftime('%B'))} {fecha_ref.year}"

        df_informe = df_informe.drop(columns=["_fecha_dt"])

    st.dataframe(df_informe, use_container_width=True, hide_index=True)

    if not df_informe.empty:
        total_i = len(df_informe)
        n_pend_i = int((df_informe["Estado"] == ESTADO_PENDIENTE).sum())
        n_avis_i = int((df_informe["Estado"] == ESTADO_AVISADO).sum())
        n_comp_i = int((df_informe["Estado"] == ESTADO_COMPLETADA).sum())
        st.caption(f"Total: {total_i} · 🔴 Pendientes: {n_pend_i} · 🟡 Avisadas: {n_avis_i} · 🟢 Completadas: {n_comp_i}")

        col_excel, col_pdf = st.columns(2)

        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
            df_informe.to_excel(writer, index=False, sheet_name="Solicitudes")
        col_excel.download_button(
            "⬇️ Descargar Excel",
            data=buffer_excel.getvalue(),
            file_name=f"informe_solicitudes_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if FPDF_DISPONIBLE:
            pdf_bytes = generar_pdf_informe(df_informe, titulo_rango)
            col_pdf.download_button(
                "⬇️ Descargar PDF",
                data=pdf_bytes,
                file_name=f"informe_solicitudes_{date.today().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            col_pdf.caption("Instala `pip install fpdf2` en el repo (requirements.txt) para activar la descarga en PDF.")
    else:
        st.caption("No hay solicitudes en ese rango para exportar.")
