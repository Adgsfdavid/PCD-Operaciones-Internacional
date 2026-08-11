# ==========================================
# Archivo: novedades.py
# Modulo: Novedades y Status Diario (Sheet nuevo, separado)
# Parte 1 (visor de novedades) + Parte 2 (pizarra de status + 2 imagenes)
# + Cierre 6:20 (historial)
# ==========================================
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import html as _html
import json as _json
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

COL_FECHA = {"Novedades_Despacho": "FECHA RECLAMO", "Encomiendas": "FECHA", "Novedades_Ruta": "FECHA"}

ZONAS_TODAS = ["Centro", "Centro-Occidente", "Occidente", "Oriente", "Transbordo", "Encomiendas"]
ZONAS_DESPACHO = ["Centro", "Centro-Occidente", "Occidente", "Oriente"]
ZONAS_EXTRA = ["Transbordo", "Encomiendas"]
STATUS_OPC = ["Despacho", "Retorno", "Resguardo"]
STATUS_COLOR = {"DESPACHO": "#1565c0", "RETORNO": "#e65100", "RESGUARDO": "#2e7d32"}

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
    return gspread.authorize(Credentials.from_service_account_info(cred, scopes=alcance))

def abrir_libro():
    return obtener_cliente_sheets().open_by_key(SHEET_ID)

def asegurar_estructura(libro):
    existentes = {w.title: w for w in libro.worksheets()}
    creadas, reparadas, ya = [], [], []
    for nombre, cols in HOJAS.items():
        if nombre not in existentes:
            ws = libro.add_worksheet(title=nombre, rows=500, cols=max(12, len(cols) + 1))
            ws.update([cols]); creadas.append(nombre)
        else:
            ws = existentes[nombre]
            if not ws.row_values(1):
                ws.update([cols]); reparadas.append(nombre)
            else:
                ya.append(nombre)
    for basura in ["Hoja 1", "Hoja1", "Sheet1"]:
        if basura in [w.title for w in libro.worksheets()] and basura not in HOJAS:
            try: libro.del_worksheet(libro.worksheet(basura))
            except Exception: pass
    return creadas, reparadas, ya

@st.cache_data(ttl=60, show_spinner=False)
def leer_hoja_df(nombre_hoja):
    try:
        return pd.DataFrame(abrir_libro().worksheet(nombre_hoja).get_all_records())
    except Exception:
        return pd.DataFrame()

def guardar_status(df):
    ws = abrir_libro().worksheet("Status_Dia")
    ws.clear()
    df = df.fillna("").astype(str)
    ws.update([HOJAS["Status_Dia"]] + df.values.tolist())

def guardar_cierre(df):
    ws = abrir_libro().worksheet("Historial_Status")
    fecha = datetime.now().strftime("%d/%m/%Y")
    corte = datetime.now().strftime("%I:%M %p")
    filas = []
    for _, r in df.fillna("").iterrows():
        filas.append([fecha, corte, r.get("ZONA", ""), r.get("UNIDAD", ""), r.get("CHOFER", ""),
                      r.get("AYUDANTE", ""), r.get("RUTA/DESPACHO", ""), r.get("HORA", ""),
                      r.get("UBICACIÓN ACTUAL", ""), r.get("STATUS", "")])
    if filas:
        ws.append_rows(filas, value_input_option="USER_ENTERED")

# ==========================================
# UTILIDADES
# ==========================================
def norm(s):
    s = str(s).upper().strip()
    for a, b in [("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U")]:
        s = s.replace(a, b)
    return s

def up(v):
    return str(v).upper() if isinstance(v, str) else v

def a_mayusculas(df):
    if df.empty: return df
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].apply(up)
    return df

def agregar_dia_mes(df, col_fecha):
    if df.empty or col_fecha not in df.columns: return df
    def _d(v):
        try: return DIAS_SEMANA[pd.to_datetime(str(v), dayfirst=True).weekday()]
        except Exception: return ""
    def _m(v):
        try: return MESES_ANO[pd.to_datetime(str(v), dayfirst=True).month - 1]
        except Exception: return ""
    df = df.copy()
    df.insert(1, "DÍA", df[col_fecha].apply(_d))
    df.insert(2, "MES", df[col_fecha].apply(_m))
    return df

def filtrar_por_dia(df, col_fecha, dia_sel):
    if df.empty or col_fecha not in df.columns: return df
    obj = pd.to_datetime(df[col_fecha].astype(str), dayfirst=True, errors="coerce")
    return df[obj.dt.date == dia_sel]

# ==========================================
# PIZARRA (IMAGEN) Y CAJA DE COPIAR
# ==========================================
def badge_status(status):
    color = STATUS_COLOR.get(norm(status), "#546e7a")
    txt = _html.escape(str(status).upper())
    return f'<span style="background:{color};color:#fff;padding:3px 12px;border-radius:12px;font-weight:bold;font-size:12px;white-space:nowrap;">{txt}</span>'

def html_pizarra_status(df, zonas, titulo, uid):
    fecha_txt = datetime.now().strftime("%A, %d/%m/%Y")
    total = 0
    cuerpo = ""
    for zona in zonas:
        sub = df[df["ZONA"].apply(norm) == norm(zona)] if not df.empty and "ZONA" in df.columns else pd.DataFrame()
        if sub.empty:
            continue
        total += len(sub)
        cuerpo += f"""<tr><td colspan="6" style="background:#0d2b57;color:#fff;font-weight:bold;padding:8px 12px;font-size:14px;">📍 ZONA {_html.escape(zona.upper())} &nbsp;—&nbsp; {len(sub)} unidad(es)</td></tr>"""
        for i, (_, r) in enumerate(sub.iterrows()):
            fondo = "#f4f6f9" if i % 2 == 0 else "#ffffff"
            personal = _html.escape(up(r.get("CHOFER", "")))
            ayud = up(r.get("AYUDANTE", ""))
            if ayud: personal += f" / {_html.escape(ayud)}"
            cuerpo += f"""<tr style="background:{fondo};">
                <td style="padding:8px 10px;font-weight:bold;color:#0d2b57;border-bottom:1px solid #e0e4ea;">{_html.escape(up(r.get('RUTA/DESPACHO','')))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #e0e4ea;">{_html.escape(up(r.get('UNIDAD','')))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #e0e4ea;font-size:13px;">{personal}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #e0e4ea;white-space:nowrap;">{_html.escape(up(r.get('HORA','')))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #e0e4ea;font-weight:bold;">{_html.escape(up(r.get('UBICACIÓN ACTUAL','')))}</td>
                <td style="padding:8px 10px;border-bottom:1px solid #e0e4ea;text-align:center;">{badge_status(r.get('STATUS',''))}</td>
            </tr>"""
    if total == 0:
        cuerpo = """<tr><td colspan="6" style="padding:18px;text-align:center;color:#888;">Sin unidades para estas zonas.</td></tr>"""

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right;margin-bottom:8px;">
      <button onclick="descargar_{uid}()" style="background:#0d47a1;color:#fff;border:none;padding:10px 18px;border-radius:6px;cursor:pointer;font-weight:bold;">⬇️ Descargar imagen</button>
    </div>
    <div id="piz-{uid}" style="background:#fff;font-family:Arial,sans-serif;border:1px solid #ccc;border-radius:8px;overflow:hidden;">
      <div style="background:#0d47a1;color:#fff;padding:14px 16px;">
        <div style="font-size:20px;font-weight:bold;">🚦 {_html.escape(titulo)}</div>
        <div style="font-size:13px;opacity:.9;">Drotaca · PERÍODO: {_html.escape(fecha_txt)} · Total: {total} unidad(es)</div>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;color:#1a1a1a;">
        <tr style="background:#1c3d6e;color:#fff;">
          <th style="padding:9px 10px;text-align:left;">RUTA / DESPACHO</th>
          <th style="padding:9px 10px;text-align:left;">UNIDAD</th>
          <th style="padding:9px 10px;text-align:left;">CHÓFER / AYUDANTE</th>
          <th style="padding:9px 10px;text-align:left;">HORA</th>
          <th style="padding:9px 10px;text-align:left;">UBICACIÓN</th>
          <th style="padding:9px 10px;text-align:center;">STATUS</th>
        </tr>
        {cuerpo}
      </table>
    </div>
    <script>
    function descargar_{uid}() {{
        html2canvas(document.getElementById('piz-{uid}'), {{scale:2}}).then(function(canvas){{
            var link=document.createElement('a');
            link.download='Pizarra_{uid}.png';
            link.href=canvas.toDataURL('image/png');
            link.click();
        }});
    }}
    </script>
    """

def html_caja_copiar(texto, uid):
    texto = "" if texto is None else str(texto)
    return f"""
    <div style="font-family:Arial,sans-serif;">
      <textarea id="cp-{uid}" readonly style="width:100%;height:240px;padding:10px;border:1px solid #444;border-radius:8px;background:#0e1117;color:#fafafa;font-size:13px;resize:vertical;box-sizing:border-box;white-space:pre-wrap;">{_html.escape(texto)}</textarea>
      <button onclick="copiar_{uid}()" style="margin-top:8px;width:100%;background:#25D366;color:#fff;border:none;padding:11px 18px;border-radius:8px;cursor:pointer;font-weight:bold;font-size:14px;">📋 Copiar mensaje</button>
      <span id="ok-{uid}" style="display:none;color:#25D366;font-weight:bold;margin-left:8px;">✓ Copiado</span>
    </div>
    <script>
    function copiar_{uid}() {{
        var t=document.getElementById('cp-{uid}'); var txt={_json.dumps(texto)};
        function ok(){{var s=document.getElementById('ok-{uid}');s.style.display='inline';setTimeout(function(){{s.style.display='none';}},1800);}}
        if(navigator.clipboard&&navigator.clipboard.writeText){{navigator.clipboard.writeText(txt).then(ok,function(){{t.removeAttribute('readonly');t.focus();t.select();document.execCommand('copy');t.setAttribute('readonly','readonly');ok();}});}}
        else{{t.removeAttribute('readonly');t.focus();t.select();document.execCommand('copy');t.setAttribute('readonly','readonly');ok();}}
    }}
    </script>
    """

def texto_novedades(dia_str, vd, ve, vr):
    L = [f"*NOVEDADES DEL DÍA {dia_str}*", ""]
    L.append(f"🚨 *DESPACHO ({len(vd)}):*")
    if vd.empty: L.append("Sin novedades.")
    else:
        for _, r in vd.iterrows():
            L.append(f"• {up(r.get('RUTA',''))} · {up(r.get('NOVEDAD',''))} · {up(r.get('CLIENTE/FARMACIA',''))} — {up(r.get('CONTEXTO',''))}")
    L.append("")
    L.append(f"📦 *ENCOMIENDAS ({len(ve)}):*")
    if ve.empty: L.append("Sin encomiendas.")
    else:
        for _, r in ve.iterrows():
            L.append(f"• {up(r.get('MOVIMIENTO',''))} · {up(r.get('RUTA',''))} · {up(r.get('TIPO DE ENCOMIENDA',''))} — {up(r.get('DETALLE',''))}")
    L.append("")
    L.append(f"🛞 *NOVEDADES POR RUTA ({len(vr)}):*")
    if vr.empty: L.append("Sin novedades.")
    else:
        for _, r in vr.iterrows():
            L.append(f"• {up(r.get('RUTA',''))} · {up(r.get('TIPO DE NOVEDAD',''))} — {up(r.get('DESCRIPCIÓN',''))}")
    return "\n".join(L)

# ==========================================
# INTERFAZ
# ==========================================
st.title("📝 Novedades y Status Diario")
usuario = str(st.session_state.get("usuario", "-")).upper()

try:
    libro = abrir_libro()
except Exception as e:
    st.error("No se pudo conectar con el Google Sheet de novedades.")
    st.info("Comparte el Sheet como **Editor** con `bot-pizarra@pcd-drotaca.iam.gserviceaccount.com`.")
    st.exception(e); st.stop()

if not st.session_state.get("nov_estructura_ok", False):
    try:
        asegurar_estructura(libro); st.session_state["nov_estructura_ok"] = True
    except Exception: pass

with st.expander("⚙️ Configuración de hojas (crear / reparar)"):
    if st.button("🛠️ Crear / reparar hojas ahora"):
        try:
            c, r, y = asegurar_estructura(libro)
            if c: st.success("✅ Creadas: " + ", ".join(c))
            if r: st.info("🔧 Encabezados repuestos: " + ", ".join(r))
            if not c and not r: st.info("Todo en orden.")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"Error: {e}")
    st.write("**Hojas actuales:**", ", ".join([w.title for w in libro.worksheets()]))

pizarra_tab, nov_tab, cierre_tab = st.tabs(["🚦 Status del Día", "🚨 Novedades", "📤 Cierre 6:20 pm"])

# ------------------------------------------
# STATUS DEL DÍA (pizarra editable + 2 imágenes)
# ------------------------------------------
with pizarra_tab:
    st.subheader("🚦 Pizarra de Status del Día")
    st.caption("Puedes llenar/editar aquí o directo en la hoja Status_Dia. Zonas de despacho: Centro, "
               "Centro-Occidente, Occidente, Oriente. Extra: Transbordo y Encomiendas.")

    if st.button("🔄 Traer del Sheet", key="refresh_status"):
        st.cache_data.clear(); st.rerun()

    df_status = leer_hoja_df("Status_Dia")
    if df_status.empty:
        df_status = pd.DataFrame(columns=HOJAS["Status_Dia"])
    for c in HOJAS["Status_Dia"]:
        if c not in df_status.columns: df_status[c] = ""
    df_status = df_status[HOJAS["Status_Dia"]]

    edit = st.data_editor(
        df_status, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "ZONA": st.column_config.SelectboxColumn("ZONA", options=ZONAS_TODAS, required=False),
            "STATUS": st.column_config.SelectboxColumn("STATUS", options=STATUS_OPC, required=False),
        }, key="editor_status"
    )

    if st.button("💾 Guardar status en el Sheet", type="primary", use_container_width=True):
        try:
            guardar_status(edit)
            st.cache_data.clear()
            st.success("✅ Status guardado en el Sheet.")
        except Exception as e:
            st.error(f"No se pudo guardar: {e}")

    st.markdown("---")
    st.markdown("### 🖼️ Imagen 1 — Despacho por zonas")
    import streamlit.components.v1 as components
    alto1 = 260 + max(1, len(edit[edit["ZONA"].apply(norm).isin([norm(z) for z in ZONAS_DESPACHO])])) * 42
    components.html(html_pizarra_status(edit, ZONAS_DESPACHO, "PIZARRA DE STATUS — DESPACHO", "desp"), height=int(alto1), scrolling=True)

    st.markdown("### 🖼️ Imagen 2 — Transbordo + Encomiendas")
    alto2 = 260 + max(1, len(edit[edit["ZONA"].apply(norm).isin([norm(z) for z in ZONAS_EXTRA])])) * 42
    components.html(html_pizarra_status(edit, ZONAS_EXTRA, "PIZARRA DE STATUS — TRANSBORDO Y ENCOMIENDAS", "extra"), height=int(alto2), scrolling=True)

# ------------------------------------------
# NOVEDADES (visor + texto WhatsApp)
# ------------------------------------------
with nov_tab:
    c1, c2, c3 = st.columns([2, 2, 1])
    dia_sel = c1.date_input("📅 Ver día", value=date.today(), format="DD/MM/YYYY", key="dia_nov")
    ver_todo = c2.checkbox("Ver todo el histórico")
    if c3.button("🔄 Actualizar", key="refresh_nov", use_container_width=True):
        st.cache_data.clear(); st.rerun()

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
    etq = "histórico" if ver_todo else dia_sel.strftime("%d/%m/%Y")
    m1.metric(f"🚨 Despacho ({etq})", len(vd))
    m2.metric(f"📦 Encomiendas ({etq})", len(ve))
    m3.metric(f"🛞 Por ruta ({etq})", len(vr))
    st.markdown("---")

    t1, t2, t3, t4 = st.tabs(["🚨 Despacho", "📦 Encomiendas", "🛞 Por Ruta", "📱 Texto WhatsApp"])
    with t1:
        st.dataframe(a_mayusculas(vd).iloc[::-1], use_container_width=True, hide_index=True) if not vd.empty else st.info("Sin registros.")
    with t2:
        st.dataframe(a_mayusculas(ve).iloc[::-1], use_container_width=True, hide_index=True) if not ve.empty else st.info("Sin registros.")
    with t3:
        st.dataframe(a_mayusculas(vr).iloc[::-1], use_container_width=True, hide_index=True) if not vr.empty else st.info("Sin registros.")
    with t4:
        import streamlit.components.v1 as components
        dia_txt = "TODO EL HISTÓRICO" if ver_todo else dia_sel.strftime("%d/%m/%Y")
        components.html(html_caja_copiar(texto_novedades(dia_txt, vd, ve, vr), "nov"), height=340)

# ------------------------------------------
# CIERRE 6:20 pm (guardar histórico)
# ------------------------------------------
with cierre_tab:
    st.subheader("📤 Cierre del día (6:20 pm)")
    st.caption("Guarda una foto del status actual en la hoja Historial_Status con la fecha y la hora de corte.")
    df_actual = leer_hoja_df("Status_Dia")
    st.write(f"Unidades en el status actual: **{len(df_actual)}**")
    if not df_actual.empty:
        st.dataframe(a_mayusculas(df_actual), use_container_width=True, hide_index=True)

    if st.button("📅 Guardar cierre en el historial", type="primary", use_container_width=True):
        if df_actual.empty:
            st.warning("No hay filas en Status_Dia para guardar.")
        else:
            try:
                guardar_cierre(df_actual)
                st.cache_data.clear()
                st.success("✅ Cierre guardado en Historial_Status.")
            except Exception as e:
                st.error(f"No se pudo guardar el cierre: {e}")

    with st.expander("📚 Ver historial guardado"):
        dfh = leer_hoja_df("Historial_Status")
        if dfh.empty:
            st.info("Aún no hay cierres guardados.")
        else:
            st.dataframe(a_mayusculas(dfh).iloc[::-1], use_container_width=True, hide_index=True)
