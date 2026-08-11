# ==========================================
# Archivo: velocidad.py
# Modulo: Exceso de Velocidad (carga Excel -> pizarra + WhatsApp + historial)
# ==========================================
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import html as _html
import json as _json
import re
import pandas as pd
from datetime import datetime, date
import streamlit.components.v1 as components

# ==========================================
# CONFIGURACION
# ==========================================
SHEET_ID = "1D7w0ABnnatGd83TpJHFeVxxLOKRqoBBFbM9FyYYdFEg"  # mismo Sheet de novedades
HOJA_HIST = "Historial_Velocidad"
HIST_COLS = ["FECHA", "PLACA", "EXCESOS", "VEL MÁXIMA", "CHOFER", "RUTA", "GUARDADO"]

# ==========================================
# CONEXION A GOOGLE SHEETS
# ==========================================
@st.cache_resource
def obtener_cliente_sheets():
    cred = dict(st.secrets["gcp_service_account"])
    lk = cred["private_key"]
    lk = (lk.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
            .replace("\\n", "").replace("\n", "").replace(" ", ""))
    cred["private_key"] = "-----BEGIN PRIVATE KEY-----\n" + "\n".join(textwrap.wrap(lk, 64)) + "\n-----END PRIVATE KEY-----\n"
    alcance = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    return gspread.authorize(Credentials.from_service_account_info(cred, scopes=alcance))

def abrir_libro():
    return obtener_cliente_sheets().open_by_key(SHEET_ID)

@st.cache_data(ttl=60, show_spinner=False)
def leer_status_map():
    """Lee Status_Dia y arma placa -> (chofer, ruta)."""
    try:
        regs = abrir_libro().worksheet("Status_Dia").get_all_records()
    except Exception:
        return {}
    mapa = {}
    for r in regs:
        unidad = r.get("UNIDAD", "")
        k = placa_key(unidad)
        if k:
            mapa[k] = (str(r.get("CHOFER", "")), str(r.get("RUTA/DESPACHO", "")))
    return mapa

def guardar_historial(fecha_txt, resumen, usuario):
    libro = abrir_libro()
    try:
        ws = libro.worksheet(HOJA_HIST)
    except Exception:
        ws = libro.add_worksheet(title=HOJA_HIST, rows=1000, cols=len(HIST_COLS) + 1)
        ws.update([HIST_COLS])
    if not ws.row_values(1):
        ws.update([HIST_COLS])
    marca = datetime.now().strftime("%d/%m/%Y %I:%M %p") + " · " + str(usuario)
    filas = []
    for _, r in resumen.iterrows():
        filas.append([fecha_txt, str(r["Unidad"]), int(r["Excesos"]), int(r["VelMax"]),
                      str(r["Chofer"]), str(r["Ruta"]), marca])
    if filas:
        ws.append_rows(filas, value_input_option="USER_ENTERED")

# ==========================================
# UTILIDADES
# ==========================================
def placa_key(v):
    """Extrae la placa base para cruzar con Status_Dia (primer token, solo alfanumérico)."""
    t = str(v).upper().strip()
    t = re.split(r'[\s\-]+', t)[0]
    return re.sub(r'[^A-Z0-9]', '', t)

def su(v):
    if v is None:
        return ""
    try:
        if isinstance(v, float) and pd.isna(v):
            return ""
    except Exception:
        pass
    t = str(v)
    return "" if t.strip().lower() == "nan" else t.upper()

def corta_dir(v):
    """Acorta la localización: quita el sufijo '~Estado' y deja las 2 primeras partes."""
    s = str(v).split("~")[0]
    partes = [p.strip() for p in s.split(",") if p.strip() and p.strip().lower() != "nan"]
    return ", ".join(partes[:2])

def cargar_excel(file):
    """Lee el Excel detectando la fila de encabezados (Unidad + Velocidad), aunque empiece en A9."""
    raw = pd.read_excel(file, header=None)
    hrow = None
    for r in range(min(40, len(raw))):
        fila = [str(x).upper() for x in raw.iloc[r].tolist()]
        if any("UNIDAD" in x for x in fila) and any("VELOCIDAD" in x for x in fila):
            hrow = r
            break
    if hrow is None:
        hrow = 8
    df = raw.iloc[hrow:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
    return df.iloc[1:].reset_index(drop=True)

def procesar(df, umbral, status_map, topn):
    c = {}
    for col in df.columns:
        u = str(col).upper()
        if "UNIDAD" in u and "Unidad" not in c: c["Unidad"] = col
        elif "FECHA" in u and "Fecha" not in c: c["Fecha"] = col
        elif "ESTADO" in u and "Estado" not in c: c["Estado"] = col
        elif "VELOCIDAD" in u and "Vel" not in c: c["Vel"] = col
        elif "LOCALIZ" in u and "Loc" not in c: c["Loc"] = col
        elif "CONDUCTOR" in u and "Cond" not in c: c["Cond"] = col

    d = pd.DataFrame()
    d["Unidad"] = df[c["Unidad"]].astype(str).str.strip()
    d["Fecha"] = pd.to_datetime(df[c["Fecha"]], errors="coerce", dayfirst=True) if "Fecha" in c else pd.NaT
    d["Estado"] = df[c["Estado"]].astype(str) if "Estado" in c else ""
    d["Vel"] = pd.to_numeric(df[c["Vel"]], errors="coerce") if "Vel" in c else None
    d["Loc"] = df[c["Loc"]].astype(str) if "Loc" in c else ""
    d["Cond"] = df[c["Cond"]].astype(str) if "Cond" in c else ""

    d = d.dropna(subset=["Vel"])
    d = d[d["Unidad"].str.upper() != "NAN"]
    d = d[d["Vel"] > umbral]

    resumen = d.groupby("Unidad").agg(Excesos=("Vel", "size"), VelMax=("Vel", "max")).reset_index()
    resumen = resumen.sort_values("Excesos", ascending=False).reset_index(drop=True)

    def chof(u):
        ch, _ = status_map.get(placa_key(u), ("", ""))
        return ch
    def rut(u):
        _, rt = status_map.get(placa_key(u), ("", ""))
        return rt
    resumen["Chofer"] = resumen["Unidad"].apply(chof)
    resumen["Ruta"] = resumen["Unidad"].apply(rut)
    # Si no hay chofer en el status, usar el 'Conductor' del propio Excel (si viene)
    cond_por_placa = d.groupby("Unidad")["Cond"].first().to_dict()
    resumen["Chofer"] = resumen.apply(
        lambda r: r["Chofer"] if str(r["Chofer"]).strip() else str(cond_por_placa.get(r["Unidad"], "")).replace("nan", "").strip(),
        axis=1)

    top = d.sort_values("Vel", ascending=False).head(topn).reset_index(drop=True)

    fecha_txt = ""
    fser = d["Fecha"].dropna()
    if not fser.empty:
        fecha_txt = fser.iloc[0].strftime("%d/%m/%Y")
    return d, resumen, top, fecha_txt

# ==========================================
# PIZARRA + TEXTO
# ==========================================
def html_pizarra_velocidad(resumen, top, fecha_txt, umbral, uid="vel"):
    bd = "border:1px solid #000;"
    total_ex = int(resumen["Excesos"].sum()) if not resumen.empty else 0
    n_placas = len(resumen)

    filas_r = ""
    for i, (_, r) in enumerate(resumen.iterrows()):
        fondo = "#f4f6f9" if i % 2 == 0 else "#ffffff"
        filas_r += f"""<tr style="background:{fondo};">
            <td style="{bd}padding:7px 10px;font-weight:bold;color:#0d2b57;">{_html.escape(su(r['Unidad']))}</td>
            <td style="{bd}padding:7px 10px;text-align:center;font-weight:bold;color:#c62828;">{int(r['Excesos'])}</td>
            <td style="{bd}padding:7px 10px;text-align:center;">{int(r['VelMax'])}</td>
            <td style="{bd}padding:7px 10px;">{_html.escape(su(r['Chofer']))}</td>
            <td style="{bd}padding:7px 10px;">{_html.escape(su(r['Ruta']))}</td>
        </tr>"""
    if not filas_r:
        filas_r = f"""<tr><td colspan="5" style="{bd}padding:16px;text-align:center;color:#888;">Sin excesos sobre {umbral} km/h.</td></tr>"""

    filas_t = ""
    for i, (_, r) in enumerate(top.iterrows()):
        fondo = "#fff3f3" if i % 2 == 0 else "#ffffff"
        fh = r["Fecha"].strftime("%d/%m %H:%M") if pd.notna(r["Fecha"]) else ""
        filas_t += f"""<tr style="background:{fondo};">
            <td style="{bd}padding:7px 10px;font-weight:bold;">{_html.escape(su(r['Unidad']))}</td>
            <td style="{bd}padding:7px 10px;white-space:nowrap;">{_html.escape(fh)}</td>
            <td style="{bd}padding:7px 10px;">{_html.escape(su(r['Estado']))}</td>
            <td style="{bd}padding:7px 10px;text-align:center;font-weight:bold;color:#c62828;">{int(r['Vel'])}</td>
            <td style="{bd}padding:7px 10px;font-size:12px;">{_html.escape(su(corta_dir(r['Loc'])))}</td>
        </tr>"""
    if not filas_t:
        filas_t = f"""<tr><td colspan="5" style="{bd}padding:16px;text-align:center;color:#888;">Sin datos.</td></tr>"""

    return f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right;margin-bottom:8px;">
      <button onclick="descargar_{uid}()" style="background:#0d47a1;color:#fff;border:none;padding:10px 18px;border-radius:6px;cursor:pointer;font-weight:bold;">⬇️ Descargar imagen</button>
    </div>
    <div id="piz-{uid}" style="background:#fff;font-family:Arial,sans-serif;{bd}">
      <div style="background:#b71c1c;color:#fff;padding:14px 16px;{bd}">
        <div style="font-size:20px;font-weight:bold;">🚨 EXCESOS DE VELOCIDAD (&gt; {umbral} km/h)</div>
        <div style="font-size:13px;opacity:.95;">Drotaca · {_html.escape(fecha_txt)} · Total excesos: {total_ex} · Placas: {n_placas}</div>
      </div>
      <div style="padding:8px 12px;font-weight:bold;background:#0d2b57;color:#fff;{bd}">📊 RANKING POR PLACA</div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;color:#1a1a1a;{bd}">
        <tr style="background:#1c3d6e;color:#fff;">
          <th style="{bd}padding:8px;text-align:left;">PLACA</th>
          <th style="{bd}padding:8px;">EXCESOS</th>
          <th style="{bd}padding:8px;">VEL. MÁX</th>
          <th style="{bd}padding:8px;text-align:left;">CHÓFER</th>
          <th style="{bd}padding:8px;text-align:left;">RUTA</th>
        </tr>
        {filas_r}
      </table>
      <div style="padding:8px 12px;font-weight:bold;background:#0d2b57;color:#fff;{bd}">🔝 TOP EXCESOS</div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;color:#1a1a1a;{bd}">
        <tr style="background:#1c3d6e;color:#fff;">
          <th style="{bd}padding:8px;text-align:left;">UNIDAD</th>
          <th style="{bd}padding:8px;text-align:left;">FECHA</th>
          <th style="{bd}padding:8px;text-align:left;">ESTADO</th>
          <th style="{bd}padding:8px;">VEL.</th>
          <th style="{bd}padding:8px;text-align:left;">LOCALIZACIÓN</th>
        </tr>
        {filas_t}
      </table>
    </div>
    <script>
    function descargar_{uid}() {{
        html2canvas(document.getElementById('piz-{uid}'), {{scale:2}}).then(function(canvas){{
            var link=document.createElement('a'); link.download='Exceso_Velocidad.png';
            link.href=canvas.toDataURL('image/png'); link.click();
        }});
    }}
    </script>
    """

def texto_velocidad(resumen, top, fecha_txt, umbral):
    total_ex = int(resumen["Excesos"].sum()) if not resumen.empty else 0
    L = [f"*EXCESOS DE VELOCIDAD {fecha_txt}* (> {umbral} km/h)",
         f"TOTAL EXCESOS: {total_ex} · PLACAS: {len(resumen)}", "", "🏎️ *POR PLACA:*", ""]
    if resumen.empty:
        L.append("Sin excesos.")
    else:
        for _, r in resumen.iterrows():
            extra = ""
            if su(r["Chofer"]): extra += f" · {su(r['Chofer'])}"
            if su(r["Ruta"]): extra += f" · {su(r['Ruta'])}"
            L.append(f"• {su(r['Unidad'])} · {int(r['Excesos'])} excesos · máx {int(r['VelMax'])}{extra}")
    L.append("")
    L.append("🔝 *TOP EXCESOS:*")
    L.append("")
    for _, r in top.iterrows():
        fh = r["Fecha"].strftime("%d/%m %H:%M") if pd.notna(r["Fecha"]) else ""
        L.append(f"• {int(r['Vel'])} km/h · {su(r['Unidad'])} · {fh} · {su(corta_dir(r['Loc']))}")
    return "\n".join(L)

def html_caja_copiar(texto, uid):
    texto = "" if texto is None else str(texto)
    return f"""
    <div style="font-family:Arial,sans-serif;">
      <textarea id="cp-{uid}" readonly style="width:100%;height:260px;padding:10px;border:1px solid #444;border-radius:8px;background:#0e1117;color:#fafafa;font-size:13px;resize:vertical;box-sizing:border-box;white-space:pre-wrap;">{_html.escape(texto)}</textarea>
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

# ==========================================
# INTERFAZ
# ==========================================
st.title("🚨 Exceso de Velocidad")
st.caption("Carga el Excel del reporte (la información empieza en la fila 9). Cuenta los excesos por placa "
           "sobre el umbral y cruza chófer/ruta con el Status del Día.")

usuario = str(st.session_state.get("usuario", "-")).upper()

c1, c2 = st.columns([2, 1])
archivo = c1.file_uploader("📂 Sube el Excel de Exceso de Velocidad", type=["xlsx", "xls"], key="vel_file")
umbral = c2.number_input("Umbral km/h (excesos por sobre)", min_value=0, max_value=300, value=106, step=1)
topn = c2.number_input("Top excesos a mostrar", min_value=1, max_value=20, value=3, step=1)

if archivo:
    try:
        status_map = leer_status_map()
    except Exception:
        status_map = {}
    try:
        df_raw = cargar_excel(archivo)
        d, resumen, top, fecha_txt = procesar(df_raw, umbral, status_map, int(topn))
        if not fecha_txt:
            fecha_txt = date.today().strftime("%d/%m/%Y")
        st.session_state["vel_data"] = {"resumen": resumen, "top": top, "fecha": fecha_txt}
    except Exception as e:
        st.error(f"No se pudo procesar el Excel: {e}")
        st.stop()

if "vel_data" in st.session_state:
    dd = st.session_state["vel_data"]
    resumen, top, fecha_txt = dd["resumen"], dd["top"], dd["fecha"]

    m1, m2, m3 = st.columns(3)
    m1.metric("Total excesos", int(resumen["Excesos"].sum()) if not resumen.empty else 0)
    m2.metric("Placas con excesos", len(resumen))
    m3.metric("Velocidad máxima", int(resumen["VelMax"].max()) if not resumen.empty else 0)

    st.markdown("### 🖼️ Pizarra")
    alto = 420 + (len(resumen) + len(top)) * 34
    components.html(html_pizarra_velocidad(resumen, top, fecha_txt, int(umbral)), height=int(alto), scrolling=True)

    st.markdown("### 📱 Texto para WhatsApp")
    components.html(html_caja_copiar(texto_velocidad(resumen, top, fecha_txt, int(umbral)), "veltxt"), height=340)

    st.markdown("---")
    if st.button("💾 Guardar en historial (Google Sheet)", type="primary", use_container_width=True):
        try:
            guardar_historial(fecha_txt, resumen, usuario)
            st.cache_data.clear()
            st.success(f"✅ Guardado en la hoja «{HOJA_HIST}».")
        except Exception as e:
            st.error(f"No se pudo guardar: {e}")

    with st.expander("📚 Ver historial guardado"):
        try:
            dfh = pd.DataFrame(abrir_libro().worksheet(HOJA_HIST).get_all_records())
            if dfh.empty:
                st.info("Aún no hay historial.")
            else:
                st.dataframe(dfh.iloc[::-1], use_container_width=True, hide_index=True)
        except Exception:
            st.info("Aún no hay historial.")
else:
    st.info("Sube un Excel para ver la pizarra.")