import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import textwrap
import traceback
import streamlit.components.v1 as components
from fpdf import FPDF
import io

# ==========================================
# CREDENCIALES Y CONEXIÓN A GOOGLE SHEETS
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

def extraer_datos_cierre_diario():
    try:
        cliente = obtener_cliente_sheets()
        doc = cliente.open_by_key("1L0N2O82bLzT1fE6kX5k-f5713tW3oN_YgOQzR5tY1R0")
        hoja = doc.worksheet("CIERRE_DIARIO")
        data = hoja.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error conectando a Google Sheets: {e}")
        return pd.DataFrame()

# ==========================================
# FUNCIONES DE APOYO (FORMATO Y WHATSAPP)
# ==========================================
def formato_numero(num):
    """Formatea un número usando puntos para los miles (Ej: 1806 -> 1.806)"""
    try:
        if pd.isna(num) or num == "": return "0"
        return f"{float(num):,.0f}".replace(",", ".")
    except:
        return "0"

def generar_ws_semanal(df, f_ini, f_fin, g_rutas, g_farm, g_bultos):
    f_ini_str = f_ini.strftime("%d/%m/%Y")
    f_fin_str = f_fin.strftime("%d/%m/%Y")
    
    msg = f"📊 *Reporte Semanal de Tráfico Drotaca* 🚚\n📅 Periodo: {f_ini_str} al {f_fin_str}\n\n"
    msg += "*ESTADÍSTICA GENERAL:*\n"
    msg += f"📍 Total Despachos: {g_rutas}\n"
    msg += f"🏥 Farmacias Entregadas: {formato_numero(g_farm)}\n"
    msg += f"📦 Total Bultos Procesados: {formato_numero(g_bultos)}\n\n"
    
    msg += "🌍 *RESUMEN POR REGIÓN:*\n"
    for reg in ["ORIENTE", "CENTRO OCCIDENTE", "OCCIDENTE"]:
        df_reg = df[df['Region_Filtro'] == reg]
        if not df_reg.empty:
            bultos = df_reg['Total_Bultos'].sum()
            porc = (bultos / g_bultos * 100) if g_bultos > 0 else 0
            msg += f"▪️ {reg}: {porc:.1f}% ({formato_numero(bultos)} Bultos)\n"
            
    msg += "\n✅ *Pizarra visual de tráfico adjunta.*"
    return msg

# ==========================================
# GENERADOR DE PDF PROFESIONAL
# ==========================================
def crear_pdf_reporte_semanal(df_filtrado, fecha_ini_str, fecha_fin_str, g_rutas, g_farmacias, g_bultos):
    pdf = FPDF()
    pdf.add_page()
    azul_drotaca = (13, 71, 161)
    
    # ENCABEZADO
    pdf.set_fill_color(*azul_drotaca)
    pdf.rect(0, 0, 210, 35, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "REPORTE SEMANAL DE TRAFICO Y DESPACHOS", ln=True, align='C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, f"Periodo Analizado: {fecha_ini_str} al {fecha_fin_str}", ln=True, align='C')
    pdf.cell(0, 6, "DEPARTAMENTO DE TRAFICO - DROTACA", ln=True, align='C')
    
    pdf.ln(15)
    pdf.set_text_color(0, 0, 0)
    
    # KPIs GENERALES
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "1. RESUMEN GENERAL DE OPERACIONES", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, f"Total Rutas/Despachos Ejecutados: {g_rutas}", ln=True)
    pdf.cell(0, 6, f"Total Farmacias Cubiertas: {formato_numero(g_farmacias)}", ln=True)
    pdf.cell(0, 6, f"Total Bultos Entregados: {formato_numero(g_bultos)}", ln=True)
    pdf.ln(10)
    
    # DETALLE POR REGIÓN
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "2. DETALLE DE DESEMPEÑO POR REGION", ln=True)
    
    orden_regiones = ["ORIENTE", "CENTRO OCCIDENTE", "OCCIDENTE"]
    
    for region in orden_regiones:
        df_region = df_filtrado[df_filtrado['Region_Filtro'] == region]
        if df_region.empty: continue
        
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(0, 8, f" MACRO ZONA: {region}", ln=True, fill=True)
        
        # CABECERAS DE TABLA
        pdf.set_font("Arial", 'B', 8)
        cols = [("FECHA", 20), ("RUTA", 50), ("CHOFER", 45), ("FARM.", 15), ("BULTOS", 15), ("SOBR", 15), ("FALT", 15), ("RECH.", 15)]
        for col_name, width in cols:
            pdf.cell(width, 7, col_name, border=1, align='C')
        pdf.ln()
        
        pdf.set_font("Arial", '', 7)
        for _, row in df_region.iterrows():
            f_str = row['Fecha_DT'].strftime("%d/%m") if pd.notnull(row['Fecha_DT']) else "-"
            pdf.cell(20, 6, f_str, border=1, align='C')
            pdf.cell(50, 6, str(row['Ruta'])[:30], border=1, align='L')
            pdf.cell(45, 6, str(row['Chofer'])[:25], border=1, align='L')
            pdf.cell(15, 6, str(row['Total_Farmacias']), border=1, align='C')
            pdf.cell(15, 6, str(row['Total_Bultos']), border=1, align='C')
            pdf.cell(15, 6, str(row['Bultos_Sobrantes']), border=1, align='C')
            pdf.cell(15, 6, str(row['Bultos_Faltantes']), border=1, align='C')
            pdf.cell(15, 6, str(row['Rechazos']), border=1, align='C')
            pdf.ln()
            
        # SUBTOTALES REGIÓN
        sub_f = df_region['Total_Farmacias'].sum()
        sub_b = df_region['Total_Bultos'].sum()
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(115, 7, f"SUBTOTAL {region}", border=1, align='R')
        pdf.cell(15, 7, formato_numero(sub_f), border=1, align='C')
        pdf.cell(15, 7, formato_numero(sub_b), border=1, align='C')
        pdf.cell(45, 7, "", border=1, align='C')
        pdf.ln()
    
    return pdf.output(dest='S').encode('latin-1', 'replace')


# ==========================================
# CREADOR DEL MEGA-HTML NACIONAL (TRÁFICO)
# ==========================================
def html_pizarra_nacional(df_global, fecha_ini_str, fecha_fin_str):
    color_header = "#0d47a1" 
    
    g_rutas = len(df_global)
    g_farmacias = df_global['Total_Farmacias'].sum()
    g_bultos = df_global['Total_Bultos'].sum()
    
    html_tablas_regiones = ""
    orden_regiones = ["ORIENTE", "CENTRO OCCIDENTE", "OCCIDENTE"]
    
    for region in orden_regiones:
        df_region = df_global[df_global['Region_Filtro'] == region]
        if df_region.empty: continue
        
        t_rutas_reg = len(df_region)
        t_farmacias_reg = df_region['Total_Farmacias'].sum()
        t_bultos_reg = df_region['Total_Bultos'].sum()
        
        # Título de Macro Zona
        html_tablas_regiones += f"""
        <div style="margin-top: 30px;">
            <h3 style="background-color: {color_header}; color: white; padding: 10px; margin-bottom: 0; font-size: 20px; text-transform: uppercase; border-radius: 5px 5px 0 0;">
                🌍 MACRO ZONA: {region}
            </h3>
            <table style="width: 100%; border-collapse: collapse; border: 2px solid #000; margin-bottom: 10px;">
                <thead>
                    <tr style="background-color: #e0e0e0; font-size: 13px; text-align: center; color: #000;">
                        <th style="border: 1px solid #000; padding: 10px;">FECHA</th>
                        <th style="border: 1px solid #000; padding: 10px; width: 20%;">RUTA</th>
                        <th style="border: 1px solid #000; padding: 10px; width: 15%;">CHOFER</th>
                        <th style="border: 1px solid #000; padding: 10px; width: 15%;">AYUDANTE</th>
                        <th style="border: 1px solid #000; padding: 10px;">FARMACIAS</th>
                        <th style="border: 1px solid #000; padding: 10px;">BULTOS</th>
                        <th style="border: 1px solid #000; padding: 10px;">FALT.</th>
                        <th style="border: 1px solid #000; padding: 10px;">SOBR.</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for i, r in df_region.iterrows():
            bg = "#f9f9f9" if i % 2 != 0 else "#ffffff"
            f_str = r['Fecha_DT'].strftime("%d/%m") if pd.notnull(r['Fecha_DT']) else "-"
            
            # Formateamos valores a puntos
            f_farm = formato_numero(r['Total_Farmacias'])
            f_bultos = formato_numero(r['Total_Bultos'])
            
            # Alertas visuales para faltantes/sobrantes
            color_falt = "#d32f2f" if r['Bultos_Faltantes'] > 0 else "#000"
            color_sobr = "#f57c00" if r['Bultos_Sobrantes'] > 0 else "#000"
            
            html_tablas_regiones += f"""
                <tr style="background-color: {bg}; text-align: center; font-size: 14px; color: #333;">
                    <td style="border: 1px solid #ccc; padding: 8px;">{f_str}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; font-weight: bold; text-align: left;">{r['Ruta']}</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">{r['Chofer']}</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">{r['Ayudante']}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; font-weight: bold; font-size: 15px;">{f_farm}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; font-weight: bold; font-size: 15px; color: #0d47a1;">{f_bultos}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; font-weight: bold; color: {color_falt};">{r['Bultos_Faltantes']}</td>
                    <td style="border: 1px solid #ccc; padding: 8px; font-weight: bold; color: {color_sobr};">{r['Bultos_Sobrantes']}</td>
                </tr>
            """
            
        html_tablas_regiones += f"""
                </tbody>
                <tfoot>
                    <tr style="background-color: #ffe082; text-align: center; font-size: 16px; color: #000;">
                        <td colspan="4" style="border: 1px solid #000; padding: 10px; text-align: right; font-weight: bold;">SUBTOTAL {region}:</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold; color: #e65100;">{formato_numero(t_farmacias_reg)}</td>
                        <td style="border: 1px solid #000; padding: 10px; font-weight: bold; color: #0d47a1;">{formato_numero(t_bultos_reg)}</td>
                        <td colspan="2" style="border: 1px solid #000; padding: 10px; font-size: 12px; color: #555;">({t_rutas_reg} Rutas Procesadas)</td>
                    </tr>
                </tfoot>
            </table>
        </div>
        """

    bloque_estadistica_general = f"""
    <div style="margin-top: 40px; padding: 25px; border: 3px solid {color_header}; border-radius: 10px; background-color: #f8f9fa;">
        <h2 style="text-align: center; color: {color_header}; margin-top: 0; margin-bottom: 20px; font-size: 26px; text-transform: uppercase; letter-spacing: 1px;">🌎 Estadística General de la Semana</h2>
        <div style="display: flex; justify-content: space-around; align-items: center;">
            <div style="text-align: center; flex: 1; border-right: 2px solid #ddd;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #666; text-transform: uppercase;">Total Despachos</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #333;">{g_rutas}</p>
            </div>
            <div style="text-align: center; flex: 1; border-right: 2px solid #ddd;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #1565c0; text-transform: uppercase;">Total Farmacias</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #1565c0;">{formato_numero(g_farmacias)}</p>
            </div>
            <div style="text-align: center; flex: 1;">
                <p style="margin: 0; font-size: 14px; font-weight: bold; color: #e65100; text-transform: uppercase;">Total Bultos</p>
                <p style="margin: 5px 0 0 0; font-size: 32px; font-weight: bold; color: #e65100;">{formato_numero(g_bultos)}</p>
            </div>
        </div>
    </div>
    """

    html_final = f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <div style="text-align:right; margin-bottom:15px;">
        <button onclick="descargar_pizarra_semanal()" style="background:#d32f2f;color:#fff;border:none;padding:12px 25px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:16px;box-shadow:0 4px 6px rgba(0,0,0,0.2);">⬇️ DESCARGAR PIZARRA SEMANAL</button>
    </div>
    
    <div id="pizarra-semanal" style="font-family: Arial, sans-serif; width: 1100px; margin: auto; background-color: #fff; border: 3px solid #000; box-sizing: border-box; padding-bottom: 20px;">
        <div style="background-color: {color_header}; color: white; padding: 25px; text-align: center; border-bottom: 5px solid #ffca28;">
            <h2 style="margin: 0; font-size: 30px; letter-spacing: 2px; text-transform: uppercase;">🌎 AUDITORÍA SEMANAL DE TRÁFICO</h2>
            <p style="margin: 5px 0 0 0; font-size: 18px; color: #bbdefb; font-weight: bold; text-transform: uppercase;">Departamento de Tráfico</p>
            <p style="margin: 10px 0 0 0; font-size: 16px;">Periodo: {fecha_ini_str} al {fecha_fin_str}</p>
        </div>
        
        <div style="padding: 0 30px;">
            {html_tablas_regiones}
            {bloque_estadistica_general}
        </div>
        <div style="margin-top: 30px; text-align: center; font-size: 14px; color: #666; font-weight: bold;">
            Control Tower - Drotaca Flota 2026
        </div>
    </div>
    <script>
    function descargar_pizarra_semanal() {{ 
        html2canvas(document.getElementById('pizarra-semanal'), {{scale: 2}}).then(canvas => {{ 
            var link = document.createElement('a'); 
            link.download = 'Pizarra_Trafico_Semanal.png'; 
            link.href = canvas.toDataURL(); link.click(); 
        }}); 
    }}
    </script>
    """
    return html_final

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("Master Reporte Semanal: Desempeño de Tráfico")
st.info("Este módulo consolida la data de despachos diarios (Farmacias, Bultos, Faltantes, etc.) para crear el Cierre Semanal Gerencial.")

# --- FILTROS DE FECHA ---
hoy = datetime.now()
lunes_pasado = hoy - timedelta(days=hoy.weekday())

c1, c2 = st.columns(2)
with c1:
    f_inicio = st.date_input("Fecha Inicio (Lunes):", lunes_pasado)
with c2:
    f_fin = st.date_input("Fecha Fin (Domingo):", hoy)

if st.button("1️⃣ Procesar Reporte Semanal", type="primary", use_container_width=True):
    with st.spinner("Extrayendo historial de la bóveda..."):
        df_db = extraer_datos_cierre_diario()
        
        if df_db.empty:
            st.error("No se pudo conectar con la base de datos o está vacía.")
        else:
            # Limpieza y preparación de datos
            df_db['Fecha_DT'] = pd.to_datetime(df_db['Fecha_Reporte'], format='%d/%m/%Y', errors='coerce')
            
            # Filtro por rango
            df_filtrado = df_db[(df_db['Fecha_DT'].dt.date >= f_inicio) & (df_db['Fecha_DT'].dt.date <= f_fin)].copy()
            
            if df_filtrado.empty:
                st.warning("No hay registros de tráfico guardados para ese rango de fechas.")
            else:
                # Asegurar columnas numéricas
                num_cols = ['Total_Farmacias', 'Total_Bultos', 'Bultos_Sobrantes', 'Bultos_Faltantes', 'Rechazos']
                for col in num_cols:
                    if col in df_filtrado.columns:
                        df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors='coerce').fillna(0)
                    else:
                        df_filtrado[col] = 0

                st.success(f"✅ ¡Datos compilados! Se encontraron {len(df_filtrado)} rutas en este periodo.")
                
                # Renderizar la Pizarra HTML
                f_ini_str = f_inicio.strftime("%d/%m/%Y")
                f_fin_str = f_fin.strftime("%d/%m/%Y")
                
                components.html(html_pizarra_nacional(df_filtrado, f_ini_str, f_fin_str), height=1400, scrolling=True)
                
                # CÁLCULO PARA BOTONES PDF Y WHATSAPP
                g_rutas = len(df_filtrado)
                g_farm = df_filtrado['Total_Farmacias'].sum()
                g_bultos = df_filtrado['Total_Bultos'].sum()
                
                st.markdown("---")
                st.subheader("📱 Mensaje para WhatsApp")
                st.code(generar_ws_semanal(df_filtrado, f_inicio, f_fin, g_rutas, g_farm, g_bultos), language="markdown")
                
                st.markdown("---")
                # Botón de Descarga PDF (Manteniendo la funcionalidad que ya te gustaba)
                pdf_bytes = crear_pdf_reporte_semanal(df_filtrado, f_ini_str, f_fin_str, g_rutas, g_farm, g_bultos)
                st.download_button(
                    label="📄 DESCARGAR INFORME SEMANAL EN PDF",
                    data=pdf_bytes,
                    file_name=f"Cierre_Semanal_Trafico_{f_fin.strftime('%d%m%y')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
