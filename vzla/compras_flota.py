# ---------------------------------------------------------
# PESTAÑA 3: HISTORIAL Y AUDITORÍA (DISEÑO MEJORADO)
# ---------------------------------------------------------
with tab_metricas:
    st.header("📊 Auditoría de Cumplimiento y Eficiencia")
    
    if not df_global.empty:
        # Procesamiento de datos para la auditoría
        df_audit = df_global.copy()
        df_audit['Estatus'] = df_audit['Estatus'].str.strip()
        df_audit['Fecha_Solicitud'] = pd.to_datetime(df_audit['Fecha_Solicitud'], format='%d/%m/%Y', errors='coerce')
        
        # Cálculos de KPIs
        df_comp = df_audit[df_audit['Estatus'] == 'COMPRADO'].copy()
        df_pend = df_audit[df_audit['Estatus'] == 'PENDIENTE'].copy()
        
        # Indicadores en la parte superior (KPIs)
        m1, m2, m3, m4 = st.columns(4)
        
        total_sol = len(df_audit)
        entregados = len(df_comp)
        pendientes = len(df_pend)
        efectividad = (entregados / total_sol * 100) if total_sol > 0 else 0
        
        # Tiempo promedio de resolución
        if not df_comp.empty:
            promedio_dias = pd.to_numeric(df_comp['Dias_Resolucion'], errors='coerce').mean()
        else:
            promedio_dias = 0

        m1.metric("Total Solicitudes", total_sol)
        m2.metric("Efectividad", f"{efectividad:.1f}%")
        m3.metric("Pendientes", pendientes, delta=f"{pendientes} ítems", delta_color="inverse")
        m4.metric("Promedio Entrega", f"{promedio_dias:.1f} días")

        # Alerta Crítica (Más de 15 pendientes)
        if pendientes >= 15:
            st.error(f"🚨 **ALERTA DE GESTIÓN:** Se han acumulado {pendientes} solicitudes sin atender. Es necesario revisar el flujo de compras.")

        st.markdown("---")
        st.subheader("📋 Detalle General de Movimientos")

        # Preparar la tabla visual
        # Para pendientes, calculamos los días que llevan esperando hasta hoy
        df_audit['Días'] = 0
        hoy = datetime.now()
        
        for idx, row in df_audit.iterrows():
            if row['Estatus'] == 'PENDIENTE' and pd.notnull(row['Fecha_Solicitud']):
                dias_espera = (hoy - row['Fecha_Solicitud']).days
                df_audit.at[idx, 'Días'] = dias_espera
            else:
                df_audit.at[idx, 'Días'] = row['Dias_Resolucion']

        # Selección de columnas para mostrar (como en la imagen)
        df_visual = df_audit[['Fecha_Solicitud', 'Descripcion', 'Cantidad', 'Usuario', 'Estatus', 'Días']].copy()
        df_visual.columns = ['FECHA', 'ITEM / DESCRIPCIÓN', 'CANT', 'SOLICITADO POR', 'ESTADO', 'DÍAS']
        
        # Formatear fecha para mostrar
        df_visual['FECHA'] = df_visual['FECHA'].dt.strftime('%d/%m/%Y')

        # Aplicar Estilos de Color (Configuración de columnas de Streamlit)
        st.dataframe(
            df_visual.sort_values('FECHA', ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "ESTADO": st.column_config.SelectboxColumn(
                    "ESTADO",
                    options=["PENDIENTE", "COMPRADO"],
                    required=True,
                ),
                "DÍAS": st.column_config.NumberColumn(
                    "DÍAS",
                    format="%d d",
                    help="Días transcurridos desde la solicitud hasta la entrega (o hasta hoy si está pendiente)"
                ),
            }
        )

        # Resumen por Categoría (Opcional, para mayor detalle)
        with st.expander("🔍 Ver resumen por tipo de solicitud"):
            resumen_cat = df_audit.groupby('Categoria').size().reset_index(name='Cantidad')
            st.table(resumen_cat)

    else:
        st.info("Aún no hay registros en el historial para mostrar métricas.")
