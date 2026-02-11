import streamlit as st
import pandas as pd
import io
import numpy as np
import plotly.express as px
import plotly.graph_objects as go # Necesario para agregar la línea al gráfico de barras
import time 

# ==============================================================================
# 1. CONFIGURACIÓN Y TÍTULO GLOBAL (NAVBAR)
# ==============================================================================
st.set_page_config(page_title="FTE App", layout="wide")

st.markdown("""
    <style>
        .block-container {padding-top: 1rem;}
        .stTabs [data-baseweb="tab-list"] {gap: 10px;}
        .stTabs [data-baseweb="tab"] {height: 50px; white-space: pre-wrap; background-color: #f0f2f6; border-radius: 4px 4px 0px 0px; gap: 1px; padding-top: 10px; padding-bottom: 10px;}
        .stTabs [aria-selected="true"] {background-color: #ffffff; border-bottom: 2px solid #ff4b4b;}
    </style>
""", unsafe_allow_html=True)


st.markdown("---") 

tab1, tab2, tab_diario, tab3 = st.tabs(["🕵️ Validación de Pesos", "⏱️ FTE MENSUAL", "⏱️ FTE DIARIO", "⚙️ Días Trabajados"])

# ==============================================================================
# BARRA LATERAL
# ==============================================================================
st.sidebar.title("📊 FTE App")
st.sidebar.markdown("---") 
st.sidebar.header("📂 Carga de Archivos")
file_solicitudes = st.sidebar.file_uploader("1. Excel Solicitudes (Power APP)", type=['xlsx'])
st.sidebar.markdown("--") 
file_pesos = st.sidebar.file_uploader("2. Excel Pesos", type=['xlsx'])
st.sidebar.markdown("--") 
file_prod = st.sidebar.file_uploader("3. Excel Días Trabajados (Contiene las vacaciones)", type=['xlsx'])

def limpiar_texto(texto):
    if pd.isna(texto):
        return ""
    texto_str = str(texto).upper().strip()
    texto_str = texto_str.replace('\xa0', ' ')
    texto_str = " ".join(texto_str.split())
    return texto_str

# ==============================================================================
# PESTAÑA 1: VALIDACIÓN
# ==============================================================================
with tab1:
    st.subheader("🕵️ Validación de Pesos y Solicitudes")
    st.info("Acá verificamos que todos los procesos del PowerAPP tienen su peso respectivo en el excel de Pesos")
    if file_solicitudes and file_pesos:
        st.info("Analizando coincidencias...")
        try:
            df_sol = pd.read_excel(file_solicitudes)
            df_pesos_data = pd.read_excel(file_pesos)
        except Exception as e:
            st.error(f"Error al leer archivos: {e}")
            st.stop()

        df_sol.columns = df_sol.columns.str.strip()
        col_resolutor = None
        for posible in ['Resolutor', 'RESOLUTOR', 'Nombre Resolutor', 'Nombre Técnico']:
            if posible in df_sol.columns:
                col_resolutor = posible
                break
        
        if col_resolutor:
            df_sol.rename(columns={col_resolutor: 'Resolutor'}, inplace=True)
            empleados_permitidos = [
                "JESSICA ACUNA VELASQUEZ", "ALEJANDRA MATUS DURAN", "DIANA CARRASCO HERRERA",
                "CLEMENTINA GALAZ MATTA", "BRENDA OLGUIN QUIROZ", 
                "STEPHANIE CIFUENTES LUENGO", "KARINNA ALVAREZ MORALES"
            ]
            df_sol['Resolutor'] = df_sol['Resolutor'].astype(str).str.upper().str.strip()
            df_sol = df_sol[df_sol['Resolutor'].isin(empleados_permitidos)].copy()
            st.success(f"✅ Filtro de personal aplicado: {len(df_sol)} registros.")
        else:
            st.warning("⚠️ No encontré columna 'Resolutor'.")

        if 'Tipo de Pedido' not in df_sol.columns:
            st.error("Falta la columna 'Tipo de Pedido'.")
            st.stop()
        
        df_sol['Tipo de Pedido Normalizado'] = df_sol['Tipo de Pedido'].apply(limpiar_texto)

        correcciones_manuales = {
            "1. SOLICITUDES NIVEL 2": "SOLICITUDES NIVEL 2",
            "ACLARACIONES DE CARGOS ABONOS": "ACLARACIONES DE CARGOS Y ABONOS",
            "CERTIFICADO DE SALDO": "CERTIFICADO DE SALDOS",
            "CONDONACION DE GASTOS": "CONDONACIÓN DE GASTOS",
            "ENTREGA DE PAGARES ABOGADO ASIGNADO": "ENTREGA DE PAGARÉS ABOGADO ASIGNADO",
            "INICIO - TERMINO DE DÍA CONTABLE": "INICIO - TÉRMINO DE DÍA CONTABLE",
            "SIMULACIÓN DE CRÉDITOS": "SIMULACIÓN DE CRÉDITO",
            "PAGO DE HONORARIOS": "PAGO HONORARIOS",
            "SOLICITUD EMISIÓN DE PAGARÉ": "SOLICITUD EMISIÓN PAGARÉ",
            "APLICACIÓN DE REMATE, DACIÓN EN PAGO O CONSIGNACIONES": "APLICACIÓN DE REMATE, DACIÓN EN PAGO O CONSIGNAC",
            "EMISIÓN DE VALE VISTA VIRTUAL O ABONO A CUENTA BCI U OTRO BANCO": "EMISIÓN DE VALE VISTA VIRTUAL O ABONO A CUENTA BCI",
            "ANB EMISIÓN DE VALE VISTA VIRTUAL O ABONO A CUENTA BCI U OTRO BANCO": "EMISIÓN DE VALE VISTA VIRTUAL O ABONO A CUENTA BCI",
            "FOGAPE: CURSES PRORROGAS , MODIFICACIONES Y SEGUIMIENTO": "FOGAPE: CURSES PRÓRROGAS, MODIFICACIONES Y SEGUI",
            "PROCESO LIR - CONDONACIÓN POR SENTENCIA DE TÉRMINO": "PROCESO LIR - CONDONACIÓN POR SENTENCIA DE TÉRMI",
            "TRASLADO DE PAGARES": "RECEPCIÓN PAGARÉS OFICINA", 
            "INICIO DE DÍA CONTABLE": "INICIO - TÉRMINO DE DÍA CONTABLE"
        }
        df_sol['Tipo de Pedido Normalizado'] = df_sol['Tipo de Pedido Normalizado'].replace(correcciones_manuales)

        df_pesos_data.columns = df_pesos_data.columns.str.strip()
        df_pesos_data['TIPO DE PEDIDO'] = df_pesos_data['TIPO DE PEDIDO'].apply(limpiar_texto)
        dict_scores = df_pesos_data.set_index('TIPO DE PEDIDO')['Score'].to_dict()

        df_sol['Score_Encontrado'] = df_sol['Tipo de Pedido Normalizado'].map(dict_scores)
        df_faltantes = df_sol[df_sol['Score_Encontrado'].isna()].copy()

        if len(df_faltantes) > 0:
            st.error(f"⛔ Faltan {len(df_faltantes)} Scores.")
            resumen_faltantes = df_faltantes['Tipo de Pedido Normalizado'].value_counts().reset_index()
            resumen_faltantes.columns = ['NOMBRE EXACTO A COPIAR', 'CANTIDAD']
            st.warning(f"""1. Copia el nombre y agrega el proceso al excel de Pesos
2. Asignale un Peso 
3. Cambiale la columna Peso y Score (IMPORTANTE CAMBIAR AMBOS)
4. Sube el archivo actualizado""")
            st.table(resumen_faltantes)
        else:
            st.success("✅ Todos los procesos tienen Score.")
            df_sol['Score_Final'] = df_sol['Score_Encontrado'] + 2.5 
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                df_sol.to_excel(writer, index=False)
            
            st.download_button("📥 Descargar Excel Scores", data=buffer, file_name="Solicitudes_Scores.xlsx")
    else:
        st.warning("👈 Carga 'Solicitudes' y 'Pesos' en la barra lateral.")

# ==============================================================================
# PESTAÑA 2: CÁLCULO DE FTE MENSUAL
# ==============================================================================
with tab2:
    st.subheader("🚀 Cálculo de FTE Mensual")
    st.markdown("Cruce de **Scores** + **Tiempos** para calcular carga laboral.")
    
    with st.form("formulario_configuracion_mensual"):
        st.write("Configuración de parámetros:")
        col_conf1, col_conf2, col_conf3 = st.columns(3)
        with col_conf1:
            OLE_USADO_M = st.number_input("Factor OLE (Eficiencia)", value=0.66, step=0.01, min_value=0.1, max_value=1.0, key="ole_m")
        with col_conf2:
            HORA_DIARIA_M = st.number_input("Horas Diarias Contrato", value=7.9, step=0.01, key="hora_m")
        with col_conf3:
            ANIO_SELECCIONADO_M = st.number_input("📅 Año a Procesar", value=2025, step=1, min_value=2023, max_value=2030, key="anio_m")
        st.caption("Ajusta los valores y presiona el botón.")
        submitted_m = st.form_submit_button("🔄 Calcular FTE Mensual", type="primary")

    if submitted_m:
        if not (file_solicitudes and file_pesos and file_prod):
            st.warning("⚠️ Faltan archivos. Por favor carga **Solicitudes**, **Pesos** y **Días Trabajados** en el menú lateral.")
        else:
            try:
                progress_text = "Iniciando motor de cálculo..."
                my_bar = st.progress(0, text=progress_text)
                
                my_bar.progress(10, text="📂 Leyendo archivos Excel...")
                df_s = pd.read_excel(file_solicitudes)
                df_p = pd.read_excel(file_pesos)
                
                my_bar.progress(30, text="🧹 Limpiando y asignando Scores...")
                col_fecha = None
                for c in df_s.columns:
                    if 'fin real' in c.lower():
                        col_fecha = c
                        break
                if col_fecha == None:
                    for c in df_s.columns:
                        if 'fecha de creación' in c.lower():
                            col_fecha = c
                            break
                if not col_fecha:
                    st.error("❌ No encontré la columna de 'Fin Real'.")
                    st.stop()
                
                df_s[col_fecha] = pd.to_datetime(df_s[col_fecha], errors='coerce')
                df_s['Mes_Num'] = df_s[col_fecha].dt.month
                df_s['Año'] = df_s[col_fecha].dt.year
                df_s = df_s.dropna(subset=['Mes_Num', 'Año']) 
                df_s['Mes_Num'] = df_s['Mes_Num'].astype(int)
                df_s['Año'] = df_s['Año'].astype(int)
                df_s = df_s[df_s['Año'] == ANIO_SELECCIONADO_M].copy()
                
                if len(df_s) == 0:
                    st.error(f"⚠️ No hay registros en Solicitudes para el año {ANIO_SELECCIONADO_M}.")
                    st.stop()

                col_res = [c for c in df_s.columns if 'Resolutor' in c or 'Técnico' in c][0]
                df_s.rename(columns={col_res: 'Resolutor'}, inplace=True)
                df_s['Resolutor'] = df_s['Resolutor'].astype(str).str.upper().str.strip()
                
                df_s['Tipo Limpio'] = df_s['Tipo de Pedido'].apply(limpiar_texto)
                df_s['Tipo Limpio'] = df_s['Tipo Limpio'].replace(correcciones_manuales)
                
                df_p['TIPO DE PEDIDO'] = df_p['TIPO DE PEDIDO'].apply(limpiar_texto)
                scores_dict = df_p.set_index('TIPO DE PEDIDO')['Score'].to_dict()
                
                df_s['Score_Unitario'] = df_s['Tipo Limpio'].map(scores_dict)
                df_s['Score_Unitario'] = df_s['Score_Unitario'].fillna(0) + 2.5 
                
                df_pedidos_resumen = df_s.groupby(['Resolutor', 'Mes_Num'])['Score_Unitario'].sum().reset_index()

                my_bar.progress(60, text="⏱️ Calculando tiempos de reuniones y chats...")
                MAPA_EMPLEADOS = {
                    "JACUNVE": "JESSICA ACUNA VELASQUEZ", "AMATUSD": "ALEJANDRA MATUS DURAN",
                    "DCARRAH": "DIANA CARRASCO HERRERA", "CGALAZ": "CLEMENTINA GALAZ MATTA",
                    "BOLGUIQ": "BRENDA OLGUIN QUIROZ", "SCIFUEN": "STEPHANIE CIFUENTES LUENGO",
                    "KARINNA": "KARINNA ALVAREZ MORALES"
                }
                
                try:
                    df_prod_raw = pd.read_excel(file_prod, sheet_name='HorasTotales')
                except:
                    df_prod_raw = pd.read_excel(file_prod)
                
                col_anio_prod = None
                for c in df_prod_raw.columns:
                    if str(c).strip().lower() in ['año', 'anio', 'year', 'ano']:
                        col_anio_prod = c
                        break
                if col_anio_prod:
                    df_prod_raw = df_prod_raw[df_prod_raw[col_anio_prod] == ANIO_SELECCIONADO_M].copy()

                if "Nombre Técnico" in df_prod_raw.columns:
                     df_prod_raw["Resolutor"] = df_prod_raw["Nombre Técnico"].map(MAPA_EMPLEADOS)
                
                df_equipo = df_prod_raw.dropna(subset=['Resolutor']).copy()
                tabla_dias = df_equipo.pivot_table(index='Resolutor', columns='Número Mes', values='Dias Trabajados', aggfunc='sum').fillna(0)
                
                REU_SEMANAL_TOTAL = (20 * 3) + 50 + 50 
                MINUTOS_REU_DIARIA = REU_SEMANAL_TOTAL / 5
                tabla_reuniones = tabla_dias * MINUTOS_REU_DIARIA
                for m in [1, 7]:
                    if m in tabla_reuniones.columns:
                        tabla_reuniones.loc[tabla_reuniones[m] > 0, m] += 60 
                        
                MIN_CHAT_STD = 47
                MIN_CHAT_ESP = 90
                tabla_chats = tabla_dias * MIN_CHAT_STD
                if "BRENDA OLGUIN QUIROZ" in tabla_chats.index:
                    tabla_chats.loc["BRENDA OLGUIN QUIROZ"] = tabla_dias.loc["BRENDA OLGUIN QUIROZ"] * MIN_CHAT_ESP

                my_bar.progress(85, text="🔄 Cruzando datos y generando KPI...")
                def pivotar_tabla(df_wide, nombre_valor):
                    df_wide = df_wide.reset_index()
                    df_wide.columns = [str(c) for c in df_wide.columns]
                    vars_cols = [c for c in df_wide.columns if c != 'Resolutor']
                    df_melted = df_wide.melt(id_vars=['Resolutor'], value_vars=vars_cols, var_name='Mes_Num', value_name=nombre_valor)
                    df_melted['Mes_Num'] = pd.to_numeric(df_melted['Mes_Num'], errors='coerce')
                    return df_melted.dropna(subset=['Mes_Num'])

                df_reuniones_long = pivotar_tabla(tabla_reuniones, 'Minutos_Reunion')
                df_dias_long = pivotar_tabla(tabla_dias, 'Dias_Trabajados')
                df_chat_long = pivotar_tabla(tabla_chats, 'Minutos_Chat')

                df_final = pd.merge(df_reuniones_long, df_pedidos_resumen, on=['Resolutor', 'Mes_Num'], how='left').fillna(0)
                df_final = pd.merge(df_final, df_dias_long, on=['Resolutor', 'Mes_Num'], how='left').fillna(0)
                df_final = pd.merge(df_final, df_chat_long, on=['Resolutor', 'Mes_Num'], how='left').fillna(0)

                df_final = df_final[ (df_final['Dias_Trabajados'] > 0) | (df_final['Score_Unitario'] > 0) ]

                def calcular_fte_row(row):
                    horas_p = HORA_DIARIA_M - 1 if row['Resolutor'] == "STEPHANIE CIFUENTES LUENGO" else HORA_DIARIA_M
                    denominador = (horas_p * 60 * row['Dias_Trabajados'] * OLE_USADO_M)
                    if denominador == 0: return 0
                    numerador = row['Score_Unitario'] + row['Minutos_Reunion'] + row['Minutos_Chat']
                    return numerador / denominador

                df_final['FTE'] = df_final.apply(calcular_fte_row, axis=1)
                df_final = df_final.sort_values(by=['Mes_Num', 'Resolutor'])

                df_fte_mes = df_final.groupby(["Mes_Num"])["FTE"].sum().reset_index()
                df_fte_mes['Capacidad_Minima_Personas'] = np.ceil(df_fte_mes['FTE']).astype(int)
                
                df_final.insert(1, 'Año', ANIO_SELECCIONADO_M)
                df_fte_mes.insert(0, 'Año', ANIO_SELECCIONADO_M)

                st.session_state['df_fte_final'] = df_final
                st.session_state['df_fte_mes'] = df_fte_mes
                st.session_state['anio_calculado'] = ANIO_SELECCIONADO_M
                st.session_state['calculo_realizado'] = True
                
                my_bar.progress(100, text="✅ ¡Cálculo completado!")
                time.sleep(1)
                my_bar.empty()

            except Exception as e:
                st.error(f"Ocurrió un error en el cálculo: {e}")
                st.write("Detalle del error:", e)
    
    if st.session_state.get('calculo_realizado'):
        df_final = st.session_state['df_fte_final']
        df_fte_mes = st.session_state['df_fte_mes']
        anio_actual = st.session_state.get('anio_calculado', '---')
        
        st.success(f"Visualizando datos del año: **{anio_actual}**")

        subtab_graf, subtab_datos = st.tabs(["📈 Visualización Gráfica", "📋 Tablas y Descarga"])
        
        with subtab_graf:
            st.markdown("#### Análisis Gráfico de FTE")
            lista_personas = ["Todos"] + list(df_final['Resolutor'].unique())
            seleccion = st.selectbox("Filtrar por:", lista_personas)
            
            if seleccion == "Todos":
                df_grafico = df_fte_mes.melt(
                    id_vars=['Mes_Num', 'Año'], 
                    value_vars=['FTE', 'Capacidad_Minima_Personas'],
                    var_name='Indicador', 
                    value_name='Valor'
                )
                df_grafico['Mes_Num'] = df_grafico['Mes_Num'].astype(str)
                
                fig = px.bar(
                    df_grafico, x='Mes_Num', y='Valor', color='Indicador',
                    barmode='group', title=f'FTE Mensual {anio_actual} vs Capacidad Mínima',
                    text_auto='.2f',
                    color_discrete_map={'FTE': '#1f77b4', 'Capacidad_Minima_Personas': '#ff7f0e'}
                )
                
                # --- NUEVO: LÍNEA DE CAPACIDAD REAL (Igual que Pestaña 4) ---
                # Calculamos la capacidad real basada en días trabajados
                # df_final ya tiene los días trabajados por persona por mes
                df_capacidad_linea = df_final[df_final['Dias_Trabajados'] > 0].groupby('Mes_Num').agg(
                    Total_Dias=('Dias_Trabajados', 'sum'),
                    Max_Dias=('Dias_Trabajados', 'max')
                ).reset_index()
                
                if not df_capacidad_linea.empty:
                    df_capacidad_linea['Capacidad_Real'] = df_capacidad_linea['Total_Dias'] / df_capacidad_linea['Max_Dias']
                    df_capacidad_linea['Mes_Num'] = df_capacidad_linea['Mes_Num'].astype(str)
                    
                    fig.add_scatter(
                        x=df_capacidad_linea['Mes_Num'],
                        y=df_capacidad_linea['Capacidad_Real'],
                        mode='lines+markers',
                        name='Capacidad Real (Personas Disponibles)',
                        line=dict(color='#00CC96', width=3, dash='dot')
                    )
                # ------------------------------------------------------------

                fig.update_layout(xaxis_title="Mes", yaxis_title="Valor FTE / Personas")
                st.plotly_chart(fig, use_container_width=True)
                
            else:
                df_persona = df_final[df_final['Resolutor'] == seleccion].copy()
                df_persona['Mes_Num'] = df_persona['Mes_Num'].astype(str)
                
                fig = px.bar(
                    df_persona, x='Mes_Num', y='FTE',
                    title=f'Evolución FTE {anio_actual}: {seleccion}',
                    text_auto='.2f',
                    color_discrete_sequence=['#2ca02c']
                )
                fig.add_hline(y=1, line_dash="dash", line_color="red", annotation_text="Límite (1.0)", annotation_position="top right")
                fig.add_hline(y=0.8, line_dash="dash", line_color="green", annotation_text="Meta (0.8)", annotation_position="bottom right")
                fig.update_layout(xaxis_title="Mes", yaxis_title="FTE (Carga Laboral)")
                st.plotly_chart(fig, use_container_width=True)

        with subtab_datos:
            st.subheader("1. Detalle por Persona y Mes")
            st.dataframe(df_final.style.format({"FTE": "{:.2f}", "Score_Unitario": "{:.0f}", "Año": "{:.0f}"}), use_container_width=True)
            st.subheader("2. Resumen Gerencial (FTE Total x Mes)")
            st.dataframe(df_fte_mes.style.format({"FTE": "{:.2f}", "Año": "{:.0f}"}), use_container_width=True)
            buffer_fte = io.BytesIO()
            with pd.ExcelWriter(buffer_fte) as writer:
                df_final.to_excel(writer, sheet_name="Detalle_FTE", index=False)
                df_fte_mes.to_excel(writer, sheet_name="Resumen_Mes", index=False)
            st.download_button("📥 Descargar Reporte FTE Completo", data=buffer_fte, file_name=f"Reporte_FTE_{anio_actual}.xlsx")

# ==============================================================================
# PESTAÑA 3: CÁLCULO DE FTE DIARIO
# ==============================================================================
with tab_diario:
    st.subheader("⏱️ Cálculo de FTE Diario")
    st.markdown("Análisis granular día por día para detectar cuellos de botella específicos.")

    with st.form("formulario_configuracion_diario"):
        st.write("Configuración de parámetros (Diario):")
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            OLE_USADO_D = st.number_input("Factor OLE (Eficiencia)", value=0.66, step=0.01, min_value=0.1, max_value=1.0, key="ole_d")
        with col_d2:
            HORA_DIARIA_D = st.number_input("Horas Diarias Contrato", value=7.9, step=0.01, key="hora_d")
        with col_d3:
            ANIO_SELECCIONADO_D = st.number_input("📅 Año a Procesar", value=2025, step=1, min_value=2023, max_value=2030, key="anio_d")
            
        st.caption("Ajusta los valores y presiona el botón.")
        submitted_d = st.form_submit_button("🔄 Calcular FTE Diario", type="primary")

    if submitted_d:
        if not (file_solicitudes and file_pesos):
            st.warning("⚠️ Faltan archivos. Por favor carga **Solicitudes** y **Pesos** en el menú lateral.")
        else:
            try:
                progress_text = "Procesando FTE Diario..."
                bar_d = st.progress(0, text=progress_text)

                bar_d.progress(10, text="Leyendo datos diarios...")
                df_s = pd.read_excel(file_solicitudes)
                df_p = pd.read_excel(file_pesos)

                col_fecha = None
                for c in df_s.columns:
                    if 'fin real' in c.lower():
                        col_fecha = c
                        break
                if not col_fecha:
                    for c in df_s.columns:
                        if 'fecha de creación' in c.lower():
                            col_fecha = c
                            break
                if not col_fecha:
                    st.error("No se encontró columna de fecha en Solicitudes.")
                    st.stop()

                df_s[col_fecha] = pd.to_datetime(df_s[col_fecha], errors='coerce')
                df_s['Fecha'] = df_s[col_fecha].dt.date
                df_s['Año'] = df_s[col_fecha].dt.year
                
                df_s = df_s.dropna(subset=['Fecha', 'Año'])
                df_s = df_s[df_s['Año'] == ANIO_SELECCIONADO_D].copy()
                
                if len(df_s) == 0:
                    st.error(f"No hay datos para el año {ANIO_SELECCIONADO_D}")
                    st.stop()

                col_res = [c for c in df_s.columns if 'Resolutor' in c or 'Técnico' in c][0]
                df_s.rename(columns={col_res: 'Resolutor'}, inplace=True)
                df_s['Resolutor'] = df_s['Resolutor'].astype(str).str.upper().str.strip()

                # --- FILTRO DE PERSONAS PERMITIDAS (COMO EN TAB 1) ---
                empleados_permitidos_d = [
                    "JESSICA ACUNA VELASQUEZ", "ALEJANDRA MATUS DURAN", "DIANA CARRASCO HERRERA",
                    "CLEMENTINA GALAZ MATTA", "BRENDA OLGUIN QUIROZ", 
                    "STEPHANIE CIFUENTES LUENGO", "KARINNA ALVAREZ MORALES"
                ]
                df_s = df_s[df_s['Resolutor'].isin(empleados_permitidos_d)].copy()
                # -----------------------------------------------------------

                df_s['Tipo Limpio'] = df_s['Tipo de Pedido'].apply(limpiar_texto)
                df_s['Tipo Limpio'] = df_s['Tipo Limpio'].replace(correcciones_manuales)
                
                df_p['TIPO DE PEDIDO'] = df_p['TIPO DE PEDIDO'].apply(limpiar_texto)
                scores_dict = df_p.set_index('TIPO DE PEDIDO')['Score'].to_dict()
                
                df_s['Score_Unitario'] = df_s['Tipo Limpio'].map(scores_dict)
                df_s['Score_Unitario'] = df_s['Score_Unitario'].fillna(0) + 2.5 

                bar_d.progress(50, text="Calculando carga diaria...")
                # Agrupación por resolutor/día
                df_diario = df_s.groupby(['Resolutor', 'Fecha'])['Score_Unitario'].sum().reset_index()

                REU_SEMANAL_TOTAL = (20 * 3) + 50 + 50 
                MINUTOS_REU_DIARIA = REU_SEMANAL_TOTAL / 5 
                MIN_CHAT_STD = 47
                MIN_CHAT_ESP = 90

                # LÓGICA CORE: Calcular tanto el FTE individual como los MINUTOS DE CARGA (Numerador)
                def calc_datos_fila(row):
                    # 1. Numerador (Carga en minutos)
                    min_chat = MIN_CHAT_ESP if row['Resolutor'] == "BRENDA OLGUIN QUIROZ" else MIN_CHAT_STD
                    carga_minutos = row['Score_Unitario'] + MINUTOS_REU_DIARIA + min_chat
                    
                    # 2. Denominador Individual (Capacidad Persona)
                    horas_p = HORA_DIARIA_D - 1 if row['Resolutor'] == "STEPHANIE CIFUENTES LUENGO" else HORA_DIARIA_D
                    capacidad_minutos_ind = horas_p * 60 * OLE_USADO_D
                    
                    fte_ind = carga_minutos / capacidad_minutos_ind if capacidad_minutos_ind > 0 else 0
                    
                    return pd.Series([carga_minutos, fte_ind])

                df_diario[['Carga_Minutos', 'FTE_Diario']] = df_diario.apply(calc_datos_fila, axis=1)
                df_diario = df_diario.sort_values(by=['Resolutor', 'Fecha'])

                st.session_state['df_diario'] = df_diario
                st.session_state['anio_diario'] = ANIO_SELECCIONADO_D
                st.session_state['calc_diario_ok'] = True
                
                bar_d.progress(100, text="✅ Terminado")
                time.sleep(1)
                bar_d.empty()
            
            except Exception as e:
                st.error(f"Error en cálculo diario: {e}")

    if st.session_state.get('calc_diario_ok'):
        df_diario = st.session_state['df_diario']
        anio_d = st.session_state['anio_diario']
        
        st.success(f"Visualizando Detalle Diario: **{anio_d}**")

        tab_d_graf, tab_d_data = st.tabs(["📈 Gráficos de Línea", "📋 Datos Diarios"])

        with tab_d_graf:
            st.markdown("#### Evolución Diaria de Carga Laboral")
            lista_personas_d = ["Todos"] + list(df_diario['Resolutor'].unique())
            seleccion_d = st.selectbox("Filtrar por Resolutor (Diario):", lista_personas_d)

            if seleccion_d == "Todos":
                # --- LÓGICA DE AGREGACIÓN "TODOS" ---
                df_total_diario = df_diario.groupby('Fecha')[['Carga_Minutos']].sum().reset_index()
                
                capacidad_estandar_minutos = HORA_DIARIA_D * 60 * OLE_USADO_D
                
                if capacidad_estandar_minutos > 0:
                    df_total_diario['FTE_Requerido'] = df_total_diario['Carga_Minutos'] / capacidad_estandar_minutos
                else:
                    df_total_diario['FTE_Requerido'] = 0

                # Gráfico 1: FTE Decimal
                fig_d = px.line(
                    df_total_diario, x='Fecha', y='FTE_Requerido',
                    title=f"FTE Total Requerido (Carga Total / Capacidad Estándar) - {anio_d}",
                    markers=True,
                    labels={'FTE_Requerido': 'FTE Necesario'}
                )
                fig_d.update_layout(yaxis_title="FTE Necesario")
                st.plotly_chart(fig_d, use_container_width=True)

                # Gráfico 2: Headcount
                st.markdown("---")
                st.markdown("#### 👥 Capacidad Real Requerida (Personas Físicas)")
                df_total_diario['Personas_Requeridas'] = np.ceil(df_total_diario['FTE_Requerido']).astype(int)

                fig_personas = px.line(
                    df_total_diario,
                    x='Fecha',
                    y='Personas_Requeridas',
                    title=f"Headcount Diario Requerido (Redondeado) - {anio_d}",
                    markers=True,
                    labels={'Personas_Requeridas': 'Cantidad de Personas'},
                    color_discrete_sequence=['#ff7f0e']
                )
                fig_personas.update_yaxes(tick0=0, dtick=1)
                st.plotly_chart(fig_personas, use_container_width=True)
                
            else:
                # Lógica Individual
                df_persona_d = df_diario[df_diario['Resolutor'] == seleccion_d].copy()
                fig_d = px.line(
                    df_persona_d, x='Fecha', y='FTE_Diario',
                    title=f"FTE Diario - {seleccion_d} ({anio_d})",
                    markers=True,
                    color_discrete_sequence=['#1f77b4'] 
                )
                fig_d.add_hline(y=1, line_dash="dash", line_color="red", annotation_text="Límite (1.0)")
                fig_d.add_hline(y=0.8, line_dash="dash", line_color="green", annotation_text="Meta (0.8)")
                fig_d.update_layout(yaxis_title="FTE Diario")
                st.plotly_chart(fig_d, use_container_width=True)

        with tab_d_data:
            st.dataframe(df_diario.style.format({"FTE_Diario": "{:.2f}", "Carga_Minutos": "{:.0f}"}), use_container_width=True)
            
            buffer_dia = io.BytesIO()
            with pd.ExcelWriter(buffer_dia) as writer:
                df_diario.to_excel(writer, sheet_name="FTE_Diario_Detalle", index=False)
            st.download_button("📥 Descargar Excel Diario", data=buffer_dia, file_name=f"FTE_Diario_{anio_d}.xlsx")

# ==============================================================================
# PESTAÑA 4: DETALLE DE TIEMPOS
# ==============================================================================
with tab3:
    st.subheader("⚙️ Días Trabajados (Detalle)")
    
    col_t3_1, col_t3_2 = st.columns([1, 3])
    with col_t3_1:
        ANIO_DETALLE = st.number_input("📅 Filtro Año", value=2025, step=1, key="anio_tab3")

    if file_prod:
        try:
            REU_SEMANAL_TOTAL = 160
            MINUTOS_REU_DIARIA = 32.0
            
            try:
                df_prod_raw = pd.read_excel(file_prod, sheet_name='HorasTotales')
            except:
                df_prod_raw = pd.read_excel(file_prod)

            col_anio_prod = None
            for c in df_prod_raw.columns:
                if str(c).strip().lower() in ['año', 'anio', 'year', 'ano']:
                    col_anio_prod = c
                    break
            
            if col_anio_prod:
                df_prod_raw = df_prod_raw[df_prod_raw[col_anio_prod] == ANIO_DETALLE].copy()
                st.info(f"Visualizando datos filtrados por año: {ANIO_DETALLE}")
            else:
                st.warning("⚠️ No se detectó columna de Año en el Excel. Se muestran todos los registros disponibles.")

            MAPA_EMPLEADOS = {
                "JACUNVE": "JESSICA ACUNA VELASQUEZ", "AMATUSD": "ALEJANDRA MATUS DURAN",
                "DCARRAH": "DIANA CARRASCO HERRERA", "CGALAZ": "CLEMENTINA GALAZ MATTA",
                "BOLGUIQ": "BRENDA OLGUIN QUIROZ", "SCIFUEN": "STEPHANIE CIFUENTES LUENGO",
                "KARINNA": "KARINNA ALVAREZ MORALES"
            }

            if "Nombre Técnico" in df_prod_raw.columns:
                df_prod_raw["Resolutor"] = df_prod_raw["Nombre Técnico"].map(MAPA_EMPLEADOS)
                df_equipo = df_prod_raw.dropna(subset=['Resolutor']).copy()
                
                tabla_dias = df_equipo.pivot_table(index='Resolutor', columns='Número Mes', values='Dias Trabajados', aggfunc='sum').fillna(0).astype(int)
                st.write("**Días Trabajados:**")
                st.dataframe(tabla_dias, use_container_width=True)
                
                # --- LÓGICA: CAPACIDAD REAL (PERSONAS DISPONIBLES) ---
                st.markdown("---")
                st.subheader("👥 Capacidad Real (Personas Disponibles)")
                st.caption("Cálculo: (Total Días Trabajados del Mes / Días Hábiles del Mes)")

                df_activos = df_equipo[df_equipo['Dias Trabajados'] > 0].copy()

                if not df_activos.empty:
                    # Agrupamos por MES
                    df_capacidad = df_activos.groupby('Número Mes').agg(
                        Dias_Totales_Trabajados=('Dias Trabajados', 'sum'),
                        Dias_Habiles_Mes=('Dias Trabajados', 'max'),
                        Personas_Activas=('Resolutor', 'nunique') # Cuenta de personas activas
                    ).reset_index()

                    df_capacidad['Personas_Reales_Disponibles'] = df_capacidad['Dias_Totales_Trabajados'] / df_capacidad['Dias_Habiles_Mes']
                    df_capacidad['Mes'] = df_capacidad['Número Mes'].astype(str)
                    
                    # Gráfico Barras + Línea
                    fig_capacidad = px.bar(
                        df_capacidad,
                        x='Mes',
                        y='Personas_Reales_Disponibles',
                        title=f"Capacidad Efectiva del Equipo (FTE Disponible) - {ANIO_DETALLE}",
                        text_auto='.2f',
                        labels={'Personas_Reales_Disponibles': 'Personas Completas (FTE)'},
                        color_discrete_sequence=['#00CC96']
                    )
                    
                    # Agregar línea de "Personas Activas" (quienes trabajaron al menos 1 día)
                    fig_capacidad.add_scatter(
                        x=df_capacidad['Mes'], 
                        y=df_capacidad['Personas_Activas'],
                        mode='lines+markers',
                        name='Personas Activas (Headcount > 0 días)',
                        line=dict(color='red', width=2, dash='dot')
                    )
                    
                    st.plotly_chart(fig_capacidad, use_container_width=True)
                    
                    # Corrección del Error de Formato
                    # Solo aplicamos formato a las columnas numéricas relevantes
                    format_dict = {
                        'Personas_Reales_Disponibles': '{:.2f}',
                        'Dias_Totales_Trabajados': '{:.0f}',
                        'Dias_Habiles_Mes': '{:.0f}',
                        'Personas_Activas': '{:.0f}'
                    }
                    st.dataframe(df_capacidad.style.format(format_dict))
                else:
                    st.info("No hay datos de días trabajados para generar el gráfico de capacidad.")
                
                st.write("**Cálculo Rápido:**")
                st.markdown(f"- Minutos Reunión Diarios: **{MINUTOS_REU_DIARIA}**")
                st.markdown(f"- Minutos Chat Estándar: **47** (Especial Brenda: 90)")
            else:
                st.error("Falta columna 'Nombre Técnico'")
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("👈 Carga 'Productividad' para ver el detalle.")