import streamlit as st
import pandas as pd
import io
import numpy as np
import plotly.express as px
import plotly.graph_objects as go 
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

# PESTAÑAS
tab1, tab2, tab_diario, tab_demanda, tab_contingencia, tab_desglose, tab3 = st.tabs([
    "✅ Validación de Pesos", 
    "⏱️ Productividad FTE MENSUAL", 
    "⏱️ Productividad FTE DIARIO", 
    "📈 Demanda FTE (Caso Estándar)", 
    "🚨 Demanda FTE (Caso Contingencia)",
    "📊 Desglose de Tiempos",
    "🗓️ Días Trabajados"
])

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

                my_bar.progress(85, text="🔄 Cruzando datos y generating KPI...")
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
                df_fte_mes['Personas Efectivas'] = np.ceil(df_fte_mes['FTE']).astype(int)
                
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
                    value_vars=['FTE', 'Personas Efectivas'],
                    var_name='Indicador', 
                    value_name='Valor'
                )
                df_grafico['Mes_Num'] = df_grafico['Mes_Num'].astype(str)
                
                fig = px.bar(
                    df_grafico, x='Mes_Num', y='Valor', color='Indicador',
                    barmode='group', title=f'FTE Mensual {anio_actual} vs Capacidad Mínima',
                    text_auto='.2f',
                    color_discrete_map={'FTE': '#1f77b4', 'Personas Efectivas': '#ff7f0e'}
                )
                
                # Línea de capacidad real
                df_capacidad_linea = df_final[df_final['Dias_Trabajados'] > 0].groupby('Mes_Num')['Resolutor'].nunique().reset_index()
                df_capacidad_linea.rename(columns={'Resolutor': 'Capacidad_Real'}, inplace=True)
                
                if not df_capacidad_linea.empty:
                    df_capacidad_linea['Mes_Num'] = df_capacidad_linea['Mes_Num'].astype(str)
                    fig.add_scatter(
                        x=df_capacidad_linea['Mes_Num'],
                        y=df_capacidad_linea['Capacidad_Real'],
                        mode='lines+markers',
                        name='Capacidad Real (Personas Activas)',
                        line=dict(color='#00CC96', width=3, dash='dot')
                    )

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

                empleados_permitidos_d = [
                    "JESSICA ACUNA VELASQUEZ", "ALEJANDRA MATUS DURAN", "DIANA CARRASCO HERRERA",
                    "CLEMENTINA GALAZ MATTA", "BRENDA OLGUIN QUIROZ", 
                    "STEPHANIE CIFUENTES LUENGO", "KARINNA ALVAREZ MORALES"
                ]
                df_s = df_s[df_s['Resolutor'].isin(empleados_permitidos_d)].copy()

                df_s['Tipo Limpio'] = df_s['Tipo de Pedido'].apply(limpiar_texto).replace(correcciones_manuales)
                df_p['TIPO DE PEDIDO'] = df_p['TIPO DE PEDIDO'].apply(limpiar_texto)
                scores_dict = df_p.set_index('TIPO DE PEDIDO')['Score'].to_dict()
                
                df_s['Score_Unitario'] = df_s['Tipo Limpio'].map(scores_dict).fillna(0) + 2.5 

                bar_d.progress(50, text="Calculando carga diaria...")
                df_diario = df_s.groupby(['Resolutor', 'Fecha'])['Score_Unitario'].sum().reset_index()

                REU_SEMANAL_TOTAL = (20 * 3) + 50 + 50 
                MINUTOS_REU_DIARIA = REU_SEMANAL_TOTAL / 5 
                MIN_CHAT_STD = 47
                MIN_CHAT_ESP = 90

                def calc_datos_fila(row):
                    min_chat = MIN_CHAT_ESP if row['Resolutor'] == "BRENDA OLGUIN QUIROZ" else MIN_CHAT_STD
                    carga_minutos = row['Score_Unitario'] + MINUTOS_REU_DIARIA + min_chat
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
                df_total_diario = df_diario.groupby('Fecha').agg(Carga_Total=('Carga_Minutos', 'sum')).reset_index()
                capacidad_estandar_minutos = HORA_DIARIA_D * 60 * OLE_USADO_D
                
                if capacidad_estandar_minutos > 0:
                    df_total_diario['FTE_Logrado'] = df_total_diario['Carga_Total'] / capacidad_estandar_minutos
                else:
                    df_total_diario['FTE_Logrado'] = 0

                df_total_diario['Personas_Necesarias'] = np.ceil(df_total_diario['FTE_Logrado'])

                fig_fte = px.line(
                    df_total_diario, x='Fecha', y='FTE_Logrado',
                    title=f" Productividad Diaria (FTE) - {anio_d}",
                    markers=True, color_discrete_sequence=['#ff7f0e'] 
                )
                fig_fte.update_layout(yaxis_title="FTE (Exacto)")
                st.plotly_chart(fig_fte, use_container_width=True)

                st.markdown("---")
                st.markdown("Si el FTE es de 4,6 lo aproximamos a 5, ya que no existen 4,6 personas. Por ende, acá podemos ver cuanta gente tuvo que trabajar ese día:")
                
                fig_comparativo = px.line(
                    df_total_diario, x='Fecha', y='Personas_Necesarias', 
                    title=f"Productividad Diaria Redondeada (FTE) - {anio_d}",
                    markers=True, color_discrete_sequence=['#d62728']
                )
                fig_comparativo.update_layout(yaxis_title="Cantidad de Personas", hovermode="x unified")
                fig_comparativo.update_yaxes(tick0=0, dtick=1)
                st.plotly_chart(fig_comparativo, use_container_width=True)

            else:
                df_persona_d = df_diario[df_diario['Resolutor'] == seleccion_d].copy()
                fig_d = px.line(
                    df_persona_d, x='Fecha', y='FTE_Diario',
                    title=f"FTE Diario - {seleccion_d} ({anio_d})",
                    markers=True, color_discrete_sequence=['#1f77b4'] 
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
# PESTAÑA 4: DEMANDA FTE (IDEAL)
# ==============================================================================
with tab_demanda:
    st.subheader("🔮 Demanda FTE (Carga Ideal según Días Hábiles)")
    st.markdown("""
    Esta simulación calcula cuántas personas se necesitan en un escenario ideal:
    - **Numerador:** Carga real (Scores) + Carga Administrativa proyectada (Reuniones/Chat x Días Hábiles x Personas).
    - **Denominador:** Capacidad ideal de una persona trabajando **todos los días hábiles** del mes.
    """)

    with st.form("form_demanda_ideal"):
        c_dem1, c_dem2, c_dem3 = st.columns(3)
        with c_dem1:
            OLE_DEMANDA = st.number_input("Factor OLE", value=0.66, step=0.01, key="ole_dem")
        with c_dem2:
            HORA_DEMANDA = st.number_input("Horas Diarias", value=7.95, step=0.01, key="hor_dem")
        with c_dem3:
            ANIO_DEMANDA = st.number_input("📅 Año a Analizar", value=2025, step=1, key="anio_dem")
            
        bt_demanda = st.form_submit_button("Calculadora de Demanda Ideal", type="primary")

    if bt_demanda:
        if not (file_solicitudes and file_pesos):
             st.warning("⚠️ Carga Solicitudes y Pesos primero.")
        else:
            try:
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
                
                df_s[col_fecha] = pd.to_datetime(df_s[col_fecha], errors='coerce')
                df_s['Mes_Num'] = df_s[col_fecha].dt.month
                df_s['Año'] = df_s[col_fecha].dt.year
                df_s = df_s[df_s['Año'] == ANIO_DEMANDA].copy()
                
                col_res = [c for c in df_s.columns if 'Resolutor' in c or 'Técnico' in c][0]
                df_s.rename(columns={col_res: 'Resolutor'}, inplace=True)
                df_s['Resolutor'] = df_s['Resolutor'].astype(str).str.upper().str.strip()
                
                empleados_permitidos_dem = [
                    "JESSICA ACUNA VELASQUEZ", "ALEJANDRA MATUS DURAN", "DIANA CARRASCO HERRERA",
                    "CLEMENTINA GALAZ MATTA", "BRENDA OLGUIN QUIROZ", 
                    "STEPHANIE CIFUENTES LUENGO", "KARINNA ALVAREZ MORALES"
                ]
                df_s = df_s[df_s['Resolutor'].isin(empleados_permitidos_dem)]

                df_s['Tipo Limpio'] = df_s['Tipo de Pedido'].apply(limpiar_texto).replace(correcciones_manuales)
                df_p['TIPO DE PEDIDO'] = df_p['TIPO DE PEDIDO'].apply(limpiar_texto)
                scores_dict = df_p.set_index('TIPO DE PEDIDO')['Score'].to_dict()
                df_s['Score_Unitario'] = df_s['Tipo Limpio'].map(scores_dict).fillna(0) + 2.5
                
                df_base = df_s.groupby(['Resolutor', 'Mes_Num'])['Score_Unitario'].sum().reset_index()
                resumen_demanda = []

                REU_SEMANAL_TOTAL = (20 * 3) + 50 + 50 
                MINUTOS_REU_DIARIA = REU_SEMANAL_TOTAL / 5 
                MIN_CHAT_STD = 47
                MIN_CHAT_ESP = 90
                
                for index, row in df_base.iterrows():
                    resolutor = row['Resolutor']
                    mes = int(row['Mes_Num'])
                    score_tickets = row['Score_Unitario']
                    
                    fecha_inicio = f"{int(ANIO_DEMANDA)}-{mes:02d}-01"
                    inicio = pd.Timestamp(fecha_inicio)
                    fin = inicio + pd.offsets.MonthEnd(0)
                    dias_habiles = np.busday_count(inicio.date(), fin.date() + pd.Timedelta(days=1))
                    
                    min_chat = MIN_CHAT_ESP if resolutor == "BRENDA OLGUIN QUIROZ" else MIN_CHAT_STD
                    carga_admin_total = (MINUTOS_REU_DIARIA + min_chat) * dias_habiles
                    numerador_total = score_tickets + carga_admin_total
                    
                    horas_p = HORA_DEMANDA - 1 if resolutor == "STEPHANIE CIFUENTES LUENGO" else HORA_DEMANDA
                    capacidad_ideal = horas_p * 60 * dias_habiles * OLE_DEMANDA
                    fte_ideal = numerador_total / capacidad_ideal if capacidad_ideal > 0 else 0
                    
                    resumen_demanda.append({
                        'Mes_Num': mes,
                        'Resolutor': resolutor,
                        'Dias_Habiles': dias_habiles,
                        'Score_Tickets': score_tickets,
                        'Carga_Admin': carga_admin_total,
                        'Carga_Total': numerador_total,
                        'Capacidad_Individual': capacidad_ideal,
                        'FTE_Ideal': fte_ideal
                    })
                
                df_demanda = pd.DataFrame(resumen_demanda)

                if not df_demanda.empty:
                    df_demanda_mes = df_demanda.groupby('Mes_Num').agg(
                        FTE_Ideal_Total=('FTE_Ideal', 'sum'),
                        Dias_Habiles_Prom=('Dias_Habiles', 'max') 
                    ).reset_index()
                    
                    FACTOR_SHRINKAGE = 0.80
                    df_demanda_mes['Headcount_Requerido'] = df_demanda_mes['FTE_Ideal_Total'] / FACTOR_SHRINKAGE
                    df_demanda_mes['Personas_A_Contratar'] = np.ceil(df_demanda_mes['Headcount_Requerido'])
                    
                    st.markdown("#### Resultado: Plantilla Necesaria ")
                    fig_dem = px.bar(
                        df_demanda_mes, x='Mes_Num', y='Headcount_Requerido',
                        title=f"Dimensionamiento FTE {int(ANIO_DEMANDA)} (Plantilla Necesaria)",
                        text_auto='.2f', labels={'Headcount_Requerido': 'Plantilla Exacta (Decimales)'},
                        color_discrete_sequence=['#ff7f0e'] 
                    )
                    
                    fig_dem.add_trace(go.Bar(
                        x=df_demanda_mes['Mes_Num'], y=df_demanda_mes['Personas_A_Contratar'],
                        name='Contratación Sugerida (Redondeo)', text=df_demanda_mes['Personas_A_Contratar'],
                        textposition='auto', marker_color='rgba(255, 0, 0, 0.3)',
                        marker_line_width=2, marker_line_color='red'
                    ))

                    fig_dem.update_layout(barmode='overlay')
                    st.plotly_chart(fig_dem, use_container_width=True)
                    
                    st.write("### Detalle de Cálculo")
                    st.info("Nota: Se muestra únicamente la **Plantilla Necesaria** (que incluye el Shrinkage de 17%).")
                    st.dataframe(df_demanda_mes)
                    
                    with st.expander("Ver desglose por Resolutor (Ideal)"):
                        st.dataframe(df_demanda)

                else:
                    st.warning("No se generaron datos. Revisa el año seleccionado.")

            except Exception as e:
                st.error(f"Error en Demanda FTE: {e}")

# ==============================================================================
# PESTAÑA 5: DEMANDA FTE (CONTINGENCIA)
# ==============================================================================
with tab_contingencia:
    st.subheader("🚨 Demanda FTE (Caso Contingencia)")
    st.markdown("""
    **Escenario de Emergencia/Contingencia:**
    - Se eliminan tiempos de reuniones y chats.
    - Se asume que el personal está **100% dedicado a Procesos (Full Proceso)**.
    """)

    with st.form("form_demanda_contingencia"):
        c_cont1, c_cont2, c_cont3 = st.columns(3)
        with c_cont1:
            OLE_CONT = st.number_input("Factor OLE", value=0.66, step=0.01, key="ole_cont")
        with c_cont2:
            HORA_CONT = st.number_input("Horas Diarias", value=7.95, step=0.01, key="hor_cont")
        with c_cont3:
            ANIO_CONT = st.number_input("📅 Año a Analizar", value=2025, step=1, key="anio_cont")
            
        bt_contingencia = st.form_submit_button("Calculadora de Contingencia", type="primary")

    if bt_contingencia:
        if not (file_solicitudes and file_pesos):
             st.warning("⚠️ Carga Solicitudes y Pesos primero.")
        else:
            try:
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
                
                df_s[col_fecha] = pd.to_datetime(df_s[col_fecha], errors='coerce')
                df_s['Mes_Num'] = df_s[col_fecha].dt.month
                df_s['Año'] = df_s[col_fecha].dt.year
                df_s = df_s[df_s['Año'] == ANIO_CONT].copy()
                
                col_res = [c for c in df_s.columns if 'Resolutor' in c or 'Técnico' in c][0]
                df_s.rename(columns={col_res: 'Resolutor'}, inplace=True)
                df_s['Resolutor'] = df_s['Resolutor'].astype(str).str.upper().str.strip()
                
                empleados_permitidos_cont = [
                    "JESSICA ACUNA VELASQUEZ", "ALEJANDRA MATUS DURAN", "DIANA CARRASCO HERRERA",
                    "CLEMENTINA GALAZ MATTA", "BRENDA OLGUIN QUIROZ", 
                    "STEPHANIE CIFUENTES LUENGO", "KARINNA ALVAREZ MORALES"
                ]
                df_s = df_s[df_s['Resolutor'].isin(empleados_permitidos_cont)]

                df_s['Tipo Limpio'] = df_s['Tipo de Pedido'].apply(limpiar_texto).replace(correcciones_manuales)
                df_p['TIPO DE PEDIDO'] = df_p['TIPO DE PEDIDO'].apply(limpiar_texto)
                scores_dict = df_p.set_index('TIPO DE PEDIDO')['Score'].to_dict()
                df_s['Score_Unitario'] = df_s['Tipo Limpio'].map(scores_dict).fillna(0) + 2.5
                
                df_base = df_s.groupby(['Resolutor', 'Mes_Num'])['Score_Unitario'].sum().reset_index()
                resumen_contingencia = []

                for index, row in df_base.iterrows():
                    resolutor = row['Resolutor']
                    mes = int(row['Mes_Num'])
                    score_tickets = row['Score_Unitario']
                    
                    fecha_inicio = f"{int(ANIO_CONT)}-{mes:02d}-01"
                    inicio = pd.Timestamp(fecha_inicio)
                    fin = inicio + pd.offsets.MonthEnd(0)
                    dias_habiles = np.busday_count(inicio.date(), fin.date() + pd.Timedelta(days=1))
                    
                    carga_admin_total = 0 
                    numerador_total = score_tickets + carga_admin_total
                    
                    horas_p = HORA_CONT - 1 if resolutor == "STEPHANIE CIFUENTES LUENGO" else HORA_CONT
                    capacidad_ideal = horas_p * 60 * dias_habiles * OLE_CONT
                    fte_ideal = numerador_total / capacidad_ideal if capacidad_ideal > 0 else 0
                    
                    resumen_contingencia.append({
                        'Mes_Num': mes,
                        'Resolutor': resolutor,
                        'Dias_Habiles': dias_habiles,
                        'Score_Tickets': score_tickets,
                        'Carga_Admin': carga_admin_total, 
                        'Carga_Total': numerador_total,
                        'Capacidad_Individual': capacidad_ideal,
                        'FTE_Ideal': fte_ideal
                    })
                
                df_contingencia = pd.DataFrame(resumen_contingencia)

                if not df_contingencia.empty:
                    df_cont_mes = df_contingencia.groupby('Mes_Num').agg(
                        FTE_Ideal_Total=('FTE_Ideal', 'sum'),
                        Dias_Habiles_Prom=('Dias_Habiles', 'max') 
                    ).reset_index()
                    
                    FACTOR_SHRINKAGE = 0.85 
                    df_cont_mes['Headcount_Requerido'] = df_cont_mes['FTE_Ideal_Total'] / FACTOR_SHRINKAGE
                    df_cont_mes['Personas_A_Contratar'] = np.ceil(df_cont_mes['Headcount_Requerido'])
                    
                    st.markdown("#### Resultado Contingencia: Plantilla Necesaria (Sin Admin Load)")
                    fig_cont = px.bar(
                        df_cont_mes, x='Mes_Num', y='Headcount_Requerido',
                        title=f"Dimensionamiento CONTINGENCIA {int(ANIO_CONT)} (Solo Procesos)",
                        text_auto='.2f', labels={'Headcount_Requerido': 'Plantilla Exacta (Decimales)'},
                        color_discrete_sequence=['#ff7f0e']
                    )
                    
                    fig_cont.add_trace(go.Bar(
                        x=df_cont_mes['Mes_Num'], y=df_cont_mes['Personas_A_Contratar'],
                        name='Contratación Sugerida (Redondeo)', text=df_cont_mes['Personas_A_Contratar'],
                        textposition='auto', marker_color='rgba(214, 39, 40, 0.3)', 
                        marker_line_width=2, marker_line_color='red'
                    ))

                    fig_cont.update_layout(barmode='overlay')
                    st.plotly_chart(fig_cont, use_container_width=True)
                    
                    st.write("### Detalle de Cálculo (Contingencia)")
                    st.dataframe(df_cont_mes)
                    
                    with st.expander("Ver desglose por Resolutor (Contingencia)"):
                        st.dataframe(df_contingencia)

                else:
                    st.warning("No se generaron datos. Revisa el año seleccionado.")

            except Exception as e:
                st.error(f"Error en Contingencia FTE: {e}")


# ==============================================================================
# PESTAÑA 6: DESGLOSE DE TIEMPOS OLE
# ==============================================================================
with tab_desglose:
    st.subheader("📊 Desglose de Tiempos (Operación, Reuniones y OLE)")
    st.markdown("""
    Esta pestaña desglosa la distribución del tiempo total utilizado por el equipo, 
    agrupando la operación y separando el OLE en sus correspondientes categorías.
    """)

    # ----------------------------------------------------------------------
    # A. PANEL DE CONFIGURACIÓN DINÁMICA DEL OLE
    # ----------------------------------------------------------------------
    with st.expander("⚙️ Configurar Minutos Diarios del OLE", expanded=False):
        with st.form("form_config_ole"):
            st.markdown("Ajusta los tiempos de pérdida diarios para ver el impacto visual y calcular tu OLE Teórico:")
            c_ole1, c_ole2, c_ole3, c_ole4 = st.columns(4)
            with c_ole1:
                val_horas = st.number_input("Horas Diarias", value=7.95, step=0.01)
                val_almuerzo = st.number_input("Cat A: Alimentación", value=40.0, step=1.0)
            with c_ole2:
                val_fisiologicas = st.number_input("Cat A: Necesidades Fisiológicas (Baño)", value=24.0, step=1.0)
                val_fatiga = st.number_input("Cat A: Fatiga Básica y Variable", value=28.0, step=1.0)
            with c_ole3:
                val_fallas = st.number_input("Cat B: Fallas de Sist.", value=10.0, step=1.0)
            with c_ole4:
                val_reu_no_est = st.number_input("Cat C: Reu. No Est.", value=30.4, step=0.1)
                val_micro = st.number_input("Cat C: MicroTareas", value=30.0, step=1.0)
            
            btn_recalcular = st.form_submit_button("🔄 Calcular y Actualizar Gráficos")
            
        # Cálculo del OLE directamente fuera del if para que siempre se muestre
        minutos_totales = val_horas * 60
        minutos_perdidos = val_fisiologicas + val_fatiga + val_almuerzo + val_fallas + val_reu_no_est + val_micro
        ole_calc = (minutos_totales - minutos_perdidos) / minutos_totales if minutos_totales > 0 else 0
        
        st.info(f"📊 **OLE Teórico Calculado:** **{ole_calc:.1%}** (Se restan {minutos_perdidos:.1f} min de un total de {minutos_totales:.0f} min diarios)")
    
    if st.session_state.get('calculo_realizado'):
        df_desglose = st.session_state['df_fte_final'].copy()
        anio_desglose = st.session_state.get('anio_calculado', '---')
        
        # 1. Filtro por Mes
        meses_disponibles = ["Todos"] + sorted(list(df_desglose['Mes_Num'].unique()))
        mes_seleccionado = st.selectbox("📅 Filtrar por Mes:", meses_disponibles)
        
        if mes_seleccionado != "Todos":
            df_desglose = df_desglose[df_desglose['Mes_Num'] == mes_seleccionado].copy()
            st.success(f"Visualizando desglose del año: **{anio_desglose}** - Mes: **{mes_seleccionado}**")
        else:
            st.success(f"Visualizando desglose de todo el año: **{anio_desglose}**")

        # 2. Agrupaciones solicitadas: Operación sumando con tickets y sin tickets
        df_desglose['Operación (Tickets + Sin Tickets)'] = df_desglose['Score_Unitario'] + df_desglose['Minutos_Chat']
        df_desglose['Reuniones Estandarizadas (Fijas)'] = df_desglose['Minutos_Reunion']
        
        # 3. Desglose del OLE dinámico basado en las variables del expander
        df_desglose['Cat. A: Necesidades Fisiológicas y Fatiga'] = df_desglose['Dias_Trabajados'] * (val_fisiologicas + val_fatiga)
        df_desglose['Cat. A: Alimentación'] = df_desglose['Dias_Trabajados'] * val_almuerzo
        
        df_desglose['Cat. B: Fallas de Sistema'] = df_desglose['Dias_Trabajados'] * val_fallas
        df_desglose['Cat. C: Reuniones No Estandarizadas'] = df_desglose['Dias_Trabajados'] * val_reu_no_est
        df_desglose['Cat. C: MicroTareas (Gestión/Cursos/Soporte/Setup)'] = df_desglose['Dias_Trabajados'] * val_micro

        # --- LÓGICA DE CAPACIDAD LIBRE ---
        # Calculamos el total de minutos teóricos por persona según sus días trabajados y las horas configuradas
        def calc_capacidad_libre(row):
            horas_p = val_horas - 1 if row['Resolutor'] == "STEPHANIE CIFUENTES LUENGO" else val_horas
            minutos_teoricos_totales = horas_p * 60 * row['Dias_Trabajados']
            
            # Sumamos todo lo que hemos "consumido"
            minutos_consumidos = (
                row['Operación (Tickets + Sin Tickets)'] +
                row['Reuniones Estandarizadas (Fijas)'] +
                row['Cat. A: Necesidades Fisiológicas y Fatiga'] +
                row['Cat. A: Alimentación'] +
                row['Cat. B: Fallas de Sistema'] +
                row['Cat. C: Reuniones No Estandarizadas'] +
                row['Cat. C: MicroTareas (Gestión/Cursos/Soporte/Setup)']
            )
            
            libre = minutos_teoricos_totales - minutos_consumidos
            return libre if libre > 0 else 0 # Si da negativo, significa que trabajaron más de su horario (horas extras)
        
        df_desglose['Capacidad Libre (Ocio / Proyectos)'] = df_desglose.apply(calc_capacidad_libre, axis=1)
        # ----------------------------------

        cols_base = [
            'Operación (Tickets + Sin Tickets)', 
            'Reuniones Estandarizadas (Fijas)',
            'Cat. C: Reuniones No Estandarizadas',
            'Cat. B: Fallas de Sistema',
            'Cat. C: MicroTareas (Gestión/Cursos/Soporte/Setup)',
            'Capacidad Libre (Ocio / Proyectos)'
        ]
        
        cols_pausas = [
            'Cat. A: Necesidades Fisiológicas y Fatiga',
            'Cat. A: Alimentación'
        ]

        # 4. Toggle para controlar la visualización
        st.markdown("---")
        mostrar_pausas = st.toggle("👁️ Mostrar tiempos de Alimentación y Necesidades Fisiológicas en el análisis", value=False)
        st.caption("Por defecto estos tiempos se ocultan para analizar estrictamente la carga y fricción operativa. Actívalo para ver la jornada en su totalidad.")
        
        if mostrar_pausas:
            cols_analisis = cols_base + cols_pausas
        else:
            cols_analisis = cols_base

        # ----------------------------------------------------------------------
        # A. Análisis General (Total Equipo)
        # ----------------------------------------------------------------------
        df_total = df_desglose[cols_analisis].sum().reset_index()
        df_total.columns = ['Categoría', 'Minutos Totales']
        df_total['Horas Totales'] = df_total['Minutos Totales'] / 60
        df_total['%'] = (df_total['Minutos Totales'] / df_total['Minutos Totales'].sum()) * 100

        col_g1, col_g2 = st.columns([1.5, 1])
        with col_g1:
            st.markdown("#### ¿En qué se va el tiempo del equipo?")
            
            # Asignamos un color específico para 'Capacidad Libre'
            color_map = {
                'Operación (Tickets + Sin Tickets)': '#636EFA',
                'Reuniones Estandarizadas (Fijas)': '#EF553B',
                'Cat. C: Reuniones No Estandarizadas': '#00CC96',
                'Cat. B: Fallas de Sistema': '#AB63FA',
                'Cat. C: MicroTareas (Gestión/Cursos/Soporte/Setup)': '#FFA15A',
                'Cat. A: Necesidades Fisiológicas y Fatiga': '#19D3F3',
                'Cat. A: Alimentación': '#FF6692',
                'Capacidad Libre (Ocio / Proyectos)': "#D3D3D3" # Color distintivo claro
            }
            
            fig_pie = px.pie(
                df_total, 
                names='Categoría', 
                values='Minutos Totales',
                hole=0.3,
                color='Categoría',
                color_discrete_map=color_map
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_g2:
            st.markdown("#### Resumen en Horas")
            st.dataframe(df_total[['Categoría', 'Horas Totales', '%']].style.format({
                'Horas Totales': "{:,.0f} hrs",
                '%': "{:.1f}%"
            }), use_container_width=True)
            
            # Tarjeta de resumen de Capacidad Libre
            if 'Capacidad Libre (Ocio / Proyectos)' in df_total['Categoría'].values:
                libre_pct = df_total.loc[df_total['Categoría'] == 'Capacidad Libre (Ocio / Proyectos)', '%'].values[0]
                libre_hrs = df_total.loc[df_total['Categoría'] == 'Capacidad Libre (Ocio / Proyectos)', 'Horas Totales'].values[0]
                st.info(f"**💡 Capacidad Libre (Ocio / Proyectos):** El equipo tiene un **{libre_pct:.1f}%** ({libre_hrs:,.0f} hrs) de su tiempo disponible o no registrado operativamente.")

        st.markdown("---")

        # ----------------------------------------------------------------------
        # B. Análisis Detallado por Persona (Resolutor)
        # ----------------------------------------------------------------------
        st.markdown("#### Desglose de Tiempos por Resolutor (En Horas)")
        df_resolutores = df_desglose.groupby('Resolutor')[cols_analisis].sum().reset_index()
        
        # Melt para gráfica apilada
        df_melt = df_resolutores.melt(id_vars='Resolutor', value_vars=cols_analisis, var_name='Categoría', value_name='Minutos')
        df_melt['Horas'] = df_melt['Minutos'] / 60

        fig_bar = px.bar(
            df_melt, x='Resolutor', y='Horas', color='Categoría',
            text_auto='.0f', title="Distribución de Horas por Empleado",
            color_discrete_map=color_map
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # ----------------------------------------------------------------------
        # C. Tabla y Descarga
        # ----------------------------------------------------------------------
        st.markdown("#### Datos Crudos (Base Mensual)")
        # Mostramos siempre todas las columnas en la tabla y descarga para no perder data
        columnas_mostrar = ['Mes_Num', 'Resolutor', 'Dias_Trabajados'] + cols_base + cols_pausas
        st.dataframe(df_desglose[columnas_mostrar].style.format(precision=1), use_container_width=True)

        buffer_ole = io.BytesIO()
        with pd.ExcelWriter(buffer_ole) as writer:
            df_desglose[columnas_mostrar].to_excel(writer, index=False, sheet_name="Desglose_Tiempos")
        st.download_button("📥 Descargar Base Completa del Desglose OLE", data=buffer_ole, file_name=f"Desglose_Tiempos_OLE_{anio_desglose}.xlsx")

    else:
        st.warning("⚠️ Debes correr primero el cálculo en la pestaña 'Productividad FTE MENSUAL' (Pestaña 2) para generar los datos de tiempo.")


# ==============================================================================
# PESTAÑA 7: DETALLE DE TIEMPOS
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
                
                st.markdown("---")
                st.subheader("👥 Capacidad Real (Personas Disponibles)")
                st.caption("Cálculo: (Total Días Trabajados del Mes / Días Hábiles del Mes)")

                df_activos = df_equipo[df_equipo['Dias Trabajados'] > 0].copy()

                if not df_activos.empty:
                    df_capacidad = df_activos.groupby('Número Mes').agg(
                        Dias_Totales_Trabajados=('Dias Trabajados', 'sum'),
                        Dias_Habiles_Mes=('Dias Trabajados', 'max'),
                        Personas_Activas=('Resolutor', 'nunique') 
                    ).reset_index()

                    df_capacidad['Personas_Reales_Disponibles'] = df_capacidad['Dias_Totales_Trabajados'] / df_capacidad['Dias_Habiles_Mes']
                    df_capacidad['Mes'] = df_capacidad['Número Mes'].astype(str)
                    
                    fig_capacidad = px.bar(
                        df_capacidad, x='Mes', y='Personas_Reales_Disponibles',
                        title=f"Capacidad Efectiva del Equipo (FTE Disponible) - {ANIO_DETALLE}",
                        text_auto='.2f', labels={'Personas_Reales_Disponibles': 'Personas Completas (FTE)'},
                        color_discrete_sequence=['#00CC96']
                    )
                    
                    fig_capacidad.add_scatter(
                        x=df_capacidad['Mes'], y=df_capacidad['Personas_Activas'],
                        mode='lines+markers', name='Personas Activas (Headcount > 0 días)',
                        line=dict(color='red', width=2, dash='dot')
                    )
                    
                    st.plotly_chart(fig_capacidad, use_container_width=True)
                    
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
                st.markdown(f"- Minutos Chats y Correos Estándar: **47** (Especial Brenda: 90)")
            else:
                st.error("Falta columna 'Nombre Técnico'")
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("👈 Carga 'Productividad' para ver el detalle.")