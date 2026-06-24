import streamlit as st
import pandas as pd
import io
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Consolidación SAP - Univalle", layout="wide")

# Título corporativo oficial
st.title("🏛️ Sistema de Consolidación Contable SAP")
st.markdown("Plataforma centralizada para la validación, unificación y generación de asientos contables financieros. Utilice el panel lateral para iniciar la ingesta de datos.")

# --- 1. INICIALIZAR MEMORIA Y ESTADOS ---
if 'plantilla_maestra' not in st.session_state:
    columnas_sap = ['DIA_ETIQUETA', 'BUKRS', 'HKONT', 'SGTXT', 'WRSOL', 'WRHAB', 'DMBTR', 'DMBE2', 'MWSKZ', 'TXJCD', 'KOSTL', 'PRCTR', 'AUFNR', 'PS_POSID', 'VALUT', 'HBKID', 'HKTID', 'ZUONR', 'VBUND', 'FIPEX']
    st.session_state.plantilla_maestra = pd.DataFrame(columns=columnas_sap)

if 'marcar_todo' not in st.session_state:
    st.session_state.marcar_todo = True

# DataFrames de control de estado interno
for key in ['df_cajas', 'df_atc', 'df_com']:
    if key not in st.session_state:
        st.session_state[key] = pd.DataFrame()

# Estabilidad de nombres de archivos cargados (ahora guarda listas de nombres)
for key in ['name_cajas', 'name_atc', 'name_com']:
    if key not in st.session_state:
        st.session_state[key] = []

# --- FUNCIÓN DE INGESTA MULTIPLE Y EXTRACCIÓN DE METADATOS ---
def procesar_subida_multiple(lista_archivos, state_name, state_df, tipo_archivo):
    # Obtenemos los nombres actuales de los archivos subidos
    nombres_actuales = [f.name for f in lista_archivos] if lista_archivos else []
    
    # Si la lista de archivos cambió, procesamos de nuevo
    if nombres_actuales != st.session_state.get(state_name, []):
        if lista_archivos:
            dfs_temporales = []
            for file_obj in lista_archivos:
                df = pd.read_excel(file_obj)
                df.columns = df.columns.str.strip().str.upper()
                
                # Guardamos el nombre del archivo limpio para usarlo como filtro dinámico
                nombre_limpio = file_obj.name.split('.')[0].upper().replace('_', ' ')
                df['FUENTE_ARCHIVO'] = nombre_limpio
                
                # Extracción automática de código de caja desde el nombre del archivo para ATC
                if tipo_archivo == 'ATC':
                    tiene_caja = any('CAJA' in col for col in df.columns)
                    if not tiene_caja:
                        match = re.search(r'(SFC\d+|SOUVENIRS?)', file_obj.name.upper())
                        df['NÚMERO DE CAJA'] = match.group(1) if match else 'GENERAL'
                
                dfs_temporales.append(df)
            
            # Unimos todos los archivos de esta categoría en una sola tabla gigante
            df_consolidado = pd.concat(dfs_temporales, ignore_index=True)
            df_consolidado['PROCESADO'] = False 
            df_consolidado['ORIGINAL_INDEX'] = df_consolidado.index 
            
            st.session_state[state_df] = df_consolidado
        else:
            # Si quitaron todos los archivos, vaciamos el dataframe
            st.session_state[state_df] = pd.DataFrame()
            
        st.session_state[state_name] = nombres_actuales

# --- PANEL LATERAL (SIDEBAR) CONTROLES Y EXPORTACIÓN ---
with st.sidebar:
    st.header("⚙️ Módulo de Ingesta")
    st.write("*(Puede seleccionar varios archivos a la vez)*")
    # ACTIVAMOS CARGA MÚLTIPLE
    file_cajas = st.file_uploader("📂 Cargar extracto CAJAS", type=["xlsx", "xls"], accept_multiple_files=True)
    file_atc = st.file_uploader("📂 Cargar extracto ATC", type=["xlsx", "xls"], accept_multiple_files=True)
    file_com = st.file_uploader("📂 Cargar extracto COMUNICACIONES", type=["xlsx", "xls"], accept_multiple_files=True)

    # Procesar subidas identificando el tipo de flujo
    procesar_subida_multiple(file_cajas, 'name_cajas', 'df_cajas', 'CAJAS')
    procesar_subida_multiple(file_atc, 'name_atc', 'df_atc', 'ATC')
    procesar_subida_multiple(file_com, 'name_com', 'df_com', 'COMUN')

    st.markdown("---")
    st.header("📦 Archivos Listos (Layout SAP)")
    
    if len(st.session_state.plantilla_maestra) > 0:
        dias_procesados = st.session_state.plantilla_maestra['DIA_ETIQUETA'].unique()
        st.success(f"📊 {len(dias_procesados)} periodos guardados en memoria:")

        for dia in sorted(dias_procesados):
            df_dia = st.session_state.plantilla_maestra[st.session_state.plantilla_maestra['DIA_ETIQUETA'] == dia].copy()
            df_export = df_dia.drop(columns=['DIA_ETIQUETA'])
            
            # Formatos de salida estándar para SAP
            df_export['VALUT'] = pd.to_datetime(df_export['VALUT']).dt.strftime('%d/%m/%Y')
            def string_format_sap(val):
                try: return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except: return val
            df_export['WRSOL'] = df_export['WRSOL'].apply(string_format_sap)
            
            csv_data = df_export.to_csv(index=False, sep='|', header=True)
            st.download_button(
                label=f"⬇️ Descargar Layout {dia}", 
                data=csv_data, 
                file_name=f"SAP_{dia.replace(' ', '_')}.csv", 
                mime="text/csv", 
                use_container_width=True
            )
    else:
        st.info("Los layouts de descarga se generarán aquí conforme valide las transacciones en la mesa de trabajo.")

    # --- BOTÓN PARA LIMPIAR Y BORRAR TODO ---
    st.markdown("---")
    if st.button("🗑️ Limpiar y Borrar Todo", type="secondary", use_container_width=True):
        for key in list(st.session_state.keys()): 
            del st.session_state[key]
        st.rerun()

# --- ÁREA PRINCIPAL (PANTALLA COMPLETA) ---
# AHORA USAMOS 'OR' PARA PERMITIR TRABAJAR AUNQUE FALTE ALGÚN ARCHIVO (EJ. Días sin Comunicaciones)
if not st.session_state.df_cajas.empty or not st.session_state.df_atc.empty or not st.session_state.df_com.empty:
    
    # Controles superiores de asignación de periodo
    col_dia, col_btn1, col_btn2 = st.columns([3, 1, 1])
    with col_dia: dia_actual = st.selectbox("📅 Asignar las transacciones marcadas al periodo:", [f"Día {i}" for i in range(1, 32)])
    with col_btn1: 
        if st.button("✅ Marcar Todo", use_container_width=True): 
            st.session_state.marcar_todo = True
            st.rerun()
    with col_btn2: 
        if st.button("⬜ Desmarcar Todo", use_container_width=True): 
            st.session_state.marcar_todo = False
            st.rerun()

    # Filtrado dinámico inicial (solo registros pendientes)
    df_pendientes_cajas = st.session_state.df_cajas[st.session_state.df_cajas['PROCESADO'] == False].copy() if not st.session_state.df_cajas.empty else pd.DataFrame()
    df_pendientes_atc = st.session_state.df_atc[st.session_state.df_atc['PROCESADO'] == False].copy() if not st.session_state.df_atc.empty else pd.DataFrame()
    df_pendientes_com = st.session_state.df_com[st.session_state.df_com['PROCESADO'] == False].copy() if not st.session_state.df_com.empty else pd.DataFrame()

    st.markdown("---")
    st.subheader("🔍 Filtros Avanzados de Operación")
    
    # --- FILTROS DESPLEGABLES DINÁMICOS ---
    col_filtro_caja, col_filtro_com = st.columns(2)
    
    # 1. Identificar la columna de Caja en Cajas Contables
    col_caja_field = None
    if not df_pendientes_cajas.empty:
        for col in df_pendientes_cajas.columns:
            if 'CAJA' in col:
                col_caja_field = col
                break

    caja_filtrada = "Mostrar Todas"
    if col_caja_field or (not df_pendientes_atc.empty and 'NÚMERO DE CAJA' in df_pendientes_atc.columns):
        # Combinar códigos de caja disponibles en Cajas y ATC
        cajas_unicas = set()
        if col_caja_field and not df_pendientes_cajas.empty:
            cajas_unicas.update(df_pendientes_cajas[col_caja_field].dropna().astype(str))
        if not df_pendientes_atc.empty and 'NÚMERO DE CAJA' in df_pendientes_atc.columns:
            cajas_unicas.update(df_pendientes_atc['NÚMERO DE CAJA'].dropna().astype(str))
            
        if cajas_unicas:
            opciones_caja = ["Mostrar Todas"] + sorted(list(cajas_unicas))
            with col_filtro_caja:
                caja_filtrada = st.selectbox("🛒 Filtrar Cajas y ATC por Código de Cajero:", opciones_caja, key="filtro_caja_master")

    # 2. Identificar Archivos de Comunicaciones cargados
    com_filtrada = "Mostrar Todas"
    if not df_pendientes_com.empty and 'FUENTE_ARCHIVO' in df_pendientes_com.columns:
        opciones_com = ["Mostrar Todas"] + sorted(list(df_pendientes_com['FUENTE_ARCHIVO'].unique()))
        with col_filtro_com:
            com_filtrada = st.selectbox("📑 Filtrar Comunicaciones Internas por Documento:", opciones_com, key="filtro_com_master")

    # --- APLICACIÓN EN TIEMPO REAL DE LOS FILTROS ---
    if caja_filtrada != "Mostrar Todas":
        if not df_pendientes_cajas.empty and col_caja_field:
            df_pendientes_cajas = df_pendientes_cajas[df_pendientes_cajas[col_caja_field].astype(str) == caja_filtrada]
        if not df_pendientes_atc.empty and 'NÚMERO DE CAJA' in df_pendientes_atc.columns:
            df_pendientes_atc = df_pendientes_atc[df_pendientes_atc['NÚMERO DE CAJA'].astype(str) == caja_filtrada]

    if com_filtrada != "Mostrar Todas" and not df_pendientes_com.empty:
        df_pendientes_com = df_pendientes_com[df_pendientes_com['FUENTE_ARCHIVO'] == com_filtrada]

    # Inyectar columnas de selección visual si hay datos
    if not df_pendientes_cajas.empty: df_pendientes_cajas.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
    if not df_pendientes_atc.empty: df_pendientes_atc.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
    if not df_pendientes_com.empty: df_pendientes_com.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)

    key_suffix = f"{dia_actual}_{st.session_state.marcar_todo}_{caja_filtrada}_{com_filtrada}"

    st.markdown("---")
    
    # --- MESA DE TRABAJO EXTENDIDA HORIZONTALMENTE ---
    if not df_pendientes_cajas.empty:
        st.subheader(f"🛒 Consolidado Cajas Diarias — ({len(df_pendientes_cajas)} registros filtrados)")
        edit_cajas = st.data_editor(df_pendientes_cajas, hide_index=True, use_container_width=True, height=350, key=f"ed_c_{key_suffix}")
    else:
        edit_cajas = pd.DataFrame()
        
    if not df_pendientes_atc.empty:
        st.subheader(f"💳 Transacciones ATC Unificado — ({len(df_pendientes_atc)} registros filtrados)")
        edit_atc = st.data_editor(df_pendientes_atc, hide_index=True, use_container_width=True, height=350, key=f"ed_a_{key_suffix}")
    else:
        edit_atc = pd.DataFrame()

    if not df_pendientes_com.empty:
        st.subheader(f"📑 Flujo Comunicaciones Internas — ({len(df_pendientes_com)} registros filtrados)")
        edit_com = st.data_editor(df_pendientes_com, hide_index=True, use_container_width=True, height=350, key=f"ed_co_{key_suffix}")
    else:
        edit_com = pd.DataFrame()

    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # Botón maestro de procesamiento
    if st.button(f"🚀 VERIFICAR Y ANEXAR TRANSACCIONES AL {dia_actual.upper()}", type="primary", use_container_width=True):
        sel_cajas = edit_cajas[edit_cajas['SELECCIONAR'] == True].copy() if not edit_cajas.empty else pd.DataFrame()
        sel_atc = edit_atc[edit_atc['SELECCIONAR'] == True].copy() if not edit_atc.empty else pd.DataFrame()
        sel_com = edit_com[edit_com['SELECCIONAR'] == True].copy() if not edit_com.empty else pd.DataFrame()

        if sel_cajas.empty and sel_atc.empty and sel_com.empty:
            st.warning("⚠️ No ha seleccionado ninguna transacción válida para procesar.")
        else:
            # Marcar filas procesadas
            if not sel_cajas.empty: st.session_state.df_cajas.loc[sel_cajas['ORIGINAL_INDEX'], 'PROCESADO'] = True
            if not sel_atc.empty: st.session_state.df_atc.loc[sel_atc['ORIGINAL_INDEX'], 'PROCESADO'] = True
            if not sel_com.empty: st.session_state.df_com.loc[sel_com['ORIGINAL_INDEX'], 'PROCESADO'] = True

            def limpiar_asig(df, col, respaldo):
                if df.empty: return df
                return df.get(col, df.get(respaldo, '')).fillna('').astype(str).replace('nan', '', regex=True)

            if not sel_cajas.empty: sel_cajas['ZUONR_FINAL'] = limpiar_asig(sel_cajas, 'ASIGNACION', 'CÓDIGO ASIENTO')
            if not sel_atc.empty: sel_atc['ZUONR_FINAL'] = limpiar_asig(sel_atc, 'ASIGNACION', 'CÓDIGO ASIENTO')
            if not sel_com.empty: sel_com['ZUONR_FINAL'] = limpiar_asig(sel_com, 'ASIGNACION', '')

            def parse_monto(df, col):
                return df[col].astype(str).str.replace(',', '.').astype(float)

            # Mapeos
            sap_cajas = pd.DataFrame() if sel_cajas.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_cajas.get('CUENTA CONTABLE', ''), 'SGTXT': sel_cajas.get('GLOSA RECORTADA', ''),
                'WRSOL': parse_monto(sel_cajas, 'CRÉDITOS'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(sel_cajas['FECHA']), 'ZUONR': sel_cajas['ZUONR_FINAL']
            })

            sap_atc = pd.DataFrame() if sel_atc.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_atc.get('CUENTA CONTABLE', ''), 'SGTXT': sel_atc.get('DETALLE', ''),
                'WRSOL': parse_monto(sel_atc, 'MONTO'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(sel_atc['FECHA']), 'ZUONR': sel_atc['ZUONR_FINAL']
            })

            sap_com = pd.DataFrame() if sel_com.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_com.get('CUENTA CONTABLE BANCO', ''), 'SGTXT': sel_com.get('GLOSA ASIENTO COMUNICACIONES INTERNAS', ''),
                'WRSOL': parse_monto(sel_com, 'TOTAL C.I.'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(sel_com['FECHA']), 'ZUONR': sel_com['ZUONR_FINAL']
            })

            # Consolidar
            bloque_dia = pd.concat([sap_cajas, sap_atc, sap_com])
            if not bloque_dia.empty:
                for col in st.session_state.plantilla_maestra.columns:
                    if col not in bloque_dia.columns: bloque_dia[col] = ''
                
                st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, bloque_dia], ignore_index=True)
                st.success(f"🎉 ¡Conciliación exitosa! Registros anexados al {dia_actual}.")
                st.rerun()
else:
    st.info("👋 ¡Bienvenido! Por favor, diríjase al panel lateral izquierdo para cargar los extractos maestros. Puede cargar varios archivos por categoría y omitir las categorías que no tengan movimientos hoy.")
