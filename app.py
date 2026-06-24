import streamlit as st
import pandas as pd
import io

# Configuración de página ancha nativa
st.set_page_config(page_title="Consolidador SAP Interactivo", layout="wide")

st.title("🗂️ Consolidador SAP - Modo Pantalla Completa")
st.write("Las tablas ahora se expanden **hasta el extremo derecho de la pantalla** para darte máxima visibilidad. Usa el panel izquierdo para cargar y exportar.")

# --- 1. INICIALIZAR MEMORIA Y ESTADOS ---
if 'plantilla_maestra' not in st.session_state:
    columnas_sap = ['DIA_ETIQUETA', 'BUKRS', 'HKONT', 'SGTXT', 'WRSOL', 'WRHAB', 'DMBTR', 'DMBE2', 'MWSKZ', 'TXJCD', 'KOSTL', 'PRCTR', 'AUFNR', 'PS_POSID', 'VALUT', 'HBKID', 'HKTID', 'ZUONR', 'VBUND', 'FIPEX']
    st.session_state.plantilla_maestra = pd.DataFrame(columns=columnas_sap)

if 'marcar_todo' not in st.session_state:
    st.session_state.marcar_todo = True

# Variables para guardar los datos pendientes
for key in ['df_cajas', 'df_atc', 'df_com']:
    if key not in st.session_state:
        st.session_state[key] = pd.DataFrame()

# Control de nombres de archivo para estabilidad de la sesión
for key in ['name_cajas', 'name_atc', 'name_com']:
    if key not in st.session_state:
        st.session_state[key] = ""

# --- FUNCIÓN PARA CARGAR Y ETIQUETAR DATOS NUEVOS ---
def procesar_subida(file_obj, state_name, state_df):
    if file_obj and st.session_state[state_name] != file_obj.name:
        df = pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()
        df['PROCESADO'] = False 
        df['ORIGINAL_INDEX'] = df.index 
        st.session_state[state_df] = df
        st.session_state[state_name] = file_obj.name

# --- PANEL LATERAL (SIDEBAR) CONTROLES Y EXPORTACIÓN ---
with st.sidebar:
    st.header("⚙️ Panel de Control")
    file_cajas = st.file_uploader("📂 Cargar CAJAS", type=["xlsx", "xls"])
    file_atc = st.file_uploader("📂 Cargar ATC", type=["xlsx", "xls"])
    file_com = st.file_uploader("📂 Cargar COMUNICACIONES", type=["xlsx", "xls"])

    # Inyectar datos a memoria de fondo
    procesar_subida(file_cajas, 'name_cajas', 'df_cajas')
    procesar_subida(file_atc, 'name_atc', 'df_atc')
    procesar_subida(file_com, 'name_com', 'df_com')

    st.markdown("---")
    st.header("📦 Archivos Listos SAP")
    
    if len(st.session_state.plantilla_maestra) > 0:
        dias_procesados = st.session_state.plantilla_maestra['DIA_ETIQUETA'].unique()
        st.success(f"📊 {len(dias_procesados)} periodos guardados:")

        for dia in sorted(dias_procesados):
            df_dia = st.session_state.plantilla_maestra[st.session_state.plantilla_maestra['DIA_ETIQUETA'] == dia].copy()
            df_export = df_dia.drop(columns=['DIA_ETIQUETA'])
            
            # Formatos de salida obligatorios para SAP
            df_export['VALUT'] = pd.to_datetime(df_export['VALUT']).dt.strftime('%d/%m/%Y')
            def string_format_sap(val):
                try: return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except: return val
            df_export['WRSOL'] = df_export['WRSOL'].apply(string_format_sap)
            
            csv_data = df_export.to_csv(index=False, sep='|', header=True)
            st.download_button(label=f"⬇️ Descargar {dia}", data=csv_data, file_name=f"SAP_{dia.replace(' ', '_')}.csv", mime="text/csv", use_container_width=True)
            
        st.markdown("---")
        if st.button("🗑️ Resetear Todo de Cero", type="secondary", use_container_width=True):
            for key in list(st.session_state.keys()): 
                del st.session_state[key]
            st.rerun()
    else:
        st.info("Los botones de descarga aparecerán aquí conforme vayas confirmando días en la mesa principal.")

# --- ÁREA PRINCIPAL (PANTALLA COMPLETA HASTA LA DERECHA) ---
if not st.session_state.df_cajas.empty and not st.session_state.df_atc.empty and not st.session_state.df_com.empty:
    
    # Controles superiores extendidos
    col_dia, col_btn1, col_btn2 = st.columns([3, 1, 1])
    with col_dia: dia_actual = st.selectbox("📅 Asignar las filas marcadas al periodo:", [f"Día {i}" for i in range(1, 32)])
    with col_btn1: 
        if st.button("✅ Marcar Todo", use_container_width=True): 
            st.session_state.marcar_todo = True
            st.rerun()
    with col_btn2: 
        if st.button("⬜ Desmarcar Todo", use_container_width=True): 
            st.session_state.marcar_todo = False
            st.rerun()

    # Filtrado dinámico de filas (Bandeja de entrada Cero)
    df_pendientes_cajas = st.session_state.df_cajas[st.session_state.df_cajas['PROCESADO'] == False].copy()
    df_pendientes_atc = st.session_state.df_atc[st.session_state.df_atc['PROCESADO'] == False].copy()
    df_pendientes_com = st.session_state.df_com[st.session_state.df_com['PROCESADO'] == False].copy()

    # Inyectar booleanos de control de selección
    df_pendientes_cajas.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
    df_pendientes_atc.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
    df_pendientes_com.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)

    key_suffix = f"{dia_actual}_{st.session_state.marcar_todo}"

    st.markdown("---")
    
    # SECCIÓN DE TABLAS TOTALMENTE EXPANDIDAS HORIZONTALMENTE
    st.subheader(f"🛒 Cajas Diarias — ({len(df_pendientes_cajas)} registros pendientes)")
    edit_cajas = st.data_editor(df_pendientes_cajas, hide_index=True, use_container_width=True, height=400, key=f"ed_c_{key_suffix}")
    
    st.subheader(f"💳 ATC Unificado — ({len(df_pendientes_atc)} registros pendientes)")
    edit_atc = st.data_editor(df_pendientes_atc, hide_index=True, use_container_width=True, height=400, key=f"ed_a_{key_suffix}")
    
    st.subheader(f"📑 Comunicaciones Internas — ({len(df_pendientes_com)} registros pendientes)")
    edit_com = st.data_editor(df_pendientes_com, hide_index=True, use_container_width=True, height=400, key=f"ed_co_{key_suffix}")

    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # Botón maestro de confirmación de filas
    if st.button(f"🚀 PROCESAR Y ENVIAR FILAS MARCADAS AL {dia_actual.upper()}", type="primary", use_container_width=True):
        sel_cajas = edit_cajas[edit_cajas['SELECCIONAR'] == True].copy()
        sel_atc = edit_atc[edit_atc['SELECCIONAR'] == True].copy()
        sel_com = edit_com[edit_com['SELECCIONAR'] == True].copy()

        if sel_cajas.empty and sel_atc.empty and sel_com.empty:
            st.warning("⚠️ No has seleccionado ninguna fila válida.")
        else:
            # Registrar filas como procesadas de forma interna
            st.session_state.df_cajas.loc[sel_cajas['ORIGINAL_INDEX'], 'PROCESADO'] = True
            st.session_state.df_atc.loc[sel_atc['ORIGINAL_INDEX'], 'PROCESADO'] = True
            st.session_state.df_com.loc[sel_com['ORIGINAL_INDEX'], 'PROCESADO'] = True

            # Limpieza quirúrgica de asignaciones
            def limpiar_asig(df, col, respaldo):
                if df.empty: return df
                return df.get(col, df.get(respaldo, '')).fillna('').astype(str).replace('nan', '', regex=True)

            if not sel_cajas.empty: sel_cajas['ZUONR_FINAL'] = limpiar_asig(sel_cajas, 'ASIGNACION', 'CÓDIGO ASIENTO')
            if not sel_atc.empty: sel_atc['ZUONR_FINAL'] = limpiar_asig(sel_atc, 'ASIGNACION', 'CÓDIGO ASIENTO')
            if not sel_com.empty: sel_com['ZUONR_FINAL'] = limpiar_asig(sel_com, 'ASIGNACION', '')

            # Parseo monetario limpio
            def parse_monto(df, col):
                return df[col].astype(str).str.replace(',', '.').astype(float)

            # Mapeo estructurado SAP
            sap_cajas = pd.DataFrame() if sel_cajas.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_cajas['CUENTA CONTABLE'], 'SGTXT': sel_cajas['GLOSA RECORTADA'],
                'WRSOL': parse_monto(sel_cajas, 'CRÉDITOS'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(sel_cajas['FECHA']), 'ZUONR': sel_cajas['ZUONR_FINAL']
            })

            sap_atc = pd.DataFrame() if sel_atc.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_atc['CUENTA CONTABLE'], 'SGTXT': sel_atc['DETALLE'],
                'WRSOL': parse_monto(sel_atc, 'MONTO'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(sel_atc['FECHA']), 'ZUONR': sel_atc['ZUONR_FINAL']
            })

            sap_com = pd.DataFrame() if sel_com.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_com['CUENTA CONTABLE BANCO'], 'SGTXT': sel_com['GLOSA ASIENTO COMUNICACIONES INTERNAS'],
                'WRSOL': parse_monto(sel_com, 'TOTAL C.I.'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(sel_com['FECHA']), 'ZUONR': sel_com['ZUONR_FINAL']
            })

            # Añadir lote procesado a la canasta maestra global
            bloque_dia = pd.concat([sap_cajas, sap_atc, sap_com])
            for col in st.session_state.plantilla_maestra.columns:
                if col not in bloque_dia.columns: bloque_dia[col] = ''
            
            st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, bloque_dia], ignore_index=True)
            st.success(f"🎉 ¡Fase completada! Filas asignadas y empaquetadas en el {dia_actual}.")
            st.rerun()
else:
    st.info("👋 ¡Bienvenido! Por favor, ve al panel lateral izquierdo para cargar tus archivos maestros de Cajas, ATC y Comunicaciones.")
