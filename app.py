import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador SAP Interactivo", layout="wide")

st.title("🗂️ Consolidador SAP - Selección por Día")
st.write("Sube tus archivos, selecciona las filas correspondientes al día actual y confírmalas. Podrás descargar un archivo individual por cada día procesado.")

# --- INICIALIZAR MEMORIA ---
if 'plantilla_maestra' not in st.session_state:
    columnas_sap = ['DIA_ETIQUETA', 'BUKRS', 'HKONT', 'SGTXT', 'WRSOL', 'WRHAB', 'DMBTR', 'DMBE2', 'MWSKZ', 'TXJCD', 'KOSTL', 'PRCTR', 'AUFNR', 'PS_POSID', 'VALUT', 'HBKID', 'HKTID', 'ZUONR', 'VBUND', 'FIPEX']
    st.session_state.plantilla_maestra = pd.DataFrame(columns=columnas_sap)

# Control de estado para los botones de Seleccionar Todo
if 'marcar_todo' not in st.session_state:
    st.session_state.marcar_todo = True

# --- FUNCIÓN PARA CARGAR LOS EXCELS ---
@st.cache_data
def leer_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.upper()
    return df

col_panel, col_resumen = st.columns([1.5, 1])

with col_panel:
    st.subheader("⚙️ 1. Carga y Selección")
    dia_actual = st.selectbox("¿Qué día estás procesando ahora mismo?", [f"Día {i}" for i in range(1, 32)])
    
    st.markdown("---")
    file_cajas = st.file_uploader("📂 Archivo CAJAS", type=["xlsx", "xls"], key="cajas")
    file_atc = st.file_uploader("📂 Archivo ATC", type=["xlsx", "xls"], key="atc")
    file_com = st.file_uploader("📂 Archivo COMUNICACIONES", type=["xlsx", "xls"], key="com")

    if file_cajas and file_atc and file_com:
        st.write("### Revisa y selecciona las filas a incluir:")
        
        # --- BOTONES DE SELECCIÓN MASIVA ---
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("✅ Marcar Todas las Filas", use_container_width=True):
                st.session_state.marcar_todo = True
        with col_btn2:
            if st.button("⬜ Desmarcar Todas las Filas", use_container_width=True):
                st.session_state.marcar_todo = False
        
        # Leemos los archivos
        df_cajas_raw = leer_excel(file_cajas).copy()
        df_atc_raw = leer_excel(file_atc).copy()
        df_com_raw = leer_excel(file_com).copy()

        # Insertamos la columna SELECCIONAR según el estado del botón
        df_cajas_raw.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
        df_atc_raw.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
        df_com_raw.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)

        # Clave dinámica para que las tablas se refresquen al tocar los botones o cambiar de día
        key_suffix = f"{dia_actual}_{st.session_state.marcar_todo}"

        # Mostramos los editores interactivos
        st.write("**Cajas Diarias**")
        edit_cajas = st.data_editor(df_cajas_raw, hide_index=True, use_container_width=True, key=f"edit_cajas_{key_suffix}")
        
        st.write("**ATC Unificado**")
        edit_atc = st.data_editor(df_atc_raw, hide_index=True, use_container_width=True, key=f"edit_atc_{key_suffix}")
        
        st.write("**Comunicaciones Internas**")
        edit_com = st.data_editor(df_com_raw, hide_index=True, use_container_width=True, key=f"edit_com_{key_suffix}")

        if st.button(f"📥 Confirmar filas marcadas para el {dia_actual}", type="primary", use_container_width=True):
            # 1. Filtrar solo las filas que el usuario dejó marcadas (True)
            df_cajas = edit_cajas[edit_cajas['SELECCIONAR'] == True].copy()
            df_atc = edit_atc[edit_atc['SELECCIONAR'] == True].copy()
            df_com = edit_com[edit_com['SELECCIONAR'] == True].copy()

            # 2. Limpieza de Asignaciones
            def limpiar_asig(df, col, respaldo):
                if df.empty: return df
                return df.get(col, df.get(respaldo, '')).fillna('').astype(str).replace('nan', '', regex=True)

            if not df_cajas.empty: df_cajas['ZUONR_FINAL'] = limpiar_asig(df_cajas, 'ASIGNACION', 'CÓDIGO ASIENTO')
            if not df_atc.empty: df_atc['ZUONR_FINAL'] = limpiar_asig(df_atc, 'ASIGNACION', 'CÓDIGO ASIENTO')
            if not df_com.empty: df_com['ZUONR_FINAL'] = limpiar_asig(df_com, 'ASIGNACION', '')

            # 3. Mapeo SAP
            def parse_monto(df, col):
                return df[col].astype(str).str.replace(',', '.').astype(float)

            sap_cajas = pd.DataFrame() if df_cajas.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': df_cajas['CUENTA CONTABLE'], 'SGTXT': df_cajas['GLOSA RECORTADA'],
                'WRSOL': parse_monto(df_cajas, 'CRÉDITOS'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(df_cajas['FECHA']), 'ZUONR': df_cajas['ZUONR_FINAL']
            })

            sap_atc = pd.DataFrame() if df_atc.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': df_atc['CUENTA CONTABLE'], 'SGTXT': df_atc['DETALLE'],
                'WRSOL': parse_monto(df_atc, 'MONTO'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(df_atc['FECHA']), 'ZUONR': df_atc['ZUONR_FINAL']
            })

            sap_com = pd.DataFrame() if df_com.empty else pd.DataFrame({
                'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': df_com['CUENTA CONTABLE BANCO'], 'SGTXT': df_com['GLOSA ASIENTO COMUNICACIONES INTERNAS'],
                'WRSOL': parse_monto(df_com, 'TOTAL C.I.'), 'WRHAB': '', 'PRCTR': '10010101',
                'VALUT': pd.to_datetime(df_com['FECHA']), 'ZUONR': df_com['ZUONR_FINAL']
            })

            # Unir y guardar en memoria global
            bloque_dia = pd.concat([sap_cajas, sap_atc, sap_com])
            
            if not bloque_dia.empty:
                for col in st.session_state.plantilla_maestra.columns:
                    if col not in bloque_dia.columns:
                        bloque_dia[col] = ''
                
                st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, bloque_dia], ignore_index=True)
                st.success(f"✅ ¡{len(bloque_dia)} filas guardadas correctamente bajo el {dia_actual}!")
            else:
                st.warning("⚠️ No seleccionaste ninguna fila para guardar.")

# --- SECCIÓN DE EXPORTACIÓN (ARCHIVOS SEPARADOS) ---
with col_resumen:
    st.subheader("📦 2. Exportación por Días")
    
    if len(st.session_state.plantilla_maestra) > 0:
        # Obtenemos qué días han sido procesados
        dias_procesados = st.session_state.plantilla_maestra['DIA_ETIQUETA'].unique()
        
        st.info(f"Tienes {len(dias_procesados)} días listos para exportar.")

        for dia in dias_procesados:
            # Filtramos la memoria para obtener solo el día actual del bucle
            df_dia = st.session_state.plantilla_maestra[st.session_state.plantilla_maestra['DIA_ETIQUETA'] == dia].copy()
            
            # Limpiamos la columna de control antes de exportar
            df_export = df_dia.drop(columns=['DIA_ETIQUETA'])
            
            # Formatos SAP
            df_export['VALUT'] = pd.to_datetime(df_export['VALUT']).dt.strftime('%d/%m/%Y')
            def string_format_sap(val):
                try: return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except: return val
            df_export['WRSOL'] = df_export['WRSOL'].apply(string_format_sap)
            
            # Generar CSV en memoria
            csv_data = df_export.to_csv(index=False, sep='|', header=True)
            
            # Crear un botón de descarga ÚNICO para este día
            st.download_button(
                label=f"⬇️ Descargar SAP {dia}",
                data=csv_data,
                file_name=f"PLANTILLA_SAP_{dia.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        st.markdown("---")
        if st.button("🗑️ Limpiar Memoria Completa", type="secondary", use_container_width=True):
            st.session_state.plantilla_maestra = pd.DataFrame(columns=columnas_sap)
            st.rerun() # Reinicia la página
    else:
        st.write("Aún no has confirmado ningún día. Los botones de descarga aparecerán aquí a medida que confirmes la selección.")
