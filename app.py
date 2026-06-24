import streamlit as st
import pandas as pd
import io

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(page_title="Consolidador Acumulativo SAP", layout="wide")

st.title("🗂️ Consolidador Stateful SAP por Días - Univalle")
st.write("Sube los archivos de forma secuencial (día por día). Los datos se irán acumulando de forma segura en la memoria de la aplicación.")

# 2. INICIALIZACIÓN DE LA MEMORIA (session_state)
if 'plantilla_maestra' not in st.session_state:
    # Creamos un DataFrame vacío con las 19 columnas oficiales de tu plantilla SAP
    columnas_sap = ['BUKRS', 'HKONT', 'SGTXT', 'WRSOL', 'WRHAB', 'DMBTR', 'DMBE2', 'MWSKZ', 'TXJCD', 'KOSTL', 'PRCTR', 'AUFNR', 'PS_POSID', 'VALUT', 'HBKID', 'HKTID', 'ZUONR', 'VBUND', 'FIPEX']
    st.session_state.plantilla_maestra = pd.DataFrame(columns=columnas_sap)

# 3. CONTROLES DE LA INTERFAZ
col_control, col_resumen = st.columns([1, 2])

with col_control:
    st.subheader("⚙️ Panel de Carga")
    # Selector de control para organizar el flujo del usuario
    dia_actual = st.selectbox("Seleccione el periodo/día a procesar:", [f"Día {i}" for i in range(1, 32)])
    
    st.markdown("---")
    file_cajas = st.file_uploader(f"📂 Archivo CAJAS ({dia_actual})", type=["xlsx", "xls"])
    file_atc = st.file_uploader(f"📂 Archivo ATC ({dia_actual})", type=["xlsx", "xls"])
    file_com = st.file_uploader(f"📂 Archivo COMUNICACIONES ({dia_actual})", type=["xlsx", "xls"])

    # Botón para inyectar los datos cargados a la memoria global
    if st.button("📥 Confirmar y Añadir a Plantilla Global", use_container_width=True, type="secondary"):
        if file_cajas and file_atc and file_com:
            try:
                # Lectura de flujos
                df_cajas = pd.read_excel(file_cajas)
                df_atc = pd.read_excel(file_atc)
                df_com = pd.read_excel(file_com)

                # Limpieza de columnas estándar
                df_cajas.columns = df_cajas.columns.str.strip().str.upper()
                df_atc.columns = df_atc.columns.str.strip().str.upper()
                df_com.columns = df_com.columns.str.strip().str.upper()

                # Control estricto de nulos en asignación
                def limpiar_asig(df, col, respaldo):
                    return df.get(col, df.get(respaldo, '')).fillna('').astype(str).replace('nan', '', regex=True)

                df_cajas['ZUONR_FINAL'] = limpiar_asig(df_cajas, 'ASIGNACION', 'CÓDIGO ASIENTO')
                df_atc['ZUONR_FINAL'] = limpiar_asig(df_atc, 'ASIGNACION', 'CÓDIGO ASIENTO')
                df_com['ZUONR_FINAL'] = limpiar_asig(df_com, 'ASIGNACION', '')

                # Conversión técnica de montos
                def parse_monto(df, col):
                    return df[col].astype(str).str.replace(',', '.').astype(float)

                # Mapeos individuales a SAP
                sap_cajas = pd.DataFrame({
                    'BUKRS': 'BO01', 'HKONT': df_cajas['CUENTA CONTABLE'], 'SGTXT': df_cajas['GLOSA RECORTADA'],
                    'WRSOL': parse_monto(df_cajas, 'CRÉDITOS'), 'WRHAB': '', 'PRCTR': '10010101',
                    'VALUT': pd.to_datetime(df_cajas['FECHA']), 'ZUONR': df_cajas['ZUONR_FINAL']
                })

                sap_atc = pd.DataFrame({
                    'BUKRS': 'BO01', 'HKONT': df_atc['CUENTA CONTABLE'], 'SGTXT': df_atc['DETALLE'],
                    'WRSOL': parse_monto(df_atc, 'MONTO'), 'WRHAB': '', 'PRCTR': '10010101',
                    'VALUT': pd.to_datetime(df_atc['FECHA']), 'ZUONR': df_atc['ZUONR_FINAL']
                })

                sap_com = pd.DataFrame({
                    'BUKRS': 'BO01', 'HKONT': df_com['CUENTA CONTABLE BANCO'], 'SGTXT': df_com['GLOSA ASIENTO COMUNICACIONES INTERNAS'],
                    'WRSOL': parse_monto(df_com, 'TOTAL C.I.'), 'WRHAB': '', 'PRCTR': '10010101',
                    'VALUT': pd.to_datetime(df_com['FECHA']), 'ZUONR': df_com['ZUONR_FINAL']
                })

                # Unión del bloque actual
                bloque_dia = pd.concat([sap_cajas, sap_atc, sap_com])
                
                # Rellenar las columnas faltantes del layout SAP
                for col in st.session_state.plantilla_maestra.columns:
                    if col not in bloque_dia.columns:
                        bloque_dia[col] = ''
                
                bloque_dia = bloque_dia[st.session_state.plantilla_maestra.columns]

                # CONCATENACIÓN ACUMULATIVA: Se añade al estado global de la sesión
                st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, bloque_dia], ignore_index=True)
                st.success(f"💥 ¡Datos de {dia_actual} acumulados con éxito!")
            
            except Exception as e:
                st.error(f"Error procesando los archivos: {e}")
        else:
            st.warning("Por favor sube los 3 archivos requeridos antes de confirmar.")

    if st.button("🗑️ Reiniciar Toda la Plantilla", use_container_width=True):
        st.session_state.plantilla_maestra = pd.DataFrame(columns=columnas_sap)
        st.experimental_rerun()

# 4. MONITOR DE ESTADO Y VISTA PREVIA GLOBAL
with col_resumen:
    st.subheader("📈 Estado del Asiento Acumulado")
    
    total_filas = len(st.session_state.plantilla_maestra)
    
    # KPI metrics en pantalla
    st.metric(label="Total de Filas en el Debe Acumuladas", value=total_filas)
    
    st.markdown("---")
    st.write("### Vista Previa de la Plantilla Global (Acumulada)")
    
    if total_filas > 0:
        # Hacemos una copia para aplicar el formato visual solicitado sin alterar los datos numéricos internos
        df_visual = st.session_state.plantilla_maestra.copy()
        
        # Aplicar formato de fecha DD/MM/YYYY
        df_visual['VALUT'] = pd.to_datetime(df_visual['VALUT']).dt.strftime('%d/%m/%Y')
        
        # Aplicar formato de número europeo/boliviano: 1.000,12
        def string_format_sap(val):
            try:
                return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            except:
                return val
                
        df_visual['WRSOL'] = df_visual['WRSOL'].apply(string_format_sap)
        
        # Mostrar grilla interactiva
        st.dataframe(df_visual, use_container_width=True)
        
        # 5. EXPORTACIÓN ÚNICA DEL CSV COMPLETO
        st.markdown("### 📥 Fase Final de Carga Masiva")
        csv_data = df_visual.to_csv(index=False, sep='|', header=True)
        
        st.download_button(
            label="🚀 Exportar Plantilla Consolidada Completa a SAP (.csv)",
            data=csv_data,
            file_name="PLANTILLA_SAP_CONSOLIDADO_DIARIO.csv",
            mime="text/csv",
            use_container_width=True,
            type="primary"
        )
    else:
        st.info("La plantilla está vacía. Comienza a confirmar periodos en el panel izquierdo para ver la acumulación de datos contables.")
