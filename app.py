import streamlit as st
import pandas as pd
import io

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Consolidador SAP Pro", layout="wide", initial_sidebar_state="expanded")

# --- MEMORIA ESTRUCTURADA ---
if 'plantilla_maestra' not in st.session_state:
    st.session_state.plantilla_maestra = pd.DataFrame()

# --- FUNCIONES AUXILIARES ---
def limpiar_asig(df, col, respaldo):
    if df.empty: return df
    return df.get(col, df.get(respaldo, '')).fillna('').astype(str).replace('nan', '', regex=True)

def parse_monto(df, col):
    if df.empty: return df
    return df[col].astype(str).str.replace(',', '.').astype(float)

# --- MAQUETACIÓN DE LA INTERFAZ ---
st.title("🗂️ Consolidador SAP - Control de Flujo")
st.write("Selecciona el día, mapea las pestañas y confirma las filas antes de generar tu reporte SAP.")

col_izq, col_der = st.columns([1.2, 2.8])

# ==========================================
# PANEL IZQUIERDO: CONTROLES Y BORRADO
# ==========================================
with col_izq:
    st.subheader("⚙️ 1. Panel de Carga")
    dia_actual = st.selectbox("Seleccione el Día a procesar:", [f"Día {i:02d}" for i in range(1, 32)])
    
    file_cajas = st.file_uploader("📂 Archivo CAJAS", type=["xlsx", "xls"])
    file_atc = st.file_uploader("📂 Archivo ATC", type=["xlsx", "xls"])
    file_com = st.file_uploader("📂 Archivo COMUNICACIONES", type=["xlsx", "xls"])

    hoja_com_seleccionada = None
    marcar_todo_com = True

    # Selector de pestañas dinámico para Comunicaciones
    if file_com:
        st.markdown("---")
        st.write("📋 **Opciones de Comunicaciones**")
        # Leer nombres de las pestañas y ordenarlos de menor a mayor
        xls_com = pd.ExcelFile(file_com)
        hojas_ordenadas = sorted(xls_com.sheet_names)
        
        hoja_com_seleccionada = st.selectbox("Selecciona la pestaña a procesar:", hojas_ordenadas)
        marcar_todo_com = st.checkbox("✅ Marcar todas las filas por defecto", value=True)

    st.markdown("---")
    st.subheader("🗑️ Historial de Días (Descartar)")
    st.write("Haz clic en la 'X' para borrar un día específico si te equivocaste.")
    
    if not st.session_state.plantilla_maestra.empty:
        # Extraer los días que ya están guardados en memoria y ordenarlos
        dias_cargados = sorted(st.session_state.plantilla_maestra['DIA_ETIQUETA'].unique())
        
        for dia in dias_cargados:
            c1, c2 = st.columns([3, 1])
            c1.write(f"📁 **{dia}**")
            # Botón de borrado único (La "X")
            if c2.button("❌", key=f"del_{dia}", help=f"Eliminar {dia}"):
                st.session_state.plantilla_maestra = st.session_state.plantilla_maestra[st.session_state.plantilla_maestra['DIA_ETIQUETA'] != dia]
                st.rerun()
    else:
        st.info("No hay días procesados en la memoria.")

# ==========================================
# PANEL DERECHO: VISTA DE DATOS Y EXPORTACIÓN
# ==========================================
with col_der:
    if file_cajas and file_atc and file_com:
        st.subheader(f"🔍 2. Auditoría del {dia_actual}")
        
        try:
            # Leer los archivos
            df_cajas = pd.read_excel(file_cajas)
            df_atc = pd.read_excel(file_atc)
            # Leer específicamente la pestaña seleccionada en el panel izquierdo
            df_com = pd.read_excel(file_com, sheet_name=hoja_com_seleccionada)

            # Normalizar nombres de columnas
            for df in [df_cajas, df_atc, df_com]:
                df.columns = df.columns.str.strip().str.upper()

            # Insertar columna de checkboxes (Comunicaciones respeta el botón de "Marcar Todo")
            df_cajas.insert(0, 'SELECCIONAR', True)
            df_atc.insert(0, 'SELECCIONAR', True)
            df_com.insert(0, 'SELECCIONAR', marcar_todo_com)

            # Editores interactivos
            st.write("**Cajas Diarias**")
            edit_cajas = st.data_editor(df_cajas, hide_index=True, use_container_width=True, key="ed_cajas")
            
            st.write("**ATC Unificado**")
            edit_atc = st.data_editor(df_atc, hide_index=True, use_container_width=True, key="ed_atc")
            
            st.write(f"**Comunicaciones Internas (Pestaña: {hoja_com_seleccionada})**")
            edit_com = st.data_editor(df_com, hide_index=True, use_container_width=True, key="ed_com")

            # Botón para inyectar a SAP
            if st.button(f"📥 Confirmar {dia_actual} y Guardar", type="primary", use_container_width=True):
                # Verificar que el día no exista ya en memoria
                if not st.session_state.plantilla_maestra.empty and dia_actual in st.session_state.plantilla_maestra['DIA_ETIQUETA'].values:
                    st.error(f"⚠️ El {dia_actual} ya existe en la memoria. Bórralo primero con la 'X' en el panel izquierdo si necesitas sobreescribirlo.")
                else:
                    # Filtrar solo lo marcado
                    f_cajas = edit_cajas[edit_cajas['SELECCIONAR'] == True].copy()
                    f_atc = edit_atc[edit_atc['SELECCIONAR'] == True].copy()
                    f_com = edit_com[edit_com['SELECCIONAR'] == True].copy()

                    # Limpieza Asignaciones
                    if not f_cajas.empty: f_cajas['ZUONR_FINAL'] = limpiar_asig(f_cajas, 'ASIGNACION', 'CÓDIGO ASIENTO')
                    if not f_atc.empty: f_atc['ZUONR_FINAL'] = limpiar_asig(f_atc, 'ASIGNACION', 'CÓDIGO ASIENTO')
                    if not f_com.empty: f_com['ZUONR_FINAL'] = limpiar_asig(f_com, 'ASIGNACION', '')

                    # Mapeo a SAP
                    sap_cajas = pd.DataFrame() if f_cajas.empty else pd.DataFrame({
                        'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': f_cajas['CUENTA CONTABLE'], 'SGTXT': f_cajas['GLOSA RECORTADA'],
                        'WRSOL': parse_monto(f_cajas, 'CRÉDITOS'), 'WRHAB': '', 'PRCTR': '10010101',
                        'VALUT': pd.to_datetime(f_cajas['FECHA'], errors='coerce'), 'ZUONR': f_cajas['ZUONR_FINAL']
                    })

                    sap_atc = pd.DataFrame() if f_atc.empty else pd.DataFrame({
                        'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': f_atc['CUENTA CONTABLE'], 'SGTXT': f_atc['DETALLE'],
                        'WRSOL': parse_monto(f_atc, 'MONTO'), 'WRHAB': '', 'PRCTR': '10010101',
                        'VALUT': pd.to_datetime(f_atc['FECHA'], errors='coerce'), 'ZUONR': f_atc['ZUONR_FINAL']
                    })

                    sap_com = pd.DataFrame() if f_com.empty else pd.DataFrame({
                        'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': f_com['CUENTA CONTABLE BANCO'], 'SGTXT': f_com['GLOSA ASIENTO COMUNICACIONES INTERNAS'],
                        'WRSOL': parse_monto(f_com, 'TOTAL C.I.'), 'WRHAB': '', 'PRCTR': '10010101',
                        'VALUT': pd.to_datetime(f_com['FECHA'], errors='coerce'), 'ZUONR': f_com['ZUONR_FINAL']
                    })

                    bloque_dia = pd.concat([sap_cajas, sap_atc, sap_com])
                    
                    if not bloque_dia.empty:
                        # Rellenar 19 columnas de SAP
                        cols_sap = ['DIA_ETIQUETA', 'BUKRS', 'HKONT', 'SGTXT', 'WRSOL', 'WRHAB', 'DMBTR', 'DMBE2', 'MWSKZ', 'TXJCD', 'KOSTL', 'PRCTR', 'AUFNR', 'PS_POSID', 'VALUT', 'HBKID', 'HKTID', 'ZUONR', 'VBUND', 'FIPEX']
                        for col in cols_sap:
                            if col not in bloque_dia.columns:
                                bloque_dia[col] = ''
                        
                        bloque_dia = bloque_dia[cols_sap]
                        
                        # Inyectar en memoria
                        st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, bloque_dia], ignore_index=True)
                        st.rerun() # Recarga la página para mostrar el éxito
                    else:
                        st.warning("⚠️ No seleccionaste ninguna fila para procesar.")
        except Exception as e:
            st.error(f"Error procesando los archivos: {e}")
    else:
        st.info("👈 Empieza subiendo los archivos del día en el panel izquierdo.")

    # --- 3. ÁREA DE DESCARGA GLOBAL ---
    if not st.session_state.plantilla_maestra.empty:
        st.markdown("---")
        st.subheader("🚀 3. Plantilla SAP Lista para Descarga")
        
        # Preparar copia para visualización y descarga
        df_export = st.session_state.plantilla_maestra.copy()
        
        # Quitar la columna interna de control de días
        df_export = df_export.drop(columns=['DIA_ETIQUETA'])

        # Aplicar Formato Fecha DD/MM/YYYY
        df_export['VALUT'] = pd.to_datetime(df_export['VALUT']).dt.strftime('%d/%m/%Y')
        
        # Aplicar Formato Numérico 1.000,12
        def string_format_sap(val):
            try: return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            except: return val
        
        df_export['WRSOL'] = df_export['WRSOL'].apply(string_format_sap)

        # Mostrar tabla final
        st.dataframe(df_export, use_container_width=True)

        # Botón de Descarga
        csv_data = df_export.to_csv(index=False, sep='|', header=True)
        st.download_button(
            label="⬇️ Descargar Archivo SAP Consolidado Total (.csv)", 
            data=csv_data, 
            file_name="PLANTILLA_SAP_TOTAL_FINAL.csv", 
            mime="text/csv", 
            type="primary",
            use_container_width=True
        )
