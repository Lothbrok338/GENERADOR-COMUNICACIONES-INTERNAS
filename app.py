import streamlit as st
import pandas as pd

st.set_page_config(page_title="Consolidador SAP Pro", layout="wide")

st.title("🗂️ Consolidador SAP - Control Total")

# --- 1. MEMORIA ESTRUCTURADA ---
if 'plantilla_maestra' not in st.session_state:
    st.session_state.plantilla_maestra = pd.DataFrame()

# --- 2. FUNCIONES AUXILIARES ---
def limpiar_asig(df, col, respaldo):
    return df.get(col, df.get(respaldo, '')).fillna('').astype(str).replace('nan', '', regex=True)

# --- 3. INTERFAZ: PANEL DE CARGA ---
col_carga, col_control = st.columns([1, 1.5])

with col_carga:
    st.subheader("⚙️ Carga de Datos")
    dia_actual = st.selectbox("Seleccione el Día:", [f"Día {i:02d}" for i in range(1, 32)])
    
    file_cajas = st.file_uploader("📂 Cajas", type=["xlsx"], key="cajas")
    file_atc = st.file_uploader("📂 ATC", type=["xlsx"], key="atc")
    file_com = st.file_uploader("📂 Comunicaciones", type=["xlsx"], key="com")

    marcar_todo_com = st.checkbox("✅ Marcar todas las filas de Comunicaciones")

with col_control:
    if file_cajas and file_atc and file_com:
        df_cajas = pd.read_excel(file_cajas)
        df_atc = pd.read_excel(file_atc)
        df_com = pd.read_excel(file_com)
        
        # Preparar tablas con columna de selección
        for df in [df_cajas, df_atc, df_com]: 
            df.columns = df.columns.str.strip().str.upper()
            df.insert(0, 'SELECCIONAR', True)
        
        # Aplicar el "Marcar Todo" solo en Comunicaciones
        if marcar_todo_com: df_com['SELECCIONAR'] = True
        else: df_com['SELECCIONAR'] = False

        st.write(f"### Edición: {dia_actual}")
        edit_cajas = st.data_editor(df_cajas, use_container_width=True)
        edit_atc = st.data_editor(df_atc, use_container_width=True)
        edit_com = st.data_editor(df_com, use_container_width=True)

        if st.button("➕ Confirmar día y agregar"):
            # Procesamiento lógico (mismo que antes)
            # ... (aquí iría el mapeo)
            # Añadimos al estado con la etiqueta 'dia_actual'
            st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, nuevo_bloque])
            st.rerun()

# --- 4. PANEL DE CONTROL Y BORRADO ---
st.markdown("---")
st.subheader("📦 Gestión de Días Procesados")

if not st.session_state.plantilla_maestra.empty:
    dias_cargados = st.session_state.plantilla_maestra['DIA_ETIQUETA'].unique()
    
    for dia in sorted(dias_cargados):
        col1, col2 = st.columns([4, 1])
        col1.write(f"✅ {dia} procesado")
        
        if col2.button(f"❌ Descartar {dia}", key=f"del_{dia}"):
            st.session_state.plantilla_maestra = st.session_state.plantilla_maestra[
                st.session_state.plantilla_maestra['DIA_ETIQUETA'] != dia
            ]
            st.rerun()

    # Botón de exportación final
    st.download_button("⬇️ Descargar Consolidado Final", data=st.session_state.plantilla_maestra.to_csv(sep='|'), file_name="SAP_FINAL.csv")
