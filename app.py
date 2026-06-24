import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Sistema Facturación Univalle", layout="wide")
st.title("📄 Sistema de Facturación Univalle")

# Listas
opciones_facturacion = ["Posgrado", "Pregrado", "Reserva Int."]
opciones_banco = [
    "BNB MN", "BNB CLINICA", "BISA MN", "BCP MN", "BUSA MN", "BANECO", 
    "BMS", "GASTO.ADM", "ALQUILERES", "BNB ME", "BISA ME", "BCP ME", 
    "BUSA ME", "POSGRADO.PLA", "BANECO AH", "BISA EURO", "BNB AH"
]

mapa_cuentas = {
    "BNB MN": "110103012", "BNB CLINICA": "110103022", "BISA MN": "110103032",
    "BCP MN": "110103042", "BUSA MN": "110103052", "BANECO": "110103062",
    "BMS": "110103072", "BNB ME": "110104012", "BISA ME": "110104022",
    "BCP ME": "110104032", "BUSA ME": "110104042", "BANECO AH": "110103722",
    "BISA EURO": "110105112", "BNB AH": "110103712"
}

uploaded_file = st.file_uploader("Cargar reporte original", type=["csv", "xls", "xlsx", "html"])

if uploaded_file:
    try:
        # Intento 1: CSV crudo (omitiendo líneas que causan error)
        try:
            df = pd.read_csv(uploaded_file, skiprows=10, engine='python', on_bad_lines='skip')
        except:
            # Intento 2: Excel
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, skiprows=10, engine='openpyxl')
        
        # Filtro de seguridad: Eliminar filas donde todas las columnas sean nulas
        df = df.dropna(how='all')
        
        # Filtro de Válidos (ajusta 'Estado' si en tu archivo se llama diferente)
        if 'Estado' in df.columns:
            df = df[df['Estado'] == 'Válido'].copy()
        else:
            st.warning("No se encontró la columna 'Estado'. Mostrando todos los datos.")

        # Preparación de columnas
        df['Facturacion'] = None
        df['Clinica'] = 0.0
        df['Total C.I.'] = df['Monto']
        df['Banco'] = None
        df['Glosa Asiento'] = ""
        df['Cuenta'] = ""
        df['Asignacion'] = ""

        edited_df = st.data_editor(df, use_container_width=True)

        if st.button("Generar Procesamiento Final"):
            edited_df.loc[edited_df['Facturacion'] == 'Reserva Int.', 'Clinica'] = 2400.0
            edited_df['Monto'] = pd.to_numeric(edited_df['Monto'], errors='coerce').fillna(0)
            edited_df['Total C.I.'] = edited_df['Monto'] + edited_df['Clinica']
            
            # Glosa truncada a 43 caracteres
            edited_df['Glosa Asiento'] = edited_df.apply(
                lambda x: f"FAC {x.get('Número Factura', '')} {x.get('Nombre Estudiante', '')} {str(x['Facturacion'] or '')}"[:43], axis=1
            )
            
            edited_df['Cuenta'] = edited_df['Banco'].map(mapa_cuentas)
            
            st.success("Reporte procesado exitosamente")
            st.dataframe(edited_df)
            
    except Exception as e:
        st.error(f"Error técnico: {e}. ¿Podrías abrir el archivo en Excel y guardarlo como .xlsx manualmente antes de subirlo?")
