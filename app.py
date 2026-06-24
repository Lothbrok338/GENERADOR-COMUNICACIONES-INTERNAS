import streamlit as st
import pandas as pd
import io

# Configuración de página
st.set_page_config(page_title="Sistema Facturación Univalle", layout="wide")

st.title("📄 Sistema de Facturación Univalle")

# Definición de listas
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

uploaded_file = st.file_uploader("Cargar reporte original", type=["csv", "xls", "xlsx"])

if uploaded_file:
    try:
        # 1. Carga robusta
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=10)
        else:
            try:
                df = pd.read_excel(uploaded_file, skiprows=10, engine='openpyxl')
            except:
                # Si falla, intentamos leerlo como tabla HTML (por si es un Excel "falso")
                uploaded_file.seek(0)
                df = pd.read_html(uploaded_file, skiprows=10)[0]
        
        # 2. Limpieza y Filtro
        df = df.dropna(how='all') # Elimina filas totalmente vacías
        df = df[df['Estado'] == 'Válido'].copy()
        
        # 3. Preparación de columnas
        df['Facturacion'] = None
        df['Clinica'] = 0.0
        df['Total C.I.'] = df['Monto']
        df['Banco'] = None
        df['Glosa Asiento'] = ""
        df['Cuenta'] = ""
        df['Asignacion'] = ""

        st.write("Edita la tabla para asignar Facturación y Banco:")
        
        # 4. Editor interactivo
        edited_df = st.data_editor(
            df,
            column_config={
                "Facturacion": st.column_config.SelectboxColumn("Facturación", options=opciones_facturacion),
                "Banco": st.column_config.SelectboxColumn("Banco", options=opciones_banco),
            },
            use_container_width=True
        )

        if st.button("Generar Procesamiento Final"):
            # Lógica: Clínica y Total CI
            edited_df.loc[edited_df['Facturacion'] == 'Reserva Int.', 'Clinica'] = 2400.0
            
            # Asegurar numéricos
            edited_df['Monto'] = pd.to_numeric(edited_df['Monto'], errors='coerce').fillna(0)
            edited_df['Clinica'] = pd.to_numeric(edited_df['Clinica'], errors='coerce').fillna(0)
            edited_df['Total C.I.'] = edited_df['Monto'] + edited_df['Clinica']
            
            # Lógica: Glosa
            edited_df['Glosa Asiento'] = edited_df.apply(
                lambda x: f"FAC {x['Número Factura']} {x['Nombre Estudiante']} {str(x['Facturacion'] or '')}"[:43], axis=1
            )
            
            # Lógica: Cuenta
            edited_df['Cuenta'] = edited_df['Banco'].map(mapa_cuentas)
            
            # Resultado
            st.success("Reporte procesado exitosamente")
            st.dataframe(edited_df)
            
            # Descarga
            csv = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button("Descargar Excel Final", csv, "reporte_final_univalle.csv", "text/csv")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
