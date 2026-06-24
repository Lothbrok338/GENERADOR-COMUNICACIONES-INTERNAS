import streamlit as st
import pandas as pd

# Configuración de página
st.set_page_config(page_title="Sistema Facturación Univalle", layout="wide")

st.title("📄 Sistema de Facturación Univalle")

# Definición de listas para desplegables
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

uploaded_file = st.file_uploader("Cargar reporte original (CSV)", type="csv")

if uploaded_file:
    # 1. Carga y filtro inicial
    df = pd.read_csv(uploaded_file, skiprows=10)
    df = df[df['Estado'] == 'Válido'].copy()
    
    # 2. Inicializar columnas para el editor
    df['Facturacion'] = None
    df['Clinica'] = 0.0
    df['Total C.I.'] = df['Monto']
    df['Banco'] = None
    df['Glosa Asiento'] = ""
    df['Cuenta'] = ""
    df['Asignacion'] = ""

    st.write("Edita la tabla para asignar Facturación y Banco:")
    
    # 3. Editor interactivo
    edited_df = st.data_editor(
        df,
        column_config={
            "Facturacion": st.column_config.SelectboxColumn("Facturación", options=opciones_facturacion),
            "Banco": st.column_config.SelectboxColumn("Banco", options=opciones_banco),
        },
        use_container_width=True
    )

    if st.button("Generar Procesamiento Final"):
        # 4. Cálculos automáticos
        # Clínica y Total CI
        edited_df.loc[edited_df['Facturacion'] == 'Reserva Int.', 'Clinica'] = 2400.0
        edited_df['Total C.I.'] = edited_df['Monto'] + edited_df['Clinica']
        
        # Glosa (Trunca a 43 caracteres)
        edited_df['Glosa Asiento'] = edited_df.apply(
            lambda x: f"FAC {x['Número Factura']} {x['Nombre Estudiante']} {x['Facturacion']}"[:43], axis=1
        )
        
        # Cuenta
        edited_df['Cuenta'] = edited_df['Banco'].map(mapa_cuentas)
        
        # Resultado
        st.success("Reporte procesado exitosamente")
        st.dataframe(edited_df)
        
        # Botón descarga
        csv = edited_df.to_csv(index=False).encode('utf-8')
        st.download_button("Descargar Excel Final", csv, "reporte_final_univalle.csv", "text/csv")
