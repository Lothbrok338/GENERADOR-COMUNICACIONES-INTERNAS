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

if 'dia_seleccionado' not in st.session_state:
    st.session_state.dia_seleccionado = "Día 1"

for key in ['df_cajas', 'df_atc', 'df_com']:
    if key not in st.session_state:
        st.session_state[key] = pd.DataFrame()

for key in ['name_cajas', 'name_atc', 'name_com']:
    if key not in st.session_state:
        st.session_state[key] = []

# --- FUNCIÓN DE INGESTA Y EXTRACCIÓN ---
def procesar_subida_multiple(lista_archivos, state_name, state_df, tipo_archivo):
    nombres_actuales = [f.name for f in lista_archivos] if lista_archivos else []
    if nombres_actuales != st.session_state.get(state_name, []):
        if lista_archivos:
            dfs_temporales = []
            for file_obj in lista_archivos:
                if tipo_archivo == 'COMUN':
                    diccionario_hojas = pd.read_excel(file_obj, sheet_name=None)
                    for nombre_pestaña, df in diccionario_hojas.items():
                        if not df.empty:
                            df.columns = df.columns.str.strip().str.upper()
                            df['FUENTE_ARCHIVO'] = f"PESTAÑA DÍA {nombre_pestaña}"
                            dfs_temporales.append(df)
                else:
                    df = pd.read_excel(file_obj)
                    df.columns = df.columns.str.strip().str.upper()
                    nombre_limpio = file_obj.name.split('.')[0].upper().replace('_', ' ')
                    df['FUENTE_ARCHIVO'] = nombre_limpio
                    if tipo_archivo == 'ATC':
                        tiene_caja = any('CAJA' in col for col in df.columns)
                        if not tiene_caja:
                            match = re.search(r'(SFC\d+|SOUVENIRS?)', file_obj.name.upper())
                            df['NÚMERO DE CAJA'] = match.group(1) if match else 'GENERAL'
                    dfs_temporales.append(df)
            
            if dfs_temporales:
                df_consolidado = pd.concat(dfs_temporales, ignore_index=True)
                df_consolidado['PROCESADO'] = False 
                df_consolidado['ORIGINAL_INDEX'] = df_consolidado.index 
                st.session_state[state_df] = df_consolidado
            else:
                st.session_state[state_df] = pd.DataFrame()
        else:
            st.session_state[state_df] = pd.DataFrame()
        st.session_state[state_name] = nombres_actuales

# --- FUNCIONES DE FORMATO EXCEL ---
def aplicar_formato_excel(writer, df, sheet_name="SAP"):
    # Obtener el workbook y el worksheet de xlsxwriter
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Definir formatos
    formato_encabezado = workbook.add_format({
        'bold': True,
        'bg_color': '#D9E1F2',  # Azul claro
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    formato_celda = workbook.add_format({
        'border': 1,
        'valign': 'vcenter'
    })
    
    formato_numero = workbook.add_format({
        'border': 1,
        'valign': 'vcenter',
        'num_format': '#,##0.00'  # Formato numérico estándar de Excel
    })

    # Escribir encabezados
    for col_num, value in enumerate(df.columns):
        worksheet.write(0, col_num, value, formato_encabezado)

    # Escribir datos
    for row_num in range(len(df)):
        for col_num in range(len(df.columns)):
            valor = df.iloc[row_num, col_num]
            # Aplicar formato de número a la columna WRSOL (índice 3, ya que DIA_ETIQUETA se elimina antes)
            if col_num == 3 and pd.notnull(valor):
                 worksheet.write(row_num + 1, col_num, valor, formato_numero)
            else:
                 # Reemplazar NaN con string vacío para Excel
                 val_str = "" if pd.isna(valor) else valor
                 worksheet.write(row_num + 1, col_num, val_str, formato_celda)

    # Ajustar anchos de columna
    # Columnas con datos: BUKRS(0), HKONT(1), SGTXT(2), WRSOL(3)
    worksheet.set_column(0, 0, 8)   # BUKRS
    worksheet.set_column(1, 1, 12)  # HKONT
    worksheet.set_column(2, 2, 40)  # SGTXT (Glosa)
    worksheet.set_column(3, 3, 15)  # WRSOL (Monto)
    
    # Columnas E a M (índices 4 a 12): WRHAB, DMBTR, DMBE2, MWSKZ, TXJCD, KOSTL, PRCTR, AUFNR, PS_POSID
    worksheet.set_column(4, 12, 3)  # Muy angostas
    
    # Resto de columnas
    worksheet.set_column(13, 13, 12) # VALUT (Fecha)
    worksheet.set_column(14, 15, 3)  # HBKID, HKTID angostas
    worksheet.set_column(16, 16, 20) # ZUONR (Asignación)
    worksheet.set_column(17, 18, 3)  # VBUND, FIPEX angostas


# --- PANEL LATERAL ---
with st.sidebar:
    st.header("⚙️ Módulo de Ingesta")
    file_cajas = st.file_uploader("📂 Cargar extracto CAJAS", type=["xlsx", "xls"], accept_multiple_files=True)
    file_atc = st.file_uploader("📂 Cargar extracto ATC", type=["xlsx", "xls"], accept_multiple_files=True)
    file_com = st.file_uploader("📂 Libro Maestro COMUNICACIONES", type=["xlsx", "xls"], accept_multiple_files=True)

    procesar_subida_multiple(file_cajas, 'name_cajas', 'df_cajas', 'CAJAS')
    procesar_subida_multiple(file_atc, 'name_atc', 'df_atc', 'ATC')
    procesar_subida_multiple(file_com, 'name_com', 'df_com', 'COMUN')

    st.markdown("---")
    if len(st.session_state.plantilla_maestra) > 0:
        dias_procesados = st.session_state.plantilla_maestra['DIA_ETIQUETA'].unique()
        st.success(f"📊 {len(dias_procesados)} periodos listos:")
        
        # --- BOTÓN DE EXPORTACIÓN CONSOLIDADA TOTAL ---
        df_total = st.session_state.plantilla_maestra.copy()
        df_total['num_dia'] = df_total['DIA_ETIQUETA'].str.replace('Día ', '').astype(int)
        
        # Ordenamos y separamos por bloques
        dias_ordenados_totales = sorted(df_total['num_dia'].unique())
        
        # Crear un buffer de Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            # Preparar un DataFrame consolidado con filas en blanco
            dfs_a_concatenar = []
            
            # Fila vacía para los separadores
            fila_vacia = pd.DataFrame([[""] * (len(df_total.columns) - 2)], columns=df_total.drop(columns=['DIA_ETIQUETA', 'num_dia']).columns)
            
            for idx, n_dia in enumerate(dias_ordenados_totales):
                dia_str = f"Día {n_dia}"
                df_dia_bloque = df_total[df_total['DIA_ETIQUETA'] == dia_str].copy()
                df_export_bloque = df_dia_bloque.drop(columns=['DIA_ETIQUETA', 'num_dia'])
                
                df_export_bloque['VALUT'] = pd.to_datetime(df_export_bloque['VALUT']).dt.strftime('%d/%m/%Y')
                
                # Convertir WRSOL a numérico para que Excel lo formatee bien
                df_export_bloque['WRSOL'] = pd.to_numeric(df_export_bloque['WRSOL'], errors='coerce')
                
                # Si no es el primer bloque, insertar filas separadoras e intentar recrear el encabezado
                if idx > 0:
                    dfs_a_concatenar.extend([fila_vacia, fila_vacia, fila_vacia])
                    # Insertar una fila que funcione como encabezado visual
                    df_header = pd.DataFrame([df_export_bloque.columns], columns=df_export_bloque.columns)
                    dfs_a_concatenar.append(df_header)
                    
                dfs_a_concatenar.append(df_export_bloque)

            df_final_excel = pd.concat(dfs_a_concatenar, ignore_index=True)
            
            # Escribir al Excel
            df_final_excel.to_excel(writer, sheet_name='SAP', index=False)
            
            # Obtener workbook y worksheet
            workbook = writer.book
            worksheet = writer.sheets['SAP']

            # Formato de celdas y numérico
            formato_celda = workbook.add_format({'valign': 'vcenter'})
            formato_numero = workbook.add_format({'valign': 'vcenter', 'num_format': '#,##0.00'})
            formato_encabezado_extra = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})

            for row_num in range(len(df_final_excel)):
                for col_num in range(len(df_final_excel.columns)):
                    valor = df_final_excel.iloc[row_num, col_num]
                    # Si el valor de la celda es igual al nombre de la columna, tratarlo como encabezado
                    if valor == df_final_excel.columns[col_num]:
                         worksheet.write(row_num + 1, col_num, valor, formato_encabezado_extra)
                    elif col_num == 3 and pd.notnull(valor) and valor != "": # Columna WRSOL
                         try:
                             worksheet.write_number(row_num + 1, col_num, float(valor), formato_numero)
                         except ValueError:
                             worksheet.write(row_num + 1, col_num, valor, formato_celda)
                    else:
                         val_str = "" if pd.isna(valor) else valor
                         worksheet.write(row_num + 1, col_num, val_str, formato_celda)

            # Formatear el primer encabezado (fila 0)
            for col_num, value in enumerate(df_final_excel.columns):
                 worksheet.write(0, col_num, value, formato_encabezado_extra)

            # Ajustar anchos
            worksheet.set_column(0, 0, 8)   # BUKRS
            worksheet.set_column(1, 1, 12)  # HKONT
            worksheet.set_column(2, 2, 40)  # SGTXT (Glosa)
            worksheet.set_column(3, 3, 15)  # WRSOL (Monto)
            worksheet.set_column(4, 12, 3)  # Columnas E a M (WRHAB a PS_POSID) angostas
            worksheet.set_column(13, 13, 12) # VALUT (Fecha)
            worksheet.set_column(14, 15, 3)  # HBKID, HKTID angostas
            worksheet.set_column(16, 16, 20) # ZUONR (Asignación)
            worksheet.set_column(17, 18, 3)  # VBUND, FIPEX angostas
            
        excel_data_total = output.getvalue()
        
        st.download_button(
            label="📦 Descargar Consolidado Total (Excel)", 
            data=excel_data_total, 
            file_name="SAP_CONSOLIDADO_COMPLETO.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            use_container_width=True,
            type="primary"
        )
        
        st.markdown("---")
        st.write("⬇️ O descargar por día individual:")

        for dia in sorted(dias_procesados, key=lambda x: int(x.replace('Día ', ''))):
            df_dia = st.session_state.plantilla_maestra[st.session_state.plantilla_maestra['DIA_ETIQUETA'] == dia].copy()
            df_export = df_dia.drop(columns=['DIA_ETIQUETA'])
            df_export['VALUT'] = pd.to_datetime(df_export['VALUT']).dt.strftime('%d/%m/%Y')
            
            # Asegurar WRSOL como número
            df_export['WRSOL'] = pd.to_numeric(df_export['WRSOL'], errors='coerce')
            
            # Generar Excel para el día individual
            output_dia = io.BytesIO()
            with pd.ExcelWriter(output_dia, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, sheet_name='SAP', index=False)
                aplicar_formato_excel(writer, df_export)
                
            excel_data_dia = output_dia.getvalue()
            
            col_d1, col_d2 = st.columns([4, 1])
            with col_d1:
                st.download_button(f"📄 Descargar {dia} (Excel)", excel_data_dia, f"SAP_{dia.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_d2:
                if st.button("❌", key=f"del_{dia}", help=f"Descartar y borrar el {dia}"):
                    st.session_state.plantilla_maestra = st.session_state.plantilla_maestra[st.session_state.plantilla_maestra['DIA_ETIQUETA'] != dia]
                    st.rerun()
            
    st.markdown("---")
    if st.button("🗑️ Limpiar y Borrar Todo", type="secondary", use_container_width=True):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

# --- ÁREA PRINCIPAL ---
if not st.session_state.df_cajas.empty or not st.session_state.df_atc.empty or not st.session_state.df_com.empty:
    col_dia, col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1, 1])
    with col_dia: 
        # Modificación para evitar StreamlitAPIException al avanzar el día
        if 'temp_next_day' in st.session_state:
            st.session_state.dia_seleccionado = st.session_state.temp_next_day
            del st.session_state.temp_next_day
            
        dia_actual = st.selectbox("📅 Asignar transacciones al periodo:", [f"Día {i}" for i in range(1, 32)], key='dia_seleccionado')
    with col_btn1: 
        if st.button("✅ Marcar Todo", use_container_width=True): 
            st.session_state.marcar_todo = True
            st.rerun()
    with col_btn2: 
        if st.button("⬜ Desmarcar Todo", use_container_width=True): 
            st.session_state.marcar_todo = False
            st.rerun()
    with col_btn3:
        if st.button("📑 Todo Com.", help="Marcar solo filas de Comunicaciones", use_container_width=True): 
            st.session_state.marcar_todo_com = True
            st.rerun()

    df_pen_cajas = st.session_state.df_cajas[st.session_state.df_cajas['PROCESADO'] == False].copy() if not st.session_state.df_cajas.empty else pd.DataFrame()
    df_pen_atc = st.session_state.df_atc[st.session_state.df_atc['PROCESADO'] == False].copy() if not st.session_state.df_atc.empty else pd.DataFrame()
    df_pen_com = st.session_state.df_com[st.session_state.df_com['PROCESADO'] == False].copy() if not st.session_state.df_com.empty else pd.DataFrame()

    st.markdown("---")
    st.subheader("🔍 Filtros de Operación")
    col_filtro_caja, col_filtro_com = st.columns(2)
    
    col_caja_field = next((c for c in df_pen_cajas.columns if 'CAJA' in c), None)
    caja_filtrada = "Mostrar Todas"
    if col_caja_field or 'NÚMERO DE CAJA' in df_pen_atc.columns:
        cajas = set(df_pen_cajas[col_caja_field].dropna().astype(str)) if col_caja_field and not df_pen_cajas.empty else set()
        if 'NÚMERO DE CAJA' in df_pen_atc.columns: cajas.update(df_pen_atc['NÚMERO DE CAJA'].dropna().astype(str))
        caja_filtrada = col_filtro_caja.selectbox("🛒 Filtrar Cajas/ATC:", ["Mostrar Todas"] + sorted(list(cajas)))

    com_filtrada = "Mostrar Todas"
    if 'FUENTE_ARCHIVO' in df_pen_com.columns:
        def natural_sort_key(s):
            return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]
        opciones_com = sorted(list(df_pen_com['FUENTE_ARCHIVO'].unique()), key=natural_sort_key)
        com_filtrada = col_filtro_com.selectbox("📑 Filtrar Comunicaciones:", ["Mostrar Todas"] + opciones_com)

    if caja_filtrada != "Mostrar Todas":
        if not df_pen_cajas.empty and col_caja_field: df_pen_cajas = df_pen_cajas[df_pen_cajas[col_caja_field].astype(str) == caja_filtrada]
        if not df_pen_atc.empty and 'NÚMERO DE CAJA' in df_pen_atc.columns: df_pen_atc = df_pen_atc[df_pen_atc['NÚMERO DE CAJA'].astype(str) == caja_filtrada]
    if com_filtrada != "Mostrar Todas" and not df_pen_com.empty: df_pen_com = df_pen_com[df_pen_com['FUENTE_ARCHIVO'] == com_filtrada]

    if not df_pen_cajas.empty: df_pen_cajas.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
    if not df_pen_atc.empty: df_pen_atc.insert(0, 'SELECCIONAR', st.session_state.marcar_todo)
    if not df_pen_com.empty: df_pen_com.insert(0, 'SELECCIONAR', st.session_state.get('marcar_todo_com', st.session_state.marcar_todo))

    for df in [df_pen_cajas, df_pen_atc, df_pen_com]: 
        if not df.empty and 'FECHA' in df.columns: df['FECHA'] = pd.to_datetime(df['FECHA']).dt.strftime('%d/%m/%Y')

    key_suffix = f"{dia_actual}_{st.session_state.marcar_todo}_{caja_filtrada}_{com_filtrada}"
    
    st.subheader(f"🛒 Cajas ({len(df_pen_cajas)}) | 💳 ATC ({len(df_pen_atc)}) | 📑 Com. ({len(df_pen_com)})")
    if not df_pen_cajas.empty: edit_cajas = st.data_editor(df_pen_cajas, hide_index=True, use_container_width=True, height=300, key=f"ed_c_{key_suffix}")
    else: edit_cajas = pd.DataFrame()
    if not df_pen_atc.empty: edit_atc = st.data_editor(df_pen_atc, hide_index=True, use_container_width=True, height=300, key=f"ed_a_{key_suffix}")
    else: edit_atc = pd.DataFrame()
    if not df_pen_com.empty: edit_com = st.data_editor(df_pen_com, hide_index=True, use_container_width=True, height=300, key=f"ed_co_{key_suffix}")
    else: edit_com = pd.DataFrame()

    if st.button(f"🚀 VERIFICAR Y ANEXAR AL {dia_actual.upper()}", type="primary", use_container_width=True):
        sel_c = edit_cajas[edit_cajas['SELECCIONAR'] == True].copy() if not edit_cajas.empty else pd.DataFrame()
        sel_a = edit_atc[edit_atc['SELECCIONAR'] == True].copy() if not edit_atc.empty else pd.DataFrame()
        sel_co = edit_com[edit_com['SELECCIONAR'] == True].copy() if not edit_com.empty else pd.DataFrame()
        if sel_c.empty and sel_a.empty and sel_co.empty: st.warning("⚠️ Selecciona transacciones.")
        else:
            if not sel_c.empty: st.session_state.df_cajas.loc[sel_c['ORIGINAL_INDEX'], 'PROCESADO'] = True
            if not sel_a.empty: st.session_state.df_atc.loc[sel_a['ORIGINAL_INDEX'], 'PROCESADO'] = True
            if not sel_co.empty: st.session_state.df_com.loc[sel_co['ORIGINAL_INDEX'], 'PROCESADO'] = True
            
            def parse_m(df, col): return df[col].astype(str).str.replace(',', '.').astype(float)
            
            s_c = pd.DataFrame({'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_c.get('CUENTA CONTABLE', ''), 'SGTXT': sel_c.get('GLOSA RECORTADA', ''), 'WRSOL': parse_m(sel_c, 'CRÉDITOS'), 'VALUT': pd.to_datetime(sel_c['FECHA'], dayfirst=True), 'ZUONR': sel_c.get('ASIGNACION', '')}) if not sel_c.empty else pd.DataFrame()
            s_a = pd.DataFrame({'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_a.get('CUENTA CONTABLE', ''), 'SGTXT': sel_a.get('DETALLE', ''), 'WRSOL': parse_m(sel_a, 'MONTO'), 'VALUT': pd.to_datetime(sel_a['FECHA'], dayfirst=True), 'ZUONR': sel_a.get('ASIGNACION', '')}) if not sel_a.empty else pd.DataFrame()
            s_co = pd.DataFrame({'DIA_ETIQUETA': dia_actual, 'BUKRS': 'BO01', 'HKONT': sel_co.get('CUENTA CONTABLE BANCO', ''), 'SGTXT': sel_co.get('GLOSA ASIENTO COMUNICACIONES INTERNAS', ''), 'WRSOL': parse_m(sel_co, 'TOTAL C.I.'), 'VALUT': pd.to_datetime(sel_co['FECHA'], dayfirst=True), 'ZUONR': sel_co.get('ASIGNACION', '')}) if not sel_co.empty else pd.DataFrame()
            
            bloque = pd.concat([s_c, s_a, s_co])
            for col in st.session_state.plantilla_maestra.columns:
                if col not in bloque.columns: bloque[col] = ''
            
            st.session_state.plantilla_maestra = pd.concat([st.session_state.plantilla_maestra, bloque], ignore_index=True)
            
            # Modificación: En lugar de asignar directo al widget en esta misma carga, lo delegamos
            current_day_num = int(dia_actual.replace("Día ", ""))
            next_day_num = min(current_day_num + 1, 31)
            st.session_state.temp_next_day = f"Día {next_day_num}"
            
            st.success("🎉 ¡Conciliación exitosa!")
            st.rerun()
else:
    st.info("Carga archivos para iniciar.")
