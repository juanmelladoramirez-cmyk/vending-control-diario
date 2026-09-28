import streamlit as st
import pandas as pd
import os
import datetime

st.set_page_config(page_title="Control de Máquinas Vending", page_icon="🥤", layout="wide")

st.title("🥤 Sistema de Control de Máquinas Vending")
st.markdown("Gestión interactiva de inventario, ubicaciones y estado de máquinas.")

excel_file = "maquinas vending.xlsx"

if not os.path.exists(excel_file):
    st.error(f"⚠️ No se encontró el archivo '{excel_file}' en la misma carpeta.")
else:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        
        selected_sheet = st.sidebar.selectbox("Seleccionar Hoja / Vista", sheet_names)
        
        df = pd.read_excel(xls, selected_sheet)
        
        # Configurar columnas de fecha para que muestren un calendario interactivo
        column_config = {}
        for col in df.columns:
            if 'FECHA' in col.upper() or 'BOLK' in col.upper() or pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
                column_config[col] = st.column_config.DateColumn(
                    col,
                    format="YYYY-MM-DD",
                    default=datetime.date.today(),
                )
        
        st.subheader(f"Vista: {selected_sheet}")
        st.info("💡 Haz clic en cualquier celda de fecha para abrir un **calendario desplegable** y seleccionar el día fácilmente. Usa el botón '+' para agregar nuevas filas.")
        
        search_query = st.text_input("🔍 Buscar en la tabla (por nombre, dirección, código, etc.):")
        if search_query:
            mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            df_view = df[mask]
        else:
            df_view = df
            
        # Editor interactivo con selector de calendario en las fechas
        edited_df = st.data_editor(
            df_view, 
            num_rows="dynamic", 
            use_container_width=True, 
            column_config=column_config,
            key=f"editor_{selected_sheet}"
        )
        
        if st.button("💾 Guardar Cambios en el Excel"):
            if search_query:
                st.warning("⚠️ Por favor limpia el cuadro de búsqueda antes de guardar para asegurar que se guarden todos los datos correctamente.")
            else:
                try:
                    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                        for s in sheet_names:
                            if s == selected_sheet:
                                edited_df.to_excel(writer, sheet_name=s, index=False)
                            else:
                                temp_df = pd.read_excel(xls, s)
                                for col in temp_df.columns:
                                    if 'FECHA' in col.upper() or 'BOLK' in col.upper() or pd.api.types.is_datetime64_any_dtype(temp_df[col]):
                                        temp_df[col] = pd.to_datetime(temp_df[col], errors='coerce').dt.date
                                temp_df.to_excel(writer, sheet_name=s, index=False)
                    st.success("✅ ¡Cambios guardados exitosamente en tu archivo Excel!")
                except Exception as ex:
                    st.error(f"Error al guardar: {ex}")
                
        st.text(f"Total de registros: {len(df_view)}")
        
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")