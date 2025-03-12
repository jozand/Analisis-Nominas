import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

def load_data(file):
    sheets = pd.read_excel(file, sheet_name=None, engine="openpyxl", dtype=str)  # Cargar todas las hojas
    clean_sheets = {}
    for sheet_name, df in sheets.items():
        sheet_name = sheet_name.strip()  # Eliminar espacios en blanco en los nombres de las hojas
        df.columns = df.columns.str.strip()  # Eliminar espacios en blanco en los nombres de las columnas
        df["Hoja"] = sheet_name  # Agregar nombre de la hoja como columna
        if "NIT" in df.columns:
            df["NIT"] = df["NIT"].astype(str)  # Asegurar que NIT sea string
        if "CODIGO" in df.columns:
            df["CODIGO"] = df["CODIGO"].astype(str)  # Asegurar que CODIGO sea string
        if len(df.columns) > 11:
            df = df.rename(columns={df.columns[11]: "UNUM"})  # Columna L es UNUM
        clean_sheets[sheet_name] = df  # Guardar la hoja con el nombre limpio
    return clean_sheets

def comparar_meses(df1, df2):
    codigos_mes_actual = set(df1["CODIGO"].astype(str))
    codigos_mes_anterior = set(df2["CODIGO"].astype(str))
    
    altas = df1[df1["CODIGO"].astype(str).isin(codigos_mes_actual - codigos_mes_anterior)]
    bajas = df2[df2["CODIGO"].astype(str).isin(codigos_mes_anterior - codigos_mes_actual)]
    
    return altas, bajas

def identificar_movimientos(sheets1, sheets2):
    hojas_excluidas = ["Gtos. de Representacion"]  # Lista de hojas a excluir
    
    df1_total = pd.concat([df[["CODIGO", "Hoja"]] for sheet, df in sheets1.items() if "CODIGO" in df and sheet not in hojas_excluidas], ignore_index=True)
    df2_total = pd.concat([df[["CODIGO", "Hoja"]] for sheet, df in sheets2.items() if "CODIGO" in df and sheet not in hojas_excluidas], ignore_index=True)
    
    merged = df1_total.merge(df2_total, on="CODIGO", suffixes=("_actual", "_anterior"), how="inner")
    
    cambios = merged[(merged["Hoja_actual"] != merged["Hoja_anterior"])]
    
    # Excluir cualquier registro donde una de las hojas sea "Gtos. de Representación"
    cambios = cambios[~cambios["Hoja_anterior"].isin(hojas_excluidas) & ~cambios["Hoja_actual"].isin(hojas_excluidas)]
    cambios = cambios[["CODIGO", "Hoja_anterior", "Hoja_actual"]]
    
    return cambios

def generar_sankey(df):
    if df.empty:
        return None
    
    source_labels = list(set(df["Hoja_anterior"].dropna().unique()).union(set(df["Hoja_actual"].dropna().unique())))
    source_target_map = {label: idx for idx, label in enumerate(source_labels)}
    
    df = df.dropna(subset=["Hoja_anterior", "Hoja_actual"])
    df_grouped = df.groupby(["Hoja_anterior", "Hoja_actual"]).size().reset_index(name="count")
    
    source_indices = df_grouped["Hoja_anterior"].map(source_target_map).tolist()
    target_indices = df_grouped["Hoja_actual"].map(source_target_map).tolist()
    values = df_grouped["count"].tolist()
    
    sankey_figure = go.Figure(go.Sankey(
        node=dict(
            pad=15,
            thickness=20,
            label=source_labels
        ),
        link=dict(
            source=source_indices,
            target=target_indices,
            value=values
        )
    ))
    
    sankey_figure.update_layout(title_text="Flujo de Movimientos entre Hojas", font_size=10)
    return sankey_figure

def main():
    st.title("📊 Comparador de Datos Mensuales por Hojas")
    st.subheader("Sube los archivos de dos meses para analizar cambios y comparar altas y bajas en cada hoja del Excel")
    
    file1 = st.file_uploader("Sube el archivo del mes actual", type=["xlsx"])
    file2 = st.file_uploader("Sube el archivo del mes anterior", type=["xlsx"])
    
    if file1 and file2:
        sheets1 = load_data(file1)
        sheets2 = load_data(file2)
        
        # Sección de comparación por hoja
        sheet_names = list(sheets1.keys())
        selected_sheet = st.selectbox("Selecciona una hoja para analizar", sheet_names)
        
        if selected_sheet in sheets2:
            df1 = sheets1[selected_sheet]
            df2 = sheets2[selected_sheet]
            
            altas, bajas = comparar_meses(df1, df2)
            
            st.subheader("📈 Altas (Nuevos registros)")
            st.dataframe(altas)
            st.subheader("📉 Bajas (Registros que desaparecieron)")
            st.dataframe(bajas)
            
            # Gráfico de comparativa de altas y bajas
            st.subheader("📊 Comparación de Altas y Bajas")
            comparacion_df = pd.DataFrame({"Tipo": ["Altas", "Bajas"], "Cantidad": [len(altas), len(bajas)]})
            fig = px.bar(comparacion_df, x="Tipo", y="Cantidad", title=f"Comparación de Altas y Bajas en {selected_sheet}", text_auto=True, color="Tipo")
            st.plotly_chart(fig)
        
        # Sección de identificación de movimientos entre hojas
        st.subheader("🔄 Movimientos entre Hojas")
        movimientos = identificar_movimientos(sheets1, sheets2)
        st.dataframe(movimientos.head(50))  # Limitar la visualización para mejorar el rendimiento
        
        # Gráfico de movimientos con Sankey
        fig_sankey = generar_sankey(movimientos)
        if fig_sankey:
            st.plotly_chart(fig_sankey)
    
if __name__ == "__main__":
    main()
