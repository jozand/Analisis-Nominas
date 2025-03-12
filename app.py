import streamlit as st
import pandas as pd
import plotly.express as px

def load_data(file):
    sheets = pd.read_excel(file, sheet_name=None, engine="openpyxl", dtype=str)  # Cargar todas las hojas
    for sheet_name, df in sheets.items():
        df.columns = df.columns.str.strip()  # Elimina espacios en los nombres de las columnas
        if "NIT" in df.columns:
            df["NIT"] = df["NIT"].astype(str)  # Asegurar que NIT sea string
        if len(df.columns) > 11:
            df = df.rename(columns={df.columns[11]: "UNUM"})  # Columna L es UNUM
    return sheets

def comparar_meses(df1, df2):
    codigos_mes_actual = set(df1["CODIGO"].astype(str))
    codigos_mes_anterior = set(df2["CODIGO"].astype(str))
    
    altas = df1[df1["CODIGO"].astype(str).isin(codigos_mes_actual - codigos_mes_anterior)]
    bajas = df2[df2["CODIGO"].astype(str).isin(codigos_mes_anterior - codigos_mes_actual)]
    
    return altas, bajas

def main():
    st.title("📊 Comparador de Datos Mensuales por Hojas")
    st.subheader("Sube los archivos de dos meses para analizar altas y bajas en cada hoja del Excel")
    
    file1 = st.file_uploader("Sube el archivo del mes actual", type=["xlsx"])
    file2 = st.file_uploader("Sube el archivo del mes anterior", type=["xlsx"])
    
    if file1 and file2:
        sheets1 = load_data(file1)
        sheets2 = load_data(file2)
        
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
    
if __name__ == "__main__":
    main()