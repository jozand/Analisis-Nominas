import streamlit as st
import pandas as pd
import plotly.express as px

def load_data(file):
    df = pd.read_excel(file, engine="openpyxl", dtype=str)  # Forzar todo como texto
    df = df.iloc[:, :59]  # Limitar a las columnas relevantes
    df.columns = df.columns.str.strip()  # Elimina espacios en los nombres de las columnas
    df = df.rename(columns={df.columns[11]: "UNUM"})  # Columna L es UNUM
    
    # Asegurar que la columna "NIT" sea de tipo string
    if "NIT" in df.columns:
        df["NIT"] = df["NIT"].astype(str)
    
    return df

def comparar_meses(df1, df2):
    """Compara dos dataframes y encuentra altas y bajas."""
    codigos_mes_actual = set(df1["CODIGO"].astype(str))
    codigos_mes_anterior = set(df2["CODIGO"].astype(str))
    
    altas = df1[df1["CODIGO"].astype(str).isin(codigos_mes_actual - codigos_mes_anterior)]
    bajas = df2[df2["CODIGO"].astype(str).isin(codigos_mes_anterior - codigos_mes_actual)]
    
    return altas, bajas

def main():
    st.title("📊 Comparador de Datos Mensuales")
    st.subheader("Sube los archivos de dos meses para analizar altas y bajas")
    
    file1 = st.file_uploader("Sube el archivo del mes actual", type=["xlsx"])
    file2 = st.file_uploader("Sube el archivo del mes anterior", type=["xlsx"])
    
    if file1 and file2:
        df1 = load_data(file1)
        df2 = load_data(file2)
        
        altas, bajas = comparar_meses(df1, df2)
        
        st.subheader("📈 Altas (Nuevos registros)")
        st.dataframe(altas)
        st.subheader("📉 Bajas (Registros que desaparecieron)")
        st.dataframe(bajas)
        
        # Gráficos
        # st.subheader("📊 Comparación de registros por categoría")
        # cat_fig = px.bar(df1, x="CATEGORIA", title="Distribución por Categoría", text_auto=True)
        # st.plotly_chart(cat_fig)
        
        # st.subheader("📊 Comparación de registros por dependencia")
        # dep_fig = px.bar(df1, x="DEPENDENCIA", title="Distribución por Dependencia", text_auto=True)
        # st.plotly_chart(dep_fig)

        # Gráfico de comparativa de altas y bajas
        st.subheader("📊 Comparación de Altas y Bajas")
        comparacion_df = pd.DataFrame({"Tipo": ["Altas", "Bajas"], "Cantidad": [len(altas), len(bajas)]})
        fig = px.bar(comparacion_df, x="Tipo", y="Cantidad", title="Comparación de Altas y Bajas", text_auto=True, color="Tipo")
        st.plotly_chart(fig)
    
if __name__ == "__main__":
    main()
