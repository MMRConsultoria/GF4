import streamlit as st
import pandas as pd

st.set_page_config(page_title="Processar Metas Dinâmico", layout="wide")
st.title("📈 Processar Metas (mesclas eliminadas via ffill, 2 linhas abaixo do META)")

uploaded_file = st.file_uploader("📁 Escolha seu arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    todas_abas = xls.sheet_names

    abas_escolhidas = st.multiselect(
        "Selecione as abas a processar:",
        options=todas_abas,
        default=[]
    )

   
    st.success("✅ Dados consolidados só com META, ignorando Consolidado:")
    st.dataframe(df_final)

    if not df_final.empty:
        csv = df_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar CSV consolidado",
            data=csv,
            file_name="metas_consolidado.csv",
            mime='text/csv'
        )
else:
    st.info("💡 Faça o upload de um arquivo Excel para começar.")

