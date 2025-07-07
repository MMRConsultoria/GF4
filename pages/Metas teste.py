import streamlit as st
import pandas as pd

st.set_page_config(page_title="Escolher Planilha", layout="wide")

st.title("📊 Importar planilha Excel")

# Upload do arquivo Excel
uploaded_file = st.file_uploader(
    "📁 Clique para selecionar ou arraste aqui seu arquivo Excel (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is not None:
    # Lê todas as abas
    xls = pd.ExcelFile(uploaded_file)
    abas = xls.sheet_names
    aba_escolhida = st.selectbox("Escolha a aba:", abas)

    # Carrega a aba escolhida
    df = pd.read_excel(xls, sheet_name=aba_escolhida)

    st.success(f"Arquivo carregado com sucesso ✅")
    st.write(df)
else:
    st.info("💡 Por favor, carregue um arquivo Excel para começar.")

