import streamlit as st
import pandas as pd

st.set_page_config(page_title="Importar Excel Multiplas Abas", layout="wide")
st.title("📊 Importar planilha Excel - escolha as abas")

uploaded_file = st.file_uploader(
    "📁 Clique para selecionar ou arraste aqui seu arquivo Excel (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is not None:
    # Carrega todas as abas
    xls = pd.ExcelFile(uploaded_file)
    todas_abas = xls.sheet_names

    st.write(f"🗂 O arquivo possui as seguintes abas: {todas_abas}")

    # Usuário escolhe quais quer carregar
    abas_escolhidas = st.multiselect(
        "Selecione as abas que deseja importar:",
        options=todas_abas,
        default=[]
    )

    # Para cada aba escolhida, mostra o DataFrame
    for aba in abas_escolhidas:
        st.subheader(f"📄 Dados da aba: {aba}")
        df = pd.read_excel(xls, sheet_name=aba)
        st.dataframe(df)

    if not abas_escolhidas:
        st.info("💡 Selecione ao menos uma aba para exibir os dados.")
else:
    st.info("🚀 Carregue um arquivo Excel para começar.")
