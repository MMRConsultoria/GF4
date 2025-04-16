import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from urllib.parse import quote

st.set_page_config(page_title="Processador de Sangria", layout="centered")
st.title("📊 Processador de Sangria")

# ID da sua planilha
sheet_id = "13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q"

# Nomes codificados das abas (sheets)
sheet_empresa = quote("Tabela_Empresa")
sheet_descricoes = quote("Tabela_Descrição_ Sangria")  # <- espaço depois do underline

# URLs formatadas
tabela_empresa_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_empresa}"
tabela_descricoes_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_descricoes}"

# Lê as tabelas auxiliares do Google Sheets público
tabela_empresa = pd.read_csv(tabela_empresa_url)
tabela_descricoes = pd.read_csv(tabela_descricoes_url)

# Upload do Excel do usuário
uploaded_file = st.file_uploader("📥 Envie seu arquivo Excel (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_dados = pd.read_excel(xls, sheet_name="Sheet")

        st.subheader("Prévia dos dados da aba 'Sheet'")
        st.dataframe(df_dados.head())

        if st.button("🚀 Processar Sangria"):
            st.info("🔄 Processando...")

            df = df_dados.copy()
            df["Data"] = pd.to_datetime(df["Data"]).dt.date
            df["Dia da Semana"] = pd.to_datetime(df["Data"]).dt.strftime("%A")

            dias_semana_pt = {
                "Monday": "segunda-feira",
                "Tuesday": "terça-feira",
                "Wednesday": "quarta-feira",
                "Thursday": "quinta-feira",
                "Friday": "sexta-feira",
                "Saturday": "sábado",
                "Sunday": "domingo"
            }
            df["Dia da Semana"] = df["Dia da Semana"].map(dias_semana_pt)

            df["Valor (R$)"] = pd.to_numeric(df["Valor (R$)"], errors="coerce").fillna(0)

            def agrupar_descricao(desc):
                for _, row in tabela_descricoes.iterrows():
                    chave = str(row["Chave"]).lower()
                    if chave in str(desc).lower():
                        return row["Descrição Agrupada"]
                return desc

            df["Descrição Base"] = df["Descrição"].apply(agrupar_descricao)
            df["Mês"] = pd.to_datetime(df["Data"]).dt.strftime("%B").str.capitalize()
            df["Ano"] = pd.to_datetime(df["Data"]).dt.year

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sangria')
                workbook = writer.book
                worksheet = writer.sheets['Sangria']
                for idx, col in enumerate(df.columns, 1):
                    worksheet.column_dimensions[chr(64 + idx)].width = 18
            output.seek(0)

            st.success("✅ Processamento finalizado!")
            st.download_button(
                label="📥 Baixar arquivo processado",
                data=output,
                file_name="sangria_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")

