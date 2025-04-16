import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from urllib.parse import quote

st.set_page_config(page_title="Processador de Sangria", layout="centered")
st.title("📊 Processador de Sangria")

# ID da planilha pública no Google Sheets
sheet_id = "13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q"

# Codifica nomes com espaços
sheet_empresa = quote("Tabela_Empresa")
sheet_descricoes = quote("Tabela_Descrição_ Sangria")  # com espaço mesmo

# Monta os links de exportação
tabela_empresa_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_empresa}"
tabela_descricoes_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_descricoes}"

# Lê as tabelas auxiliares públicas
tabela_empresa = pd.read_csv(tabela_empresa_url)
tabela_descricoes = pd.read_csv(tabela_descricoes_url)

# Upload do Excel
uploaded_file = st.file_uploader("📥 Envie seu arquivo Excel (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        st.write("📄 Abas encontradas:", xls.sheet_names)
        df_dados = pd.read_excel(xls, sheet_name=0)  # Lê a primeira aba automaticamente
        st.subheader("Prévia dos dados carregados")
        st.dataframe(df_dados.head())
        st.write("🧾 Colunas encontradas:", df_dados.columns.tolist())

        if st.button("🚀 Processar Sangria"):
            st.info("🔄 Processando...")

            df = df_dados.copy()

            # Tenta encontrar uma coluna que pareça ser 'Data'
            col_data = next((col for col in df.columns if "data" in col.lower()), None)
            col_descr = next((col for col in df.columns if "descr" in col.lower()), None)
            col_valor = next((col for col in df.columns if "valor" in col.lower()), None)

            if not all([col_data, col_descr, col_valor]):
                st.error("❌ Não foi possível identificar as colunas 'Data', 'Descrição' e 'Valor'.")
                st.stop()

            df[col_data] = pd.to_datetime(df[col_data], errors="coerce")
            df["Data"] = df[col_data].dt.date
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

            df["Valor (R$)"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0)
            df["Descrição"] = df[col_descr]

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

            st.success("✅ Arquivo processado com sucesso!")
            st.download_button(
                label="📥 Baixar arquivo processado",
                data=output,
                file_name="sangria_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")


