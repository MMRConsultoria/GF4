import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import numbers
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import locale

# Configurações iniciais
st.set_page_config(page_title="Processador de Sangria", layout="centered")
st.title("📊 Processador de Sangria")

# Conexão com o Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
import json
credentials_dict = json.loads(st.secrets["GCP_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

# Abrindo a planilha correta
spreadsheet = gc.open("Tabela")
tabela_empresa = pd.DataFrame(spreadsheet.worksheet("Tabela Empresa").get_all_records())
tabela_descricoes = pd.DataFrame(spreadsheet.worksheet("Tabela Descrição Sangria").get_all_records())

uploaded_file = st.file_uploader("📥 Envie seu arquivo Excel (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_dados = pd.read_excel(xls, sheet_name="Sheet")

        st.subheader("Prévia dos dados da aba 'Sheet'")
        st.dataframe(df_dados.head())

        if st.button("🚀 Processar Sangria"):
            st.info("🔄 Processando arquivo...")

            # Processamento
            df = df_dados.copy()
            df["Data"] = pd.to_datetime(df["Data"]).dt.date
            df["Dia da Semana"] = pd.to_datetime(df["Data"]).dt.strftime("%A")

            # Traduz dia da semana
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

            # Formatação moeda
            df["Valor (R$)"] = pd.to_numeric(df["Valor (R$)"], errors="coerce")
            df["Valor (R$)"] = df["Valor (R$)"].fillna(0)

            # Agrupamento por descrição
            def agrupar_descricao(desc):
                for i, row in tabela_descricoes.iterrows():
                    chave = str(row["Chave"]).lower()
                    if chave in str(desc).lower():
                        return row["Descrição Agrupada"]
                return desc

            df["Descrição Base"] = df["Descrição"].apply(agrupar_descricao)

            # Mês e Ano
            df["Mês"] = pd.to_datetime(df["Data"]).dt.strftime("%B").str.capitalize()
            df["Ano"] = pd.to_datetime(df["Data"]).dt.year

            # Exportar para Excel
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
