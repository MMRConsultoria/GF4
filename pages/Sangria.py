# pages/Sangria.py

import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Relatório de Sangria", layout="wide")
st.title("💵 Relatório de Sangria")

uploaded_file = st.file_uploader("Envie seu arquivo Excel (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_dados = pd.read_excel(xls, sheet_name="Sheet")

        # Processamento básico: adicionando colunas Data, Mês, Ano, Dia da Semana
        df_dados["Data"] = pd.to_datetime(df_dados["Data"], errors="coerce")
        df_dados = df_dados.dropna(subset=["Data"]).copy()
        df_dados["Mês"] = df_dados["Data"].dt.strftime("%b")
        df_dados["Ano"] = df_dados["Data"].dt.year

        dias_traducao = {
            "Monday": "segunda-feira",
            "Tuesday": "terça-feira",
            "Wednesday": "quarta-feira",
            "Thursday": "quinta-feira",
            "Friday": "sexta-feira",
            "Saturday": "sábado",
            "Sunday": "domingo"
        }
        df_dados["Dia da Semana"] = df_dados["Data"].dt.day_name().map(dias_traducao)

        # Conectar à Tabela Empresa no Google Sheets
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open("Tabela").worksheet("Tabela Empresa")
        tabela_empresa = pd.DataFrame(sheet.get_all_records())

        df_dados = df_dados.merge(tabela_empresa, left_on="Loja", right_on="Empresa", how="left")

        # Reorganizar colunas para exibição
        colunas_exibir = [
            "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
            "Funcionário", "Descrição", "Valor (R$)", "Mês", "Ano"
        ]
        colunas_existentes = [col for col in colunas_exibir if col in df_dados.columns]
        df_resultado = df_dados[colunas_existentes]

        st.success("✅ Arquivo processado com sucesso!")
        st.dataframe(df_resultado.head(50))

        # Exportar para Excel
        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sangria')
            writer.close()
            output.seek(0)
            return output

        excel_data = to_excel(df_resultado)
        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=excel_data,
            file_name="relatorio_sangria.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
