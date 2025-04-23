import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Faturamento por Serviço", layout="wide")
st.title("📋 Relatório de Faturamento por Serviço")

uploaded_file = st.file_uploader("Envie o arquivo Excel com a aba 'FaturamentoDiarioPorLoja'", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

        if str(df_raw.iloc[0, 1]).strip().lower() != "faturamento diário sintético multi-loja":
            st.error("❌ A célula B1 não contém o texto esperado: 'faturamento diário sintético multi-loja'")
        else:
            df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
            lojas = df_raw.iloc[3, 4:].dropna()

            registros = []
            for idx_loja, nome_loja in lojas.items():
                if "totais" in str(nome_loja).lower() or "#" not in str(nome_loja):
                    continue

                base_col = idx_loja
                if str(df.iloc[0, base_col]) != "Fat.Total":
                    continue

                for i in range(2, df.shape[0]):
                    linha = df.iloc[i]
                    data_texto = str(df.iloc[i, 2]).strip()
                    if data_texto.lower() == "subtotal" or str(df.iloc[i, 1]).lower() == "total":
                        continue

                    try:
                        data = pd.to_datetime(data_texto, dayfirst=True)
                    except:
                        continue

                    valores = linha[base_col:base_col+5].values
                    if pd.isna(valores).all():
                        continue

                    registros.append([
                        data,
                        nome_loja,
                        *valores,
                        data.strftime("%b"),
                        data.year
                    ])

            df_final = pd.DataFrame(registros, columns=[
                "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ])

            df_final.insert(1, "Dia da Semana", df_final["Data"].dt.day_name(locale='pt_BR'))

            # Conectar com a aba "Tabela" do Google Sheets (igual sangria)
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
            client = gspread.authorize(creds)
            sheet = client.open("Tabela").worksheet("Tabela Empresa")
            tabela_empresa = pd.DataFrame(sheet.get_all_records())

            df_final = df_final.merge(tabela_empresa, left_on="Loja", right_on="Empresa", how="left")

            df_final = df_final[[
                "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ]]

            st.success("✅ Relatório processado com sucesso!")
            st.dataframe(df_final.head(50))

            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
                writer.close()
                output.seek(0)
                return output

            excel_data = to_excel(df_final)
            st.download_button(
                label="📥 Baixar Relatório Excel",
                data=excel_data,
                file_name="faturamento_servico.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
