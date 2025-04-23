# pages/FaturamentoServico.py (reescrito com nova lógica baseada na sua descrição)

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re

st.set_page_config(page_title="Faturamento por Serviço", layout="wide")
st.title("📋 Relatório de Faturamento por Serviço")

uploaded_file = st.file_uploader("Envie o arquivo Excel com a aba 'FaturamentoDiarioPorLoja'", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

        # Validação da célula B1
        texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
        if texto_b1 != "faturamento diário sintético multi-loja":
            st.error("❌ ERRO: A célula B1 deve conter 'Faturamento diário sintético multi-loja'. Verifique o arquivo.")
            st.stop()

        df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
        linha_lojas = df_raw.iloc[4, 4:].dropna()

        registros = []
        col = 4
        while col < df.shape[1]:
            nome_loja = str(df_raw.iloc[4, col]).strip()
            if re.match(r"^\d+\s*-", nome_loja):
                nome_loja = nome_loja.split("-", 1)[-1].strip()

                header_col = str(df.iloc[0, col]).strip().lower()
                st.write(f"Coluna {col}: header -> {header_col}")  # DEBUG
                if "fat.total" in header_col:
                    for i in range(1, df.shape[0]):
                        linha = df.iloc[i]
                        valor_data = str(df.iloc[i, 2]).strip().lower()
                        valor_check = str(df.iloc[i, 1]).strip().lower()

                        if valor_data in ["subtotal"] or valor_check in ["total"]:
                            continue

                        try:
                            data = pd.to_datetime(valor_data, dayfirst=True)
                        except:
                            continue

                        valores = linha[col:col+5].values
                        if pd.isna(valores).all():
                            continue

                        registros.append([
                            data,
                            nome_loja,
                            *valores,
                            data.strftime("%b"),
                            data.year
                        ])

                col += 5  # pula para o próximo bloco
            else:
                col += 1

        df_final = pd.DataFrame(registros, columns=[
            "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
        ])

        dias_traducao = {
            "Monday": "segunda-feira",
            "Tuesday": "terça-feira",
            "Wednesday": "quarta-feira",
            "Thursday": "quinta-feira",
            "Friday": "sexta-feira",
            "Saturday": "sábado",
            "Sunday": "domingo"
        }

        df_final["Data"] = pd.to_datetime(df_final["Data"], errors='coerce')
        df_final.insert(1, "Dia da Semana", df_final["Data"].dt.day_name().map(dias_traducao))

        st.success("✅ Relatório processado com sucesso!")
        st.dataframe(df_final.head(50))

        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
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
