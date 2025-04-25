import streamlit as st
import pandas as pd
import numpy as np
import re
import json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Relatório de Faturamento", layout="wide")

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento</h1>
    </div>
""", unsafe_allow_html=True)

# Conexão com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha.worksheet("Tabela_Empresa").get_all_records())

# Upload do arquivo Excel
uploaded_file = st.file_uploader(
    label="📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento",
    type=["xlsx", "xlsm"],
    help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_raw = pd.read_excel(xls, sheet_name="FaturamentoPorMeioDePagamento", header=None)
    except Exception as e:
        st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes: {e}")
    else:
        if str(df_raw.iloc[0, 1]).strip().lower() != "faturamento diário por meio de pagamento":
            st.error("❌ A célula B1 deve conter 'Faturamento diário por meio de pagamento'.")
            st.stop()

        linha_inicio_dados = 6
        blocos = []

        for col in range(3, df_raw.shape[1]):
            loja_nome = df_raw.iloc[3, col]
            meio_pagamento = df_raw.iloc[4, col]

            if pd.isna(loja_nome) or pd.isna(meio_pagamento):
                continue

            if any(palavra in str(loja_nome).lower() for palavra in ["total", "serv", "real"]) or \
               any(palavra in str(meio_pagamento).lower() for palavra in ["total", "serv", "real"]):
                continue

            df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
            df_temp.columns = ["Data", "Valor (R$)"]
            df_temp.insert(1, "Meio de Pagamento", meio_pagamento)
            df_temp.insert(2, "Loja", loja_nome)
            blocos.append(df_temp)

        if not blocos:
            st.error("❌ Nenhum dado válido encontrado na planilha.")
        else:
            df = pd.concat(blocos, ignore_index=True)
            df =
