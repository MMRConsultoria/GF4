import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# ================================
# Configuração da página Streamlit
# ================================
st.set_page_config(page_title="Importador de Metas", layout="wide")
st.title("📊 Importador de Metas para Google Sheets")

# ================================
# 1. Conexão Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# 🔥 Aqui pega as credenciais do secrets ou local
try:
    credentials_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
except:
    import json
    with open("seu_arquivo_credenciais.json") as f:
        credentials_dict = json.load(f)

credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

# Abre a sua planilha pelo ID do link
planilha = gc.open_by_key("1ZaRXVZyv7WZ8xJ8yGEViRibZ-sGoilGO")
aba_metas = planilha.worksheet("Metas 1")

# ================================
# 2. Upload do arquivo
# ================================
uploaded_file = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("📥 Arquivo carregado com sucesso!")

    # ================================
    # 3. Consolidar abas
    # ================================
    xls = pd.ExcelFile(uploaded_file)
    st.write("✅ Abas encontradas:", xls.sheet_names)

    dfs = []
    for aba in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=aba)
        df["Grupo"] = aba
        dfs.append(df)

    df_total = pd.concat(dfs, ignore_index=True)
    st.write("📈 Consolidado:", df_total.head())

   
