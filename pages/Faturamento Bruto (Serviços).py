# pages/FaturamentoServico.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela_Empresa").get_all_records())

# ================================
# 2. Configuração inicial do app
# ================================
st.set_page_config(page_title="Faturamento por Serviço", layout="wide")
st.title("📋 Relatório de Faturamento por Serviço")

# 🎨 BÔNUS: Estilizar as abas para ficarem mais bonitas
st.markdown("""
    <style>
    .stApp {
        background-color: #f9f9f9;
    }
    div[data-baseweb="tab-list"] {
        margin-top: 20px;
    }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover {
        background-color: #dce0ea;
        color: black;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0366d6;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# ================================
# 3. Separação em ABAS
# ================================
aba1, aba2, aba3 = st.tabs(["📄 Upload e Processamento", "📥 Download Excel", "🔄 Atualizar Google Sheets"])

# Variável para guardar o dataframe final
df_final = None
