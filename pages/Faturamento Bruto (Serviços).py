import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(page_title="Relatório de Faturamento", layout="wide")
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento</h1>
    </div>
""", unsafe_allow_html=True)


# Conexão com Google Sheets via secrets (correção: usar from_json_keyfile_dict)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

#json da Google Shett consultoriaGF4
import json
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)



gc = gspread.authorize(credentials)
planilha = gc.open("Tabela")

df_empresa = pd.DataFrame(planilha.worksheet("Tabela_Empresa").get_all_records())






uploaded_file = st.file_uploader(
    label="📁 Envie o arquivo de faturamento",
    type=["xlsx", "xlsm"],
    help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])#pega sempre a primeira planilha
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo enviado: {e}")
    else:
        df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
        df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
        df = pd.merge(df, df_empresa, on="Loja", how="left")

        # Dia da Semana, Mês e Ano
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
        dias_semana = {
            0: 'segunda-feira', 1: 'terça-feira', 2: 'quarta-feira',
            3: 'quinta-feira', 4: 'sexta-feira', 5: 'sábado', 6: 'domingo'
        }
        df["Dia da Semana"] = df["Data"].dt.dayofweek.map(dias_semana)
        df["Mês"] = df["Data"].dt.month.map({
            1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
            7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
        })
        df["Ano"] = df["Data"].dt.year
        df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

        # Reordenar colunas
        colunas_ordenadas = [
            "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
            "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket", "Mês", "Ano"
        ]
        df = df[colunas_ordenadas]

        df = df.sort_values(by=["Data", "Loja"])

        periodo_min = pd.to_datetime(df["Data"], dayfirst=True).min().strftime("%d/%m/%Y")
        periodo_max = pd.to_datetime(df["Data"], dayfirst=True).max().strftime("%d/%m/%Y")

        col1, col2 = st.columns(2)
        col1.metric("📅 Período processado", f"{periodo_min} até {periodo_max}")
        col2.metric("📈 Total de linhas", f"{len(df):,}".replace(",", "."))

        st.success("✅ Relatório gerado com sucesso!")

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Faturamento")
        output.seek(0)

        st.download_button("📥 Baixar relatório de faturamento", data=output, file_name="Faturamento.xlsx")
