import streamlit as st
import pandas as pd
import numpy as np
import json
import gspread
from io import BytesIO
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Relatório de Faturamento", layout="wide")
st.title("📊 Relatório de Faturamento - Multi-loja")

# Autenticação Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha.worksheet("Tabela_Empresa").get_all_records())

uploaded_file = st.file_uploader("📁 Envie o arquivo de faturamento", type=["xlsx", "xlsm"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

    # Colunas fixas: A, B, C (índices 0, 1, 2)
    col_fixas = list(range(3))
    linha_inicio_dados = 4 - 1  # Python começa do 0

    blocos = []
    col = 3
    while col < df_raw.shape[1]:
        if pd.notna(df_raw.iloc[1, col]):
            nome_loja = df_raw.iloc[1, col]
            bloco_cols = col_fixas + [col + i for i in range(5)]
            df_bloco = df_raw.iloc[linha_inicio_dados:, bloco_cols].copy()
            df_bloco.columns = ["Data", "Código", "Grupo", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]
            df_bloco.insert(2, "Loja", nome_loja)
            blocos.append(df_bloco)
            col += 5
        else:
            break

    df = pd.concat(blocos, ignore_index=True)
    df.dropna(subset=["Data"], inplace=True)

    # Converter datas
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Data"].notna()]

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

    # Merge com dados da empresa
    df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
    df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
    df = pd.merge(df, df_empresa, on="Loja", how="left")

    # Ordenar colunas finais
    colunas_finais = [
        "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
        "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
        "Ticket", "Mês", "Ano"
    ]
    df = df[colunas_finais]
    df = df.sort_values(by=["Data", "Loja"])

    periodo_min = df["Data"].min()
    periodo_max = df["Data"].max()

    col1, col2 = st.columns(2)
    col1.metric("📅 Período processado", f"{periodo_min} até {periodo_max}")
    col2.metric("🏪 Total de lojas", df["Loja"].nunique())

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Faturamento")
    output.seek(0)

    st.success("✅ Relatório de faturamento gerado com sucesso!")
    st.download_button("📥 Baixar relatório", data=output, file_name="Faturamento_transformado.xlsx")
