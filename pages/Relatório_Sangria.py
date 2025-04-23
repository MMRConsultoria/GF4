import streamlit as st
import pandas as pd
import numpy as np
import json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Relatório de Sangria", layout="wide")
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/clipboard-list.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Sangria</h1>
    </div>
""", unsafe_allow_html=True)

# Conexão com Google Sheets via secrets (correção: usar from_json_keyfile_dict)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Tabela")

df_empresa = pd.DataFrame(planilha.worksheet("Tabela_Empresa").get_all_records())
df_descricoes = pd.DataFrame(
    planilha.worksheet("Tabela_Descrição_Sangria").get_all_values(),
    columns=["Palavra-chave", "Descrição Agrupada"]
)

uploaded_file = st.file_uploader(
    label="📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de sangria",
    type=["xlsx", "xlsm"],
    help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_dados = pd.read_excel(xls, sheet_name="Sheet")
    except Exception as e:
        st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes: {e}")
    else:
        df = df_dados.copy()
        df["Loja"] = np.nan
        df["Data"] = np.nan
        df["Funcionário"] = np.nan

        data_atual = None
        funcionario_atual = None
        loja_atual = None
        linhas_validas = []

        for i, row in df.iterrows():
            valor = str(row["Hora"]).strip()
            if valor.startswith("Loja:"):
                loja = valor.split("Loja:")[1].split("(Total")[0].strip()
                if "-" in loja:
                    loja = loja.split("-", 1)[1].strip()
                loja_atual = loja or "Loja não cadastrada"
            elif valor.startswith("Data:"):
                try:
                    data_atual = pd.to_datetime(valor.split("Data:")[1].split("(Total")[0].strip(), dayfirst=True)
                except:
                    data_atual = pd.NaT
            elif valor.startswith("Funcionário:"):
                funcionario_atual = valor.split("Funcionário:")[1].split("(Total")[0].strip()
            else:
                if pd.notna(row["Valor(R$)"]) and pd.notna(row["Hora"]):
                    df.at[i, "Data"] = data_atual
                    df.at[i, "Funcionário"] = funcionario_atual
                    df.at[i, "Loja"] = loja_atual
                    linhas_validas.append(i)

        df = df.loc[linhas_validas].copy()
        df.ffill(inplace=True)

        df["Descrição"] = df["Descrição"].astype(str).str.strip().str.lower()
        df["Funcionário"] = df["Funcionário"].astype(str).str.strip()
        df["Valor(R$)"] = pd.to_numeric(df["Valor(R$)"], errors="coerce")

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

        df = pd.merge(df, df_empresa, on="Loja", how="left")

        def mapear_descricao(desc):
            desc_lower = str(desc).lower()
            for _, row in df_descricoes.iterrows():
                if str(row["Palavra-chave"]).lower() in desc_lower:
                    return row["Descrição Agrupada"]
            return "Outros"

        df["Descrição Agrupada"] = df["Descrição"].apply(mapear_descricao)

        df = df.sort_values(by=["Data", "Loja"])

        periodo_min = pd.to_datetime(df["Data"], dayfirst=True).min().strftime("%d/%m/%Y")
        periodo_max = pd.to_datetime(df["Data"], dayfirst=True).max().strftime("%d/%m/%Y")
        valor_total = df["Valor(R$)"].sum()

        col1, col2 = st.columns(2)
        col1.metric("📅 Período processado", f"{periodo_min} até {periodo_max}")
        col2.metric("💰 Valor total de sangria", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.success("✅ Relatório gerado com sucesso!")

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Sangria")
        output.seek(0)

        st.download_button("📥 Baixar relatório de sangria", data=output, file_name="Sangria_estruturada.xlsx")
