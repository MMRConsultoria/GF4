import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="📊 Processador de Sangria", layout="centered")
st.title("📊 Processador de Sangria")

# 🔐 Autenticação com Google Sheets via st.secrets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gspread"], scope)
gc = gspread.authorize(credentials)

# 🧾 Abre a planilha e aba
spreadsheet = gc.open("Nome da SUA Planilha")  # <-- Troque aqui pelo nome exato
tabela_empresa = spreadsheet.worksheet("Tabela Empresa")  # <-- Nome da aba com dados da empresa
df_empresa = pd.DataFrame(tabela_empresa.get_all_records())

# 📤 Upload do Excel do usuário
uploaded_file = st.file_uploader("Envie seu arquivo Excel (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        df_dados = pd.read_excel(uploaded_file, sheet_name="Sheet")
    except Exception as e:
        st.error(f"Erro ao ler o Excel: {e}")
    else:
        st.subheader("Prévia dos dados enviados")
        st.dataframe(df_dados.head())

        if st.button("Processar Sangria"):
            st.info("🔄 Processando arquivo...")

            # 📅 Converte coluna de data (assumindo que ela existe)
            df_dados["Data"] = pd.to_datetime(df_dados["Data"])
            df_dados["Dia da Semana"] = df_dados["Data"].dt.day_name(locale="pt_BR")

            # 🔗 Exemplo de merge com a Tabela Empresa
            if "Loja" in df_dados.columns and "Loja" in df_empresa.columns:
                df_final = pd.merge(df_dados, df_empresa, on="Loja", how="left")
            else:
                df_final = df_dados  # se não puder mesclar, segue com os dados originais

            # 💰 Formatações
            df_final["Data"] = df_final["Data"].dt.strftime("%d/%m/%Y")
            if "Valor" in df_final.columns:
                df_final["Valor (R$)"] = df_final["Valor"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # 📤 Download
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, sheet_name="Sangria", index=False)
            st.success("✅ Processamento concluído!")
            st.download_button("📥 Baixar resultado Excel", output.getvalue(), file_name="Sangria_Processada.xlsx")


     
