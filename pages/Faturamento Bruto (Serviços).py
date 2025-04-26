import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# Função para gerar o Excel em memória
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    output.seek(0)
    return output

# Configurações iniciais do app
st.set_page_config(page_title="Relatório de Faturamento", layout="wide")
st.title("📊 Relatório de Faturamento")

# =============================
# Começo do app - Upload e processamento
# =============================

uploaded_file = st.file_uploader("📄 Envie seu arquivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        # (1) Carregar o arquivo
        df = pd.read_excel(uploaded_file)

        # (2) Processar para gerar o df_final
        # Aqui você faz seus tratamentos — exemplo:
        df_final = df.copy()  # (troque pelo seu processamento real)

        # =============================
        # Gerar o arquivo Excel para download
        # =============================

        excel_data = to_excel(df_final)
        st.download_button(
            label="📅 Baixar Relatório Excel",
            data=excel_data,
            file_name="faturamento_servico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # =============================
        # Botão para Atualizar Google Sheets
        # =============================

        st.markdown("---")  # Linha de separação
        st.subheader("🔄 Atualizar Google Sheets?")

        if st.button("📤 Atualizar tabela 'Fat Sistema Externo' no Google Sheets"):
            with st.spinner('🔄 Atualizando a planilha no Google Sheets...'):
                try:
                    # Conexão com Google Sheets
                    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                    credentials_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
                    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                    gc = gspread.authorize(credentials)

                    # Abre a planilha e aba
                    planilha = gc.open("Faturamento Sistema Externo")
                    aba = planilha.worksheet("Fat Sistema Externo")

                    # Achar a primeira linha vazia
                    valores_existentes = aba.get_all_values()
                    primeira_linha_vazia = len(valores_existentes) + 1

                    # Preparar os dados para colar
                    rows = df_final.values.tolist()

                    # Atualizar o Google Sheets
                    aba.update(f"A{primeira_linha_vazia}", rows)

                    st.success("✅ Dados atualizados com sucesso no Google Sheets!")

                except Exception as e:
                    st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
