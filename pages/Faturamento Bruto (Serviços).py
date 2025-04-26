import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# Função para gerar Excel em memória
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
        writer.save()
    output.seek(0)
    return output

# Começo do app
st.set_page_config(page_title="Relatório de Faturamento", layout="wide")
st.title("📊 Relatório de Faturamento")

# Aqui você deve carregar/processar seu arquivo e gerar o df_final
# Exemplo de simulação (você já deve ter seu df_final real antes desse ponto):
# df_final = seu_dataframe_processado

# ---------------------------
# Bloco principal
# ---------------------------
try:
    # Processo de geração do arquivo para download
    excel_data = to_excel(df_final)
    st.download_button(
        label="📅 Baixar Relatório Excel",
        data=excel_data,
        file_name="faturamento_servico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as e:
    st.error(f"Erro ao processar o arquivo: {e}")

# ---------------------------
# Bloco adicional: Atualizar Google Sheets
# ---------------------------

st.markdown("---")  # Linha para separar visualmente

st.subheader("🔄 Atualizar Google Sheets?")

# Botão para atualizar o Google Sheets
if st.button("📤 Atualizar tabela 'Fat Sistema Externo' no Google Sheets"):
    with st.spinner('🔄 Atualizando a planilha no Google Sheets...'):
        try:
            # Configurar conexão com o Google Sheets
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            credentials_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
            credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
            gc = gspread.authorize(credentials)

            # Abrir a planilha e a aba
            planilha = gc.open("Faturamento Sistema Externo")
            aba = planilha.worksheet("Fat Sistema Externo")

            # Descobrir a próxima linha vazia
            valores_existentes = aba.get_all_values()
            primeira_linha_vazia = len(valores_existentes) + 1

            # Preparar os dados do df_final
            rows = df_final.values.tolist()

            # Atualizar a planilha a partir da linha vazia
            aba.update(f"A{primeira_linha_vazia}", rows)

            st.success("✅ Dados atualizados com sucesso no Google Sheets!")

        except Exception as e:
            st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")
