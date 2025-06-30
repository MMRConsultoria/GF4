import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import json
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Faturamento por Meio de Pagamento", layout="wide")

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# ================================
# 2. Upload do arquivo Excel
# ================================
st.title("📈 Faturamento por Meio de Pagamento")
uploaded_file = st.file_uploader("Carregue o Excel (FaturamentoPorMeioPgto)", type=["xls", "xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    if "FaturamentoPorMeioPgto" not in xls.sheet_names:
        st.error("❌ A aba 'FaturamentoPorMeioPgto' não foi encontrada.")
        st.stop()

    df_raw = pd.read_excel(xls, sheet_name="FaturamentoPorMeioPgto")
    df_raw.columns = df_raw.columns.str.strip()

    # 🔧 Ajuste da Data
    if "Data" in df_raw.columns:
        df_raw["Data"] = pd.to_datetime(df_raw["Data"], dayfirst=True, errors="coerce")
        ultima_data = df_raw["Data"].dropna().max()
        st.markdown(f"📅 Última data carregada: **{ultima_data.strftime('%d/%m/%Y')}**" if pd.notnull(ultima_data) else "⚠️ Nenhuma data válida encontrada")
    else:
        st.error("⚠️ Coluna 'Data' não encontrada.")
        st.stop()

    # 🔗 Merge com Tabela Empresa
    df_raw["Loja"] = df_raw["Loja"].astype(str).str.strip().str.lower()
    df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
    df_final = pd.merge(df_raw, df_empresa, on="Loja", how="left")

    # ✅ Criar coluna K = Data + Loja + Meio + Valor
    df_final["K"] = df_final["Data"].dt.strftime("%Y-%m-%d") + \
                    df_final["Loja"].astype(str) + \
                    df_final["Meio"].astype(str) + \
                    df_final["Valor"].astype(str)

    # 🔍 Preview
    st.session_state.df_final_meio = df_final
    st.success("✅ Dados carregados e coluna K criada para deduplicação.")
    st.dataframe(df_final.head(), use_container_width=True)

    # 📥 Download local
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Faturamento Meio')
        output.seek(0)
        return output

    st.download_button(
        label="📥 Baixar Excel Gerado",
        data=to_excel(df_final),
        file_name="faturamento_meio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ================================
# 3. Atualizar Google Sheets - ABA correta
# ================================
st.header("📤 Atualizar Google Sheets - Faturamento Meio Pagamento")

if "df_final_meio" in st.session_state:
    df_final = st.session_state.df_final_meio.copy()

    aba_destino = planilha_empresa.worksheet("Faturamento Meio Pagamento")
    dados_existentes = aba_destino.get_all_values()

    # Criar conjunto de chaves existentes pela coluna K (11ª coluna)
    set_existentes = set([linha[10] for linha in dados_existentes[1:] if len(linha) > 10])

    novos_dados = []
    duplicados = []
    for linha in df_final.fillna("").values.tolist():
        chave_k = linha[10]  # coluna K
        if chave_k not in set_existentes:
            novos_dados.append(linha)
            set_existentes.add(chave_k)
        else:
            duplicados.append(linha)

    if st.button("📥 Enviar dados ao Google Sheets"):
        with st.spinner("🔄 Enviando dados para o Google Sheets..."):
            if novos_dados:
                inicio = len(dados_existentes) + 1
                aba_destino.update(f"A{inicio}", novos_dados)
                st.success(f"✅ {len(novos_dados)} registro(s) enviados com sucesso!")
            else:
                st.info("✅ Nenhum novo registro para enviar (tudo já está no Google Sheets).")
            if duplicados:
                st.warning(f"⚠️ {len(duplicados)} registro(s) duplicados não foram enviados.")
else:
    st.info("⚠️ Primeiro faça o upload do arquivo acima para preparar os dados.")
