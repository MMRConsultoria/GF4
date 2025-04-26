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

# 🎨 Estilizar as abas
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] { margin-top: 20px; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }
    </style>
""", unsafe_allow_html=True)

# ================================
# 3. Separação em ABAS
# ================================
aba1, aba2, aba3 = st.tabs(["📄 Upload e Processamento", "📥 Download Excel", "🔄 Atualizar Google Sheets"])

# Variável para guardar o dataframe final
df_final = None
uploaded_file = None

# ================================
# 📄 Aba 1 - Upload e Processamento
# ================================
with aba1:
    st.header("📄 Upload e Processamento")

    uploaded_file = st.file_uploader("Envie o arquivo Excel com a aba 'FaturamentoDiarioPorLoja'", type=["xlsx"])

    if uploaded_file is None:
        st.info("📂 Envie um arquivo para iniciar o processamento.")
    else:
        st.success("✅ Arquivo enviado!")

        try:
            xls = pd.ExcelFile(uploaded_file)
            df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

            # Validação da célula B1
            texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
            if texto_b1 != "faturamento diário sintético multi-loja":
                st.error(f"❌ A célula B1 está com '{texto_b1}'. Corrija para 'Faturamento diário sintético multi-loja'.")
                st.stop()

            df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
            df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], dayfirst=True, errors='coerce')

            registros = []
            col = 3  # Começa na coluna D

            while col < df.shape[1]:
                nome_loja = str(df_raw.iloc[3, col]).strip()
                if re.match(r"^\d+\s*-?\s*", nome_loja):
                    nome_loja = nome_loja.split("-", 1)[-1].strip()

                    header_col = str(df.iloc[0, col]).strip().lower()
                    if "fat.total" in header_col:
                        for i in range(1, df.shape[0]):
                            linha = df.iloc[i]
                            valor_data = df.iloc[i, 2]
                            valor_check = str(df.iloc[i, 1]).strip().lower()

                            if pd.isna(valor_data) or valor_check in ["total", "subtotal"]:
                                continue

                            data = valor_data
                            valores = linha[col:col+5].values

                            if pd.isna(valores).all():
                                continue

                            registros.append([
                                data,
                                nome_loja,
                                *valores,
                                data.strftime("%b"),
                                data.year
                            ])
                    col += 5
                else:
                    col += 1

            if len(registros) == 0:
                st.warning("⚠️ Nenhum registro encontrado.")

            # Montar o df_final
            df_final = pd.DataFrame(registros, columns=[
                "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ])

            df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
            df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
            df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")

            empresas_nao_localizadas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()

            if len(empresas_nao_localizadas) > 0:
                st.warning(f"⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s):")
                for loja in empresas_nao_localizadas:
                    st.text(f"🔎 {loja}")
            else:
                st.success("✅ Todas as empresas foram localizadas!")

            dias_traducao = {
                "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
                "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
            }
            df_final.insert(1, "Dia da Semana", df_final["Data"].dt.day_name().map(dias_traducao))
            df_final["Data"] = df_final["Data"].dt.strftime("%d/%m/%Y")

            for col_val in ["Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas"]:
                df_final[col_val] = pd.to_numeric(df_final[col_val], errors="coerce").round(2)

            meses = {
                "jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"
            }
            df_final["Mês"] = df_final["Mês"].str.lower().map(meses)

            df_final["Data_Ordenada"] = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce')
            df_final = df_final.sort_values(by=["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")

            colunas_finais = [
                "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
                "Ticket", "Mês", "Ano"
            ]
            df_final = df_final[colunas_finais]

            st.success(f"✅ {len(df_final)} registros processados com sucesso!")

        except Exception as e:
            st.error(f"❌ Erro no processamento: {e}")

# ================================
# 📥 Aba 2 - Download Excel
# ================================
with aba2:
    st.header("📥 Download Relatório Excel")

    if uploaded_file and df_final is not None:
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
            output.seek(0)
            return output

        excel_data = to_excel(df_final)

        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=excel_data,
            file_name="faturamento_servico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo na aba anterior.")

# ================================
# 🔄 Aba 3 - Atualizar Google Sheets
# ================================
with aba3:
    st.header("🔄 Atualizar Google Sheets")

    if uploaded_file and df_final is not None:
        if 'atualizou_google' not in st.session_state:
            st.session_state.atualizou_google = False

        if not st.session_state.atualizou_google:
            if st.button("📤 Atualizar no Google Sheets"):
                with st.spinner('🔄 Atualizando...'):
                    try:
                        planilha_destino = gc.open("Faturamento Sistema Externo")
                        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                        dados_existentes = pd.DataFrame(aba_destino.get_all_records())

                        if not dados_existentes.empty:
                            for col in ["Data", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
                                if col in dados_existentes.columns:
                                    dados_existentes[col] = dados_existentes[col].astype(str).str.strip()

                        novos_dados = df_final.copy()
                        for col in ["Data", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
                            novos_dados[col] = novos_dados[col].astype(str).str.strip()

                        colunas_chave = ["Data", "Loja", "Código Everest", "Grupo", "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket", "Mês", "Ano"]

                        merged = pd.merge(novos_dados, dados_existentes, on=colunas_chave, how="left", indicator=True)
                        registros_novos = merged[merged["_merge"] == "left_only"].drop(columns="_merge")

                        if registros_novos.empty:
                            st.info("✅ Nenhum novo registro para atualizar.")
                            st.session_state.atualizou_google = True
                        else:
                            rows = registros_novos.fillna("").values.tolist()
                            primeira_linha_vazia = len(dados_existentes) + 2
                            aba_destino.update(f"A{primeira_linha_vazia}", rows)

                            st.success(f"✅ {len(rows)} novo(s) registro(s) atualizado(s) no Google Sheets!")
                            st.session_state.atualizou_google = True

                            registros_ignorados = len(novos_dados) - len(rows)
                            if registros_ignorados > 0:
                                st.warning(f"⚠️ {registros_ignorados} registro(s) já existiam e foram ignorados.")

                    except Exception as e:
                        st.error(f"❌ Erro ao atualizar: {e}")
                        st.session_state.atualizou_google = False
        else:
            st.info("✅ Dados já foram atualizados nesta sessão.")
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo.")
