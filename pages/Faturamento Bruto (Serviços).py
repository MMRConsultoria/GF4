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

# 🎨 Estilizar abas
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

# ================================
# 📋 Aba 1 - Upload e Processamento (com ícone no título)
# ================================

with aba1:
    st.markdown("""
        <div style='display: flex; align-items: center; gap: 10px;'>
            <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
            <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento por Serviço</h1>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # 🔹 Upload do Arquivo
    uploaded_file = st.file_uploader("📄 Envie o arquivo Excel com a aba 'FaturamentoDiarioPorLoja'", type=["xlsx"])

    if uploaded_file:
        # 🔹 Resetar atualização Google Sheets
        st.session_state.atualizou_google = False

        try:
            xls = pd.ExcelFile(uploaded_file)
            df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

            # Validação da célula B1
            texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
            if texto_b1 != "faturamento diário sintético multi-loja":
                st.error("❌ ERRO: A célula B1 deve conter 'Faturamento diário sintético multi-loja'. Verifique o arquivo.")
                st.stop()

            df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)

            registros = []
            col = 3  # Começa na coluna D (índice 3)

            while col < df.shape[1]:
                nome_loja = str(df_raw.iloc[3, col]).strip()
                if re.match(r"^\d+\s*-?\s*", nome_loja):
                    nome_loja = nome_loja.split("-", 1)[-1].strip()

                    header_col = str(df.iloc[0, col]).strip().lower()
                    if "fat.total" in header_col:
                        for i in range(1, df.shape[0]):
                            linha = df.iloc[i]
                            valor_data = str(df.iloc[i, 2]).strip().lower()
                            valor_check = str(df.iloc[i, 1]).strip().lower()

                            if valor_data in ["subtotal"] or valor_check in ["total"]:
                                continue

                            try:
                                data = pd.to_datetime(valor_data, dayfirst=True)
                            except:
                                continue

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

            # Montar o df_final
            df_final = pd.DataFrame(registros, columns=[
                "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ])

            # Ajustes de nomes e merge
            df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
            df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
            df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")

            # Guardar no session_state para as próximas abas
            st.session_state.df_final = df_final
            st.session_state.atualizou_google = False

            st.success("✅ Arquivo processado com sucesso!")

        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {e}")


# ================================
# 📥 Aba 2 - Download Excel
# ================================
with aba2:
    st.header("📥 Download Relatório Excel")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final

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
# 🔄 Aba 3 - Atualizar Google Sheets (versão final corrigida, com tratamento de NaN e Ano)
# ================================
import math

with aba3:
    st.header("🔄 Atualizar Google Sheets")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final

        if 'atualizou_google' not in st.session_state:
            st.session_state.atualizou_google = False

        if not st.session_state.atualizou_google:
            if st.button("📤 Atualizar no Google Sheets"):
                with st.spinner('🔄 Atualizando...'):
                    try:
                        # Abrir a planilha e aba de destino
                        planilha_destino = gc.open("Faturamento Sistema Externo")
                        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                        # Ler dados existentes
                        dados_raw = aba_destino.get_all_values()

                        # Preparar dados existentes (linhas sem cabeçalho)
                        dados_existentes = [ [str(cell).strip() for cell in row] for row in dados_raw[1:] ]

                        # Preparar novos dados
                        novos_dados_raw = df_final.values.tolist()

                        novos_dados = []
                        for linha in novos_dados_raw:
                            nova_linha = []
                            for idx, valor in enumerate(linha):
                                if idx in [6, 7, 8, 9]:  # Fat.Total, Serv/Tx, Fat.Real, Ticket
                                    if isinstance(valor, (int, float)) and not math.isnan(valor):
                                        valor = round(valor, 2)  # número real
                                    else:
                                        valor = ""  # vazio se NaN
                                elif idx in [3, 5, 11]:  # Código Everest, Código Grupo Everest, Ano
                                    if isinstance(valor, (int, float)) and not math.isnan(valor):
                                        valor = int(valor)  # número inteiro
                                    else:
                                        valor = ""  # vazio se NaN
                                else:
                                    valor = str(valor).strip()
                                nova_linha.append(valor)
                            novos_dados.append(nova_linha)

                        # Verificar novos registros
                        registros_novos = [linha for linha in novos_dados if linha not in dados_existentes]

                        total_novos = len(registros_novos)
                        total_existentes = len(novos_dados) - total_novos

                        if total_novos == 0:
                            st.info(f"✅ Nenhum novo registro para atualizar. {total_existentes} registro(s) já existiam no Google Sheets.")
                            st.session_state.atualizou_google = True
                        else:
                            # Descobrir onde colar
                            primeira_linha_vazia = len(dados_raw) + 1  # linha após os dados

                            # Atualizar
                            aba_destino.update(f"A{primeira_linha_vazia}", registros_novos)

                            st.success(f"✅ {total_novos} novo(s) registro(s) enviado(s) para o Google Sheets!")
                            if total_existentes > 0:
                                st.warning(f"⚠️ {total_existentes} registro(s) já existiam e não foram importados.")
                            st.session_state.atualizou_google = True

                    except Exception as e:
                        st.error(f"❌ Erro ao atualizar: {e}")
                        st.session_state.atualizou_google = False
        else:
            st.info("✅ Dados já foram atualizados nesta sessão.")
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo na aba anterior.")
