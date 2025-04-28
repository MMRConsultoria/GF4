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
#st.title("📋 Relatório de Faturamento por Serviço")

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


# Cabeçalho bonito (depois do estilo)
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento por Serviço</h1>
    </div>
""", unsafe_allow_html=True)


# ================================
# 3. Separação em ABAS
# ================================
aba1, aba2, aba3 = st.tabs(["📄 Upload e Processamento", "📥 Download Excel", "🔄 Atualizar Google Sheets"])

# ================================
# 📄 Aba 1 - Upload e Processamento
# ================================
with aba1:
    uploaded_file = st.file_uploader("📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento", type=["xlsx"])

    if uploaded_file:
        try:
            # 🔹 Carregar o arquivo
            xls = pd.ExcelFile(uploaded_file)
            df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

            # 🔹 Validar B1
            texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
            if texto_b1 != "faturamento diário sintético multi-loja":
                st.error(f"❌ A célula B1 está com '{texto_b1}'. Corrija para 'Faturamento diário sintético multi-loja'.")
                st.stop()

            df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
            df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], dayfirst=True, errors='coerce')

            # 🔹 Processamento dos registros
            registros = []
            col = 3
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

            # 🔹 Montar df_final
            df_final = pd.DataFrame(registros, columns=[
                "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ])

            df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
            df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
            df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")

            # 🔹 Ajustar dados
            dias_traducao = {
                "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
                "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
            }
            df_final.insert(1, "Dia da Semana", pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.day_name().map(dias_traducao))
            df_final["Data"] = pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.strftime("%d/%m/%Y")

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

            st.session_state.df_final = df_final
            st.session_state.atualizou_google = False

            # 🔥 Agora exibir:
            # 📄 Nome do Arquivo
            st.markdown(f"""
                <div style='font-size:15px; font-weight: bold; margin-bottom:10px;'>
                    📄 Arquivo selecionado: {uploaded_file.name}
                </div>
            """, unsafe_allow_html=True)

            # 📅 e 💰 Período e Valor Total
            datas_validas = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce').dropna()

            if not datas_validas.empty:
                data_inicial = datas_validas.min().strftime("%d/%m/%Y")
                data_final = datas_validas.max().strftime("%d/%m/%Y")
                
                valor_total = df_final["Fat.Total"].sum().round(2)
                valor_total_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown(f"""
                        <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>📅 Período processado</div>
                        <div style='font-size:30px; color:#000;'>{data_inicial} até {data_final}</div>
                    """, unsafe_allow_html=True)

                with col2:
                    st.markdown(f"""
                        <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>💰 Valor total</div>
                        <div style='font-size:30px; color:green;'>{valor_total_formatado}</div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ Não foi possível identificar o período de datas.")

           # 🔎 Empresas não localizadas
            empresas_nao_localizadas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()

            if len(empresas_nao_localizadas) > 0:
                # Construir a mensagem com o link direto
                mensagem = f"""
                ⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s): 
                ✏️ Atualize a tabela clicando 
                <a href='https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link' target='_blank'><strong>aqui</strong></a>.
                """

                st.markdown(mensagem, unsafe_allow_html=True)


            else:
                st.success("✅ Todas as empresas foram localizadas na Tabela_Empresa!")

            # 🔗 Links úteis
#            st.markdown("""
#🔗 [Clique aqui para abrir a **Tabela_Empresa**](https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link)

#🔗 [Clique aqui para abrir o **Faturamento Sistema Externo**](https://docs.google.com/spreadsheets/d/1_3uX7dlvKefaGDBUhWhyDSLbfXzAsw8bKRVvfiIz8ic/edit?usp=sharing)
#""")

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
# 🔄 Aba 3 - Atualizar Google Sheets (Definitivo e Corrigido)
# ================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# 🔹 Funções auxiliares

def normalizar_linha(data, loja, fat_total):
    """Normaliza os campos para gerar chave segura."""
    # Data
    try:
        data_formatada = pd.to_datetime(data, dayfirst=True, errors='coerce')
        if pd.isna(data_formatada):
            data_final = str(data).strip()
        else:
            data_final = data_formatada.strftime("%d/%m/%Y")
    except:
        data_final = str(data).strip()

    # Loja
    try:
        loja_final = str(loja).strip().lower()
    except:
        loja_final = ""

    # Fat.Total
    try:
        fat = str(fat_total).strip()
        fat = fat.replace(".", "").replace(",", ".")
        fat_float = float(fat)
        fat_final = f"{fat_float:.2f}".replace(".", ",")
    except:
        fat_final = "0,00"

    return f"{data_final}{loja_final}{fat_final}"

# 🔹 ABA 3

with aba3:
    st.header("🔄 Atualizar Banco de Dados - Apenas Novos Registros (Corrigido)")

    st.markdown("""
    🔗 [Clique aqui para abrir o **Faturamento Sistema Externo**](https://docs.google.com/spreadsheets/d/1_3uX7dlvKefaGDBUhWhyDSLbfXzAsw8bKRVvfiIz8ic/edit?usp=sharing)
    """, unsafe_allow_html=True)

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final.copy()

        if st.button("🔍 Verificar novos registros"):
            with st.spinner('🔄 Buscando dados atuais no Google Sheets...'):
                try:
                    # Conectar ao Google Sheets
                    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
                    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                    gc = gspread.authorize(credentials)

                    planilha = gc.open("Faturamento Sistema Externo")
                    aba = planilha.worksheet("Fat Sistema Externo")

                    dados_sheets = aba.get_all_values()

                    if not dados_sheets or len(dados_sheets) < 2:
                        st.warning("⚠️ Banco de dados vazio ou sem estrutura correta no Google Sheets.")
                        dados_existentes = pd.DataFrame()
                    else:
                        dados_existentes = pd.DataFrame(dados_sheets[1:], columns=dados_sheets[0])

                    # 🔥 Normalizar e gerar chaves no banco existente
                    if not dados_existentes.empty:
                        dados_existentes['Chave'] = dados_existentes.apply(
                            lambda x: normalizar_linha(x['Data'], x['Loja'], x['Fat.Total']),
                            axis=1
                        )
                        chaves_existentes = set(dados_existentes['Chave'])
                    else:
                        chaves_existentes = set()

                    # 🔥 Normalizar e gerar chaves no df_final
                    df_final['Chave'] = df_final.apply(
                        lambda x: normalizar_linha(x['Data'], x['Loja'], x['Fat.Total']),
                        axis=1
                    )

                    # 🔥 Selecionar apenas registros novos
                    novos_registros = df_final[~df_final['Chave'].isin(chaves_existentes)].drop(columns=['Chave'])

                    total_novos = len(novos_registros)

                    if total_novos == 0:
                        st.success("✅ Nenhum novo registro para adicionar. Banco de dados já atualizado.")
                    else:
                        st.success(f"✅ Encontrados {total_novos} registro(s) novo(s) para adicionar.")
                        st.dataframe(novos_registros)

                        st.session_state['novos_registros'] = novos_registros
                        st.session_state['linha_inicio'] = len(dados_existentes) + 2 if not dados_existentes.empty else 2

                except Exception as e:
                    st.error(f"❌ Erro ao buscar dados: {e}")

        # Se temos novos registros guardados, mostrar botão para confirmar
        if 'novos_registros' in st.session_state:
            confirmar = st.checkbox("✅ Confirmo que desejo adicionar os novos registros.")

            if confirmar:
                if st.button("📤 Atualizar Google Sheets"):
                    with st.spinner('📤 Atualizando o Google Sheets...'):
                        try:
                            # Conectar novamente
                            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                            credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
                            credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                            gc = gspread.authorize(credentials)

                            planilha = gc.open("Faturamento Sistema Externo")
                            aba = planilha.worksheet("Fat Sistema Externo")

                            novos_registros = st.session_state['novos_registros']
                            linha_inicio = st.session_state['linha_inicio']

                            aba.update(f"A{linha_inicio}", novos_registros.values.tolist())

                            st.success(f"🚀 {len(novos_registros)} registro(s) novo(s) adicionado(s) ao Google Sheets com sucesso!")

                            # Limpar session_state
                            del st.session_state['novos_registros']
                            del st.session_state['linha_inicio']

                        except Exception as e:
                            st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")

    else:
        st.warning("⚠️ Primeiro faça o upload e o processamento do arquivo na Aba 1.")

