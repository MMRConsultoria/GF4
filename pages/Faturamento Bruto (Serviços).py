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

# Upload do arquivo
uploaded_file = st.file_uploader("📄 Envie o arquivo Excel com a aba 'FaturamentoDiarioPorLoja'", type=["xlsx"])

if uploaded_file:
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

        # 👉 Trata a coluna de datas logo aqui
        df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], dayfirst=True, errors='coerce')

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

        # Montar o df_final
        df_final = pd.DataFrame(registros, columns=[
            "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
        ])

        # Ajustes de nomes e merge com empresas
        df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
        df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
        df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")

        # 🎯 Empresas não localizadas
        empresas_nao_localizadas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()

        st.markdown("---")
        st.subheader("🏢 Empresas não localizadas")

        if len(empresas_nao_localizadas) > 0:
            st.warning(f"⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s) na Tabela_Empresa:")
            for loja in empresas_nao_localizadas:
                st.text(f"🔎 {loja}")
        else:
            st.success("✅ Todas as empresas foram localizadas na Tabela_Empresa!")

        # 🔗 Links para as planilhas
        st.markdown("""
🔗 [Clique aqui para abrir a **Tabela_Empresa** no Google Sheets](https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link)
""")
        st.markdown("""
🔗 [Clique aqui para abrir a **Faturamento Sistema Externo** no Google Sheets](https://docs.google.com/spreadsheets/d/1_3uX7dlvKefaGDBUhWhyDSLbfXzAsw8bKRVvfiIz8ic/edit?usp=sharing)
""")

        # Continuação do processamento
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

        df_final["Data_Ordenada"] = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors="coerce")
        df_final = df_final.sort_values(by=["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")

        colunas_finais = [
            "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
            "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
            "Ticket", "Mês", "Ano"
        ]
        df_final = df_final[colunas_finais]

        # 📢 Mensagens de sucesso
        st.success("✅ Relatório processado com sucesso!")

        datas_validas = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce').dropna()
        if not datas_validas.empty:
            data_inicial = datas_validas.min().strftime("%d/%m/%Y")
            data_final = datas_validas.max().strftime("%d/%m/%Y")
            st.info(f"📅 Período processado: **{data_inicial}** até **{data_final}**")
        else:
            st.warning("⚠️ Não foi possível identificar o período de datas.")

        totalizador = df_final[["Fat.Total", "Serv/Tx", "Fat.Real"]].sum().round(2)
        totalizador_formatado = totalizador.apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.subheader("💰 Totais Gerais (R$)")
        st.dataframe(pd.DataFrame([totalizador_formatado]))

        # Função para gerar o Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
            output.seek(0)
            return output

        excel_data = to_excel(df_final)
        st.download_button(
            label="📅 Baixar Relatório Excel",
            data=excel_data,
            file_name="faturamento_servico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # =============================
        # 🔄 Atualizar Google Sheets (COM CHECAGEM DE DUPLICADOS)
        # =============================
        st.markdown("---")
        st.subheader("🔄 Atualizar Google Sheets?")

        if 'atualizou_google' not in st.session_state:
            st.session_state.atualizou_google = False

        if not st.session_state.atualizou_google:
            if st.button("📤 Atualizar tabela 'Fat Sistema Externo' no Google Sheets"):
                with st.spinner('🔄 Atualizando a planilha no Google Sheets...'):
                    try:
                        planilha_destino = gc.open("Faturamento Sistema Externo")
                        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                        # Baixar dados existentes
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
                            st.info("✅ Nenhum novo registro para atualizar no Google Sheets (todos já existem).")
                            st.session_state.atualizou_google = True
                        else:
                            rows = registros_novos.fillna("").values.tolist()
                            primeira_linha_vazia = len(dados_existentes) + 2
                            aba_destino.update(f"A{primeira_linha_vazia}", rows)

                            st.success(f"✅ {len(rows)} novo(s) registro(s) atualizados no Google Sheets!")
                            st.session_state.atualizou_google = True

                            registros_ignorados = len(novos_dados) - len(rows)
                            if registros_ignorados > 0:
                                st.warning(f"⚠️ {registros_ignorados} registro(s) já existiam e foram ignorados para evitar duplicação.")

                    except Exception as e:
                        st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")
                        st.session_state.atualizou_google = False
        else:
            st.info("✅ Dados já foram atualizados no Google Sheets nesta sessão.")

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
