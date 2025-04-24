pages/FaturamentoServico.py (corrigido: primeira loja começa na coluna D / índice 3)

import streamlit as st
import pandas as pd
import numpy as np
import json
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials


st.set_page_config(page_title="Faturamento por Serviço", layout="wide")
st.title("📋 Relatório de Faturamento por Serviço")

uploaded_file = st.file_uploader("Envie o arquivo Excel com a aba 'FaturamentoDiarioPorLoja'", type=["xlsx"])

# Conexão com Google Sheets via secrets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha.worksheet("Tabela_Empresa").get_all_records())

uploaded_file = st.file_uploader(
    label="📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento",
    type=["xlsx", "xlsm"],
    help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
)






if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

        # Validação da célula B1
        texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
        if texto_b1 != "faturamento diário sintético multi-loja":
            st.error("❌ ERRO: A célula B1 deve conter 'Faturamento diário sintético multi-loja'. Verifique o arquivo.")
            st.stop()

        df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
        linha_lojas = df_raw.iloc[3, 3:].dropna()  # primeira loja começa na coluna D (índice 3)

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

        df_final = pd.DataFrame(registros, columns=[
            "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
        ])

        dias_traducao = {
            "Monday": "segunda-feira",
            "Tuesday": "terça-feira",
            "Wednesday": "quarta-feira",
            "Thursday": "quinta-feira",
            "Friday": "sexta-feira",
            "Saturday": "sábado",
            "Sunday": "domingo"
        }

        df_final["Data"] = pd.to_datetime(df_final["Data"], errors='coerce')
        df_final.insert(1, "Dia da Semana", df_final["Data"].dt.day_name().map(dias_traducao))
        df_final["Data"] = df_final["Data"].dt.strftime("%d/%m/%Y")

        # Formatar valores numéricos com duas casas decimais
        colunas_valores = ["Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas"]
        for col_val in colunas_valores:
            df_final[col_val] = pd.to_numeric(df_final[col_val], errors="coerce").round(2)

        # Traduzir o mês para abreviado em português
        meses = {
            "jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
            "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"
        }
        df_final["Mês"] = df_final["Mês"].str.lower().map(meses)

        # Ordenar por Data e Loja
        df_final["Data_Ordenada"] = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors="coerce")
        df_final = df_final.sort_values(by=["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")

        st.success("✅ Relatório processado com sucesso!")

        # Mostrar total dos valores na tela (sem incluir Ticket)
        totalizador = df_final[["Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas"]].sum().round(2)
        st.subheader("📊 Totais Gerais")
        st.dataframe(pd.DataFrame(totalizador).transpose())

        st.dataframe(df_final.head(50))

        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
            writer.close()
            output.seek(0)
            return output

        excel_data = to_excel(df_final)
        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=excel_data,
            file_name="faturamento_servico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")GOOGLE_SERVICE_ACCOUNT = '''
{\n  "type": "service_account",\n  "project_id": "projetosangria",\n  "private_key_id": "30ff2b50039e7e052194b3ca0c07cbe16c997ee1",\n  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCrjO2oXZfSkcli\nTXdmwElySdPwJBuweqHHo4WgJwhzyoKlH81G3Hiqeh65tdyku09Z6LvkmRIAncWu\nXZiv2AH6w049nSTIYLk8QY4yeN0oTEvoxMhsiEjbapIMikob0eTwncQuogDc+xjM\nkAh/cnnyMxhxkQiV/qhw5SnDKunGwtczAT/47dR6zgICYxIDSZjAw8kQGHi2ncHB\nbMkpGUiYc9UDAlNPI5pz4JOAuPBEKAxx/tINw3r+ieo0K1vwJTP46BAGl0ui7oGs\npfDOmf7/Ki4OQoW/wW6FpLRjIYApUiv4gHMvPdihDAdJJCD2sFIlVQxrAc+N9PXx\no4CPp1NbAgMBAAECggEAMF+c5Ez/8rCJSN7vPFocW83VpXGJHUp3NAQ4OeDH0V7Z\nWBaPm9uvauLkpxlRDmKDDadC1EMVgHb8tx5NX8hZRoysE1Ose5RKp2MU5caPFj3t\nacWTfocvhC+Y31BfdVjKZc8W8s9bzvQ/Ge/DdayiFlmyGtP6x9D3Tl5QWGhRY2o3\nOoyzJLAZEUtfTg4SbNGlz+huTf0tzKvTygVbn68TWSwzxsYZr+g6jbJiwnUUY9bJ\nz5V4860S9ZE20+YtUl7+UCYpbgFDJBU13bS/hotLdfR6uivaRb0UB0Ar+dI2YliE\nSC3LOaYbrulP/BZiAYBb+ObAYxo6iwkXOiEx1v/HAQKBgQDe9/GfqR7jdFEaigp0\nTLvf6i05IrK5OlVgCYJEMNBpljlE5a6uy3xWUPq06n0HMHNMsdSqiCkd3YgqiGS6\n057sjSp+zslmgttaPVcEq4Ju6/DdwcVSNPAx2MQeOpc/iwv+cMPzphFf56cLztnn\nW04E+xMb88Sga3t9vZGFTpKtowKBgQDE9vYrQ5KvZi1TR9ecrPdpBkQikUF04FU7\nEjh11AIf1iCsOOVBzIxNXQLmDkHMzfRkybSQdHPtYhzW33XvAxYd/DDpX/7c2bKZ\nw2Lop+3doi+NRRVo78kfW7Jwj9oKsgQJK7mR1Er8FxTDGuT1rKO4wZp3+dzNGDp9\nsy2Fcnwu6QKBgFW9ADPGA7OxObK73DNGgoQi94rd9d3WOZg5b9cq8il388Ozko1y\nf/htIUrKVJOcJOocA8wjmbP7SO4aDqns0FLkIbArcjSyIwk7Ryfrj7d5kOClL2xi\nIO76DMgW/awYmt8Xm7IobMv1Nz4KJ66YZJLzvTBld3m8evsqFzgss6iDAoGBAI1W\nekeJcdUTiSrdvsbbB58BtBrqCQKJiB+fb4x92hhmW4O6QCj7UyKgv9e+G1GP6PP6\nGewe5KFPakp2h/Y/TLuvoJmXHRf1z8evAWbaOkJ0g5LoI/jtOHGcJ2vGjjxmiabb\nDYqrwDMtjOSEhBAXFlUZ+HJnhh5WaIKS9PNxt9MhAoGBAN0I6j7RBQgUkYALtvTR\n4ZW44Iiq9kAiE7dIwJwIsfe3p+En43g/zaxEARN3OAFLiqlNtxLsP9e1ZLYxiqcd\niJu0qFQP6XQy/5dQBwTk/OfKyq0Ln+6aCgtXwvbXqH1q0WXcxNg5Sx5ppPCkF7/g\n6DIHieYAcQGeijg1/VB4Y+ij\n-----END PRIVATE KEY-----\n",\n  "client_email": "sangria@projetosangria.iam.gserviceaccount.com",\n  "client_id": "101891116808969008137",\n  "auth_uri": "https://accounts.google.com/o/oauth2/auth",\n  "token_uri": "https://oauth2.googleapis.com/token",\n  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",\n  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/sangria%40projetosangria.iam.gserviceaccount.com",\n  "universe_domain": "googleapis.com"\n}
'''
