import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import numbers
from openpyxl import load_workbook
from urllib.parse import quote

st.set_page_config(page_title="Processador de Sangria", layout="centered")
st.title("📊 Processador de Sangria")

sheet_id = "13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q"
sheet_empresa = quote("Tabela_Empresa")
tabela_empresa_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_empresa}"
df_empresa_raw = pd.read_csv(tabela_empresa_url, header=None)
cabecalho = df_empresa_raw.iloc[0].fillna("").astype(str).str.strip()
df_empresa = df_empresa_raw[1:].copy()
df_empresa.columns = cabecalho
df_empresa = df_empresa.loc[:, df_empresa.columns != ""]

uploaded_file = st.file_uploader("Envie seu arquivo Excel (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_dados = pd.read_excel(xls, sheet_name="Sheet")
    except Exception as e:
        st.error(f"Erro ao ler o Excel: {e}")
    else:
        st.subheader("Prévia dos dados da aba 'Sheet'")
        st.dataframe(df_dados.head())

        if st.button("Processar Sangria"):
            st.info("🔄 Processando arquivo... Aguarde...")
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
                    loja_atual = loja or "Loja nao cadastrada"
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

            df["Loja"] = df["Loja"].astype(str).str.strip().str.upper()
            df_empresa[df_empresa.columns[0]] = df_empresa[df_empresa.columns[0]].astype(str).str.strip().str.upper()

            st.subheader("🔎 Visualização da Tabela Empresa (Google Sheets)")
            st.write(df_empresa.head())

            df_final = df.merge(df_empresa, left_on="Loja", right_on=df_empresa.columns[0], how="left")

            lojas_nao_cadastradas = df_final[df_final[df_empresa.columns[2]].isna() | df_final[df_empresa.columns[3]].isna()]["Loja"].unique()
            if len(lojas_nao_cadastradas) > 0:
                st.warning("⚠️ As seguintes lojas não foram encontradas na Tabela Empresa:")
                for loja in lojas_nao_cadastradas:
                    st.text(f"- {loja}")

            def mapear_resumo(desc):
                desc = desc.lower()
                if any(p in desc for p in ["meta", "prem"]):
                    return "Premiação/Meta"
                elif "motiv" in desc:
                    return "Motivacional"
                elif "sangr" in desc:
                    return "Sangria"
                elif "dep" in desc:
                    return "Deposito"
                else:
                    return ""

            df_final["Resumo Descrição"] = df_final["Descrição"].apply(mapear_resumo)

            ordem_colunas = [
                "Data", "Dia da Semana", "Loja", df_empresa.columns[2], df_empresa.columns[1], df_empresa.columns[3],
                "Meio de recebimento", "Funcionário", "Hora", "Descrição", "Resumo Descrição", "Valor(R$)", "Mês", "Ano"
            ]
            df_final = df_final[[col for col in ordem_colunas if col in df_final.columns]]

            df_final["Data"] = pd.to_datetime(df_final["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
            periodo_min = pd.to_datetime(df_final["Data"], dayfirst=True, errors="coerce").min().strftime("%d/%m/%Y")
            periodo_max = pd.to_datetime(df_final["Data"], dayfirst=True, errors="coerce").max().strftime("%d/%m/%Y")
            valor_total = df_final["Valor(R$)"].sum()

            st.subheader("📅 Período e Valor Processado")
            st.markdown(f"**Período:** {periodo_min} até {periodo_max}")
            st.markdown(f"**Valor Total:** R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.subheader("🔍 Tabela final antes de exportar")
            st.write(f"Total de linhas: {len(df_final)}")
            st.dataframe(df_final.head(30))

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Sangria")
                workbook = writer.book
                worksheet = writer.sheets["Sangria"]
                valor_col_idx = df_final.columns.get_loc("Valor(R$)") + 1
                for row in range(2, len(df_final) + 2):
                    cell = worksheet.cell(row=row, column=valor_col_idx)
                    cell.number_format = '#,##0.00\ [$R$-pt-BR]'
            output.seek(0)

            st.success("✅ Sangria processada com sucesso!")
            st.download_button("📥 Baixar resultado", data=output, file_name="Sangria_estruturada.xlsx")
