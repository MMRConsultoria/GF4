import streamlit as st
import pandas as pd
import numpy as np
import re
import json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Relatório de Faturamento", layout="wide")

st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento</h1>
    </div>
""", unsafe_allow_html=True)

# Conexão com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha.worksheet("Tabela_Empresa").get_all_records())

# Upload do arquivo Excel
uploaded_file = st.file_uploader(
    label="📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento",
    type=["xlsx", "xlsm"],
    help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_raw = pd.read_excel(xls, sheet_name="FaturamentoPorMeioDePagamento", header=None)
    except Exception as e:
        st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes: {e}")
    else:
        if str(df_raw.iloc[0, 1]).strip().lower() != "faturamento diário por meio de pagamento":
            st.error("❌ A célula B1 deve conter 'Faturamento diário por meio de pagamento'.")
            st.stop()

      # Nova leitura de blocos com limpeza de linhas da coluna B
        df_raw = df_raw[~df_raw.iloc[:, 1].astype(str).str.lower().str.contains("total|subtotal", na=False)]
        
        linha_inicio_dados = 5  # Linha 6 (index 5)
        blocos = []
        col = 3  # Começa na coluna D
        loja_atual = None

        while col < df_raw.shape[1]:
            valor_linha4 = str(df_raw.iloc[3, col]).strip()
            match = re.match(r"^\d+\s*-\s*(.+)$", valor_linha4)
            if match:
                loja_atual = match.group(1).strip().lower()

            meio_pgto = str(df_raw.iloc[4, col]).strip()
            if not loja_atual or not meio_pgto or meio_pgto.lower() in ["nan", ""]:
                col += 1
                continue

            linha3 = str(df_raw.iloc[2, col]).strip().lower()
            linha5 = meio_pgto.lower()

            if any(palavra in texto for texto in [linha3, valor_linha4.lower(), linha5] for palavra in ["total", "serv/tx", "total real"]):
                col += 1
                continue

            try:
                df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
                df_temp.columns = ["Data", "Valor (R$)"]
                df_temp = df_temp[~df_temp["Data"].astype(str).str.lower().str.contains("total|subtotal")]
                df_temp.insert(1, "Meio de Pagamento", meio_pgto)
                df_temp.insert(2, "Loja", loja_atual)
                blocos.append(df_temp)
            except Exception as e:
                st.warning(f"⚠️ Erro ao processar coluna {col}: {e}")

            col += 1

        if not blocos:
            st.error("❌ Nenhum dado válido encontrado na planilha.")
        else:
            df = pd.concat(blocos, ignore_index=True)
            # 🔥 Remove linhas com qualquer célula em branco
            df = df.dropna(how="any")

            # Converte a coluna Data para datetime
        if  df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
            df = df[df["Data"].notna()]

            # Gera a coluna 'Dia da Semana' antes de formatar a data
            dias_semana = {
                'Monday': 'segunda-feira',
                'Tuesday': 'terça-feira',
                'Wednesday': 'quarta-feira',
                'Thursday': 'quinta-feira',
                'Friday': 'sexta-feira',
                'Saturday': 'sábado',
                'Sunday': 'domingo'
            }
            df["Dia da Semana"] = df["Data"].dt.day_name().map(dias_semana)

            # Ordena por Data e Loja
            df = df.sort_values(by=["Data", "Loja"])

            # Formata a data para exibição
            df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")


            

            # 📅 Ordena por Data e Loja
            #df = df.sort_values(by=["Data", "Loja"])
            
            #df = df[df["Data"].notna() & ~df["Data"].astype(str).str.lower().str.contains("total|subtotal")]
            #df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
            #df = df[df["Data"].notna()]

            # Remove linhas com "total" ou "subtotal" na coluna Data
            df = df[~df["Data"].astype(str).str.lower().str.contains("total|subtotal")]
    
            # Converte a coluna Data
            df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")

            # Remove datas inválidas
            df = df[df["Data"].notna()]

            # Ordena agora sim por Data real e Loja
            df = df.sort_values(by=["Data", "Loja"])

            # (Opcional) Formata a data para exibição
            df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

        
            # Padroniza nome da loja (remove prefixos como "1 - ", etc.)
            df["Loja"] = (
                df["Loja"]
                .astype(str)
                .str.strip()
                .str.replace(r"^\d+\s*-\s*", "", regex=True)
                .str.lower()
            )

            df_empresa["Loja"] = (
                df_empresa["Loja"]
                .astype(str)
                .str.strip()
                .str.lower()
            )

            df = pd.merge(df, df_empresa, on="Loja", how="left")

            # Tradução manual dos dias da semana
            dias_semana = {
                'Monday': 'segunda-feira',
                'Tuesday': 'terça-feira',
                'Wednesday': 'quarta-feira',
                'Thursday': 'quinta-feira',
                'Friday': 'sexta-feira',
                'Saturday': 'sábado',
                'Sunday': 'domingo'
            }

            # Tradução dos meses para abreviações em português
            meses = {
                1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
                7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
            }

            df["Dia da Semana"] = df["Data"].dt.day_name().map(dias_semana)
            df["Mês"] = df["Data"].dt.month.map(meses)
            df["Ano"] = df["Data"].dt.year
            df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

            # Garante que colunas do merge estejam presentes
            for col in ["Código Everest", "Grupo", "Código Grupo Everest"]:
                if col not in df.columns:
                    df[col] = np.nan

            df = df[[
                "Data", "Dia da Semana", "Meio de Pagamento", "Loja",
                "Código Everest", "Grupo", "Código Grupo Everest",
                "Valor (R$)", "Mês", "Ano"
            ]]

            df = df.sort_values(by=["Data", "Loja"])

            periodo_min = df["Data"].min()
            periodo_max = df["Data"].max()
            valor_total = df["Valor (R$)"].sum()

            col1, col2 = st.columns(2)
            col1.metric("📅 Período processado", f"{periodo_min} até {periodo_max}")
            col2.metric("💰 Valor total", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            lojas_sem_codigo = df[df["Código Everest"].isna()]["Loja"].unique()
            if len(lojas_sem_codigo) > 0:
                st.warning(
                    f"⚠️ Lojas sem código Everest cadastrado: {', '.join(lojas_sem_codigo)}\n\n"
                    "🔗 Atualize os dados na [planilha de empresas](https://docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit)"
                )

            st.success("✅ Relatório de faturamento por meio de pagamento gerado com sucesso!")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="FaturamentoPorMeio")
            output.seek(0)

            st.download_button("📥 Baixar relatório", data=output, file_name="FaturamentoPorMeio_transformado.xlsx")
