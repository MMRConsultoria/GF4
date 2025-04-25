import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Teste de Blocos - Faturamento", layout="wide")
st.title("🧪 Teste de Identificação de Blocos")

# Upload do arquivo
uploaded_file = st.file_uploader("📁 Envie o arquivo de Faturamento (.xlsx)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo: {e}")
    else:
        linha_inicio_dados = 6
        blocos = []
        col = 3  # Começa na coluna D (índice 3)
        loja_atual = None

        while col < df_raw.shape[1]:
            valor_linha4 = str(df_raw.iloc[3, col]).strip()

            # Verifica se é uma nova loja no formato "número - nome"
            match = re.match(r"^\d+\s*-\s*(.+)$", valor_linha4)
            if match:
                loja_atual = match.group(1).strip().lower()

            meio_pgto = str(df_raw.iloc[4, col]).strip()
            if not loja_atual or not meio_pgto or meio_pgto.lower() in ["nan", ""]:
                col += 1
                continue

            # Verifica palavras proibidas nas linhas 3, 4 e 5
            linha3 = str(df_raw.iloc[2, col]).strip().lower()
            linha5 = meio_pgto.lower()

            if any(palavra in texto for texto in [linha3, valor_linha4.lower(), linha5] for palavra in ["total", "serv/tx", "total real"]):
                col += 1
                continue

            try:
                df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
                df_temp.columns = ["Data", "Valor (R$)"]
                df_temp.insert(1, "Meio de Pagamento", meio_pgto)
                df_temp.insert(2, "Loja", loja_atual)
                blocos.append(df_temp)
            except Exception as e:
                st.warning(f"⚠️ Erro ao processar coluna {col}: {e}")

            col += 1

        if blocos:
            df_final = pd.concat(blocos, ignore_index=True)
            st.success("✅ Blocos identificados com sucesso!")
            st.dataframe(df_final, use_container_width=True)
        else:
            st.warning("⚠️ Nenhum dado válido foi identificado.")
