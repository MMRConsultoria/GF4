import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Relatório de Faturamento por Meio de Pagamento", layout="wide")
st.title("📊 Identificação de Blocos - Faturamento por Meio")

uploaded_file = st.file_uploader("📁 Envie o arquivo Excel com os dados", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo: {e}")
    else:
        linha_inicio_dados = 5  # Linha 6 da planilha (index 5)
        blocos = []
        col = 3  # Começa na coluna D (índice 3)
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
                df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()  # Coluna C (Data) e atual
                df_temp.columns = ["Data", "Valor (R$)"]

                # Alinha índice da coluna A com df_temp
                coluna_a = df_raw.iloc[linha_inicio_dados:, 0].reset_index(drop=True)
                df_temp = df_temp.reset_index(drop=True)

                # Remove linhas com coluna A vazia ou nula
                filtro_linhas_validas = ~coluna_a.astype(str).str.strip().isin(["", "nan"])
                df_temp = df_temp[filtro_linhas_validas]

                # Remove linhas com "total" ou "subtotal" na coluna Data
                df_temp = df_temp[~df_temp["Data"].astype(str).str.lower().str.contains("total|subtotal")]

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

            # Botão para baixar Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="FaturamentoPorMeio")
            output.seek(0)

            st.download_button(
                label="📥 Baixar resultado em Excel",
                data=output,
                file_name="FaturamentoPorMeio_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Nenhum dado válido foi identificado.")
