import pandas as pd
import re
import math
import time

# Cabeçalho bonito ANTES das abas
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento por Serviço</h1>
    </div>
""", unsafe_allow_html=True)

# Agora sim, criar as abas
aba1, aba2, aba3 = st.tabs(["📤 Upload e Processamento", "⬇️ Download Excel", "🔄 Atualizar Google Sheets"])

# Conteúdo de cada aba
# ================================

with aba1:
    # 🔹 Upload do Arquivo
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
