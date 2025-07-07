import streamlit as st
import pandas as pd

st.set_page_config(page_title="Processar Metas", layout="wide")
st.title("📈 Processar apenas colunas de META")

uploaded_file = st.file_uploader("📁 Escolha seu arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    todas_abas = xls.sheet_names

    abas_escolhidas = st.multiselect(
        "Selecione as abas a processar:",
        options=todas_abas,
        default=[]
    )

    # Mapa para converter meses
    mapa_meses = {
        "janeiro": "Jan", "fevereiro": "Fev", "março": "Mar", "abril": "Abr",
        "maio": "Mai", "junho": "Jun", "julho": "Jul", "agosto": "Ago",
        "setembro": "Set", "outubro": "Out", "novembro": "Nov", "dezembro": "Dez"
    }

    df_final = pd.DataFrame(columns=["Mês", "Ano", "Grupo", "Loja", "Fat.Total"])

    for aba in abas_escolhidas:
        st.subheader(f"📝 Aba: {aba}")
        df_raw = pd.read_excel(xls, sheet_name=aba, header=None)

        # Nome do grupo (A1)
        nome_grupo = df_raw.iloc[0, 0]

        # Linha das lojas (linha 2, index=1)
        linha_lojas = df_raw.iloc[1, :]
        colunas_lojas = {}

        for col_idx, val in linha_lojas.items():
            if isinstance(val, str) and val.strip() != "" and col_idx >= 2:
                colunas_lojas.setdefault(val, []).append(col_idx)

        # Montar consolidado
        for loja, cols in colunas_lojas.items():
            headers_detectados = [str(df_raw.iloc[1, c]) for c in cols]
            st.write(f"🚀 Loja: {loja} | Headers encontrados: {headers_detectados}")

            # Agora só pega colunas com 'meta' no header
            metas_cols = [c for c in cols if 'meta' in str(df_raw.iloc[1, c]).lower()]
            st.write(f"✅ Colunas selecionadas com 'meta' para loja {loja}: {metas_cols}")

            if not metas_cols:
                continue

            for idx in range(2, len(df_raw)):
                mes_bruto = str(df_raw.iloc[idx, 1]).strip().lower()
                mes = mapa_meses.get(mes_bruto, mes_bruto)

                for c in metas_cols:
                    valor = df_raw.iloc[idx, c]
                    if isinstance(valor, str):
                        valor = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
                        try:
                            valor = float(valor)
                        except:
                            valor = None

                    linha = {
                        "Mês": mes,
                        "Ano": 2025,
                        "Grupo": nome_grupo,
                        "Loja": loja,
                        "Fat.Total": valor
                    }
                    df_final = pd.concat([df_final, pd.DataFrame([linha])], ignore_index=True)

    st.success("✅ Dados consolidados somente com colunas META:")
    st.dataframe(df_final)

    if not df_final.empty:
        csv = df_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar CSV consolidado",
            data=csv,
            file_name="metas_somente_meta.csv",
            mime='text/csv'
        )
else:
    st.info("💡 Faça o upload de um arquivo Excel para começar.")
