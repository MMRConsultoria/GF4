import streamlit as st
import pandas as pd

st.set_page_config(page_title="Processar Metas", layout="wide")
st.title("📈 Processar Metas - Formato consolidado")

uploaded_file = st.file_uploader("📁 Escolha seu arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    todas_abas = xls.sheet_names

    abas_escolhidas = st.multiselect(
        "Selecione as abas a processar:",
        options=todas_abas,
        default=[]
    )

    # DataFrame final
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
            metas_cols = [c for c in cols if '2025' in str(df_raw.iloc[1, c])]
            if not metas_cols:
                continue

            for idx in range(2, len(df_raw)):
                mes = df_raw.iloc[idx, 1]
                for c in metas_cols:
                    valor = df_raw.iloc[idx, c]
                    linha = {
                        "Mês": mes,
                        "Ano": 2025,
                        "Grupo": nome_grupo,
                        "Loja": loja,
                        "Fat.Total": valor
                    }
                    df_final = pd.concat([df_final, pd.DataFrame([linha])], ignore_index=True)

    st.success("✅ Dados consolidados no formato desejado:")
    st.dataframe(df_final)

    # Botão para exportar
    if not df_final.empty:
        csv = df_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar CSV consolidado",
            data=csv,
            file_name="metas_consolidado.csv",
            mime='text/csv'
        )
else:
    st.info("💡 Faça o upload de um arquivo Excel para começar.")
