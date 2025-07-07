import streamlit as st
import pandas as pd

st.set_page_config(page_title="Processar Metas Dinâmico", layout="wide")
st.title("📈 Processar Metas (Grupo em A1, Mês em B, valores 2 linhas abaixo do META)")

uploaded_file = st.file_uploader("📁 Escolha seu arquivo Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    todas_abas = xls.sheet_names

    abas_escolhidas = st.multiselect(
        "Selecione as abas a processar:",
        options=todas_abas,
        default=[]
    )

    mapa_meses = {
        "janeiro": "Jan", "fevereiro": "Fev", "março": "Mar", "abril": "Abr",
        "maio": "Mai", "junho": "Jun", "julho": "Jul", "agosto": "Ago",
        "setembro": "Set", "outubro": "Out", "novembro": "Nov", "dezembro": "Dez"
    }

    df_final = pd.DataFrame(columns=["Mês", "Ano", "Grupo", "Loja", "Meta"])

    for aba in abas_escolhidas:
        st.subheader(f"📝 Aba: {aba}")
        df_raw = pd.read_excel(xls, sheet_name=aba, header=None)

        grupo = df_raw.iloc[0,0]  # Grupo na célula A1

        # Encontrar linha do cabeçalho que tenha 'META'
        linha_header = None
        for idx in range(0, len(df_raw)):
            linha_textos = df_raw.iloc[idx,:].astype(str).str.lower().str.replace(" ", "")
            if linha_textos.str.contains("meta").any():
                linha_header = idx
                break

        if linha_header is None:
            st.warning(f"⚠️ Não encontrou linha com 'META' na aba {aba}.")
            continue

        # Encontrar coluna META
        meta_col = None
        for col in range(df_raw.shape[1]):
            texto = str(df_raw.iloc[linha_header, col]).lower().replace(" ", "")
            if "meta" in texto:
                meta_col = col
                break

        if meta_col is None:
            st.warning(f"⚠️ Não encontrou coluna 'META' na aba {aba}.")
            continue

        st.write(f"✅ Cabeçalho META detectado na linha {linha_header}, coluna {meta_col}")

        linha_dados_inicio = linha_header + 2  # 2 linhas abaixo

        for idx in range(linha_dados_inicio, len(df_raw)):
            if pd.isna(df_raw.iloc[idx, 1]):
                continue  # garante que tem mês na coluna B

            mes_bruto = str(df_raw.iloc[idx, 1]).strip().lower()
            mes = mapa_meses.get(mes_bruto, mes_bruto)

            loja = df_raw.iloc[linha_header - 1, meta_col]
            if pd.isna(loja) or "consolidado" in str(loja).lower():
                continue

            valor = df_raw.iloc[idx, meta_col]
            if isinstance(valor, str):
                valor = valor.replace('R$', '').replace('.', '').replace(',', '.').strip()
                try:
                    valor = float(valor)
                except:
                    valor = None

            linha = {
                "Mês": mes,
                "Ano": 2025,
                "Grupo": grupo,
                "Loja": loja,
                "Meta": valor
            }
            df_final = pd.concat([df_final, pd.DataFrame([linha])], ignore_index=True)

    st.success("✅ Dados consolidados só com META, ignorando Consolidado:")
    st.dataframe(df_final)

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
