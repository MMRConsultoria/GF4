import streamlit as st
import pandas as pd

from io import BytesIO
import openpyxl

st.set_page_config(page_title="Processar Metas Dinâmico", layout="wide")
st.title("📈 Processar Metas - Excel contábil e visualização formatada (240.000,00)")

uploaded_file = st.file_uploader("📁 Escolha seu arquivo Excel", type=["xlsx"])

def formatar_excel_contabil(df, nome_aba="Metas"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nome_aba)
        workbook = writer.book
        worksheet = writer.sheets[nome_aba]
        
        # Formatar coluna Meta no Excel
        for idx, cell in enumerate(worksheet[1], 1):
            if cell.value == "Meta":
                col_meta_idx = idx
                for row in worksheet.iter_rows(min_row=2, min_col=col_meta_idx, max_col=col_meta_idx):
                    for cell in row:
                        cell.number_format = '#,##0.00'
                break
    output.seek(0)
    return output

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

    ordem_meses = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

    df_final = pd.DataFrame(columns=["Mês", "Ano", "Grupo", "Loja", "Meta"])

    for aba in abas_escolhidas:
        df_raw_ffill = pd.read_excel(xls, sheet_name=aba, header=None)
        df_raw_ffill = df_raw_ffill.ffill(axis=0)
        
        # Aba original sem ffill
        df_raw_original = pd.read_excel(xls, sheet_name=aba, header=None)

        grupo = df_raw_ffill.iloc[0,0]

        linha_header = None
        for idx in range(0, len(df_raw_ffill)):
            linha_textos = df_raw_ffill.iloc[idx,:].astype(str).str.lower().str.replace(" ", "")
            if linha_textos.str.contains("meta").any():
                linha_header = idx
                break

        if linha_header is None:
            continue

        metas_cols = []
        for col in range(df_raw_ffill.shape[1]):
            texto = str(df_raw_ffill.iloc[linha_header, col]).lower().replace(" ", "")
            loja_na_col_anterior = str(df_raw_ffill.iloc[linha_header - 1, col - 1]).lower()
            if "meta" in texto and all(x not in loja_na_col_anterior for x in ["total", "subtotal", "média"]):
                metas_cols.append(col)

        linha_dados_inicio = linha_header + 2

        for idx in range(linha_dados_inicio, len(df_raw_ffill)):
            mes_original = str(df_raw_original.iloc[idx, 1]).strip().lower()
            mes_original = mes_original.replace("marco", "março")
            if mes_original not in mapa_meses:
                continue

            mes = mapa_meses[mes_original]

            for c in metas_cols:
                loja = df_raw_ffill.iloc[linha_header - 1, c-1]
                if pd.isna(loja) or "consolidado" in str(loja).lower():
                    continue

                valor = df_raw_ffill.iloc[idx, c]
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

    df_final = df_final.drop_duplicates()
    if not df_final.empty:
        df_final["Meta"] = df_final["Meta"].fillna(0)
        df_final["Mês"] = pd.Categorical(df_final["Mês"], categories=ordem_meses, ordered=True)
        df_final = df_final.sort_values(["Ano", "Mês", "Loja"])

        # Formatar a exibição na tela do Streamlit
        df_final_fmt = df_final.copy()
        df_final_fmt["Meta"] = df_final_fmt["Meta"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.success("✅ Dados prontos, Meta exibida como 240.000,00 na tela e Excel contábil:")
        st.dataframe(df_final_fmt)

        excel_file = formatar_excel_contabil(df_final)
        st.download_button(
            label="📥 Baixar Excel consolidado (.xlsx)",
            data=excel_file,
            file_name="metas_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Nenhum dado encontrado. Verifique se as abas selecionadas contêm dados válidos.")
else:
    st.info("💡 Faça o upload de um arquivo Excel para começar.")
