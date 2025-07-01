import streamlit as st
import pandas as pd
import numpy as np
import re
import json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Meio de Pagamento", layout="wide")

# 🔥 CSS para estilizar as abas
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] { margin-top: 20px; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }
    </style>
""", unsafe_allow_html=True)

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# 🔥 Título
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Meio de Pagamento</h1>
    </div>
""", unsafe_allow_html=True)

# 🔌 Conexão com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha = gc.open("Vendas diarias")
df_empresa = pd.DataFrame(planilha.worksheet("Tabela Empresa").get_all_records())

# ========================
# 🗂️ Abas
# ========================
tab1, tab2, tab3 = st.tabs([
    "📥 Upload e Processamento",
    "🔄 Atualizar Google Sheets",
    "📝 Desenvolvimento"
])

# ======================
# 📥 Aba 1
# ======================
with tab1:
    uploaded_file = st.file_uploader(
        label="📁 Clique para selecionar ou arraste aqui o arquivo Excel",
        type=["xlsx", "xlsm"]
    )

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            abas_disponiveis = xls.sheet_names

            aba_escolhida = abas_disponiveis[0] if len(abas_disponiveis) == 1 else st.selectbox(
                "Escolha a aba para processar", abas_disponiveis)

            df_raw = pd.read_excel(xls, sheet_name=aba_escolhida, header=None)
            df_raw = df_raw[~df_raw.iloc[:, 1].astype(str).str.lower().str.contains("total|subtotal", na=False)]
        except Exception as e:
            st.error(f"❌ Não foi possível ler o arquivo enviado: {e}")
        else:
            if str(df_raw.iloc[0, 1]).strip().lower() != "faturamento diário por meio de pagamento":
                st.error("❌ A célula B1 deve conter 'Faturamento diário por meio de pagamento'.")
                st.stop()

            linha_inicio_dados, blocos, col, loja_atual = 5, [], 3, None

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
                if any(p in t for t in [linha3, valor_linha4.lower(), linha5] for p in ["total", "serv/tx", "total real"]):
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
                st.error("❌ Nenhum dado válido encontrado.")
            else:
                df_meio_pagamento = pd.concat(blocos, ignore_index=True).dropna()
                df_meio_pagamento = df_meio_pagamento[~df_meio_pagamento["Data"].astype(str).str.lower().str.contains("total|subtotal")]
                df_meio_pagamento["Data"] = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True, errors="coerce")
                df_meio_pagamento = df_meio_pagamento[df_meio_pagamento["Data"].notna()]
                dias_semana = {'Monday': 'segunda-feira','Tuesday': 'terça-feira','Wednesday': 'quarta-feira',
                               'Thursday': 'quinta-feira','Friday': 'sexta-feira','Saturday': 'sábado','Sunday': 'domingo'}
                df_meio_pagamento["Dia da Semana"] = df_meio_pagamento["Data"].dt.day_name().map(dias_semana)
                df_meio_pagamento = df_meio_pagamento.sort_values(by=["Data", "Loja"])
                periodo_min = df_meio_pagamento["Data"].min().strftime("%d/%m/%Y")
                periodo_max = df_meio_pagamento["Data"].max().strftime("%d/%m/%Y")
                df_meio_pagamento["Data"] = df_meio_pagamento["Data"].dt.strftime("%d/%m/%Y")

                df_meio_pagamento["Loja"] = df_meio_pagamento["Loja"].str.strip().str.replace(r"^\d+\s*-\s*", "", regex=True).str.lower()
                df_empresa["Loja"] = df_empresa["Loja"].str.strip().str.lower()
                df_meio_pagamento = pd.merge(df_meio_pagamento, df_empresa, on="Loja", how="left")

                df_meio_pagamento["Mês"] = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True).dt.month.map({
                    1:'jan',2:'fev',3:'mar',4:'abr',5:'mai',6:'jun',7:'jul',8:'ago',9:'set',10:'out',11:'nov',12:'dez'})
                df_meio_pagamento["Ano"] = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True).dt.year

                df_meio_pagamento = df_meio_pagamento[[
                    "Data","Dia da Semana","Meio de Pagamento","Loja","Código Everest",
                    "Grupo","Código Grupo Everest","Valor (R$)","Mês","Ano"
                ]]

                st.session_state.df_meio_pagamento = df_meio_pagamento

                col1, col2 = st.columns(2)
                col1.markdown(f"<div style='font-size:1.2rem;'>📅 Período processado<br>{periodo_min} até {periodo_max}</div>", unsafe_allow_html=True)
                valor_total = f"R$ {df_meio_pagamento['Valor (R$)'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                col2.markdown(f"<div style='font-size:1.2rem;'>💰 Valor total<br><span style='color:green;'>{valor_total}</span></div>", unsafe_allow_html=True)

                # Só permite download se todas lojas estiverem cadastradas
                empresas_nao_localizadas = df_meio_pagamento[df_meio_pagamento["Código Everest"].isna()]["Loja"].unique()
                if len(empresas_nao_localizadas) > 0:
                    empresas_nao_localizadas_str = "<br>".join(empresas_nao_localizadas)
                    mensagem = f"""
                    ⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s), cadastre e reprocesse novamente!<br>
                    {empresas_nao_localizadas_str}
                    <br>✏️ Atualize a tabela clicando 
                    <a href='https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU' target='_blank'><strong>aqui</strong></a>.
                    """
                    st.markdown(mensagem, unsafe_allow_html=True)
                else:
                    st.success("✅ Todas as empresas foram localizadas na Tabela_Empresa!")

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_meio_pagamento.to_excel(writer, index=False, sheet_name="FaturamentoPorMeio")
                    output.seek(0)

                    st.download_button(
                        "📥 Baixar relatório Excel",
                        data=output,
                        file_name="FaturamentoPorMeio.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

# ======================
# 🔄 Aba 2
# ======================
with tab2:
    st.markdown("🔗 [Abrir planilha Faturamento Meio Pagamento](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU)")

    if 'df_final' not in st.session_state:
        st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")
    elif not all(c in st.session_state.df_final.columns for c in ["Meio de Pagamento","Loja","Data"]):
        st.warning("⚠️ O arquivo processado não tem as colunas necessárias.")
    else:
        df_final = st.session_state.df_final.copy()
        lojas_nao_cadastradas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()
        todas_lojas_ok = len(lojas_nao_cadastradas) == 0

        df_final['M'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d') + \
                        df_final['Meio de Pagamento'] + df_final['Loja']
        df_final['Valor (R$)'] = df_final['Valor (R$)'].apply(lambda x: float(str(x).replace(',','.')))
        df_final['Data'] = (pd.to_datetime(df_final['Data'], dayfirst=True) - pd.Timestamp("1899-12-30")).dt.days
        for col in ["Código Everest","Código Grupo Everest","Ano"]:
            df_final[col] = df_final[col].apply(lambda x: int(x) if pd.notnull(x) and str(x).strip() != "" else "")

        aba_destino = gc.open("Vendas diarias").worksheet("Faturamento Meio Pagamento")
        valores_existentes = aba_destino.get_all_values()
        dados_existentes = set([linha[9] for linha in valores_existentes[1:] if len(linha) > 9])

        novos_dados, duplicados = [], []
        for linha in df_final.fillna("").values.tolist():
            chave_m = linha[-1]
            if chave_m not in dados_existentes:
                novos_dados.append(linha)
                dados_existentes.add(chave_m)
            else:
                duplicados.append(linha)

        if todas_lojas_ok and st.button("📥 Enviar dados para o Google Sheets"):
            with st.spinner("🔄 Atualizando..."):
                aba_destino.append_rows(novos_dados)
                st.success(f"✅ {len(novos_dados)} novos registros enviados!")
                if duplicados:
                    st.warning(f"⚠️ {len(duplicados)} registros duplicados não foram enviados.")

# ======================
# 📝 Aba 3
# ======================
with tab3:
    st.info("🔍 Desenvolvimento")
