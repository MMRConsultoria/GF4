# pages/FaturamentoServico.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Tabela")
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela_Empresa").get_all_records())

# ================================
# 2. Configuração inicial do app
# ================================
st.set_page_config(page_title="Faturamento por Serviço", layout="wide")
#st.title("📋 Relatório de Faturamento por Serviço")

# 🎨 Estilizar abas
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


# Cabeçalho bonito (depois do estilo)
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório de Faturamento por Serviço</h1>
    </div>
""", unsafe_allow_html=True)


# ================================
# 3. Separação em ABAS
# ================================
aba1, aba2, aba3, aba4 = st.tabs(["📄 Upload e Processamento", "📥 Download Excel", "🔄 Atualizar Google Sheets","📊 Relatórios Gerenciais"])

# ================================
# 📄 Aba 1 - Upload e Processamento
# ================================
with aba1:
    uploaded_file = st.file_uploader("📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento", type=["xlsx"])

    if uploaded_file:
        try:
            # 🔹 Carregar o arquivo
            xls = pd.ExcelFile(uploaded_file)
            df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)

            # 🔹 Validar B1
            texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
            if texto_b1 != "faturamento diário sintético multi-loja":
                st.error(f"❌ A célula B1 está com '{texto_b1}'. Corrija para 'Faturamento diário sintético multi-loja'.")
                st.stop()

            df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
            df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], dayfirst=True, errors='coerce')

            # 🔹 Processamento dos registros
            registros = []
            col = 3
            while col < df.shape[1]:
                nome_loja = str(df_raw.iloc[3, col]).strip()
                if re.match(r"^\d+\s*-?\s*", nome_loja):
                    nome_loja = nome_loja.split("-", 1)[-1].strip()

                    header_col = str(df.iloc[0, col]).strip().lower()
                    if "fat.total" in header_col:
                        for i in range(1, df.shape[0]):
                            linha = df.iloc[i]
                            valor_data = df.iloc[i, 2]
                            valor_check = str(df.iloc[i, 1]).strip().lower()

                            if pd.isna(valor_data) or valor_check in ["total", "subtotal"]:
                                continue

                            data = valor_data
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

            if len(registros) == 0:
                st.warning("⚠️ Nenhum registro encontrado.")

            # 🔹 Montar df_final
            df_final = pd.DataFrame(registros, columns=[
                "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
            ])

            df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
            df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
            df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")

            # 🔹 Ajustar dados
            dias_traducao = {
                "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
                "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
            }
            df_final.insert(1, "Dia da Semana", pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.day_name().map(dias_traducao))
            df_final["Data"] = pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.strftime("%d/%m/%Y")

            for col_val in ["Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas"]:
                df_final[col_val] = pd.to_numeric(df_final[col_val], errors="coerce").round(2)

            meses = {
                "jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"
            }
            df_final["Mês"] = df_final["Mês"].str.lower().map(meses)

            df_final["Data_Ordenada"] = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce')
            df_final = df_final.sort_values(by=["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")

            colunas_finais = [
                "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
                "Ticket", "Mês", "Ano"
            ]
            df_final = df_final[colunas_finais]

            st.session_state.df_final = df_final
            st.session_state.atualizou_google = False

            # 🔥 Agora exibir:
            # 📄 Nome do Arquivo
           # st.markdown(f"""
              #  <div style='font-size:15px; font-weight: bold; margin-bottom:10px;'>
                   # 📄 Arquivo selecionado: {uploaded_file.name}
              #  </div>
            #""", unsafe_allow_html=True)

            # 📅 e 💰 Período e Valor Total
            datas_validas = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce').dropna()

            if not datas_validas.empty:
                data_inicial = datas_validas.min().strftime("%d/%m/%Y")
                data_final = datas_validas.max().strftime("%d/%m/%Y")
                
                valor_total = df_final["Fat.Total"].sum().round(2)
                valor_total_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown(f"""
                        <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>📅 Período processado</div>
                        <div style='font-size:30px; color:#000;'>{data_inicial} até {data_final}</div>
                    """, unsafe_allow_html=True)

                with col2:
                    st.markdown(f"""
                        <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>💰 Valor total</div>
                        <div style='font-size:30px; color:green;'>{valor_total_formatado}</div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ Não foi possível identificar o período de datas.")

           # 🔎 Empresas não localizadas
            empresas_nao_localizadas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()

            if len(empresas_nao_localizadas) > 0:
                # Listar as empresas não localizadas
                empresas_nao_localizadas_str = "<br>".join(empresas_nao_localizadas)
                
                # Construir a mensagem com o link direto
                mensagem = f"""
                ⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s): 
                <br>{empresas_nao_localizadas_str}
                <br>
                ✏️ Atualize a tabela clicando 
                <a href='https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link' target='_blank'><strong>aqui</strong></a>.
                """

                st.markdown(mensagem, unsafe_allow_html=True)


            else:
                st.success("✅ Todas as empresas foram localizadas na Tabela_Empresa!")

           # 🔗 Links úteis
           # st.markdown("""
#🔗 [Link **Tabela_Empresa**](https://docs.google.com/spreadsheets/d/13BvAIzgp7w7wrfkwM_MOnHqHYol-dpWiEZBjyODvI4Q/edit?usp=drive_link)

#""")

        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {e}")

# ================================
# 📥 Aba 2 - Download Excel
# ================================
with aba2:
    st.header("📥 Download Relatório Excel")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
            output.seek(0)
            return output

        excel_data = to_excel(df_final)

        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=excel_data,
            file_name="faturamento_servico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo na aba anterior.")
# =======================================
# Atualizar Google Sheets (Evitar duplicação)
# =======================================

with aba3:

     
    #st.header("📤 Atualizar Banco de Dados (Evitar duplicação usando coluna M)")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final.copy()
        
        # 🔗 Links úteis
        st.markdown("""
          🔗 [Link  **Faturamento Sistema Externo**](https://docs.google.com/spreadsheets/d/1_3uX7dlvKefaGDBUhWhyDSLbfXzAsw8bKRVvfiIz8ic/edit?usp=sharing)
        """, unsafe_allow_html=True)

        # Criar a coluna "M" com a concatenação de "Data", "Fat.Total" e "Loja" como string para verificação de duplicação
        df_final['M'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d') + \
                        df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)

        #df_final['M'] = df_final['Data'] + df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)


        # Não converter para string, apenas utilizar "M" para verificação de duplicação
        df_final['M'] = df_final['M'].apply(str)

        # Converter o restante do DataFrame para string, mas mantendo as colunas numéricas com seu formato correto
        df_final = df_final.applymap(str)
        

      

        #TIRAR ASPAS DOS VALORES, DATA E NUMEROS

        
        
        # Formatando os valores monetários (não convertendo para string, mantendo como numérico)
        df_final['Fat.Total'] = df_final['Fat.Total'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
        df_final['Serv/Tx'] = df_final['Serv/Tx'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
        df_final['Fat.Real'] = df_final['Fat.Real'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
        df_final['Ticket'] = df_final['Ticket'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)

        # Garantir datetime sem aspas
        df_final['Data'] = pd.to_datetime(df_final['Data'].astype(str).str.replace("'", "").str.strip(), dayfirst=True)

        # Converter para número serial (dias desde 1899-12-30, padrão do Excel/Sheets)
        df_final['Data'] = (df_final['Data'] - pd.Timestamp("1899-12-30")).dt.days
       
        # Corrigir coluna Ano: remover aspas, espaços e garantir que seja inteiro
        df_final['Ano'] = df_final['Ano'].apply(
        lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
        )

        # Corrigir colunas D e F (Código Everest e Código Grupo Everest): remover aspas e converter para número inteiro
        df_final['Código Everest'] = df_final['Código Everest'].apply(
        lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
        )

        df_final['Código Grupo Everest'] = df_final['Código Grupo Everest'].apply(
        lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
        )
        
        
        # Conectar ao Google Sheets
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
        gc = gspread.authorize(credentials)

        planilha_destino = gc.open("Faturamento Sistema Externo")
        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

        # Obter dados já existentes na aba
        valores_existentes = aba_destino.get_all_values()

        # Criar um conjunto de linhas existentes na coluna M (usada para verificar duplicação)
        dados_existentes = set([linha[12] for linha in valores_existentes[1:]])  # Ignorando cabeçalho, coluna M é a 13ª (índice 12)

        novos_dados = []
        duplicados = []  # Armazenar os registros duplicados
        rows = df_final.fillna("").values.tolist()

       

        
        # Verificar duplicação somente na coluna "M"
        for linha in rows:
            chave_m = linha[-1]  # A chave da coluna M (última coluna)
            if chave_m not in dados_existentes:
                novos_dados.append(linha)
                dados_existentes.add(chave_m)  # Adiciona a chave da linha para não enviar novamente
            else:
                duplicados.append(linha)  # Adiciona a linha duplicada à lista

        # Adicionar o botão de atualização do Google Sheets
        if st.button("📥 Enviar dados para o Google Sheets"):
            with st.spinner("🔄 Atualizando o Google Sheets..."):
                try:
                    if novos_dados:
                        # Manter a primeira linha vazia para começar a inserção
                        primeira_linha_vazia = len(valores_existentes) + 1
                        
                        # Enviar os novos dados para o Google Sheets
                        aba_destino.update(f"A{primeira_linha_vazia}", novos_dados)

# ASPAS RESOLVIDO
                        
                        # 🔧 Aplicar formatação de data na coluna A (Data) - prbblema de aspas resolvido
                        from gspread_formatting import CellFormat, NumberFormat, format_cell_range

                        data_format = CellFormat(
                            numberFormat=NumberFormat(type='DATE', pattern='dd/mm/yyyy')
                        )

                        # 🔢 Formato para coluna Ano como número sem aspas
                        numero_format = CellFormat(
                        numberFormat=NumberFormat(type='NUMBER', pattern='0')
                        )
                      
                        
                        # Considerando que a coluna A é onde está a data
                        format_cell_range(aba_destino, f"A2:A{primeira_linha_vazia + len(novos_dados)}", data_format)
                        format_cell_range(aba_destino, f"L2:L{primeira_linha_vazia + len(novos_dados)}", numero_format)  
                        format_cell_range(aba_destino, f"D2:D{primeira_linha_vazia + len(novos_dados)}", numero_format)
                        format_cell_range(aba_destino, f"F2:F{primeira_linha_vazia + len(novos_dados)}", numero_format)




                        
                        st.success(f"✅ {len(novos_dados)} novo(s) registro(s) enviado(s) com sucesso para o Google Sheets!")

                    if duplicados:
                        st.warning(f"⚠️ {len(duplicados)} registro(s) foram duplicados e não foram enviados para o Google Sheets.")
                        # Exibir as linhas duplicadas para o usuário
                   #     st.write("Registros Duplicados:", duplicados)

                   # else:
                    #    st.info("✅ Dados atualizados google sheets.")
                except Exception as e:
                    st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")

    else:
        st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")



# ================================
# 📈 Relatórios Gerenciais (Painel Interativo)
# ================================

with aba4:
    # Conectar ao Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)

    # Carregar dados da aba
    planilha = gc.open("Faturamento Sistema Externo")
    aba = planilha.worksheet("Fat Sistema Externo")
    dados = aba.get_all_records()
    df = pd.DataFrame(dados)



    # Tratamento de dados
    def limpar_valor(x):
        try:
            if isinstance(x, str):
                return float(x.replace(".", "").replace(",", "."))
            elif isinstance(x, (int, float)):
                return x
            else:
                return None
        except:
            return None

    # Aplicar para as colunas de valor
    for coluna in ["Fat.Total", "Serv/Tx", "Fat.Real"]:
        if coluna in df.columns:
            df[coluna] = df[coluna].apply(limpar_valor)

    df_filtrado = df.copy()  # usar todo o DataFrame

    # Gráfico 5: Comparativo de Faturamento Real 2024 vs 2025
    df["Ano"] = pd.to_numeric(df["Ano"], errors="coerce")
    df["Mês"] = pd.to_numeric(df["Mês"], errors="coerce")

    # Filtrar 2024 e 2025
    df_anos = df[df["Ano"].isin([2024, 2025])].dropna(subset=["Mês", "Ano", "Fat.Real"])

    # Criar coluna 'Mês-Ano'
    meses_nome = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    def mes_ano_formatado(row):
        try:
            mes = int(row["Mês"])
            ano = int(row["Ano"])
            return f"{meses_nome.get(mes, str(mes))} {ano}"
        except:
            return "Desconhecido"

    df_anos["Mês-Ano"] = df_anos.apply(mes_ano_formatado, axis=1)

    # Agrupar por Mês-Ano
    fat_mensal = df_anos.groupby("Mês-Ano")["Fat.Real"].sum().reset_index()

    # Total por ano
    total_por_ano = df_anos.groupby("Ano")["Fat.Real"].sum().reset_index()
    total_por_ano["Mês-Ano"] = total_por_ano["Ano"].astype(str) + " (Total)"
    total_por_ano = total_por_ano[["Mês-Ano", "Fat.Real"]]

    # Concatenar final
    df_barras = pd.concat([fat_mensal, total_por_ano], ignore_index=True)

    # Ordenar corretamente
    ordem_meses = [f"{meses_nome[m]} 2024" for m in range(1, 13)] + \
                  [f"{meses_nome[m]} 2025" for m in range(1, 13)] + \
                  ["2024 (Total)", "2025 (Total)"]
    df_barras["Mês-Ano"] = pd.Categorical(df_barras["Mês-Ano"], categories=ordem_meses, ordered=True)
    df_barras = df_barras.sort_values("Mês-Ano")

    # Exibir tabelas para validação
    #st.subheader("📊 Tabela de Faturamento para Verificação")
    #st.dataframe(df_barras)

    #st.subheader("🔍 Dados brutos de 2024 e 2025")
    #st.dataframe(df_anos[["Ano", "Mês", "Fat.Real"]].head(20))

    #st.write("Tipo da coluna Fat.Real:", df["Fat.Real"].dtype)
    #st.write("Primeiros valores:", df["Fat.Real"].head(10))

    # Plotar o gráfico final
    fig5 = px.bar(df_barras, x="Mês-Ano", y="Fat.Real", title="Faturamento Real Mensal - 2024 vs 2025")
    fig5.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig5, use_container_width=True)

