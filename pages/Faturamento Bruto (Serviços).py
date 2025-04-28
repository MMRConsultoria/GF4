# ================================
# 🔄 Aba 3 - Atualizar Google Sheets (Evitar duplicação e erro de Timestamp)
# ================================

with aba3:
    st.header("📤 Atualizar Banco de Dados (Evitar duplicação e erro de Timestamp)")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final.copy()

        # Garantir que todas as colunas de 'Data' sejam convertidas para string antes de enviar
        df_final['Data'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')

        # Função para garantir que os valores sejam números reais com vírgula como separador decimal
        def format_monetary(value):
            try:
                # Verificar se o valor é numérico antes de aplicar a formatação
                if isinstance(value, str):
                    value = float(value.replace(',', '.'))  # Convertendo para número com ponto
                    # Formatando para garantir que tenha vírgula sem ponto de milhar
                    return f"{value:.2f}".replace(".", ",")
                else:
                    # Se o valor não for string, retornar como está
                    return value
            except (ValueError, TypeError):
                # Se não puder converter, retornar o valor original
                return value

        # Formatando as colunas monetárias
        df_final['Fat.Total'] = df_final['Fat.Total'].apply(format_monetary)
        df_final['Serv/Tx'] = df_final['Serv/Tx'].apply(format_monetary)
        df_final['Fat.Real'] = df_final['Fat.Real'].apply(format_monetary)
        df_final['Ticket'] = df_final['Ticket'].apply(format_monetary)

        # Não estamos mais aplicando `applymap(str)` a todo o DataFrame. Só formatamos as colunas monetárias.

        if st.button("📥 Enviar dados para o Google Sheets"):
            with st.spinner("🔄 Atualizando o Google Sheets..."):
                try:
                    # Conectar ao Google Sheets
                    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
                    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                    gc = gspread.authorize(credentials)

                    planilha_destino = gc.open("Faturamento Sistema Externo")
                    aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                    # Obter dados já existentes no Google Sheets
                    valores_existentes = aba_destino.get_all_values()

                    # Criar um conjunto de linhas já existentes para comparação
                    dados_existentes = set([tuple(linha) for linha in valores_existentes[1:]])  # Ignorando cabeçalho

                    novos_dados = []
                    rows = df_final.fillna("").values.tolist()

                    for linha in rows:
                        if tuple(linha) not in dados_existentes:
                            novos_dados.append(linha)
                            dados_existentes.add(tuple(linha))  # Adiciona a linha para não enviar novamente

                    if novos_dados:
                        primeira_linha_vazia = len(valores_existentes) + 1
                        aba_destino.update(f"A{primeira_linha_vazia}", novos_dados)

                        st.success(f"✅ {len(novos_dados)} novo(s) registro(s) enviado(s) com sucesso para o Google Sheets!")
                    else:
                        st.info("✅ Não há novos dados para atualizar.")

                except Exception as e:
                    st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")

    else:
        st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")
