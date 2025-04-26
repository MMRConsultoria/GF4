# ================================
# 🔄 Aba 3 - Atualizar Google Sheets
# ================================
with aba3:
    st.header("🔄 Atualizar Google Sheets")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final

        if 'atualizou_google' not in st.session_state:
            st.session_state.atualizou_google = False

        if not st.session_state.atualizou_google:
            if st.button("📤 Atualizar no Google Sheets"):
                with st.spinner('🔄 Atualizando...'):
                    try:
                        # Abrir a planilha e a aba
                        planilha_destino = gc.open("Faturamento Sistema Externo")
                        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                        # 🔥 CORRIGIDO: ler manualmente os dados (sem erro de header)
                        dados_raw = aba_destino.get_all_values()
                        dados_existentes = pd.DataFrame(dados_raw[1:], columns=dados_raw[0])

                        # Limpar espaços das colunas críticas
                        if not dados_existentes.empty:
                            for col in ["Data", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
                                if col in dados_existentes.columns:
                                    dados_existentes[col] = dados_existentes[col].astype(str).str.strip()

                        # Prepara os novos dados (sem cabeçalho)
                        novos_dados = df_final.copy()
                        for col in ["Data", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
                            novos_dados[col] = novos_dados[col].astype(str).str.strip()

                        # Chaves para detectar duplicados
                        colunas_chave = [
                            "Data", "Loja", "Código Everest", "Grupo",
                            "Código Grupo Everest", "Fat.Total", "Serv/Tx",
                            "Fat.Real", "Ticket", "Mês", "Ano"
                        ]

                        # Merge para encontrar dados novos
                        merged = pd.merge(novos_dados, dados_existentes, on=colunas_chave, how="left", indicator=True)
                        registros_novos = merged[merged["_merge"] == "left_only"].drop(columns="_merge")

                        if registros_novos.empty:
                            st.info("✅ Nenhum novo registro para atualizar.")
                            st.session_state.atualizou_google = True
                        else:
                            # Só os valores (sem cabeçalho)
                            rows = registros_novos.fillna("").values.tolist()

                            # Descobrir a primeira linha vazia (considerando cabeçalho na linha 1)
                            primeira_linha_vazia = len(dados_existentes) + 2

                            # Atualizar no Google Sheets sem tocar o cabeçalho
                            aba_destino.update(f"A{primeira_linha_vazia}", rows)

                            st.success(f"✅ {len(rows)} novo(s) registro(s) enviado(s) para o Google Sheets!")
                            st.session_state.atualizou_google = True

                            registros_ignorados = len(novos_dados) - len(rows)
                            if registros_ignorados > 0:
                                st.warning(f"⚠️ {registros_ignorados} registro(s) já existiam e foram ignorados.")

                    except Exception as e:
                        st.error(f"❌ Erro ao atualizar: {e}")
                        st.session_state.atualizou_google = False
        else:
            st.info("✅ Dados já foram atualizados nesta sessão.")
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo na aba anterior.")
