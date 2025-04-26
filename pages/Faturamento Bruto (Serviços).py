# ================================
# 🔄 Aba 3 - Atualizar Google Sheets (com contagem de registros novos e duplicados)
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
                        # Abrir a planilha e aba de destino
                        planilha_destino = gc.open("Faturamento Sistema Externo")
                        aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                        # Ler dados existentes
                        dados_raw = aba_destino.get_all_values()

                        # Preparar dados existentes (linhas existentes sem o cabeçalho)
                        dados_existentes = [ [str(cell).strip() for cell in row] for row in dados_raw[1:] ]

                        # Preparar novos dados
                        novos_dados = [ [str(cell).strip() for cell in row] for row in df_final.values.tolist() ]

                        # Verificar novos registros
                        registros_novos = [linha for linha in novos_dados if linha not in dados_existentes]

                        total_novos = len(registros_novos)
                        total_existentes = len(novos_dados) - total_novos

                        if total_novos == 0:
                            st.info(f"✅ Nenhum novo registro para atualizar. {total_existentes} registro(s) já existiam no Google Sheets.")
                            st.session_state.atualizou_google = True
                        else:
                            # Descobrir onde colar
                            primeira_linha_vazia = len(dados_raw) + 1  # linha após os dados

                            # Atualizar
                            aba_destino.update(f"A{primeira_linha_vazia}", registros_novos)

                            st.success(f"✅ {total_novos} novo(s) registro(s) enviado(s) para o Google Sheets!")
                            if total_existentes > 0:
                                st.warning(f"⚠️ {total_existentes} registro(s) já existiam e não foram importados.")
                            st.session_state.atualizou_google = True

                    except Exception as e:
                        st.error(f"❌ Erro ao atualizar: {e}")
                        st.session_state.atualizou_google = False
        else:
            st.info("✅ Dados já foram atualizados nesta sessão.")
    else:
        st.info("⚠️ Primeiro, faça o upload e processamento do arquivo na aba anterior.")
