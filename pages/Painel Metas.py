with aba1:
    st.header("📊 Comparativo Metas vs Realizado")

    # Carrega aba de metas
    df_metas = pd.DataFrame(planilha_empresa.worksheet("Metas").get_all_records())

    # Aplica De Para Metas
    df_depara = df_empresa[["Loja", "De Para Metas"]].dropna()
    df_metas = df_metas.merge(df_depara, left_on="Loja", right_on="De Para Metas", how="left")
    df_metas["Loja Final"] = df_metas["Loja_y"].fillna(df_metas["Loja_x"])

    # Agrupa metas
    df_metas_grouped = df_metas.groupby(["Ano", "Mês", "Loja Final"])["Meta"].sum().reset_index()

    # Agrupa realizado (substitua df_fat por seu DataFrame real de vendas)
    df_realizado_grouped = df_fat_sistema_externo.groupby(["Ano", "Mês", "Loja"])["Fat.Total"].sum().reset_index()

    # Aplica De Para também no realizado (se necessário)
    df_realizado_grouped = df_realizado_grouped.merge(df_depara, left_on="Loja", right_on="Loja", how="left")
    df_realizado_grouped["Loja Final"] = df_realizado_grouped["De Para Metas"].fillna(df_realizado_grouped["Loja"])

    df_realizado_grouped = df_realizado_grouped.groupby(["Ano", "Mês", "Loja Final"])["Fat.Total"].sum().reset_index()

    # Junta metas e realizado
    df_comp = df_metas_grouped.merge(df_realizado_grouped, on=["Ano", "Mês", "Loja Final"], how="outer")
    df_comp = df_comp.fillna(0)
    df_comp["% Atingimento"] = np.where(
        df_comp["Meta"] > 0,
        df_comp["Fat.Total"] / df_comp["Meta"],
        np.nan
    )

    # Exibe tabela
    st.dataframe(df_comp.style.format({
        "Meta": "R$ {:,.2f}",
        "Fat.Total": "R$ {:,.2f}",
        "% Atingimento": "{:.2%}"
    }))
