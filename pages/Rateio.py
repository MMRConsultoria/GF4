# ==== Calcula % Total e Rateio proporcional ao subtotal do Tipo ====
df_final["% Total"] = 0.0
df_final["Rateio"] = 0.0

for tipo in df_final["Tipo"].unique():
    # Máscara das linhas normais do tipo (exclui subtotal e TOTAL)
    mask_tipo = (
        (df_final["Tipo"] == tipo) &
        (~df_final["Grupo"].str.startswith("Subtotal")) &
        (df_final["Grupo"] != "TOTAL")
    )
    
    # Subtotal do tipo
    subtotal_tipo = df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "Total"].sum()
    
    if subtotal_tipo > 0:
        # % Total relativo ao subtotal do tipo
        df_final.loc[mask_tipo, "% Total"] = (
            df_final.loc[mask_tipo, "Total"] / subtotal_tipo
        ) * 100
    
    # Linha subtotal do tipo = 100%
    df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "% Total"] = 100
    
    # Rateio = % da linha × valor digitado para o tipo
    valor_rateio = valores_rateio_por_tipo.get(tipo, 0.0)
    df_final.loc[mask_tipo, "Rateio"] = (
        df_final.loc[mask_tipo, "% Total"] / 100 * valor_rateio
    )
    
    # Linha subtotal do tipo recebe soma das lojas
    df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "Rateio"] = \
        df_final.loc[mask_tipo, "Rateio"].sum()

# 🔹 Linha TOTAL
# TOTAL no % Total = sempre 100%
df_final.loc[df_final["Grupo"] == "TOTAL", "% Total"] = 100

# TOTAL no Rateio = soma apenas dos subtotais
df_final.loc[df_final["Grupo"] == "TOTAL", "Rateio"] = \
    df_final.loc[df_final["Grupo"].str.startswith("Subtotal"), "Rateio"].sum()
