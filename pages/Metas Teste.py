import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Oculta a janela principal do Tkinter
root = tk.Tk()
root.withdraw()

# Abre diálogo para selecionar o arquivo
arquivo_excel = filedialog.askopenfilename(
    title="Selecione o arquivo Excel",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

print(f"Arquivo escolhido: {arquivo_excel}")

# Lê todas as abas
xls = pd.ExcelFile(arquivo_excel)
print(f"Abas encontradas: {xls.sheet_names}")

dfs = []
for aba in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=aba)
    df["Grupo"] = aba
    dfs.append(df)

# Consolida
df_total = pd.concat(dfs, ignore_index=True)
print("Consolidado com sucesso.")

# Salva só para testar
df_total.to_excel("Importador.xlsx", index=False)
print("Arquivo Importador.xlsx criado.")
