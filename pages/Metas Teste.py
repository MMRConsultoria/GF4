import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ================================
# Configuração Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

import json
with open("seu_arquivo_credenciais.json") as f:
    credentials_dict = json.load(f)

credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

# Abre a planilha pelo ID
planilha = gc.open_by_key("1ZaRXVZyv7WZ8xJ8yGEViRibZ-sGoilGO")
aba = planilha.worksheet("Metas 1")

# Teste: lê as primeiras 5 linhas
dados = aba.get_all_values()
for linha in dados[:5]:
    print(linha)
