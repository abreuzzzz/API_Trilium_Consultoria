import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 🔐 Lê o segredo e salva como credentials.json
gdrive_credentials = os.getenv("GDRIVE_SERVICE_ACCOUNT")
with open("credentials.json", "w") as f:
    json.dump(json.loads(gdrive_credentials), f)

# 📌 Autenticação com Google
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# === IDs das planilhas ===
planilhas_ids = {
    "Financeiro_contas_a_receber_Trilium": "1TmlB3my6KAb-6CRXUtZPJaOmTL-yW_SJou0akLPyj7A",
    "Financeiro_contas_a_pagar_Trilium": "107Gjl8BZ6kWNguIM15wCsM6scaCM1pz-kHoseJ-zAVM",
    "Financeiro_Completo_Trilium": "1A08gZWPn0N9OIQXPsuOoHqah_IycFXGBWcwRVR3NOCE"
}

print("🗑️ Iniciando exclusão COMPLETA de todas as linhas das planilhas...")

# 1. Limpa TUDO de Contas a Receber
print("\n📋 Limpando: Financeiro_contas_a_receber_Trilium")
planilha_receber = client.open_by_key(planilhas_ids["Financeiro_contas_a_receber_Trilium"])
aba_receber = planilha_receber.sheet1
aba_receber.clear()
print("  ✅ Todas as linhas excluídas (incluindo cabeçalho)")

# 2. Limpa TUDO de Contas a Pagar
print("\n📋 Limpando: Financeiro_contas_a_pagar_Trilium")
planilha_pagar = client.open_by_key(planilhas_ids["Financeiro_contas_a_pagar_Trilium"])
aba_pagar = planilha_pagar.sheet1
aba_pagar.clear()
print("  ✅ Todas as linhas excluídas (incluindo cabeçalho)")

# 3. Limpa TUDO de Financeiro Completo - Aba principal (sheet1)
print("\n📋 Limpando: Financeiro_Completo_Trilium (sheet1)")
planilha_completo = client.open_by_key(planilhas_ids["Financeiro_Completo_Trilium"])
aba_completo = planilha_completo.sheet1
aba_completo.clear()
print("  ✅ Todas as linhas excluídas (incluindo cabeçalho)")

# 4. Limpa TUDO de Financeiro Completo - Aba Dados_Pivotados (se existir)
print("\n📋 Limpando: Financeiro_Completo_Trilium (Dados_Pivotados)")
try:
    aba_pivotada = planilha_completo.worksheet("Dados_Pivotados")
    aba_pivotada.clear()
    print("  ✅ Todas as linhas excluídas (incluindo cabeçalho)")
except:
    print("  ⚠️ Aba 'Dados_Pivotados' não encontrada")

print("\n🎉 Limpeza completa concluída com sucesso!")
print("⚠️ ATENÇÃO: Todas as linhas foram removidas, incluindo os cabeçalhos")
