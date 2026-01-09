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

def limpar_aba_completa(aba, nome_aba):
    """Limpa conteúdo E formatação de uma aba"""
    print(f"  🗑️ Limpando conteúdo de {nome_aba}...")
    aba.clear()
    
    print(f"  🎨 Removendo formatação de {nome_aba}...")
    aba.format('A:ZZ', {
        "numberFormat": {"type": "TEXT"},  # Força formato texto
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},  # Branco
        "textFormat": {
            "bold": False,
            "italic": False,
            "foregroundColor": {"red": 0, "green": 0, "blue": 0}
        }
    })
    print(f"  ✅ {nome_aba} limpa e formatação resetada")

print("🗑️ Iniciando exclusão COMPLETA de todas as linhas das planilhas...")

# 1. Limpa Contas a Receber
print("\n📋 Processando: Financeiro_contas_a_receber_Trilium")
planilha_receber = client.open_by_key(planilhas_ids["Financeiro_contas_a_receber_Trilium"])
limpar_aba_completa(planilha_receber.sheet1, "Contas a Receber")

# 2. Limpa Contas a Pagar
print("\n📋 Processando: Financeiro_contas_a_pagar_Trilium")
planilha_pagar = client.open_by_key(planilhas_ids["Financeiro_contas_a_pagar_Trilium"])
limpar_aba_completa(planilha_pagar.sheet1, "Contas a Pagar")

# 3. Limpa Financeiro Completo - Aba principal
print("\n📋 Processando: Financeiro_Completo_Trilium (sheet1)")
planilha_completo = client.open_by_key(planilhas_ids["Financeiro_Completo_Trilium"])
limpar_aba_completa(planilha_completo.sheet1, "Financeiro Completo - Principal")

# 4. Limpa Dados_Pivotados
print("\n📋 Processando: Financeiro_Completo_Trilium (Dados_Pivotados)")
try:
    aba_pivotada = planilha_completo.worksheet("Dados_Pivotados")
    limpar_aba_completa(aba_pivotada, "Dados Pivotados")
except:
    print("  ⚠️ Aba 'Dados_Pivotados' não encontrada")

print("\n🎉 Limpeza completa concluída com sucesso!")
print("⚠️ ATENÇÃO: Conteúdo e formatação removidos. Células agora estão em formato TEXTO")
