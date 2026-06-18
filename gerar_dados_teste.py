"""
Gerador de Dados Fictícios para Teste — Versão Local
=====================================================
Cria todos os arquivos necessários para testar o projeto
sem Databricks e sem dados reais.

USO:
    python gerar_dados_teste_local.py
"""

import json
import os
from pathlib import Path

import numpy as np
import pandas as pd

BASE_PATH = Path(__file__).parent

print("=" * 60)
print("  Gerador de Dados Fictícios")
print("=" * 60)

# =============================================================================
# 1. Planilha de Apuração Fictícia
# =============================================================================

ENTRADA_PATH = BASE_PATH / "Entrada"
ENTRADA_PATH.mkdir(exist_ok=True)

# Associados fictícios
associados = [
    {"nome": "Ana Silva", "email": "ana.silva@exemplo.com", "regional": "SUDESTE"},
    {"nome": "Carlos Souza", "email": "carlos.souza@exemplo.com", "regional": "SUL"},
    {"nome": "Maria Santos", "email": "maria.santos@exemplo.com", "regional": "NORTE"},
    {"nome": "Pedro Oliveira", "email": "pedro.oliveira@exemplo.com", "regional": "NORDESTE"},
    {"nome": "Julia Costa", "email": "julia.costa@exemplo.com", "regional": "CENTRO-OESTE"},
]

# Indicadores fictícios
indicadores = [
    ("Vendas - Sell In GSV", "Vendas - Sell In GSV - Categoria A"),
    ("Vendas - Sell In GSV", "Vendas - Sell In GSV - Categoria B"),
    ("Vendas - Sell In GSV", "Vendas - Sell In GSV - TOTAL"),
    ("Execução - Sales Strike", "Execução - Sales Strike - Indicador 1"),
    ("Execução - Sales Strike", "Execução - Sales Strike - Indicador 2"),
    ("Execução - Sales Strike", "Execução - Sales Strike - TOTAL"),
]

np.random.seed(42)
rows = []
for assoc in associados:
    for ind1, ind2 in indicadores:
        row = {
            "Indicador_1": ind1,
            "Associado": assoc["nome"],
            "Email": assoc["email"],
            "Regional": assoc["regional"],
            "Indicador_2": ind2,
        }
        actual = np.random.randint(50000, 200000)
        meta = np.random.randint(80000, 180000)
        pct_ating = actual / meta if meta > 0 else 0
        repres = np.random.uniform(0.05, 0.30)
        variavel = np.random.uniform(0, 0.15)
        const_marca = np.random.uniform(0, 0.05)
        ating_regional = np.random.uniform(0.7, 1.2)
        pagamento = variavel * pct_ating * ating_regional

        row["P04_Actual"] = actual
        row["P04_Meta"] = meta
        row["P04_%Ating"] = round(pct_ating, 4)
        row["P04_%Repres"] = round(repres, 4)
        row["P04_Variavel"] = round(variavel, 4)
        row["P04_ConstMarca"] = round(const_marca, 4)
        row["P04_AtingRegional"] = round(ating_regional, 4)
        row["P04_Pagamento"] = round(pagamento, 4)

        # Quarter Q2
        q_actual = np.random.randint(200000, 600000)
        q_meta = np.random.randint(300000, 550000)
        q_resultado = q_actual / q_meta if q_meta > 0 else 0
        q_recup_vendas = np.random.uniform(0, 0.08)
        q_recup_marca = np.random.uniform(0, 0.05)
        q_recup_regional = np.random.uniform(0, 0.06)
        q_total = q_recup_vendas + q_recup_marca + q_recup_regional

        row["Q2_Actual"] = q_actual
        row["Q2_Meta"] = q_meta
        row["Q2_Resultado"] = round(q_resultado, 4)
        row["Q2_RecupVendas"] = round(q_recup_vendas, 4)
        row["Q2_RecupMarca"] = round(q_recup_marca, 4)
        row["Q2_RecupRegional"] = round(q_recup_regional, 4)
        row["Q2_Total"] = round(q_total, 4)

        rows.append(row)

df = pd.DataFrame(rows)

# Monta planilha com estrutura multi-header
final_rows = []

# Linha 0: períodos
periodo_row = [""] * 5 + ["P04"] * 8 + ["Quarter"] * 7
final_rows.append(periodo_row)

# Linha 1: sub-labels
sub_row = [""] * 5
sub_row += ["Actual", "Meta", "% Ating Meta", "%Repres.",
            "Variável (%)", "Const. Marca", "Ating Regional", "Pagamento"]
sub_row += ["Actual", "Meta", "% Ating Meta",
            "Recuperação Vendas", "Recuperação Marca",
            "Recuperação Regional", "Pagamento"]
final_rows.append(sub_row)

# Linha 2: header principal
header_row = [
    "Indicador", "Associado", "Email", "Regional", "Indicador",
    "Actual", "TGT", "% Ating Meta", "%Repres.",
    "Variável (%)", "Const. Marca", "Atingimento Regional",
    "TOTAL VARIÁVEL ATINGIMENTO INDIVIDUAL",
    "Actual", "Meta", "Resultado",
    "Recuperação Vendas", "Recuperação Marca",
    "Recuperação Regional", "% Recuperação Total"
]
final_rows.append(header_row)

# Dados
for _, r in df.iterrows():
    data_row = [
        r["Indicador_1"], r["Associado"], r["Email"],
        r["Regional"], r["Indicador_2"],
        r["P04_Actual"], r["P04_Meta"], r["P04_%Ating"],
        r["P04_%Repres"], r["P04_Variavel"],
        r["P04_ConstMarca"], r["P04_AtingRegional"], r["P04_Pagamento"],
        r["Q2_Actual"], r["Q2_Meta"], r["Q2_Resultado"],
        r["Q2_RecupVendas"], r["Q2_RecupMarca"],
        r["Q2_RecupRegional"], r["Q2_Total"]
    ]
    final_rows.append(data_row)

df_final = pd.DataFrame(final_rows)
arquivo_excel = ENTRADA_PATH / "apuracao_Q2_teste.xlsx"
df_final.to_excel(
    arquivo_excel, sheet_name="RESULTADO VARI\u00c1VEL",
    index=False, header=False, engine="openpyxl"
)
print(f"\n\u2705 Planilha: {arquivo_excel}")
print(f"   {len(associados)} associados \u00d7 {len(indicadores)} indicadores = {len(rows)} linhas")

# =============================================================================
# 2. Mapeamento de Gestores Regionais
# =============================================================================

df_cc = pd.DataFrame({
    "Regional": ["NORTE", "NORDESTE", "CENTRO-OESTE", "SUDESTE", "SUL"],
    "Email_CC": [
        "gestor.norte@exemplo.com",
        "gestor.nordeste@exemplo.com",
        "gestor.centrooeste@exemplo.com",
        "gestor.sudeste@exemplo.com",
        "gestor.sul@exemplo.com"
    ]
})
arquivo_cc = BASE_PATH / "gestores_regionais_cc.xlsx"
df_cc.to_excel(arquivo_cc, index=False, engine="openpyxl")
print(f"\u2705 Gestores: {arquivo_cc}")

# =============================================================================
# 3. Arquivo de Segredos (se não existir)
# =============================================================================

secrets_path = BASE_PATH / ".secrets_config.json"
if not secrets_path.exists():
    secrets = {
        "power_automate_url": "https://prod-00.westus.logic.azure.com:443/workflows/EXEMPLO/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ASSINATURA_EXEMPLO",
        "assinatura_nome": "Nome Teste",
        "assinatura_cargo": "Cargo / Departamento Teste",
        "assinatura_email": "teste@exemplo.com",
        "email_gestor_excecoes": "gestor.excecoes@exemplo.com",
        "nome_gestor_excecoes": "Gestor Teste"
    }
    with open(secrets_path, "w", encoding="utf-8") as f:
        json.dump(secrets, f, ensure_ascii=False, indent=2)
    print(f"\u2705 Secrets: {secrets_path}")
else:
    print(f"\u26a0\ufe0f  Secrets j\u00e1 existe: {secrets_path} (n\u00e3o sobrescrito)")

# =============================================================================
# 4. Pasta de Saída
# =============================================================================

(BASE_PATH / "Saida").mkdir(exist_ok=True)
print(f"\u2705 Pasta Saida/ criada")

# =============================================================================
# Resumo
# =============================================================================

print(f"\n{'='*60}")
print("  \ud83c\udf89 SETUP COMPLETO!")
print(f"{'='*60}")
print()
print("Pr\u00f3ximos passos:")
print()
print("  1. Preview (sem enviar):")
print("     python email_metas_local.py --preview")
print()
print("  2. Teste (envia para seu e-mail):")
print("     python email_metas_local.py --email-teste seu@email.com")
print()
print("  3. Abra Saida/preview.html no navegador para ver o resultado.")
print()
print("  \u26a0\ufe0f  O envio s\u00f3 funciona se a URL do Power Automate")
print("     em .secrets_config.json apontar para um flow real.")
