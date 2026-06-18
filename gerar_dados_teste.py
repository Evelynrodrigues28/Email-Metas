# Databricks notebook source
# =============================================================================
# Script para gerar dados fictícios de teste
# Execute este script ANTES do notebook principal para criar:
#   - Entrada/apuracao_Q2_teste.xlsx (planilha com dados fictícios)
#   - gestores_regionais_cc.xlsx (mapeamento de regionais)
#   - .secrets_config.json (configuração com dados de teste)
# =============================================================================

# COMMAND ----------
# %pip install openpyxl

import os
import json
import pandas as pd
import numpy as np
from pathlib import Path

# --- Caminho Base ---
_nb_path = (
    dbutils.notebook.entry_point.getDbutils()
    .notebook().getContext().notebookPath().get()
)
BASE_PATH = "/Workspace" + "/".join(_nb_path.split("/")[:-1])

print(f"📁 Base: {BASE_PATH}")
print("="*60)

# COMMAND ----------
# Gera planilha de apuração fictícia

ENTRADA_PATH = f"{BASE_PATH}/Entrada"
os.makedirs(ENTRADA_PATH, exist_ok=True)

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

# Períodos mensais
periodos = [f"P{str(i).zfill(2)}" for i in range(1, 14)]

# Gera os dados
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
        # Gera métricas para cada período
        for p in periodos:
            actual = np.random.randint(50000, 200000)
            meta = np.random.randint(80000, 180000)
            pct_ating = actual / meta if meta > 0 else 0
            repres = np.random.uniform(0.05, 0.30)
            variavel = np.random.uniform(0, 0.15)
            const_marca = np.random.uniform(0, 0.05)
            ating_regional = np.random.uniform(0.7, 1.2)
            pagamento = variavel * pct_ating * ating_regional

            row[f"{p}_Actual"] = actual
            row[f"{p}_Meta"] = meta
            row[f"{p}_%Ating"] = round(pct_ating, 4)
            row[f"{p}_%Repres"] = round(repres, 4)
            row[f"{p}_Variavel"] = round(variavel, 4)
            row[f"{p}_ConstMarca"] = round(const_marca, 4)
            row[f"{p}_AtingRegional"] = round(ating_regional, 4)
            row[f"{p}_Pagamento"] = round(pagamento, 4)

        # Gera métricas de quarter (Q1-Q4)
        for q in ["Q1", "Q2", "Q3", "Q4"]:
            q_actual = np.random.randint(200000, 600000)
            q_meta = np.random.randint(300000, 550000)
            q_resultado = q_actual / q_meta if q_meta > 0 else 0
            q_recup_vendas = np.random.uniform(0, 0.08)
            q_recup_marca = np.random.uniform(0, 0.05)
            q_recup_regional = np.random.uniform(0, 0.06)
            q_total = q_recup_vendas + q_recup_marca + q_recup_regional

            row[f"{q}_Actual"] = q_actual
            row[f"{q}_Meta"] = q_meta
            row[f"{q}_Resultado"] = round(q_resultado, 4)
            row[f"{q}_RecupVendas"] = round(q_recup_vendas, 4)
            row[f"{q}_RecupMarca"] = round(q_recup_marca, 4)
            row[f"{q}_RecupRegional"] = round(q_recup_regional, 4)
            row[f"{q}_Total"] = round(q_total, 4)

        rows.append(row)

df = pd.DataFrame(rows)

# Monta a planilha no formato esperado pelo notebook
# (com headers multi-nível simulando o layout real)
with pd.ExcelWriter(
    f"{ENTRADA_PATH}/apuracao_Q2_teste.xlsx",
    engine="openpyxl"
) as writer:
    # Cria a aba com o nome esperado
    # Linha 0: vazia
    # Linha 1: períodos (P04 repetido 8x, depois Quarter repetido 7x)
    # Linha 2: sub-headers
    # Linha 3: header principal (Indicador, Associado, Email, Regional, etc)
    # Linha 4+: dados

    header_cols = ["Indicador", "Associado", "Email", "Regional", "Indicador"]
    periodo_cols = ["Actual", "Meta", "%Ating", "%Repres", "Variavel",
                    "ConstMarca", "AtingRegional", "Pagamento"]
    quarter_cols = ["Actual", "Meta", "Resultado", "RecupVendas",
                    "RecupMarca", "RecupRegional", "Total"]

    # Monta DataFrame final com estrutura correta
    final_rows = []

    # Linha de períodos (row 0)
    periodo_row = [""] * 5
    periodo_row += ["P04"] * 8
    periodo_row += ["Quarter"] * 7
    final_rows.append(periodo_row)

    # Linha de sub-labels (row 1)
    sub_row = [""] * 5
    sub_row += ["Actual", "Meta", "% Ating Meta", "%Repres.",
                "Variável (%)", "Const. Marca", "Ating Regional", "Pagamento"]
    sub_row += ["Actual", "Meta", "% Ating Meta",
                "Recuperação Vendas", "Recuperação Marca",
                "Recuperação Regional", "Pagamento"]
    final_rows.append(sub_row)

    # Linha de header (row 2)
    header_row = ["Indicador", "Associado", "Email", "Regional", "Indicador",
                  "Actual", "TGT", "% Ating Meta", "%Repres.",
                  "Variável (%)", "Const. Marca", "Atingimento Regional",
                  "TOTAL VARIÁVEL ATINGIMENTO INDIVIDUAL",
                  "Actual", "Meta", "Resultado",
                  "Recuperação Vendas", "Recuperação Marca",
                  "Recuperação Regional", "% Recuperação Total"]
    final_rows.append(header_row)

    # Dados (row 3+)
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
    df_final.to_excel(
        writer, sheet_name="RESULTADO VARIÁVEL",
        index=False, header=False
    )

print(f"✅ Planilha criada: {ENTRADA_PATH}/apuracao_Q2_teste.xlsx")
print(f"   {len(associados)} associados × {len(indicadores)} indicadores = {len(rows)} linhas")

# COMMAND ----------
# Gera gestores_regionais_cc.xlsx

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
df_cc.to_excel(f"{BASE_PATH}/gestores_regionais_cc.xlsx", index=False)
print(f"✅ Gestores regionais criado: gestores_regionais_cc.xlsx")

# COMMAND ----------
# Gera .secrets_config.json para teste

_secrets_path = f"{BASE_PATH}/.secrets_config.json"
if not os.path.exists(_secrets_path):
    secrets_teste = {
        "power_automate_url": "https://prod-00.westus.logic.azure.com:443/workflows/EXEMPLO/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ASSINATURA_EXEMPLO",
        "assinatura_nome": "Nome Teste",
        "assinatura_cargo": "Cargo / Departamento Teste",
        "assinatura_email": "teste@exemplo.com",
        "email_gestor_excecoes": "gestor.excecoes@exemplo.com",
        "nome_gestor_excecoes": "Gestor Teste"
    }
    with open(_secrets_path, "w", encoding="utf-8") as f:
        json.dump(secrets_teste, f, ensure_ascii=False, indent=2)
    print(f"✅ Secrets criado: .secrets_config.json")
else:
    print(f"⚠️  .secrets_config.json já existe — não sobrescrito")

# COMMAND ----------
# Resumo

print("\n" + "="*60)
print("🎉 SETUP COMPLETO!")
print("="*60)
print()
print("Arquivos criados:")
print(f"  📊 {ENTRADA_PATH}/apuracao_Q2_teste.xlsx")
print(f"  📋 {BASE_PATH}/gestores_regionais_cc.xlsx")
print(f"  🔐 {BASE_PATH}/.secrets_config.json")
print()
print("Próximos passos:")
print("  1. Abra o notebook 'email_metas_apuracao_variavel'")
print("  2. Configure os widgets:")
print("     • Período: P04")
print("     • Modo de Envio: TESTE")
print("     • E-mail de Teste: seu e-mail real")
print("  3. Execute célula por célula")
print("  4. Na aba Preview, confira o e-mail gerado")
print("  5. Na aba Enviar, clique para enviar")
print()
print("⚠️  NOTA: O envio só funciona se a URL do Power Automate")
print("   em .secrets_config.json apontar para um flow real.")
print("   Para testar apenas a geração do HTML, basta usar o Preview.")
