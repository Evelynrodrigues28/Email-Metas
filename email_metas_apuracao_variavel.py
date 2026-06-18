# Databricks notebook source
# MAGIC %pip install pyxlsb openpyxl

# COMMAND ----------
# Célula 1: Widgets de Configuração

periodos = [f"P{str(i).zfill(2)}" for i in range(1, 14)]
dbutils.widgets.dropdown("periodo", "P04", periodos, "1. Período")
dbutils.widgets.dropdown("modo_envio", "TESTE", ["TESTE", "PRODUCAO"], "2. Modo de Envio")
dbutils.widgets.text("email_teste", "seu.email@empresa.com", "3. E-mail de Teste")
dbutils.widgets.text("data_retorno", "20/01/2025", "5. Data Retorno (dd/mm/aaaa)")
dbutils.widgets.text("mes_pagamento", "fevereiro", "6. Mês Pagamento")

# COMMAND ----------
# Célula 2: Upload do Arquivo Excel

import os
from pathlib import Path

_nb_path_upload = (
    dbutils.notebook.entry_point.getDbutils()
    .notebook().getContext().notebookPath().get()
)
_base_upload = "/Workspace" + "/".join(_nb_path_upload.split("/")[:-1])
_entrada_upload = f"{_base_upload}/Entrada"
os.makedirs(_entrada_upload, exist_ok=True)

_arquivos_existentes = sorted(
    Path(_entrada_upload).glob("*.xlsb"),
    key=lambda p: p.stat().st_mtime, reverse=True
)
_arquivos_xlsx = sorted(
    Path(_entrada_upload).glob("*.xlsx"),
    key=lambda p: p.stat().st_mtime, reverse=True
)
_todos = _arquivos_existentes + _arquivos_xlsx

print(f"📁 Pasta: {_entrada_upload}")
print()
if _todos:
    print(f"✅ {len(_todos)} arquivo(s) encontrado(s):")
    for a in _todos:
        print(f"   • {a.name} ({a.stat().st_size/1024/1024:.1f} MB)")
    print()
    print("→ Execute a próxima célula para processar.")
else:
    raise FileNotFoundError(
        f"\n❌ Nenhum arquivo .xlsb encontrado em:\n"
        f"   {_entrada_upload}\n\n"
        "   Coloque o arquivo na pasta Entrada e re-execute esta célula."
    )

# COMMAND ----------
# Célula 3: Processamento

import pandas as pd
import json
import base64
import os
from pathlib import Path
from datetime import datetime, timedelta

# --- Caminhos ---
_nb_path = (
    dbutils.notebook.entry_point.getDbutils()
    .notebook().getContext().notebookPath().get()
)
BASE_PATH = "/Workspace" + "/".join(_nb_path.split("/")[:-1])
ENTRADA_PATH = f"{BASE_PATH}/Entrada"
os.makedirs(ENTRADA_PATH, exist_ok=True)

# --- Leitura dos Widgets ---
PERIODO = dbutils.widgets.get("periodo")
TEST_MODE = dbutils.widgets.get("modo_envio") == "TESTE"
TEST_EMAIL = dbutils.widgets.get("email_teste")

# --- Configuração de Quarters ---
PERIODOS_QUARTER = ["P03", "P06", "P09", "P13"]
IS_QUARTER = PERIODO in PERIODOS_QUARTER
QUARTER_MAP = {"P03": "Q1", "P06": "Q2", "P09": "Q3", "P13": "Q4"}
QUARTER_NAME = QUARTER_MAP.get(PERIODO, "")

# --- Localização do Arquivo Excel ---
arquivos_xlsb = sorted(
    Path(ENTRADA_PATH).glob("*.xlsb"),
    key=lambda p: p.stat().st_mtime, reverse=True
)
if not arquivos_xlsb:
    raise FileNotFoundError(
        f"\n\n❌ Nenhum .xlsb encontrado em: {ENTRADA_PATH}\n"
    )

QUARTER_PERIODOS = {
    "Q1": ["P01", "P02", "P03"],
    "Q2": ["P04", "P05", "P06"],
    "Q3": ["P07", "P08", "P09"],
    "Q4": ["P10", "P11", "P12", "P13"]
}

arquivo_excel = None
for q_name, periodos_do_quarter in QUARTER_PERIODOS.items():
    if PERIODO in periodos_do_quarter:
        for f in arquivos_xlsb:
            if q_name in f.name.upper():
                arquivo_excel = str(f)
                break
        break

if arquivo_excel is None:
    for f in arquivos_xlsb:
        try:
            _df_tmp = pd.read_excel(
                f, header=None,
                sheet_name="RESULTADO VARIÁVEL",
                engine="pyxlsb", nrows=6
            )
            _df_tmp = _df_tmp.fillna("")
            for i in range(_df_tmp.shape[0]):
                if PERIODO in _df_tmp.iloc[i].astype(str).values:
                    arquivo_excel = str(f)
                    break
            if arquivo_excel:
                break
        except Exception:
            continue

if arquivo_excel is None:
    arquivo_excel = str(arquivos_xlsb[0])

arquivo_html = f"{BASE_PATH}/template_email.html"
arquivo_saida = f"{BASE_PATH}/Saida/saida.json"

# --- Carrega Segredos e Assinatura ---
arquivo_imagem_email = f"{BASE_PATH}/baixados.png"
_config_path = f"{BASE_PATH}/.secrets_config.json"
with open(_config_path, "r", encoding="utf-8") as _cf:
    _secrets = json.loads(_cf.read())

_assinatura_nome = _secrets["assinatura_nome"]
_assinatura_cargo = _secrets["assinatura_cargo"]
_assinatura_email = _secrets["assinatura_email"]

_logo_email_html = ""
if os.path.exists(arquivo_imagem_email):
    with open(arquivo_imagem_email, "rb") as _f:
        _imagem_email_b64 = base64.b64encode(_f.read()).decode("utf-8")
    _logo_email_html = (
        f'<div style="margin:0 0 16px 0;">'
        f'<img src="data:image/png;base64,{_imagem_email_b64}" '
        f'style="max-width:100%;height:auto;display:block;"></div>'
    )

tag_imagem = f"""
<div>
  {_logo_email_html}
  <table cellpadding="0" cellspacing="0" border="0"
         style="font-family: Segoe UI, Arial, sans-serif;
                border-collapse: collapse; margin: 4px 0;">
    <tr>
      <td style="vertical-align: top;">
        <p style="margin: 0 0 2px 0; font-size: 14px;
                  font-weight: bold; color: #1a1a1a;">
          {_assinatura_nome}
        </p>
        <p style="margin: 0 0 3px 0; font-size: 11px; color: #666666;">
          {_assinatura_cargo}
        </p>
        <p style="margin: 0; font-size: 11px;">
          <a href="mailto:{_assinatura_email}"
             style="color: #1155CC; text-decoration: none;">
            {_assinatura_email}
          </a>
        </p>
      </td>
    </tr>
  </table>
</div>
"""

# --- Leitura e Processamento do Excel ---
excel_path = Path(arquivo_excel)
if not excel_path.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_excel}")

df = pd.read_excel(
    excel_path, header=None,
    sheet_name="RESULTADO VARIÁVEL", engine="pyxlsb"
)
df = df.fillna("")

# Encontra a linha de header
header_row = None
for i in range(len(df)):
    linha = df.iloc[i].astype(str).tolist()
    if "Associado" in linha and "Actual" in linha:
        header_row = i
        break
if header_row is None:
    raise Exception("Header não encontrado na aba 'RESULTADO VARIÁVEL'")

# Encontra a coluna do período selecionado
col_inicio_periodo = None
for linha_offset in range(1, 6):
    linha_busca = df.iloc[header_row - linha_offset]
    for i, v in enumerate(linha_busca):
        if str(v).strip() == PERIODO:
            col_inicio_periodo = i
            break
    if col_inicio_periodo is not None:
        break
if col_inicio_periodo is None:
    raise Exception(f"Período '{PERIODO}' não encontrado")

# Encontra colunas do quarter (se aplicável)
col_inicio_quarter = None
if IS_QUARTER:
    for linha_offset in range(1, 6):
        linha_busca = df.iloc[header_row - linha_offset]
        for i, v in enumerate(linha_busca):
            if str(v).strip() == "Quarter":
                col_inicio_quarter = i
                break
        if col_inicio_quarter is not None:
            break

# Renomeia colunas e filtra dados
df.columns = df.iloc[header_row]
df = df.iloc[header_row + 1:].copy()
df.columns = [f"{col}_{i}" for i, col in enumerate(df.columns)]


def find_col(nome):
    """Encontra a primeira coluna cujo nome contém o texto."""
    return next((c for c in df.columns if nome in str(c)), None)


def find_cols(nome):
    """Encontra todas as colunas cujo nome contém o texto."""
    return [c for c in df.columns if nome in str(c)]


col_email = find_col("Email")
col_associado = find_col("Associado")
col_regional = find_col("Regional")
col_ind1, col_ind2 = find_cols("Indicador")[:2]
cols = df.columns[col_inicio_periodo:col_inicio_periodo + 8]

# Mapeamento de CC por regional
_cc_map_path = f"{BASE_PATH}/gestores_regionais_cc.xlsx"
df_cc_map = pd.read_excel(_cc_map_path, engine="openpyxl")
CC_MAP = dict(
    zip(
        df_cc_map["Regional"].str.strip().str.upper(),
        df_cc_map["Email_CC"].str.strip()
    )
)

# Colunas do quarter
cols_quarter = None
if IS_QUARTER and col_inicio_quarter is not None:
    _all_q = df.columns[col_inicio_quarter:col_inicio_quarter + 14]
    _indices_quarter = [0, 1, 2, 6, 9, 12, 13]
    cols_quarter = [_all_q[i] for i in _indices_quarter]

df = df[(df[col_email] != "") & (df[col_associado] != "")]


# --- Funções Auxiliares ---
def decimal(v):
    """Formata números com separador de milhar (ponto brasileiro)."""
    try:
        return f"{round(float(v)):,}".replace(",", ".")
    except (ValueError, TypeError):
        return ""


def percent(v):
    """Formata valores como percentual (vírgula como separador decimal)."""
    try:
        return f"{round(float(v)*100, 1):.1f}%".replace(".", ",")
    except (ValueError, TypeError):
        return ""


def data_limite():
    """Retorna dia da semana e data formatada do prazo de contestação."""
    _data_str = dbutils.widgets.get("data_retorno").strip()
    d = datetime.strptime(_data_str, "%d/%m/%Y")
    dias = [
        "segunda-feira", "terça-feira", "quarta-feira",
        "quinta-feira", "sexta-feira", "sábado", "domingo"
    ]
    return dias[d.weekday()], d.strftime("%d/%m")


def mes_pagamento():
    """Retorna o mês de pagamento configurado no widget."""
    return dbutils.widgets.get("mes_pagamento").strip()


def quarter_cells(row, is_bold=False):
    """Gera células HTML das colunas de quarter."""
    if cols_quarter is None:
        return ""
    PAYMENT_STYLE = 'class="payment-result-cell"'
    bold_attr = ' style="font-weight:bold;"' if is_bold else ''
    cells = ""
    for i, col in enumerate(cols_quarter):
        val = row[col]
        formatted = decimal(val) if i < 2 else percent(val)
        if i == 6:
            bold_pay = ' style="font-weight:bold;"' if is_bold else ''
            cells += f'<td {PAYMENT_STYLE}{bold_pay}>{formatted}</td>\n'
        else:
            cells += f'<td{bold_attr}>{formatted}</td>\n'
    return cells


# --- Widget de Associados ---
associados_disponiveis = sorted(df[col_associado].unique().tolist())
associados_disponiveis = [
    str(a) for a in associados_disponiveis if str(a).strip() != ""
]
_precisa_recriar = True
try:
    _selecao_atual = dbutils.widgets.get("associados")
    _precisa_recriar = False
except Exception:
    pass
if _precisa_recriar:
    dbutils.widgets.multiselect(
        "associados",
        associados_disponiveis[0],
        associados_disponiveis,
        "4. Associados (selecione quem recebe)"
    )

print(
    f"✅ {os.path.basename(arquivo_excel)} | "
    f"{len(df)} linhas | {len(associados_disponiveis)} associados"
)
print("   Selecione associados no widget '4. Associados' e execute a próxima célula.")

# COMMAND ----------
# Célula 4: Dashboard - Preview e Envio

import zlib

_widget_val = dbutils.widgets.get("associados").strip()
if _widget_val == "":
    associados_selecionados = associados_disponiveis
else:
    associados_selecionados = [a.strip() for a in _widget_val.split(",")]

df_envio = df[df[col_associado].isin(associados_selecionados)].copy()
if len(df_envio) == 0:
    raise Exception(
        "Nenhum associado encontrado! Verifique o campo '4. Associados'."
    )

_email_gestor_excecoes = _secrets["email_gestor_excecoes"]
_nome_gestor_excecoes = _secrets.get(
    "nome_gestor_excecoes", _email_gestor_excecoes
)
template = open(arquivo_html, encoding="utf-8").read()

# --- Geração dos Headers de Quarter ---
if IS_QUARTER and cols_quarter is not None:
    _q_widths = ['5%', '5%', '5%', '2%', '2%', '2%', '2%']
    quarter_colgroup = '\n'.join(
        [f'        <col style="width: {w};">' for w in _q_widths]
    )
    r1 = ['<td class="p04-header">Quarter</td>' for _ in range(7)]
    quarter_row1 = '\n            '.join(r1)
    sub_labels = [
        ('Actual', False), ('Meta', False), ('% Ating Meta', False),
        ('Recuperação Vendas', False), ('Recuperação Marca', False),
        ('Recuperação Regional', False), ('Pagamento', True)
    ]
    quarter_row2 = '\n            '.join([
        f'<td class="{"payment-header" if p else "p04-header"}">{l}</td>'
        for l, p in sub_labels
    ])
    quarter_row3 = (
        '<td class="vertical-text-cell">'
        '<div class="excel-style-cell">GSV<br>ACTUAL</div></td>\n'
        '            <td class="vertical-text-cell">'
        '<div class="excel-style-cell">GSV<br>FCST</div></td>\n'
        '            <td class="vertical-text-cell">'
        '<div class="excel-style-cell">Realizado<br>X<br>Plano</div></td>\n'
        '            <td class="vertical-text-cell">'
        '<div class="excel-style-cell">Recuperação<br>Vendas</div></td>\n'
        '            <td class="vertical-text-cell">'
        '<div class="excel-style-cell">Recuperação<br>Marcas</div></td>\n'
        '            <td class="vertical-text-cell">'
        '<div class="excel-style-cell">Recuperação<br>Regional</div></td>\n'
        '            <td class="vertical-empty-cell"></td>'
    )
    headers_row4 = [
        ('Actual', False, False), ('Meta', False, False),
        ('Resultado', False, False),
        ('Recuperação Vendas', False, True),
        ('Recuperação Marca', False, True),
        ('Recuperação Regional', False, True),
        ('% Recuperação Total', True, True)
    ]
    r4 = []
    for label, is_pay, wrap in headers_row4:
        cls = ' class="total-variable-header"' if is_pay else ''
        style = ' style="white-space:normal;"' if wrap else ''
        r4.append(f'<th{cls}{style}>{label}</th>')
    quarter_row4 = '\n            '.join(r4)
else:
    quarter_colgroup = ""
    quarter_row1 = ""
    quarter_row2 = ""
    quarter_row3 = ""
    quarter_row4 = ""

# --- Geração dos E-mails ---
saida = []
for associado, dados in df_envio.groupby(col_associado):
    email = str(dados.iloc[0][col_email]).strip()
    if "@" not in email:
        continue
    primeiro_nome = associado.split()[0].capitalize()

    BOLD_INDICATORS = [
        "Vendas - Sell In GSV - TOTAL",
        "Execução - Sales Strike - TOTAL"
    ]

    linhas = ""
    for _, r in dados.iterrows():
        is_bold = str(r[col_ind2]).strip() in BOLD_INDICATORS
        bold_style = ' style="font-weight:bold;"' if is_bold else ''
        linhas += f"""<tr class="data-row">
<td style="font-weight:bold;">{r[col_ind1]}</td>
<td style="font-weight:bold;">{associado}</td>
<td{bold_style}>{r[col_ind2]}</td>
<td{bold_style}>{decimal(r[cols[0]])}</td>
<td{bold_style}>{decimal(r[cols[1]])}</td>
<td{bold_style}>{percent(r[cols[2]])}</td>
<td{bold_style}>{percent(r[cols[3]])}</td>
<td{bold_style}>{percent(r[cols[4]])}</td>
<td{bold_style}>{percent(r[cols[5]])}</td>
<td{bold_style}>{percent(r[cols[6]])}</td>
<td style="background-color:#595959; color:#e5a913; font-weight:bold;">
  {percent(r[cols[7]])}
</td>
{quarter_cells(r, is_bold) if IS_QUARTER and cols_quarter is not None else ""}
</tr>
"""

    dia, data = data_limite()
    mes_ref = mes_pagamento()
    texto = f"""
<br><br>
<p style="font-family:Segoe UI; font-size:10pt;"><b>
Caso tenha dúvidas quanto ao resultado ou pedido de exceção,
por favor retornar até <span style="color:red;">{dia} ({data})</span>,
mantendo <a href="mailto:{_email_gestor_excecoes}">
{_nome_gestor_excecoes}</a> em cópia.
</b></p>
<ul style="font-family:Segoe UI; font-size:10pt;">
<li>Lembrando que todos os pedidos devem partir dos associados elegíveis.</li>
<li><b>Pagamento será realizado na folha de {mes_ref}.</b></li>
</ul><br>
<div style="text-align:left; margin-top:16px;">{tag_imagem}</div>
"""

    html = (
        template
        .replace("{nome}", primeiro_nome)
        .replace("{periodo}", PERIODO)
        .replace("{quarter_colgroup}", quarter_colgroup)
        .replace("{quarter_row1}", quarter_row1)
        .replace("{quarter_row2}", quarter_row2)
        .replace("{quarter_row3}", quarter_row3)
        .replace("{quarter_row4}", quarter_row4)
        .replace("{linhas}", linhas)
    )
    html = html.replace("</table>", "</table>" + texto)

    _regional_val = (
        str(dados.iloc[0][col_regional]).strip().upper()
        if col_regional else ""
    )
    _cc_email = CC_MAP.get(_regional_val, "")
    saida.append({
        "Email": email,
        "CC": _cc_email,
        "Assunto": f"Apuração Variável {PERIODO}",
        "Body": html
    })

# --- Salva JSON de Saída ---
os.makedirs(os.path.dirname(arquivo_saida), exist_ok=True)
with open(arquivo_saida, "w", encoding="utf-8") as f:
    json.dump(saida, f, ensure_ascii=False, indent=2)

# --- Prepara Envio via Power Automate ---
POWER_AUTOMATE_URL = _secrets["power_automate_url"]
IMG_PLACEHOLDER = "__IMG_PLACEHOLDER__"
payloads = []
for item in saida:
    payloads.append({
        "to": TEST_EMAIL if TEST_MODE else item["Email"],
        "cc": "" if TEST_MODE else item.get("CC", ""),
        "subject": item["Assunto"],
        "body": item["Body"].replace(tag_imagem, IMG_PLACEHOLDER)
    })

_payloads_json = json.dumps(payloads, ensure_ascii=False)
_payloads_compressed = zlib.compress(_payloads_json.encode("utf-8"))
_payloads_b64 = base64.b64encode(_payloads_compressed).decode("ascii")
_img_compressed = zlib.compress(tag_imagem.encode("utf-8"))
_img_b64 = base64.b64encode(_img_compressed).decode("ascii")
_url_b64 = base64.b64encode(POWER_AUTOMATE_URL.encode()).decode()

print(f"✅ {len(saida)} e-mail(s) prontos para envio")
print(f"   Modo: {'TESTE' if TEST_MODE else 'PRODUÇÃO'}")
print(f"   Período: {PERIODO}")

# --- Dashboard Interativo ---
# O displayHTML abaixo renderiza um dashboard com:
#   - Aba "Preview": visualização do e-mail HTML renderizado
#   - Aba "Enviar": botão de envio em lote com:
#       - Barra de progresso
#       - Retry automático (3x) para erros 502/503/429
#       - Delays progressivos (5s, 10s)
#       - Botão "Parar" para interrupção
#
# Variáveis utilizadas no JS:
#   _payloads_b64 -> payloads comprimidos em Base64
#   _img_b64      -> imagem de assinatura comprimida em Base64
#   _url_b64      -> URL do Power Automate em Base64
#
# Biblioteca JS: pako (para descompressão zlib no browser)
# CDN: https://cdn.jsdelivr.net/npm/pako@2.1.0/dist/pako.min.js

import re as _re_preview
_preview_html = saida[0]["Body"]
_preview_html = _re_preview.sub(
    r'<script>[\s\S]*?</script>', '', _preview_html
)
_preview_html = _re_preview.sub(
    r'<style>[\s\S]*?</style>', '', _preview_html
)
_preview_html = _re_preview.sub(
    r'<!DOCTYPE[^>]*>', '', _preview_html, flags=_re_preview.IGNORECASE
)
_preview_html = _re_preview.sub(
    r'</?html[^>]*>', '', _preview_html, flags=_re_preview.IGNORECASE
)
_preview_html = _re_preview.sub(
    r'</?head[^>]*>', '', _preview_html, flags=_re_preview.IGNORECASE
)
_preview_html = _re_preview.sub(
    r'</?body[^>]*>', '', _preview_html, flags=_re_preview.IGNORECASE
)

_modo_cor = '#e67e00' if TEST_MODE else '#28a745'
_modo_txt = f'TESTE -> {TEST_EMAIL}' if TEST_MODE else 'PRODUCAO'

# displayHTML(f'''
# <script src="https://cdn.jsdelivr.net/npm/pako@2.1.0/dist/pako.min.js"></script>
# ... (dashboard HTML com abas Preview/Enviar, barra de progresso,
#      lógica de fetch com retry, botão Parar) ...
# ''')
#
# Para ver a implementação completa do displayHTML,
# consulte o notebook original no Databricks.
