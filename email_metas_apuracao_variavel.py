"""
Email Metas Apuracao Variavel - Versao Local (sem Databricks)
==============================================================
Executa em qualquer ambiente Python (VS Code, Jupyter, terminal).

DEPENDENCIAS:
    pip install pyxlsb openpyxl pandas requests numpy

USO:
    python email_metas_local_v2.py
    python email_metas_local_v2.py --periodo P06 --modo producao
    python email_metas_local_v2.py --preview
"""

import argparse
import base64
import json
import os
import sys
import time
import zlib
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests


# =============================================================================
# CONFIGURACAO VIA ARGUMENTOS
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Gera e envia e-mails de apuracao variavel."
    )
    parser.add_argument("--periodo", default="P04",
                        help="Periodo (P01 a P13). Default: P04")
    parser.add_argument("--modo", default="teste",
                        choices=["teste", "producao"],
                        help="Modo: teste ou producao. Default: teste")
    parser.add_argument("--email-teste", default="seu.email@empresa.com",
                        help="Destinatario no modo teste.")
    parser.add_argument("--data-retorno", default="20/01/2025",
                        help="Prazo para contestacoes (dd/mm/aaaa).")
    parser.add_argument("--mes-pagamento", default="fevereiro",
                        help="Mes de referencia para pagamento.")
    parser.add_argument("--associados", default="",
                        help="Nomes separados por virgula (vazio=todos).")
    parser.add_argument("--preview", action="store_true",
                        help="Apenas gera HTML, nao envia.")
    return parser.parse_args()


# =============================================================================
# CAMINHOS
# =============================================================================

BASE_PATH = Path(__file__).parent
ENTRADA_PATH = BASE_PATH / "Entrada"
SAIDA_PATH = BASE_PATH / "Saida"
SAIDA_PATH.mkdir(exist_ok=True)


# =============================================================================
# FUNCOES AUXILIARES
# =============================================================================

def decimal_fmt(v):
    """Formata numeros com separador de milhar (ponto)."""
    try:
        return f"{round(float(v)):,}".replace(",", ".")
    except (ValueError, TypeError):
        return ""


def percent_fmt(v):
    """Formata como percentual (virgula decimal)."""
    try:
        return f"{round(float(v)*100, 1):.1f}%".replace(".", ",")
    except (ValueError, TypeError):
        return ""


def data_limite(data_str):
    """Retorna dia da semana e data formatada."""
    d = datetime.strptime(data_str.strip(), "%d/%m/%Y")
    dias = ["segunda-feira", "terca-feira", "quarta-feira",
            "quinta-feira", "sexta-feira", "sabado", "domingo"]
    return dias[d.weekday()], d.strftime("%d/%m")


def find_col(df, nome):
    return next((c for c in df.columns if nome in str(c)), None)


def find_cols(df, nome):
    return [c for c in df.columns if nome in str(c)]


# =============================================================================
# LOCALIZACAO DO ARQUIVO
# =============================================================================

def localizar_arquivo(periodo):
    QUARTER_PERIODOS = {
        "Q1": ["P01", "P02", "P03"],
        "Q2": ["P04", "P05", "P06"],
        "Q3": ["P07", "P08", "P09"],
        "Q4": ["P10", "P11", "P12", "P13"]
    }
    arquivos = sorted(
        list(ENTRADA_PATH.glob("*.xlsb")) + list(ENTRADA_PATH.glob("*.xlsx")),
        key=lambda p: p.stat().st_mtime, reverse=True
    )
    if not arquivos:
        print(f"\n[ERRO] Nenhum arquivo .xlsb/.xlsx encontrado em: {ENTRADA_PATH}")
        print("       Coloque o arquivo na pasta Entrada/ e tente novamente.")
        sys.exit(1)

    for q_name, periodos_q in QUARTER_PERIODOS.items():
        if periodo in periodos_q:
            for f in arquivos:
                if q_name in f.name.upper():
                    return f
            break

    for f in arquivos:
        try:
            engine = "pyxlsb" if f.suffix == ".xlsb" else "openpyxl"
            tmp = pd.read_excel(f, header=None,
                                sheet_name="RESULTADO VARI\u00c1VEL",
                                engine=engine, nrows=6).fillna("")
            for i in range(tmp.shape[0]):
                if periodo in tmp.iloc[i].astype(str).values:
                    return f
        except Exception:
            continue

    return arquivos[0]


# =============================================================================
# MAIN
# =============================================================================

def main():
    args = parse_args()

    PERIODO = args.periodo.upper()
    TEST_MODE = args.modo == "teste"
    TEST_EMAIL = args.email_teste
    PERIODOS_QUARTER = ["P03", "P06", "P09", "P13"]
    IS_QUARTER = PERIODO in PERIODOS_QUARTER

    print("=" * 60)
    print(f"  Email Metas Apuracao Variavel")
    print(f"  Periodo: {PERIODO} | Modo: {'TESTE' if TEST_MODE else 'PRODUCAO'}")
    print("=" * 60)

    # --- Secrets ---
    config_path = BASE_PATH / ".secrets_config.json"
    if not config_path.exists():
        print("\n[ERRO] .secrets_config.json nao encontrado!")
        print("       Copie .secrets_config.json.example e preencha.")
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        secrets = json.load(f)

    # --- Assinatura ---
    _nome = secrets["assinatura_nome"]
    _cargo = secrets["assinatura_cargo"]
    _email_assin = secrets["assinatura_email"]

    arquivo_imagem = BASE_PATH / "baixados.png"
    _logo_html = ""
    if arquivo_imagem.exists():
        _img_b64 = base64.b64encode(arquivo_imagem.read_bytes()).decode("utf-8")
        _logo_html = (
            f'<div style="margin:0 0 16px 0;">'
            f'<img src="data:image/png;base64,{_img_b64}" '
            f'style="max-width:100%;height:auto;display:block;"></div>'
        )

    tag_imagem = f"""
<div>
  {_logo_html}
  <table cellpadding="0" cellspacing="0" border="0"
         style="font-family:Segoe UI,Arial,sans-serif;border-collapse:collapse;margin:4px 0;">
    <tr><td style="vertical-align:top;">
        <p style="margin:0 0 2px 0;font-size:14px;font-weight:bold;color:#1a1a1a;">{_nome}</p>
        <p style="margin:0 0 3px 0;font-size:11px;color:#666666;">{_cargo}</p>
        <p style="margin:0;font-size:11px;"><a href="mailto:{_email_assin}" style="color:#1155CC;text-decoration:none;">{_email_assin}</a></p>
    </td></tr>
  </table>
</div>
"""

    # --- Arquivo Excel ---
    arquivo_excel = localizar_arquivo(PERIODO)
    print(f"\n[INFO] Arquivo: {arquivo_excel.name}")

    engine = "pyxlsb" if arquivo_excel.suffix == ".xlsb" else "openpyxl"
    df = pd.read_excel(arquivo_excel, header=None,
                       sheet_name="RESULTADO VARI\u00c1VEL",
                       engine=engine).fillna("")

    # --- Header ---
    header_row = None
    for i in range(len(df)):
        linha = df.iloc[i].astype(str).tolist()
        if "Associado" in linha and "Actual" in linha:
            header_row = i
            break
    if header_row is None:
        print("\n[ERRO] Header nao encontrado na aba RESULTADO VARIAVEL")
        sys.exit(1)

    # --- Coluna do Periodo ---
    col_inicio_periodo = None
    for offset in range(1, 6):
        linha_busca = df.iloc[header_row - offset]
        for i, v in enumerate(linha_busca):
            if str(v).strip() == PERIODO:
                col_inicio_periodo = i
                break
        if col_inicio_periodo is not None:
            break
    if col_inicio_periodo is None:
        print(f"\n[ERRO] Periodo '{PERIODO}' nao encontrado na planilha")
        sys.exit(1)

    # --- Quarter ---
    col_inicio_quarter = None
    if IS_QUARTER:
        for offset in range(1, 6):
            linha_busca = df.iloc[header_row - offset]
            for i, v in enumerate(linha_busca):
                if str(v).strip() == "Quarter":
                    col_inicio_quarter = i
                    break
            if col_inicio_quarter is not None:
                break

    # --- DataFrame ---
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].copy()
    df.columns = [f"{col}_{i}" for i, col in enumerate(df.columns)]

    col_email = find_col(df, "Email")
    col_associado = find_col(df, "Associado")
    col_regional = find_col(df, "Regional")
    col_ind1, col_ind2 = find_cols(df, "Indicador")[:2]
    cols = df.columns[col_inicio_periodo:col_inicio_periodo + 8]

    # --- CC Map ---
    CC_MAP = {}
    cc_path = BASE_PATH / "gestores_regionais_cc.xlsx"
    if cc_path.exists():
        df_cc = pd.read_excel(cc_path, engine="openpyxl")
        CC_MAP = dict(zip(
            df_cc["Regional"].str.strip().str.upper(),
            df_cc["Email_CC"].str.strip()
        ))

    # --- Quarter Cols ---
    cols_quarter = None
    if IS_QUARTER and col_inicio_quarter is not None:
        _all_q = df.columns[col_inicio_quarter:col_inicio_quarter + 14]
        cols_quarter = [_all_q[i] for i in [0, 1, 2, 6, 9, 12, 13]]

    df = df[(df[col_email] != "") & (df[col_associado] != "")]

    # --- Filtro ---
    if args.associados:
        sel = [a.strip() for a in args.associados.split(",")]
        df_envio = df[df[col_associado].isin(sel)].copy()
    else:
        df_envio = df.copy()

    if len(df_envio) == 0:
        print("\n[ERRO] Nenhum associado encontrado!")
        sys.exit(1)

    print(f"   {df_envio[col_associado].nunique()} associado(s)")

    # --- Quarter Cells ---
    def quarter_cells(row, is_bold=False):
        if cols_quarter is None:
            return ""
        bold = ' style="font-weight:bold;"' if is_bold else ''
        cells = ""
        for i, col in enumerate(cols_quarter):
            val = row[col]
            fmt = decimal_fmt(val) if i < 2 else percent_fmt(val)
            if i == 6:
                cells += f'<td class="payment-result-cell"{bold}>{fmt}</td>\n'
            else:
                cells += f'<td{bold}>{fmt}</td>\n'
        return cells

    # --- Quarter Headers ---
    if IS_QUARTER and cols_quarter is not None:
        _q_widths = ['5%', '5%', '5%', '2%', '2%', '2%', '2%']
        quarter_colgroup = '\n'.join(
            [f'        <col style="width: {w};">' for w in _q_widths]
        )
        quarter_row1 = '\n            '.join(
            ['<td class="p04-header">Quarter</td>' for _ in range(7)]
        )
        sub_labels = [
            ('Actual', False), ('Meta', False), ('% Ating Meta', False),
            ('Recuperacao Vendas', False), ('Recuperacao Marca', False),
            ('Recuperacao Regional', False), ('Pagamento', True)
        ]
        quarter_row2 = '\n            '.join([
            f'<td class="{"payment-header" if p else "p04-header"}">{l}</td>'
            for l, p in sub_labels
        ])
        quarter_row3 = (
            '<td class="vertical-text-cell"><div class="excel-style-cell">GSV<br>ACTUAL</div></td>\n'
            '            <td class="vertical-text-cell"><div class="excel-style-cell">GSV<br>FCST</div></td>\n'
            '            <td class="vertical-text-cell"><div class="excel-style-cell">Realizado<br>X<br>Plano</div></td>\n'
            '            <td class="vertical-text-cell"><div class="excel-style-cell">Recuperacao<br>Vendas</div></td>\n'
            '            <td class="vertical-text-cell"><div class="excel-style-cell">Recuperacao<br>Marcas</div></td>\n'
            '            <td class="vertical-text-cell"><div class="excel-style-cell">Recuperacao<br>Regional</div></td>\n'
            '            <td class="vertical-empty-cell"></td>'
        )
        headers_row4 = [
            ('Actual', False, False), ('Meta', False, False),
            ('Resultado', False, False), ('Recuperacao Vendas', False, True),
            ('Recuperacao Marca', False, True), ('Recuperacao Regional', False, True),
            ('% Recuperacao Total', True, True)
        ]
        r4 = []
        for label, is_pay, wrap in headers_row4:
            cls = ' class="total-variable-header"' if is_pay else ''
            style = ' style="white-space:normal;"' if wrap else ''
            r4.append(f'<th{cls}{style}>{label}</th>')
        quarter_row4 = '\n            '.join(r4)
    else:
        quarter_colgroup = quarter_row1 = quarter_row2 = ""
        quarter_row3 = quarter_row4 = ""

    # --- Template ---
    template_path = BASE_PATH / "template_email.html"
    if not template_path.exists():
        print("\n[ERRO] template_email.html nao encontrado!")
        sys.exit(1)
    template = template_path.read_text(encoding="utf-8")

    # --- Gera Emails ---
    _email_gestor = secrets["email_gestor_excecoes"]
    _nome_gestor = secrets.get("nome_gestor_excecoes", _email_gestor)

    BOLD_INDICATORS = [
        "Vendas - Sell In GSV - TOTAL",
        "Execucao - Sales Strike - TOTAL"
    ]

    saida = []
    for associado, dados in df_envio.groupby(col_associado):
        email = str(dados.iloc[0][col_email]).strip()
        if "@" not in email:
            continue
        primeiro_nome = associado.split()[0].capitalize()

        linhas = ""
        for _, r in dados.iterrows():
            is_bold = str(r[col_ind2]).strip() in BOLD_INDICATORS
            bold_style = ' style="font-weight:bold;"' if is_bold else ''
            linhas += f"""<tr class="data-row">
<td style="font-weight:bold;">{r[col_ind1]}</td>
<td style="font-weight:bold;">{associado}</td>
<td{bold_style}>{r[col_ind2]}</td>
<td{bold_style}>{decimal_fmt(r[cols[0]])}</td>
<td{bold_style}>{decimal_fmt(r[cols[1]])}</td>
<td{bold_style}>{percent_fmt(r[cols[2]])}</td>
<td{bold_style}>{percent_fmt(r[cols[3]])}</td>
<td{bold_style}>{percent_fmt(r[cols[4]])}</td>
<td{bold_style}>{percent_fmt(r[cols[5]])}</td>
<td{bold_style}>{percent_fmt(r[cols[6]])}</td>
<td style="background-color:#595959; color:#e5a913; font-weight:bold;">{percent_fmt(r[cols[7]])}</td>
{quarter_cells(r, is_bold) if IS_QUARTER and cols_quarter is not None else ""}
</tr>
"""

        dia, data = data_limite(args.data_retorno)
        texto = f"""
<br><br>
<p style="font-family:Segoe UI; font-size:10pt;"><b>
Caso tenha duvidas quanto ao resultado ou pedido de excecao,
por favor retornar ate <span style="color:red;">{dia} ({data})</span>,
mantendo <a href="mailto:{_email_gestor}">{_nome_gestor}</a> em copia.
</b></p>
<ul style="font-family:Segoe UI; font-size:10pt;">
<li>Lembrando que todos os pedidos devem partir dos associados elegiveis.</li>
<li><b>Pagamento sera realizado na folha de {args.mes_pagamento}.</b></li>
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

        _reg = str(dados.iloc[0][col_regional]).strip().upper() if col_regional else ""
        _cc = CC_MAP.get(_reg, "")
        saida.append({
            "Email": email,
            "CC": _cc,
            "Assunto": f"Apuracao Variavel {PERIODO}",
            "Body": html
        })

    # --- Salva ---
    arquivo_saida = SAIDA_PATH / "saida.json"
    with open(arquivo_saida, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)

    preview_path = SAIDA_PATH / "preview.html"
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(saida[0]["Body"])

    print(f"\n[OK] {len(saida)} e-mail(s) gerados")
    print(f"     Saida: {arquivo_saida}")
    print(f"     Preview: {preview_path} (abra no navegador)")

    if args.preview:
        print("\n[PREVIEW] E-mails NAO enviados. Abra preview.html no navegador.")
        return

    # --- Envio ---
    PA_URL = secrets["power_automate_url"]
    if "EXEMPLO" in PA_URL or "SEU_WORKFLOW" in PA_URL:
        print("\n[!] URL do Power Automate com valor de exemplo!")
        print("    Atualize .secrets_config.json com a URL real.")
        return

    print(f"\n[ENVIO] Enviando {len(saida)} e-mail(s)...")
    print(f"        Modo: {'TESTE -> ' + TEST_EMAIL if TEST_MODE else 'PRODUCAO'}")
    print()

    ok, fail = 0, 0
    for i, item in enumerate(saida):
        payload = {
            "to": TEST_EMAIL if TEST_MODE else item["Email"],
            "cc": "" if TEST_MODE else item.get("CC", ""),
            "subject": item["Assunto"],
            "body": item["Body"]
        }
        sucesso = False
        for attempt in range(3):
            try:
                resp = requests.post(PA_URL, json=payload, timeout=30)
                if resp.status_code in (200, 202):
                    sucesso = True
                    break
                elif resp.status_code in (429, 502, 503):
                    wait = 5 * (attempt + 1)
                    print(f"   [!] Status {resp.status_code} - retry em {wait}s...")
                    time.sleep(wait)
                else:
                    print(f"   [FALHA] [{i+1}] status {resp.status_code}")
                    break
            except requests.RequestException as e:
                print(f"   [ERRO] [{i+1}] {e}")
                break

        dest = TEST_EMAIL if TEST_MODE else item["Email"]
        if sucesso:
            ok += 1
            print(f"   [OK] [{i+1}/{len(saida)}] {dest}")
        else:
            fail += 1

    print(f"\n{'='*60}")
    print(f"  RESULTADO: {ok} enviados | {fail} falhas")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
