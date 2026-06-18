# 📧 Email Metas Apuração Variável

Sistema automatizado para geração e envio de e-mails individuais de apuração de remuneração variável, desenvolvido em Databricks Notebooks com integração ao Power Automate.

## 📋 Visão Geral

Este projeto automatiza o processo de comunicação dos resultados de remuneração variável para associados, incluindo:

- Leitura de planilhas Excel (.xlsb/.xlsx) com dados de apuração
- Geração de e-mails HTML personalizados por associado
- Preview interativo no notebook antes do envio
- Envio em lote via Power Automate (Microsoft 365)
- Suporte a períodos mensais (P01–P13) e trimestrais (Quarter)
- Cópia automática para gestores regionais

## 🏗️ Arquitetura

```
┌─────────────────────────────────────────────────────────┐
│                    Databricks Notebook                    │
├─────────────────────────────────────────────────────────┤
│  1. Widgets (parâmetros de execução)                     │
│  2. Leitura do Excel (.xlsb) com dados de apuração       │
│  3. Processamento e formatação dos dados                 │
│  4. Geração de HTML personalizado por associado          │
│  5. Dashboard interativo (Preview + Envio)               │
└────────────────────────┬────────────────────────────────┘
                         │ HTTP POST (JSON comprimido)
                         ▼
┌─────────────────────────────────────────────────────────┐
│              Power Automate (Microsoft 365)               │
│         Envio de e-mail via Outlook/Exchange              │
└─────────────────────────────────────────────────────────┘
```

## 📁 Estrutura de Pastas

```
email-metas-apuracao-variavel/
├── email_metas_apuracao_variavel.py   # Notebook principal (Databricks)
├── template_email.html                 # Template HTML do corpo do e-mail
├── gestores_regionais_cc_exemplo.csv   # Exemplo: mapeamento Regional → Email CC
├── .secrets_config.json.example        # Exemplo: configurações sensíveis
├── requirements.txt                    # Dependências Python
├── .gitignore                          # Arquivos ignorados pelo Git
├── README.md                           # Este arquivo
├── Entrada/                            # Pasta para upload dos arquivos Excel
│   └── <arquivo_apuracao>.xlsb         # Planilha de resultados (NÃO versionar)
└── Saida/                              # Pasta de saída gerada automaticamente
    └── saida.json                      # JSON com e-mails gerados (NÃO versionar)
```

## ⚙️ Pré-requisitos

- **Databricks Runtime** 13.x ou superior
- **Bibliotecas Python**: `pyxlsb`, `openpyxl` (instaladas automaticamente via `%pip`)
- **Power Automate**: Flow HTTP configurado para receber e enviar e-mails
- **Planilha Excel**: Arquivo `.xlsb` com aba `RESULTADO VARIÁVEL`

## 🔧 Configuração

### 1. Arquivo de Segredos

Copie `.secrets_config.json.example` para `.secrets_config.json` e preencha com seus dados:

```bash
cp .secrets_config.json.example .secrets_config.json
```

Campos:
- `power_automate_url` — URL do flow HTTP do Power Automate
- `assinatura_nome` — Nome exibido na assinatura do e-mail
- `assinatura_cargo` — Cargo/departamento na assinatura
- `assinatura_email` — E-mail na assinatura
- `email_gestor_excecoes` — E-mail do gestor para contestações
- `nome_gestor_excecoes` — Nome do gestor de exceções

> ⚠️ **IMPORTANTE**: Este arquivo contém dados sensíveis. Nunca versione no Git.

### 2. Mapeamento de Gestores Regionais

Crie `gestores_regionais_cc.xlsx` (formato Excel) baseado no exemplo CSV fornecido.
Duas colunas: `Regional` e `Email_CC`.

### 3. Template HTML

O arquivo `template_email.html` contém o layout do e-mail com placeholders:

| Placeholder | Descrição |
| --- | --- |
| `{nome}` | Primeiro nome do associado |
| `{periodo}` | Período atual (ex: P04) |
| `{linhas}` | Linhas da tabela de resultados |
| `{quarter_colgroup}` | Colunas do quarter (quando aplicável) |
| `{quarter_row1}` a `{quarter_row4}` | Headers do quarter |

### 4. Planilha de Apuração

Coloque o arquivo `.xlsb` na pasta `Entrada/`. Requisitos:
- Aba: `RESULTADO VARIÁVEL`
- Colunas: `Associado`, `Email`, `Regional`, `Indicador` (x2), métricas
- Períodos identificados por `P01`–`P13` no header

## 🚀 Como Usar

### Passo 1: Configurar Widgets

| Widget | Descrição | Exemplo |
| --- | --- | --- |
| 1. Período | Período da apuração | P04 |
| 2. Modo de Envio | TESTE ou PRODUCAO | TESTE |
| 3. E-mail de Teste | Destinatário no modo teste | usuario@empresa.com |
| 4. Associados | Selecione quem receberá o e-mail | (multiselect) |
| 5. Data Retorno | Prazo para contestações | 20/01/2025 |
| 6. Mês Pagamento | Mês de referência do pagamento | fevereiro |

### Passo 2: Upload do Arquivo

Coloque o `.xlsb` na pasta `Entrada/`. O notebook detecta automaticamente o arquivo pelo período/quarter.

### Passo 3: Processamento

Execute a célula de processamento. Ela:
1. Localiza o arquivo correto baseado no período/quarter
2. Lê a aba `RESULTADO VARIÁVEL`
3. Identifica headers e colunas de métricas
4. Popula o widget de associados

### Passo 4: Preview e Envio

O dashboard interativo oferece:
- **Aba Preview**: Visualização do e-mail renderizado
- **Aba Enviar**: Envio em lote com barra de progresso e retry automático

## 📊 Estrutura do E-mail

Cada e-mail contém:

1. **Tabela de resultados** com colunas:
   - Indicador (categoria + detalhamento)
   - Actual / Meta / % Atingimento
   - Recuperação (Vendas, Marca, Regional)
   - % Variável (resultado final)
2. **Seção Quarter** (apenas em P03, P06, P09, P13):
   - GSV Actual / FCST / Realizado × Plano
   - Recuperações trimestrais
3. **Rodapé** com prazo de contestação e assinatura

## 🔒 Segurança

- **Modo TESTE**: E-mails direcionados ao endereço de teste
- **Modo PRODUÇÃO**: E-mails aos destinatários reais + CC gestor regional
- **Compressão**: Payloads comprimidos (zlib + Base64) antes do envio
- **Segredos**: Dados sensíveis em `.secrets_config.json` (não versionado)

## 🔄 Fluxo de Envio

1. Payloads JSON comprimidos com `zlib` e codificados em Base64
2. JavaScript no dashboard decodifica no browser
3. Envio sequencial via `fetch()` ao Power Automate
4. Retry automático (3 tentativas) em erros 502/503/429
5. Delays progressivos entre retries (5s, 10s)
6. Botão "Parar" para interromper envio

## 📝 Detecção Automática de Arquivo

Prioridade de seleção:
1. Nome do arquivo contém o quarter (Q1, Q2, Q3, Q4)
2. Conteúdo da aba contém o período solicitado
3. Arquivo mais recente (fallback)

## 🎨 Formatação Visual

- **Linhas em negrito**: Indicadores totalizadores
- **Célula de resultado**: Fundo cinza (#595959) + texto dourado (#e5a913)
- **Headers Quarter**: Amarelo (#FFCC5B) para métricas, verde (#00B050) para pagamento
- **Faixa azul**: Header principal (#000050)

## 🛠️ Manutenção

| Tarefa | Como fazer |
| --- | --- |
| Novo período | Automático — basta ter os dados na planilha |
| Alterar template | Edite `template_email.html` |
| Atualizar regionais | Edite `gestores_regionais_cc.xlsx` |
| Mudar assinatura | Edite `.secrets_config.json` |

## 📜 Licença

Projeto interno — uso restrito à organização.
