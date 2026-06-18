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
Email Variavel/
├── Email Metas Apuracao Variavel.py   # Notebook principal
├── template_email.html                 # Template HTML do corpo do e-mail
├── baixados.png                        # Imagem/banner do cabeçalho do e-mail
├── gestores_regionais_cc.xlsx          # Mapeamento Regional → Email CC
├── .secrets_config.json                # Configurações sensíveis (NÃO versionar)
├── Entrada/                            # Pasta para upload dos arquivos Excel
│   └── <arquivo_apuracao>.xlsb         # Planilha de resultados variável
├── Saida/                              # Pasta de saída gerada automaticamente
│   └── saida.json                      # JSON com todos os e-mails gerados
└── README.md                           # Este arquivo
```

## ⚙️ Pré-requisitos

- **Databricks Runtime** 13.x ou superior
- **Bibliotecas Python**: `pyxlsb`, `openpyxl` (instaladas automaticamente via `%pip`)
- **Power Automate**: Flow HTTP configurado para receber e enviar e-mails
- **Planilha Excel**: Arquivo `.xlsb` com aba `RESULTADO VARIÁVEL`

## 🔧 Configuração

### 1. Arquivo de Segredos (`.secrets_config.json`)

Crie o arquivo `.secrets_config.json` na raiz do projeto com a seguinte estrutura:

```json
{
  "power_automate_url": "https://prod-XX.westus.logic.azure.com:443/workflows/...",
  "assinatura_nome": "Nome do Remetente",
  "assinatura_cargo": "Cargo / Departamento",
  "assinatura_email": "remetente@empresa.com",
  "email_gestor_excecoes": "gestor@empresa.com",
  "nome_gestor_excecoes": "Nome do Gestor"
}
```

> ⚠️ **IMPORTANTE**: Este arquivo contém dados sensíveis. Adicione ao `.gitignore`.

### 2. Mapeamento de Gestores Regionais (`gestores_regionais_cc.xlsx`)

Planilha com duas colunas:

| Regional | Email_CC |
| --- | --- |
| NORTE | gestor.norte@empresa.com |
| SUL | gestor.sul@empresa.com |
| SUDESTE | gestor.sudeste@empresa.com |

### 3. Template HTML (`template_email.html`)

Template com placeholders que são substituídos dinamicamente:

- `{nome}` → Primeiro nome do associado
- `{periodo}` → Período atual (ex: P04)
- `{linhas}` → Linhas da tabela de resultados
- `{quarter_colgroup}` → Colunas do quarter (quando aplicável)
- `{quarter_row1}` a `{quarter_row4}` → Headers do quarter

### 4. Planilha de Apuração (Entrada)

Arquivo `.xlsb` ou `.xlsx` contendo:
- Aba: `RESULTADO VARIÁVEL`
- Colunas esperadas: `Associado`, `Email`, `Regional`, `Indicador` (x2), colunas de métricas
- Linhas de período identificadas por `P01`–`P13` no header

## 🚀 Como Usar

### Passo 1: Configurar Widgets

Ao executar a primeira célula de widgets, configure:

| Widget | Descrição | Exemplo |
| --- | --- | --- |
| 1. Período | Período da apuração | P04 |
| 2. Modo de Envio | TESTE ou PRODUCAO | TESTE |
| 3. E-mail de Teste | Destinatário no modo teste | usuario@empresa.com |
| 4. Associados | Selecione quem receberá o e-mail | (multiselect) |
| 5. Data Retorno | Prazo para contestações | 20/01/2025 |
| 6. Mês Pagamento | Mês de referência do pagamento | fevereiro |

### Passo 2: Upload do Arquivo

Coloque o arquivo `.xlsb` na pasta `Entrada/`. O notebook detecta automaticamente o arquivo mais recente e tenta associar ao quarter correto pelo nome do arquivo.

### Passo 3: Processamento

A célula de processamento:
1. Localiza o arquivo correto baseado no período/quarter
2. Lê a aba `RESULTADO VARIÁVEL`
3. Identifica headers e colunas de métricas
4. Popula o widget de associados disponíveis

### Passo 4: Preview e Envio

O dashboard interativo oferece:
- **Aba Preview**: Visualização completa do e-mail (HTML renderizado)
- **Aba Enviar**: Botão de envio em lote com barra de progresso

## 📊 Estrutura do E-mail

Cada e-mail contém:

1. **Banner/Imagem** de cabeçalho
2. **Tabela de resultados** com colunas:
   - Indicador (categoria)
   - Associado
   - Indicador (detalhamento)
   - Actual / Meta / % Atingimento
   - Recuperação (Vendas, Marca, Regional)
   - % Variável (resultado final)
3. **Seção Quarter** (apenas em P03, P06, P09, P13):
   - GSV Actual / FCST
   - Realizado × Plano
   - Recuperações trimestrais
   - Pagamento acumulado
4. **Rodapé** com prazo de contestação e assinatura

## 🔒 Segurança

- **Modo TESTE**: Todos os e-mails são direcionados ao endereço de teste configurado
- **Modo PRODUÇÃO**: E-mails enviados aos destinatários reais com CC ao gestor regional
- **Compressão**: Payloads são comprimidos (zlib) e codificados em Base64 antes do envio
- **Segredos**: URL do Power Automate e dados de assinatura armazenados em arquivo local não versionado

## 🔄 Fluxo de Envio

1. Payloads JSON são comprimidos com `zlib` e codificados em Base64
2. JavaScript no dashboard decodifica e descomprime no browser
3. Envio sequencial via `fetch()` para o endpoint do Power Automate
4. Retry automático (até 3 tentativas) em caso de erro 502/503/429
5. Delays progressivos entre retries (5s, 10s)
6. Botão "Parar" para interromper envio em andamento

## 📝 Lógica de Detecção de Arquivo

O notebook utiliza a seguinte prioridade para selecionar o arquivo Excel:

1. Busca arquivo cujo nome contenha o quarter correspondente (Q1, Q2, Q3, Q4)
2. Se não encontrar, abre cada `.xlsb` e verifica se a aba contém o período solicitado
3. Em último caso, utiliza o arquivo mais recente

## 🎨 Formatação Condicional

- **Linhas em negrito**: Indicadores totalizadores ("Vendas - Sell In GSV - TOTAL", "Execução - Sales Strike - TOTAL")
- **Célula de resultado**: Fundo cinza escuro (#595959) com texto dourado (#e5a913)
- **Headers Quarter**: Fundo amarelo (#FFCC5B) para métricas, verde (#00B050) para pagamento

## 🛠️ Manutenção

### Adicionar novo período
Os períodos P01–P13 são gerados dinamicamente. Basta que a planilha contenha os dados.

### Alterar template
Edite o arquivo `template_email.html`. Os placeholders `{...}` são substituídos em runtime.

### Atualizar gestores regionais
Edite o arquivo `gestores_regionais_cc.xlsx` com os novos mapeamentos.

## 📄 .gitignore Recomendado

```gitignore
# Segredos e configurações locais
.secrets_config.json

# Arquivos de dados (Excel)
Entrada/
Saida/

# Imagens locais
*.png

# Databricks
.databricks/
__pycache__/
```

## 📌 Notas

- O notebook foi desenvolvido para execução no Databricks com cluster pessoal (single-user)
- A integração com Power Automate requer que o flow esteja ativo e acessível
- Em períodos trimestrais (P03, P06, P09, P13), colunas adicionais de quarter são automaticamente incluídas
- O widget de associados é populado dinamicamente a partir dos dados da planilha

## 📜 Licença

Projeto interno — uso restrito à organização.
