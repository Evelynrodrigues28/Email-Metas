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

## ⚡ Configurando o Power Automate

O sistema envia e-mails através de um Flow HTTP no Power Automate (Microsoft 365). Siga os passos abaixo para configurar:

### Passo 1: Criar o Flow

1. Acesse [Power Automate](https://make.powerautomate.com)
2. Clique em **Criar** > **Fluxo de nuvem instantâneo**
3. Selecione o gatilho: **Quando uma solicitação HTTP é recebida**

### Passo 2: Configurar o Gatilho HTTP

No gatilho "Quando uma solicitação HTTP é recebida", cole este esquema JSON:

```json
{
  "type": "object",
  "properties": {
    "to": { "type": "string" },
    "cc": { "type": "string" },
    "subject": { "type": "string" },
    "body": { "type": "string" }
  },
  "required": ["to", "subject", "body"]
}
```

### Passo 3: Adicionar Ação de Envio

1. Clique em **+ Nova etapa**
2. Busque: **Office 365 Outlook — Enviar um email (V2)**
3. Configure os campos:

| Campo | Valor (conteúdo dinâmico) |
| --- | --- |
| Para | `to` |
| CC | `cc` |
| Assunto | `subject` |
| Corpo | `body` |
| É HTML | **Sim** |

> 💡 Se preferir usar uma conta compartilhada (shared mailbox), use a ação **"Enviar um email de uma caixa de correio compartilhada (V2)"** e preencha o campo "Caixa de correio original" com o endereço.

### Passo 4: (Opcional) Adicionar Condição para CC

Como o campo CC pode vir vazio, adicione uma condição:

```
Se   cc   é diferente de   (vazio)
  → Enviar email COM cc
Senão
  → Enviar email SEM cc
```

Ou simplesmente ignore — o Outlook aceita CC vazio sem erro.

### Passo 5: Salvar e Copiar a URL

1. Clique em **Salvar**
2. Volte ao gatilho HTTP — a **URL do HTTP POST** será gerada automaticamente
3. Copie essa URL (formato: `https://prod-XX.westus.logic.azure.com:443/workflows/...`)
4. Cole no arquivo `.secrets_config.json` no campo `power_automate_url`

### Passo 6: Testar Manualmente

Para validar o flow antes de usar no notebook, envie um POST manual:

```bash
curl -X POST "SUA_URL_AQUI" \
  -H "Content-Type: application/json" \
  -d '{
    "to": "seu.email@empresa.com",
    "cc": "",
    "subject": "Teste Power Automate",
    "body": "<h1>Funcionou!</h1><p>Integra\u00e7\u00e3o OK.</p>"
  }'
```

Se receber status `202 Accepted`, o flow está pronto.

### Diagrama do Flow

```
┌──────────────────────────────────────────────┐
│  Gatilho: Solicitação HTTP recebida        │
│  Método: POST                              │
│  Payload: { to, cc, subject, body }       │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│  (Opcional) Condição: CC não é vazio?     │
└───────────────────────┬──────────────────────┘
                        │
                        ▼
┌──────────────────────────────────────────────┐
│  Ação: Enviar email (V2) - Outlook 365    │
│  Para: to | CC: cc                        │
│  Assunto: subject                         │
│  Corpo: body (É HTML = Sim)               │
└──────────────────────────────────────────────┘
```

### Limites e Considerações

| Item | Limite |
| --- | --- |
| Chamadas/minuto (plano gratuito) | 100 |
| Chamadas/minuto (plano premium) | 600 |
| Tamanho máx. do body | 100 MB (na prática, e-mails HTML ficam em ~50-200 KB) |
| Timeout HTTP | 120 segundos |

O notebook já trata erros 429 (rate limit), 502 e 503 com retry automático e delays progressivos.

---

## 🧪 Como Testar (Dados Fictícios)

O repositório inclui um script que gera todos os dados necessários para teste, sem precisar de nenhum arquivo real.

### Passo 1: Executar o gerador de dados

```
1. Abra o notebook `gerar_dados_teste.py` no Databricks
2. Execute todas as células
```

Isso cria automaticamente:
- `Entrada/apuracao_Q2_teste.xlsx` — planilha com 5 associados e 6 indicadores fictícios
- `gestores_regionais_cc.xlsx` — mapeamento de 5 regionais com e-mails de exemplo
- `.secrets_config.json` — configuração com dados placeholder

### Passo 2: Executar o notebook principal

Abra `email_metas_apuracao_variavel.py` e configure os widgets:

| Widget | Valor |
| --- | --- |
| Período | **P04** |
| Modo de Envio | **TESTE** |
| E-mail de Teste | **seu-email@real.com** |
| Data Retorno | Qualquer data futura |
| Mês Pagamento | Qualquer mês |

### Passo 3: Validar

- **Sem Power Automate**: O Preview funciona normalmente. Você verá o e-mail HTML completo com dados fictícios.
- **Com Power Automate**: Substitua a `power_automate_url` no `.secrets_config.json` pela URL real do seu flow. O e-mail será enviado ao endereço de teste.

### Associados Fictícios Gerados

| Nome | E-mail | Regional |
| --- | --- | --- |
| Ana Silva | ana.silva@exemplo.com | SUDESTE |
| Carlos Souza | carlos.souza@exemplo.com | SUL |
| Maria Santos | maria.santos@exemplo.com | NORTE |
| Pedro Oliveira | pedro.oliveira@exemplo.com | NORDESTE |
| Julia Costa | julia.costa@exemplo.com | CENTRO-OESTE |

## 📜 Licença

Projeto interno — uso restrito à organização.

