# Email Metas Apuracao Variavel

Sistema automatizado para geracao e envio de e-mails individuais de apuracao de remuneracao variavel.

Funciona **localmente** (VS Code, terminal, Jupyter) ou no **Databricks**.

---

## Visao Geral

- Leitura de planilhas Excel (.xlsb/.xlsx) com dados de apuracao
- Geracao de e-mails HTML personalizados por associado
- Preview do e-mail no navegador antes do envio
- Envio em lote via Power Automate (Microsoft 365)
- Suporte a periodos mensais (P01-P13) e trimestrais (Quarter)
- Copia automatica para gestores regionais

---

## Arquitetura

```
+----------------------------+
|   Script Python            |
|   (local ou Databricks)    |
+-------------+--------------+
              | HTTP POST (JSON)
              v
+----------------------------+
|   Power Automate           |
|   (Microsoft 365)          |
|   Envio via Outlook        |
+----------------------------+
```

---

## Estrutura de Arquivos

```
email-metas-apuracao-variavel/
|
|-- email_metas_local_v2.py          # Script principal (roda local - VS Code/terminal)
|-- email_metas_apuracao_variavel.py # Versao Databricks (com widgets/displayHTML)
|-- gerar_dados_teste_local.py       # Gera dados ficticios para teste (roda local)
|-- gerar_dados_teste.py             # Gera dados ficticios (versao Databricks)
|-- template_email.html              # Template HTML do corpo do e-mail
|-- requirements.txt                 # Dependencias Python
|-- .secrets_config.json.example     # Modelo de configuracao (sem dados reais)
|-- gestores_regionais_cc_exemplo.csv# Exemplo do mapeamento Regional -> Email CC
|-- .gitignore                       # Protecao de dados sensiveis
|-- README.md                        # Este arquivo
|-- Entrada/                         # Pasta para planilhas Excel (NAO versionar)
|   +-- <arquivo>.xlsb ou .xlsx
+-- Saida/                           # Gerada automaticamente (NAO versionar)
    +-- saida.json
    +-- preview.html
```

---

## Inicio Rapido (Local - VS Code / Terminal)

### 1. Clone o repositorio

```bash
git clone https://github.com/seu-usuario/email-metas-apuracao-variavel.git
cd email-metas-apuracao-variavel
```

### 2. Crie um ambiente virtual

```bash
python -m venv venv

# Windows (PowerShell):
venv\Scripts\activate

# Mac/Linux:
source venv/bin/activate
```

### 3. Instale as dependencias

```bash
pip install pyxlsb openpyxl pandas requests numpy
```

### 4. Gere dados ficticios para teste

```bash
python gerar_dados_teste_local.py
```

Isso cria:
- `Entrada/apuracao_Q2_teste.xlsx` (planilha com 5 associados ficticios)
- `gestores_regionais_cc.xlsx` (mapeamento de regionais)
- `.secrets_config.json` (configuracao com valores de exemplo)

### 5. Execute no modo preview

```bash
python email_metas_local_v2.py --preview
```

### 6. Veja o resultado

Abra `Saida/preview.html` no navegador para ver o e-mail gerado.

---

## Uso Completo (Linha de Comando)

```bash
python email_metas_local_v2.py [opcoes]
```

### Parametros disponiveis:

| Parametro        | Descricao                          | Default             |
| ---------------- | ---------------------------------- | ------------------- |
| --periodo        | Periodo da apuracao (P01 a P13)    | P04                 |
| --modo           | teste ou producao                  | teste               |
| --email-teste    | Destinatario no modo teste         | seu.email@empresa.com |
| --data-retorno   | Prazo para contestacoes            | 20/01/2025          |
| --mes-pagamento  | Mes de referencia                  | fevereiro           |
| --associados     | Filtrar por nomes (virgula)        | todos               |
| --preview        | Apenas gera HTML, nao envia        | false               |

### Exemplos:

```bash
# Apenas preview (nao envia)
python email_metas_local_v2.py --preview

# Teste: envia para seu e-mail
python email_metas_local_v2.py --email-teste meuemail@gmail.com

# Producao: envia para todos
python email_metas_local_v2.py --modo producao --periodo P04

# Filtrar associados especificos
python email_metas_local_v2.py --preview --associados "Ana Silva,Carlos Souza"
```

### Saidas geradas:

- `Saida/saida.json` - JSON com todos os e-mails gerados
- `Saida/preview.html` - Abra no navegador para visualizar

---

## Uso no Databricks

Se preferir usar no Databricks:

1. Importe `email_metas_apuracao_variavel.py` como notebook
2. Execute `gerar_dados_teste.py` para gerar dados de teste
3. Configure os widgets na interface:
   - Periodo: P04
   - Modo de Envio: TESTE
   - E-mail de Teste: seu e-mail
4. Execute celula por celula
5. Use a aba Preview do dashboard interativo para conferir
6. Use a aba Enviar para disparar os e-mails

---

## Configurando o Power Automate

O envio de e-mails funciona via Power Automate (Microsoft 365). Siga os passos:

### Passo 1: Criar o Flow

1. Acesse https://make.powerautomate.com
2. Clique em **Criar** > **Fluxo de nuvem instantaneo**
3. Selecione o gatilho: **Quando uma solicitacao HTTP e recebida**

### Passo 2: Configurar o Gatilho HTTP

No gatilho, cole este esquema JSON:

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

### Passo 3: Adicionar Acao de Envio

1. Clique em **+ Nova etapa**
2. Busque: **Office 365 Outlook - Enviar um email (V2)**
3. Configure:

| Campo    | Valor (conteudo dinamico) |
| -------- | ------------------------- |
| Para     | `to`                      |
| CC       | `cc`                      |
| Assunto  | `subject`                 |
| Corpo    | `body`                    |
| E HTML   | **Sim**                   |

### Passo 4: Salvar e Copiar a URL

1. Salve o flow
2. Volte ao gatilho HTTP - a URL sera gerada
3. Copie a URL (formato: `https://prod-XX.westus.logic.azure.com:443/workflows/...`)
4. Cole no `.secrets_config.json` no campo `power_automate_url`

### Passo 5: Testar o Flow

Voce pode testar com curl:

```bash
curl -X POST "SUA_URL_AQUI" \
  -H "Content-Type: application/json" \
  -d "{\"to\": \"seu@email.com\", \"cc\": \"\", \"subject\": \"Teste\", \"body\": \"<h1>OK</h1>\"}"
```

Se retornar status 202, esta pronto.

### Diagrama do Flow

```
+----------------------------------------------+
|  Gatilho: Solicitacao HTTP recebida          |
|  Metodo: POST                                |
|  Payload: { to, cc, subject, body }          |
+----------------------+-----------------------+
                       |
                       v
+----------------------------------------------+
|  Acao: Enviar email (V2) - Outlook 365       |
|  Para: to | CC: cc                           |
|  Assunto: subject                            |
|  Corpo: body (E HTML = Sim)                  |
+----------------------------------------------+
```

### Limites

| Item                               | Limite     |
| ---------------------------------- | ---------- |
| Chamadas/minuto (plano gratuito)   | 100        |
| Chamadas/minuto (plano premium)    | 600        |
| Tamanho max. do body               | 100 MB     |
| Timeout HTTP                       | 120s       |

O script ja trata erros 429 (rate limit), 502 e 503 com retry automatico.

---

## Configuracao

### .secrets_config.json

Copie o exemplo e preencha:

```bash
cp .secrets_config.json.example .secrets_config.json
```

Campos:

| Campo                  | Descricao                              |
| ---------------------- | -------------------------------------- |
| power_automate_url     | URL do flow HTTP do Power Automate     |
| assinatura_nome        | Nome na assinatura do e-mail           |
| assinatura_cargo       | Cargo/departamento                     |
| assinatura_email       | E-mail na assinatura                   |
| email_gestor_excecoes  | E-mail para contestacoes               |
| nome_gestor_excecoes   | Nome do gestor de excecoes             |

### gestores_regionais_cc.xlsx

Crie baseado no CSV de exemplo (`gestores_regionais_cc_exemplo.csv`).
Duas colunas: `Regional` e `Email_CC`.

### Planilha de Apuracao (Entrada/)

Requisitos do arquivo Excel:
- Formato: `.xlsb` ou `.xlsx`
- Aba: `RESULTADO VARIAVEL` (com acento: VARIAVEL)
- Colunas: Associado, Email, Regional, Indicador (x2), metricas
- Periodos identificados por P01-P13 no header

---

## Teste com Dados Ficticios

O script `gerar_dados_teste_local.py` cria automaticamente:

| Associado       | E-mail                        | Regional      |
| --------------- | ----------------------------- | ------------- |
| Ana Silva       | ana.silva@exemplo.com         | SUDESTE       |
| Carlos Souza    | carlos.souza@exemplo.com      | SUL           |
| Maria Santos    | maria.santos@exemplo.com      | NORTE         |
| Pedro Oliveira  | pedro.oliveira@exemplo.com    | NORDESTE      |
| Julia Costa     | julia.costa@exemplo.com       | CENTRO-OESTE  |

Indicadores ficticios de Vendas (GSV) e Execucao (Sales Strike).

---

## Estrutura do E-mail

Cada e-mail contem:

1. **Tabela de resultados:**
   - Indicador (categoria + detalhamento)
   - Actual / Meta / % Atingimento
   - Recuperacao (Vendas, Marca, Regional)
   - % Variavel (resultado final - destaque em dourado)

2. **Secao Quarter** (apenas em P03, P06, P09, P13):
   - GSV Actual / FCST / Realizado x Plano
   - Recuperacoes trimestrais

3. **Rodape:**
   - Prazo de contestacao
   - Assinatura personalizada

---

## Seguranca

- **Modo TESTE**: Todos os e-mails vao apenas para o endereco de teste
- **Modo PRODUCAO**: Envia para destinatarios reais + CC gestor regional
- **Secrets**: Dados sensiveis em `.secrets_config.json` (protegido pelo .gitignore)
- **Retry**: 3 tentativas com delays progressivos (5s, 10s) em erros HTTP

---

## Manutencao

| Tarefa              | Como fazer                                  |
| ------------------- | ------------------------------------------- |
| Novo periodo        | Automatico (basta ter dados na planilha)    |
| Alterar template    | Edite `template_email.html`                 |
| Atualizar regionais | Edite `gestores_regionais_cc.xlsx`          |
| Mudar assinatura    | Edite `.secrets_config.json`                |

---

## Requisitos

- Python 3.9+
- Bibliotecas: pyxlsb, openpyxl, pandas, requests, numpy
- Power Automate (para envio real de e-mails)

---

## Licenca

Projeto interno - uso restrito a organizacao.
