# Email Metas Apuracao Variavel

Sistema automatizado para geracao e envio de e-mails individuais de apuracao de remuneracao variavel, com integracao ao Power Automate.

Funciona **localmente** (VS Code, terminal, Jupyter) ou no **Databricks**.

---

## Visao Geral

- Leitura de planilhas Excel (.xlsb/.xlsx) com dados de apuracao
- Geracao de e-mails HTML personalizados por associado
- Preview do e-mail no navegador
- Envio em lote via Power Automate (Microsoft 365)
- Suporte a periodos mensais (P01-P13) e trimestrais (Quarter)
- Copia automatica para gestores regionais

---

## Arquitetura

```
+----------------------------+
|   email_metas_apuracao     |
|   _variavel.py             |
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
|-- .secrets_config.json              # Configuracao (preencha com seus dados)
|-- README.md                         # Este arquivo
|-- email_metas_apuracao_variavel.py  # Script principal
|-- gerar_dados_teste.py              # Gera dados ficticios para teste
|-- gestores_regionais_cc.xlsx        # Mapeamento Regional -> Email CC
|-- gestores_regionais_cc_exemplo.csv # Exemplo do mapeamento (referencia)
|-- preview.html                      # Preview do ultimo e-mail gerado
|-- requirements.txt                  # Dependencias Python
|-- template_email.html               # Template HTML do corpo do e-mail
```

---

## Inicio Rapido (VS Code / Terminal)

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
python gerar_dados_teste.py
```

Isso cria:
- `Entrada/apuracao_Q2_teste.xlsx` (planilha com 5 associados ficticios)
- `gestores_regionais_cc.xlsx` (mapeamento de regionais)
- `.secrets_config.json` (configuracao com valores de exemplo)

### 5. Execute no modo preview

```bash
python email_metas_apuracao_variavel.py --preview
```

### 6. Veja o resultado

Abra `preview.html` no navegador para ver o e-mail gerado.

---

## Parametros de Linha de Comando

```bash
python email_metas_apuracao_variavel.py [opcoes]
```

| Parametro        | Descricao                          | Default               |
| ---------------- | ---------------------------------- | --------------------- |
| --periodo        | Periodo da apuracao (P01 a P13)    | P04                   |
| --modo           | teste ou producao                  | teste                 |
| --email-teste    | Destinatario no modo teste         | seu.email@empresa.com |
| --data-retorno   | Prazo para contestacoes            | 20/01/2025            |
| --mes-pagamento  | Mes de referencia                  | fevereiro             |
| --associados     | Filtrar por nomes (virgula)        | todos                 |
| --preview        | Apenas gera HTML, nao envia        | false                 |

### Exemplos:

```bash
# Apenas preview
python email_metas_apuracao_variavel.py --preview

# Teste: envia para seu e-mail
python email_metas_apuracao_variavel.py --email-teste meuemail@gmail.com

# Producao: envia para todos
python email_metas_apuracao_variavel.py --modo producao --periodo P04

# Filtrar associados
python email_metas_apuracao_variavel.py --preview --associados "Ana Silva,Carlos Souza"
```

---

## Configurando o Power Automate

### Passo 1: Criar o Flow

1. Acesse https://make.powerautomate.com
2. Clique em **Criar** > **Fluxo de nuvem instantaneo**
3. Selecione o gatilho: **Quando uma solicitacao HTTP e recebida**

### Passo 2: Configurar o Gatilho HTTP

Cole este esquema JSON no gatilho:

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
2. Volte ao gatilho HTTP - a URL sera gerada automaticamente
3. Copie a URL (formato: `https://prod-XX.westus.logic.azure.com:443/workflows/...`)
4. Cole no `.secrets_config.json` no campo `power_automate_url`

### Passo 5: Testar o Flow

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

### Limites do Power Automate

| Item                               | Limite     |
| ---------------------------------- | ---------- |
| Chamadas/minuto (plano gratuito)   | 100        |
| Chamadas/minuto (plano premium)    | 600        |
| Tamanho max. do body               | 100 MB     |
| Timeout HTTP                       | 120s       |

O script ja trata erros 429 (rate limit), 502 e 503 com retry automatico.

---

## Configuracao do .secrets_config.json

Preencha o arquivo com seus dados:

```json
{
  "power_automate_url": "SUA_URL_DO_FLOW_AQUI",
  "assinatura_nome": "Seu Nome",
  "assinatura_cargo": "Seu Cargo",
  "assinatura_email": "seu@email.com",
  "email_gestor_excecoes": "gestor@email.com",
  "nome_gestor_excecoes": "Nome do Gestor"
}
```

| Campo                  | Descricao                              |
| ---------------------- | -------------------------------------- |
| power_automate_url     | URL do flow HTTP do Power Automate     |
| assinatura_nome        | Nome na assinatura do e-mail           |
| assinatura_cargo       | Cargo/departamento                     |
| assinatura_email       | E-mail na assinatura                   |
| email_gestor_excecoes  | E-mail para contestacoes               |
| nome_gestor_excecoes   | Nome do gestor                         |

---

## Teste com Dados Ficticios

Execute `python gerar_dados_teste.py` para criar dados automaticamente:

| Associado       | E-mail                        | Regional      |
| --------------- | ----------------------------- | ------------- |
| Ana Silva       | ana.silva@exemplo.com         | SUDESTE       |
| Carlos Souza    | carlos.souza@exemplo.com      | SUL           |
| Maria Santos    | maria.santos@exemplo.com      | NORTE         |
| Pedro Oliveira  | pedro.oliveira@exemplo.com    | NORDESTE      |
| Julia Costa     | julia.costa@exemplo.com       | CENTRO-OESTE  |

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
- **Secrets**: Dados sensiveis em `.secrets_config.json`
- **Retry**: 3 tentativas com delays progressivos em erros HTTP

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
- Power Automate (para envio real)


