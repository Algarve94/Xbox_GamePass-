# Xbox_GamePass Python × Excel Integration
Planilha em Excel da Xbox com integração em Python

<div align="center">

![Xbox](https://img.shields.io/badge/Xbox-107C10?style=for-the-badge&logo=xbox&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1+-green?style=for-the-badge)
![pandas](https://img.shields.io/badge/pandas-2.0+-150458?style=for-the-badge&logo=pandas&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

**Pipeline de integração Python → Excel para análise de assinaturas do Xbox Game Pass.**  
Gera automaticamente KPIs, tabelas dinâmicas, gráficos e um dashboard profissional a partir de dados brutos.

[Funcionalidades](#-funcionalidades) • [Instalação](#-instalação) • [Como usar](#-como-usar) • [Estrutura](#-estrutura-do-projeto) • [Exemplos](#-exemplos-de-código) • [Contribuição](#-contribuição)

</div>

---

## Visão Geral

Este projeto automatiza a transformação de uma planilha de dados brutos de assinaturas Xbox Game Pass em um workbook Excel profissional e interativo, com:

- Processamento automatizado de assinaturas Xbox Game Pass com geração de KPIs e dashboards Excel.
- **12 KPIs** calculados via fórmulas nativas do Excel
- **4 análises de negócio** respondidas com SUMIF, COUNTIF e SUMPRODUCT
- **3 gráficos** gerados e embutidos no Excel
- **Tabela Excel estruturada** (`TabelaDados`) com auto-expansão
- **Script Python reutilizável** para atualização incremental de dados

---

## Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| **KPI Dashboard** | Receita bruta, ticket médio, taxa de renovação, assinantes por plano |
| **Análise de Negócio** | Faturamento por tipo de assinatura, planos anuais × auto-renovação |
| **Season Passes** | Receita de EA Play e Minecraft Season Pass separados por plano |
| **Gráficos Excel** | Barras (receita por tipo), pizza (distribuição por plano), barras (season passes) |
| **Auto-atualização** | Fórmulas reativas — basta adicionar linhas na tabela |
| **API Python** | Leitura, escrita e integração com banco de dados via script |
| **Documentação embutida** | Aba `INSTRUCOES` com guia de integração dentro da própria planilha |

---

## Estrutura do Projeto

```
xbox-gamepass-excel/
│
├── xbox_excel_integration.py    # Script principal de geração
├── Planilha_xbox.xlsx           # Dados brutos de entrada
├── Xbox_GamePass_Integrado.xlsx # Workbook gerado (output)
├── README.md                    # Este arquivo
└── requirements.txt             # Dependências Python
```

### Estrutura do Workbook Gerado

```
Xbox_GamePass_Integrado.xlsx
├── DADOS        → Base limpa com TabelaDados (295 linhas)
├── KPIs         → 12 indicadores com fórmulas dinâmicas
├── PIVOT        → 4 análises de negócio (P1–P4)
├── DASHBOARD    → 3 gráficos Excel embutidos
└── INSTRUCOES   → Guia de integração e manutenção
```

---

## Instalação

### Pré-requisitos

- Python **3.9** ou superior
- pip atualizado

### Passo a passo

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/xbox-gamepass-excel.git
cd xbox-gamepass-excel

# 2. (Opcional) Crie um ambiente virtual
python -m venv .venv
source .venv/bin/activate        # Linux/macOS
.venv\Scripts\activate           # Windows

# 3. Instale as dependências
pip install -r requirements.txt
```

### `requirements.txt`

```
openpyxl>=3.1.0
pandas>=2.0.0
```

---

## Como Usar

### Geração completa (dados brutos → workbook profissional)

```bash
python xbox_excel_integration.py
```

O script lê `Planilha_xbox.xlsx` e gera `Xbox_GamePass_Integrado.xlsx` com todas as abas, fórmulas e gráficos.

### Atualizar dados sem re-rodar o script

1. Abra `Xbox_GamePass_Integrado.xlsx` no Excel
2. Vá para a aba **DADOS**
3. Clique na última linha da tabela e pressione `Tab`
4. Preencha os campos — os KPIs e gráficos se atualizam automaticamente

---

## Exemplos de Código

### Leitura da planilha gerada

```python
import pandas as pd

df = pd.read_excel("Xbox_GamePass_Integrado.xlsx", sheet_name="DADOS")

# Visualizar estrutura
print(df.shape)        # (295, 13)
print(df.dtypes)
print(df.head())

# Filtrar apenas assinantes Ultimate com auto-renovação
df_premium = df[(df["Plan"] == "Ultimate") & (df["Auto Renewal"] == "Yes")]
print(df_premium[["Name", "Total Value ($)"]].to_string())
```

### Adicionar novos registros via Python

```python
from openpyxl import load_workbook
import datetime

wb = load_workbook("Xbox_GamePass_Integrado.xlsx")
ws = wb["DADOS"]

novo_assinante = [
    3526,                          # Subscriber ID
    "Carlos Lima",                 # Name
    "Ultimate",                    # Plan
    datetime.date(2024, 6, 1),     # Start Date
    "Yes",                         # Auto Renewal
    15,                            # Subscription Price ($)
    "Monthly",                     # Subscription Type
    "Yes",                         # EA Play Season Pass
    30,                            # EA Play Season Pass Price ($)
    "Yes",                         # Minecraft Season Pass
    20,                            # Minecraft Pass Price ($)
    0,                             # Coupon Value ($)
    65,                            # Total Value ($)
]

ws.append(novo_assinante)
wb.save("Xbox_GamePass_Integrado.xlsx")
print("Registro adicionado com sucesso!")
```

### Integração com banco de dados SQLite

```python
import sqlite3
import pandas as pd

# Exportar planilha → banco de dados
df = pd.read_excel("Xbox_GamePass_Integrado.xlsx", sheet_name="DADOS")
conn = sqlite3.connect("xbox_gamepass.db")
df.to_sql("assinantes", conn, if_exists="replace", index=False)
conn.close()
print("Base exportada para SQLite")

# Consultar e gerar novo relatório
conn = sqlite3.connect("xbox_gamepass.db")
df_ultimate = pd.read_sql(
    'SELECT * FROM assinantes WHERE Plan = "Ultimate"', conn
)
df_ultimate.to_excel("relatorio_ultimate.xlsx", index=False)
conn.close()
print(f" Relatório gerado: {len(df_ultimate)} assinantes Ultimate")
```

### Ler KPIs calculados pelo Excel

```python
from openpyxl import load_workbook

wb = load_workbook("Xbox_GamePass_Integrado.xlsx", data_only=True)
ws = wb["KPIs"]

# Os valores calculados ficam nas linhas ímpares a partir da linha 5
print("Total assinantes:", ws["A5"].value)
print("Receita bruta:   ", ws["D5"].value)
print("Ticket médio:    ", ws["G5"].value)
```

---

## KPIs Disponíveis

| # | Indicador | Fórmula Excel |
|---|---|---|
| 1 | Total de Assinantes | `=COUNTA(TabelaDados[Subscriber ID])` |
| 2 | Receita Bruta Total ($) | `=SUM(TabelaDados[Total Value ($)])` |
| 3 | Ticket Médio ($) | `=AVERAGE(TabelaDados[Total Value ($)])` |
| 4 | Taxa de Auto Renovação | `=COUNTIF(...,"Yes")/COUNTA(...)` |
| 5 | Assinantes Ultimate | `=COUNTIF(TabelaDados[Plan],"Ultimate")` |
| 6 | Assinantes Standard | `=COUNTIF(TabelaDados[Plan],"Standard")` |
| 7 | Assinantes Core | `=COUNTIF(TabelaDados[Plan],"Core")` |
| 8 | Receita EA Play ($) | `=SUMIF(...,"Yes",...)` |
| 9 | Receita Minecraft ($) | `=SUMIF(...,"Yes",...)` |
| 10 | Receita Planos Mensais ($) | `=SUMIF(...,"Monthly",...)` |
| 11 | Receita Planos Anuais ($) | `=SUMIF(...,"Annual",...)` |
| 12 | Total Cupons Utilizados ($) | `=SUM(TabelaDados[Coupon Value ($)])` |

---

## Análises de Negócio (Aba PIVOT)

| Pergunta | Análise |
|---|---|
| **P1** | Faturamento total por tipo de assinatura (Monthly / Quarterly / Annual) |
| **P2** | Faturamento de planos anuais separado por auto renovação (Yes / No) |
| **P3** | Receita e quantidade de assinantes com EA Play Season Pass por plano |
| **P4** | Receita e quantidade de assinantes com Minecraft Season Pass por plano |

---

## Paleta de Cores

O projeto segue a identidade visual oficial do Xbox:

| Cor | Hex | Uso |
|---|---|---|
| Verde escuro | `#107C10` | Cabeçalhos, títulos |
| Verde médio | `#52B043` | Acentos, sub-headers |
| Verde claro | `#9BC848` | Destaques de linha |
| Verde pastel | `#D6EFC7` | Linhas alternadas |

---

## Boas Práticas e Avisos

> **Importante:** Não renomeie a tabela `TabelaDados` dentro do Excel.  
> Todas as fórmulas nas abas KPIs, PIVOT e DASHBOARD referenciam este nome.

- Sempre use `datetime.date(YYYY, MM, DD)` ao inserir datas via Python
- Ao abrir com `data_only=True` e salvar, as fórmulas são perdidas permanentemente
- Para forçar recálculo no Excel: `Ctrl + Alt + F9`
- O script pode ser agendado com **cron** (Linux/macOS) ou **Agendador de Tarefas** (Windows)

---

## Contribuição

## Próximos passos

- Exportação para Google Sheets
- Integração PostgreSQL
- Dashboard Streamlit
- Testes automatizados
  

### Ideias para contribuição

- [ ] Suporte a múltiplos arquivos de entrada (batch processing)
- [ ] Exportação automática para Google Sheets via API
- [ ] Geração de relatório PDF a partir dos KPIs
- [ ] Dashboard interativo com Streamlit
- [ ] Testes automatizados com pytest
- [ ] Integração com PostgreSQL / MySQL

---

## 📄 Licença

Distribuído sob a licença **MIT**. Veja [`LICENSE`](LICENSE) para mais informações.

---

## Autor

Desenvolvido com Python + openpyxl  
Análise de dados de assinaturas **Xbox Game Pass**

---
## Dashboard

<p align="center">
  <img src="assets/dashboard.png" width="900">
</p>

<div align="center">



** Se este projeto foi útil, deixe uma estrela no repositório!**

![Xbox](https://img.shields.io/badge/Powered_by-Python_&_openpyxl-107C10?style=flat-square&logo=python)

</div>
