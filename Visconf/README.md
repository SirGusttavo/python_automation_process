# VISCONF Engine

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![SAP GUI](https://img.shields.io/badge/SAP_GUI-0FA6D4?style=flat-square&logo=sap&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white) ![BOM](https://img.shields.io/badge/BOM-6B8E23?style=flat-square)

> Planejamento de insumos e apoio ao cálculo de necessidades para produção.

[← Voltar ao portfólio Flow](../README.md)

## 📌 Visão geral

**Desafio.** Conferir produção, estoques e listas técnicas para calcular necessidades de insumos e organizar reservas exigia cruzamentos repetidos entre diferentes fontes.

**Solução.** Ferramenta em Python para leitura e comparação de dados SAP/Excel, explosão de listas técnicas (BOM), cálculo de necessidades de materiais e pallets e geração de relatórios.

**Resultado observado.** **~40 minutos → ~2 minutos** na execução do processo original (redução aproximada de 95%).

## ⚙️ Como funciona

- Lê e cruza dados de planejamento, produção e estoque.
- Consulta informações SAP usando **SAP GUI Scripting**.
- Realiza **explosão de BOM** e cálculo de necessidades de insumos/pallets.
- Estrutura saídas em Excel com **openpyxl** e **xlwings**.
- Apresenta interface Tkinter e processamento em segundo plano.

```text
Planilhas e SAP → Validação → Explosão de BOM → Cálculo de necessidades → Relatórios Excel
```

## 🧩 Pré-requisitos

- **Windows**, Python 3 e `pip`.
- **Microsoft Excel desktop** e arquivos-base no layout esperado.
- **SAP GUI for Windows** com scripting habilitado e acesso às transações necessárias.
- Conexão SAP configurada e permissões de leitura/escrita adequadas. Identificadores como `SAP_CONNECTION_NAME` são placeholders.
- Planos de produção, listas técnicas e fontes de estoque compatíveis com a aplicação; dados corporativos reais não acompanham a versão pública.

> **Limitação:** a versão pública foi adaptada para portfólio; não constitui ambiente de demonstração independente do SAP e dos arquivos originais. A execução em ambiente real não foi validada nesta revisão de documentação.

## 🚀 Instalação

No terminal do Windows, dentro da pasta `Visconf`:

```powershell
py -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

O Tkinter costuma fazer parte da instalação padrão do Python no Windows. Excel e SAP GUI precisam ser instalados separadamente.

## ▶️ Execução

```powershell
python "visconf.py"
```

1. Selecione as planilhas e demais entradas necessárias.
2. Configure período e parâmetros disponibilizados na interface.
3. Execute a coleta e o cálculo de necessidades.
4. Revise os relatórios gerados antes de qualquer decisão operacional.

## 🧰 Arquivos

- [`visconf.py`](./visconf.py) — aplicação principal.
- [`requirements.txt`](./requirements.txt) — bibliotecas Python usadas pelo projeto.

## 🔒 Contexto e limitações

O projeto foi desenvolvido a partir de necessidades operacionais reais, em colaboração com profissionais de Controle de Produção. Dados, endereços internos e identificadores foram substituídos ou omitidos.

A anonimização do conteúdo **não equivale à autorização de publicação** nem substitui a revisão de propriedade intelectual e confidencialidade. Os resultados são históricos e aproximados; dependem das condições e do volume de operação.
