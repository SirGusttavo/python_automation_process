# Lot Miner

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=flat-square&logo=selenium&logoColor=white) ![pandas](https://img.shields.io/badge/pandas-150458?style=flat-square&logo=pandas&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)

> Extração de relatórios SSRS e consolidação de dados de lotes em Excel.

[← Voltar ao portfólio Flow](../README.md)

## 📌 Visão geral

**Desafio.** Consultas manuais repetidas a relatórios de lotes tornavam a conferência e a investigação de divergências demoradas.

**Solução.** Extração de relatórios SSRS e consolidação de dados de lotes em Excel.

**Resultado observado.** **~10–15 segundos → ~2 segundos por lote**, com processos de até aproximadamente **1 hora** reduzidos para **menos de 10 minutos** conforme volume e condições de uso.01

## ⚙️ Como funciona

- Automação da navegação e extração de tabelas em **relatórios SSRS**, com **Selenium/Edge**.
- Consultas paralelas em navegadores e abas, configuradas pela interface.
- Integração **opcional** com SAP GUI Scripting para complementar informações de estoque (MB52).
- Tratamento de dados via **pandas** e exportação de planilhas formatadas com **openpyxl**.
- Acompanhamento do processamento pela interface Tkinter.

```text
Entradas → Extração e validação → Processamento → Excel/relatório → Conferência
```

## 🧩 Pré-requisitos

- **Windows** com Python 3 e `pip`.
- **Microsoft Edge** e WebDriver compatível.
- Acesso autorizado ao **servidor de relatórios SSRS**; o domínio `reports.example.com` é ilustrativo.
- Planilha ou lista de lotes com as colunas/formato esperados.
- Para a integração **opcional** com SAP: **SAP GUI for Windows**, scripting habilitado e acesso autorizado.
- **Microsoft Excel desktop** para os fluxos que usam xlwings/COM.

**Observação:** a versão pública foi adaptada para fins de portfólio. Ela **não é uma aplicação de demonstração independente dos sistemas corporativos**. Requer configurações próprias e não foi validada aqui em ambiente SAP/SSRS real.

## 🚀 Instalação

Dentro da pasta deste módulo, em um terminal do Windows:

```powershell
py -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

> O módulo **Tkinter** costuma acompanhar a instalação padrão do Python para Windows. Os componentes externos (SAP GUI, Excel e Edge, quando aplicáveis) são instalados separadamente e não pelo `pip`.

## ▶️ Execução

```powershell
python "lot_miner.py"
```

- Informe a lista/planilha de lotes e as configurações de consulta.
- Configure o paralelismo e, se aplicável, a etapa opcional SAP.
- Execute as consultas e acompanhe os lotes processados.
- Revise os relatórios Excel e os alertas gerados.

## 🧰 Tecnologias e código

O ponto de entrada está em [`lot_miner.py`](./lot_miner.py); as dependências Python declaradas estão em [`requirements.txt`](./requirements.txt).

## 🔒 Contexto e limitações

Projeto originado de uma necessidade operacional real durante minha atuação na Suzano (2025–2026). Valores, conexões, endereços e dados específicos foram substituídos ou omitidos na publicação. **A anonimização não substitui avaliação de direitos de uso, confidencialidade ou autorização da organização.**

Resultados de desempenho são históricos e aproximados; podem variar conforme volumes, rede, permissões e condições do ambiente. Este README descreve o fluxo, mas não constitui garantia de execução em infraestrutura não configurada.
