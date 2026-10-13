# PGR Orchestrator

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=flat-square&logo=selenium&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white) ![Tkinter](https://img.shields.io/badge/Tkinter-795548?style=flat-square&logo=python&logoColor=white)

> Lançamento automatizado de indicadores operacionais em portal web.

[← Voltar ao portfólio Flow](../README.md)

## 📌 Visão geral

**Desafio.** Indicadores já calculados em planilhas exigiam lançamentos repetitivos em um portal corporativo.

**Solução.** Lançamento automatizado de indicadores operacionais em portal web.

**Resultado observado.** **~2min30 → ~15 segundos por execução** (redução aproximada de 90% no cenário original).

## ⚙️ Como funciona

- Leitura das informações da planilha via **xlwings / Excel COM**.
- Automação do portal com **Selenium WebDriver e Microsoft Edge**.
- Execução paralela de fluxos de indicadores, com acompanhamento de progresso e estimativa de tempo.
- Interface gráfica para configurar arquivo, período e categorias.

```text
Entradas → Extração e validação → Processamento → Excel/relatório → Conferência
```

## 🧩 Pré-requisitos

- **Windows** com Python 3 e `pip`.
- **Microsoft Excel desktop** instalado e acesso à planilha no formato esperado.
- **Microsoft Edge** e WebDriver compatível (o Selenium recente pode gerenciá-lo automaticamente, conforme o ambiente).
- Acesso autorizado ao **portal web de destino**; a URL pública no código é demonstrativa (`portal.example.com`).
- Autenticação corporativa e estrutura de indicadores compatíveis com o cenário original.

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
python "pgr_orchestrator.py"
```

- Selecione a planilha-base e informe o período.
- Defina os indicadores/categorias que deseja processar.
- Inicie o processamento e acompanhe as barras de progresso.
- Confira os lançamentos no sistema de destino.

## 🧰 Tecnologias e código

O ponto de entrada está em [`pgr_orchestrator.py`](./pgr_orchestrator.py); as dependências Python declaradas estão em [`requirements.txt`](./requirements.txt).

## 🔒 Contexto e limitações

Projeto originado de uma necessidade operacional real durante minha atuação na Suzano (2025–2026). Valores, conexões, endereços e dados específicos foram substituídos ou omitidos na publicação. **A anonimização não substitui avaliação de direitos de uso, confidencialidade ou autorização da organização.**

Resultados de desempenho são históricos e aproximados; podem variar conforme volumes, rede, permissões e condições do ambiente. Este README descreve o fluxo, mas não constitui garantia de execução em infraestrutura não configurada.
