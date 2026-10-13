# Techlyst

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![SAP](https://img.shields.io/badge/SAP_GUI-0FA6D4?style=flat-square&logo=sap&logoColor=white) ![pandas](https://img.shields.io/badge/pandas-150458?style=flat-square&logo=pandas&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)

> Extração automatizada de listas técnicas (Bill of Materials — BOM) do SAP.

[← Voltar ao portfólio Flow](../README.md)

## 📌 Visão geral

**Desafio.** A consulta individual de listas técnicas e sua organização em planilhas consumiam tempo, especialmente para grandes volumes de materiais.

**Solução.** Extração automatizada de listas técnicas (Bill of Materials — BOM) do SAP.

**Resultado observado.** **Mais de 180 materiais em aproximadamente 1 hora**, com quatro sessões SAP no cenário original.

## ⚙️ Como funciona

- Automação de consultas SAP (incluindo C203 e MM03) usando **SAP GUI Scripting**.
- Extração paralela com múltiplas sessões e sincronização entre trabalhadores.
- Tratamento e classificação dos materiais com **pandas**.
- Geração de arquivos `.xlsx` com **openpyxl**; preferências/classificações auxiliares armazenadas em JSON.
- Interface gráfica com controle de execução e logs.

```text
Entradas → Extração e validação → Processamento → Excel/relatório → Conferência
```

## 🧩 Pré-requisitos

- **Windows** com Python 3 e `pip`.
- **SAP GUI for Windows** com scripting habilitado e acesso autorizado às consultas de listas técnicas.
- Conexão SAP configurada; placeholders como `SAP_CONNECTION_NAME` não funcionam sem adaptação.
- Materiais de entrada válidos e permissões para gravar a planilha de saída.
- Para executar múltiplas sessões, o ambiente precisa permitir a abertura e utilização simultânea delas.

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
python "techlyst.py"
```

- Informe os materiais e os parâmetros solicitados na interface.
- Selecione a configuração de sessões compatível com seu ambiente.
- Inicie a extração e acompanhe o status.
- Confira as classificações e a planilha Excel exportada.

## 🧰 Tecnologias e código

O ponto de entrada está em [`techlyst.py`](./techlyst.py); as dependências Python declaradas estão em [`requirements.txt`](./requirements.txt).

## 🔒 Contexto e limitações

Projeto originado de uma necessidade operacional real durante minha atuação na Suzano (2025–2026). Valores, conexões, endereços e dados específicos foram substituídos ou omitidos na publicação. **A anonimização não substitui avaliação de direitos de uso, confidencialidade ou autorização da organização.**

Resultados de desempenho são históricos e aproximados; podem variar conforme volumes, rede, permissões e condições do ambiente. Este README descreve o fluxo, mas não constitui garantia de execução em infraestrutura não configurada.
