# SAP Bridge

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![SAP](https://img.shields.io/badge/SAP_GUI-0FA6D4?style=flat-square&logo=sap&logoColor=white) ![pandas](https://img.shields.io/badge/pandas-150458?style=flat-square&logo=pandas&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)

> Coleta paralela de relatórios SAP e consolidação em uma planilha mestre.

[← Voltar ao portfólio Flow](../README.md)

## 📌 Visão geral

**Desafio.** A extração de diversos relatórios SAP e sua consolidação em Excel exigiam consultas repetidas e tempo operacional elevado.

**Solução.** Coleta paralela de relatórios SAP e consolidação em uma planilha mestre.

**Resultado observado.** **~1h30 → menos de 15 minutos** no fluxo original, com redução superior a 83% no tempo.

## ⚙️ Como funciona

- Automação de transações SAP por **SAP GUI Scripting / win32com**.
- Uso de **1 a 4 sessões** em consultas compatíveis com paralelismo.
- Consolidação e atualização de dados em Excel, com **pandas** e automação COM.
- Interface Tkinter, configurações persistidas em JSON e acompanhamento de progresso.

```text
Entradas → Extração e validação → Processamento → Excel/relatório → Conferência
```

## 🧩 Pré-requisitos

- **Windows** com Python 3 e `pip`.
- **SAP GUI for Windows**, scripting habilitado no cliente e permitido no ambiente SAP.
- Credenciais e autorizações para as transações SAP necessárias.
- **Microsoft Excel desktop** para a consolidação via COM.
- Arquivo Excel mestre no layout esperado, pasta de saída acessível e conexão SAP configurada. `SAP_CONNECTION_NAME` é apenas um placeholder.

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
python "sap_bridge.py"
```

- Selecione a pasta de destino e a planilha mestre.
- Defina período e número de sessões para os relatórios compatíveis.
- Escolha os relatórios disponíveis na interface e execute.
- Acompanhe o progresso e confira a consolidação final.

## 🧰 Tecnologias e código

O ponto de entrada está em [`sap_bridge.py`](./sap_bridge.py); as dependências Python declaradas estão em [`requirements.txt`](./requirements.txt).

## 🔒 Contexto e limitações

Projeto originado de uma necessidade operacional real durante minha atuação na Suzano (2025–2026). Valores, conexões, endereços e dados específicos foram substituídos ou omitidos na publicação. **A anonimização não substitui avaliação de direitos de uso, confidencialidade ou autorização da organização.**

Resultados de desempenho são históricos e aproximados; podem variar conforme volumes, rede, permissões e condições do ambiente. Este README descreve o fluxo, mas não constitui garantia de execução em infraestrutura não configurada.
