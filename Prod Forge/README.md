# Prod Forge

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![SAP](https://img.shields.io/badge/SAP_GUI-0FA6D4?style=flat-square&logo=sap&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white) ![COM](https://img.shields.io/badge/Windows_COM-0078D4?style=flat-square&logo=windows&logoColor=white)

> Atualização de controles de produção a partir de consultas SAP e planilhas Excel.

[← Voltar ao portfólio Flow](../README.md)

## 📌 Visão geral

**Desafio.** Atualizações periódicas de controles de produção dependiam de consultas manuais e tratamento de arquivos Excel.

**Solução.** Atualização de controles de produção a partir de consultas SAP e planilhas Excel.

**Resultado observado.** **~10 minutos → ~2 minutos** por atualização, no processo original (redução aproximada de 80%).

## ⚙️ Como funciona

- Extração e conferência de dados por **SAP GUI Scripting** (transações como MB51, COR3 e MM03).
- Tratamento e consolidação de informações em planilhas Excel.
- Rotinas de manutenção de tabelas dinâmicas, segmentadores e estruturas do arquivo.
- Interface Tkinter com etapas do processamento e terminal de logs.

```text
Entradas → Extração e validação → Processamento → Excel/relatório → Conferência
```

## 🧩 Pré-requisitos

- **Windows** com Python 3 e `pip`.
- **SAP GUI for Windows** instalado, scripting habilitado e acesso autorizado às transações utilizadas.
- **Microsoft Excel desktop** com suporte à automação COM e planilha-base compatível.
- Conexão SAP e caminhos locais **configurados para o seu ambiente**; valores demonstrativos do código não representam infraestrutura funcional.
- Permissões para ler os relatórios e gravar a planilha de saída.

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
python "prod_forge.py"
```

- Abra ou selecione a base de produção esperada pela aplicação.
- Configure período, arquivos e locais de saída, conforme os controles disponíveis.
- Execute as consultas e acompanhe cada etapa pelos logs.
- Valide a atualização da planilha e das tabelas dinâmicas.

## 🧰 Tecnologias e código

O ponto de entrada está em [`prod_forge.py`](./prod_forge.py); as dependências Python declaradas estão em [`requirements.txt`](./requirements.txt).

## 🔒 Contexto e limitações

Projeto originado de uma necessidade operacional real durante minha atuação na Suzano (2025–2026). Valores, conexões, endereços e dados específicos foram substituídos ou omitidos na publicação. **A anonimização não substitui avaliação de direitos de uso, confidencialidade ou autorização da organização.**

Resultados de desempenho são históricos e aproximados; podem variar conforme volumes, rede, permissões e condições do ambiente. Este README descreve o fluxo, mas não constitui garantia de execução em infraestrutura não configurada.
