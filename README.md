# Flow · Python Automation Process

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white) ![SAP](https://img.shields.io/badge/SAP_GUI-0FA6D4?style=flat-square&logo=sap&logoColor=white) ![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=flat-square&logo=selenium&logoColor=white) ![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)

**Seis automações em Python** desenvolvidas durante minha experiência na **Suzano (2025–2026)**, focadas em reduzir tarefas manuais, integrar sistemas e facilitar a análise de dados operacionais.

## 📂 Projetos

| Módulo | Aplicação | Resultado observado |
|---|---|---|
| [PGR Orchestrator](./PGR%20Orchestrator/) | Lançamento de indicadores web | ~2min30 → ~15s |
| [SAP Bridge](./SAP%20Bridge/) | Extração e consolidação SAP | ~1h30 → menos de 15min |
| [Lot Miner](./Lot%20Miner/) | Análise de lotes via SSRS | até ~1h → menos de 10min |
| [Prod Forge](./Prod%20Forge/) | Atualização de controles Excel | ~10min → ~2min |
| [VISCONF Engine](./Visconf/) | Planejamento de insumos e BOM | ~40min → ~2min |
| [Techlyst](./Techlyst/) | Extração de listas técnicas SAP | 180+ materiais em ~1h |

*Resultados aproximados dos processos originais. Cada módulo possui README com funcionamento, requisitos e execução.*

## 🛠️ Tecnologias

**Python**, Selenium, SAP GUI Scripting, `pywin32`/COM, pandas, openpyxl, xlwings, Tkinter e processamento paralelo.

## ▶️ Pré-requisitos

**Windows e Python 3**. Conforme o módulo, também são necessários **Excel desktop**, **SAP GUI com scripting habilitado**, **Microsoft Edge**, acesso ao SSRS/portal e arquivos de entrada compatíveis.

Na pasta do módulo escolhido:

```powershell
py -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

A execução é feita pelo script `.py` indicado no README de cada módulo.

> **Portfólio:** código e configurações adaptados, sem os acessos e dados corporativos originais. Não é uma demonstração independente do ambiente SAP/SSRS. A publicação deve respeitar confidencialidade e direitos de uso do código.
