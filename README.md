# Python Automation Process 🐍

Este diretório reúne projetos de **automação de processos desenvolvidos em Python durante minha experiência na Suzano S.A.**

Os projetos foram desenvolvidos no contexto do **Flow**, um projeto interno da Suzano voltado à automação de processos e à aplicação de programação para reduzir atividades manuais, integrar sistemas e estruturar o tratamento de dados.

Cada projeto abaixo possui um objetivo específico e diferentes abordagens técnicas, incluindo automação de SAP, sistemas web, Excel, processamento de dados, interfaces gráficas e execução paralela.

> **Importante:** os projetos foram originalmente desenvolvidos para uso interno na Suzano. As versões disponibilizadas neste repositório foram higienizadas e adaptadas para portfólio, com informações corporativas, dados sensíveis, caminhos, servidores, conexões e identificadores internos substituídos ou removidos.

---

## 📖 Índice Rápido

- [🔄 Sobre o Flow](#-sobre-o-flow)
- [📂 Projetos](#-projetos)
- [🛠️ Tecnologias](#️-tecnologias)
- [📊 Resultados](#-resultados)
- [🔒 Higienização](#-higienização)
- [📁 Estrutura](#-estrutura)

---

## 🔄 Sobre o Flow

O **Flow** foi um projeto interno da Suzano desenvolvido com o objetivo de identificar oportunidades de automação em processos operacionais.

A proposta envolvia analisar atividades que dependiam de tarefas repetitivas, consultas em sistemas corporativos, manipulação de planilhas, tratamento de dados e transferência de informações entre diferentes sistemas.

A partir dessas oportunidades, foram desenvolvidas soluções em Python capazes de automatizar diferentes etapas dos processos.

De forma geral, o trabalho seguia uma abordagem semelhante a:

**Processo → identificação da atividade manual → extração → tratamento → processamento → resultado**

Os projetos reunidos neste diretório representam diferentes aplicações dessa abordagem.

Entre as principais atividades automatizadas estão:

- Extração de informações do SAP;
- Extração de informações de sistemas web;
- Automação de planilhas Excel;
- Tratamento e transformação de dados;
- Geração de relatórios;
- Atualização de controles;
- Processamento de grandes volumes de informações;
- Execução paralela de tarefas;
- Integração entre diferentes fontes de dados.

---

## 📂 Projetos

### [PGR Orchestrator](./PGR%20Orchestrator/)

Automação desenvolvida em Python para preenchimento de indicadores em um portal corporativo.

Utiliza **Selenium WebDriver (Edge)** para interação com o sistema web e **xlwings** para leitura de dados diretamente de planilhas Excel.

### Principais recursos

- Automação web com Selenium;
- Extração de dados do Excel via `xlwings` e COM;
- Execução paralela utilizando dois drivers;
- Interface gráfica em Tkinter;
- Barras de progresso e cálculo de ETA;
- Fluxo de login SSO;
- Notificações do Windows;
- Tratamento de erros e retentativas.

**Tecnologias:** Python, Selenium WebDriver, Edge, xlwings, Tkinter, threading e `concurrent.futures`.

[**→ Ver README completo do PGR Orchestrator**](./PGR%20Orchestrator/)

---

### [Lot Miner](./Lot%20Miner/)

Ferramenta de análise em lote desenvolvida em Python para extrair e processar informações de lotes de produtos.

O projeto utiliza automação web para consultar relatórios SSRS e possui integração opcional com SAP para complementar os dados obtidos.

### Principais recursos

- Extração de dados via SSRS;
- Selenium WebDriver com Microsoft Edge;
- Processamento paralelo com múltiplas janelas e abas;
- Integração opcional com SAP MB52;
- Tratamento de dados com `pandas`;
- Geração de relatórios Excel;
- Formatação condicional;
- Interface gráfica em Tkinter;
- Monitoramento do processamento em tempo real.

**Tecnologias:** Python, Selenium, SAP GUI Scripting, pandas, BeautifulSoup, `pd.read_html`, openpyxl, xlwings, Tkinter, threading e queue.

[**→ Ver README completo do Lot Miner**](./Lot%20Miner/)

---

### [Prod Forge](./Prod%20Forge/)

Automação em Python desenvolvida para criação e manutenção de planilhas de acompanhamento da produção.

Utiliza SAP GUI Scripting por meio de COM para extrair informações de produção e estruturar um arquivo Excel consolidado.

### Principais recursos

- Registro automático de ordens de produção;
- Consulta e validação de ordens no SAP;
- Busca de pesos de materiais;
- Extração de dados de produção;
- Correção de cache XML de tabelas dinâmicas;
- Gerenciamento de segmentadores (Slicers);
- Geração e atualização de planilhas;
- Interface gráfica em Tkinter;
- Pipeline de execução;
- Terminal de logs integrado.

**Tecnologias:** Python, SAP GUI Scripting, `win32com.client`, openpyxl, pandas, Tkinter, threading e zipfile.

[**→ Ver README completo do Prod Forge**](./Prod%20Forge/)

---

### [SAP Bridge](./SAP%20Bridge/)

Ferramenta de automação multi-thread para SAP GUI Scripting.

O projeto extrai informações de diferentes transações do SAP e consolida os resultados em uma planilha Excel mestre utilizando `pandas` e automação COM.

### Principais recursos

- Interface gráfica em Tkinter;
- Tema escuro;
- Acompanhamento de progresso;
- Extração paralela com múltiplas janelas SAP;
- Suporte a 1–4 sessões para determinadas consultas;
- Notificações nativas do Windows;
- Persistência das configurações em JSON;
- Controle de concorrência entre threads;
- Consolidação dos resultados em Excel.

**Tecnologias:** Python, SAP GUI Scripting, COM, `win32com`, `pythoncom`, pandas, openpyxl, Tkinter, tkcalendar e threading.

[**→ Ver README completo do SAP Bridge**](./SAP%20Bridge/)

---

### [Techlyst](./Techlyst/)

Aplicação desenvolvida em Python para automatizar a extração de **Listas Técnicas (Bill of Materials — BOM)** diretamente do SAP.

O projeto foi desenvolvido para lidar com processos de extração em grande volume, utilizando múltiplas sessões do SAP para aumentar a capacidade de processamento.

### Principais recursos

- Extração multi-threaded;
- Execução simultânea em múltiplas sessões do SAP;
- Sincronização entre workers;
- Automação da SAP GUI via COM;
- Tratamento e limpeza dos dados com `pandas`;
- Classificação de materiais;
- Persistência de dados em JSON;
- Geração de arquivos `.xlsx` formatados;
- Filtros e validação visual dos dados;
- Interface gráfica para configuração e acompanhamento;
- Logs e status em tempo real.

**Tecnologias:** Python, SAP GUI Scripting, `win32com`, pandas, openpyxl, Tkinter e threading.

[**→ Ver README completo do Techlyst**](./Techlyst/)

---

### [VISCONF Engine](./VISCONF%20Engine/)

Aplicação desenvolvida em Python para planejamento e análise de necessidades de insumos para produção.

O projeto utiliza informações provenientes do Excel e do SAP para realizar o tratamento dos dados, explosão de BOM, cálculo de necessidades e geração de relatórios.

### Principais recursos

- Pipeline completo de processamento;
- Interface gráfica configurável;
- Integração com SAP;
- Extração automatizada de dados;
- Explosão de listas de materiais (BOM);
- Cálculo de necessidades de insumos;
- Alocação de pallets;
- Geração de relatórios estruturados em Excel;
- Processamento em background.

**Tecnologias:** Python, xlwings, openpyxl, SAP GUI Scripting, Tkinter e threading.

[**→ Ver README completo do VISCONF Engine**](./VISCONF%20Engine/)

---

## 🛠️ Tecnologias

Os projetos utilizam diferentes tecnologias de acordo com as necessidades de cada processo.

### Linguagem

- **Python**

### Automação

- **SAP GUI Scripting**
- **Selenium WebDriver**
- **COM / win32com**
- **Automação do Excel**

### Tratamento de dados

- **pandas**
- **openpyxl**
- **xlwings**
- **BeautifulSoup**
- **pd.read_html**

### Interface e execução

- **Tkinter / ttk**
- **threading**
- **concurrent.futures**
- **queue**

### Sistemas envolvidos

- **SAP**
- **Microsoft Excel**
- **Microsoft Edge**
- **Sistemas web corporativos**
- **Relatórios SSRS**

---

## 📊 Resultados

Os projetos apresentaram ganhos diferentes de acordo com o processo automatizado.

### Lot Miner

A implementação do processamento paralelo reduziu o tempo médio de processamento de aproximadamente **6–7 segundos para cerca de 2 segundos por lote**.

Em processos que poderiam levar até aproximadamente **1 hora**, a execução passou a ocorrer em **menos de 10 minutos**, dependendo da quantidade de bobinas ou lotes.

### Techlyst

O processamento paralelo permitiu processar **mais de 180 SKUs em aproximadamente 1 hora utilizando 4 threads**.

A extração direta do SAP também reduziu problemas relacionados a etapas intermediárias de cópia e tratamento manual dos dados.

### VISCONF Engine

O tempo de execução do processo foi reduzido de aproximadamente **40 minutos para cerca de 2 minutos**, representando uma redução aproximada de **95%**.

### PGR Orchestrator

- Redução de atividades manuais;
- Padronização do preenchimento;
- Execução paralela de operações;
- Acompanhamento do processamento em tempo real.

### Prod Forge

- Automação de etapas de controle da produção;
- Padronização da atualização das planilhas;
- Integração entre dados do SAP e Excel;
- Automatização da manutenção de estruturas utilizadas nos relatórios.

### SAP Bridge

- Automação da coleta de informações de diferentes consultas;
- Consolidação dos dados em um único arquivo;
- Suporte à execução paralela;
- Redução da necessidade de consultas manuais repetitivas.

---

## 🔒 Higienização

Os códigos foram originalmente desenvolvidos para utilização em ambiente corporativo.

Para publicação no GitHub, os projetos passaram por um processo de higienização e anonimização.

Foram removidos ou substituídos:

- Nomes de servidores;
- Caminhos de rede;
- Endereços e URLs internos;
- Conexões SAP;
- Dados de produção;
- Dados de materiais;
- Nomes de empresas, áreas e usuários;
- Centros e depósitos;
- Identificadores de máquinas;
- Configurações específicas do ambiente;
- Credenciais, tokens e informações sensíveis.

Quando necessário, valores reais foram substituídos por exemplos ou placeholders genéricos, preservando a estrutura necessária para demonstrar o funcionamento do código.

> A versão pública não tem como objetivo reproduzir o ambiente interno da Suzano. O objetivo é demonstrar as soluções técnicas, as abordagens de automação e os conhecimentos aplicados no desenvolvimento.

---

## 📁 Estrutura

```text
python_automation_process/
│
├── PGR Orchestrator/
│   └── README.md
│
├── Lot Miner/
│   └── README.md
│
├── Prod Forge/
│   └── README.md
│
├── SAP Bridge/
│   └── README.md
│
├── Techlyst/
│   └── README.md
│
├── VISCONF Engine/
│   └── README.md
│
└── README.md
```

Cada projeto possui um README próprio contendo informações mais detalhadas sobre:

- Funcionamento;
- Estrutura do código;
- Tecnologias utilizadas;
- Requisitos;
- Instalação;
- Execução;
- Resultados;
- Processo de higienização.

---

## 📌 Contexto profissional

Os projetos apresentados neste diretório representam parte da experiência prática adquirida durante minha atuação como **Jovem Aprendiz na Suzano S.A.**, trabalhando com desenvolvimento de soluções, automação de processos e tratamento de dados em ambiente corporativo.

O conjunto demonstra a aplicação prática de Python em diferentes cenários, desde automações simples de tarefas repetitivas até soluções envolvendo **integração de sistemas, processamento paralelo, manipulação de grandes volumes de dados e geração automatizada de relatórios**.
