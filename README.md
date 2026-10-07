# Python Automation Process

## Sobre o projeto

**Python Automation Process** é a organização, para fins de portfólio, de soluções de automação desenvolvidas durante minha experiência na **Suzano S.A.**

Os projetos apresentados nesta pasta tiveram origem em um projeto interno denominado **Flow**, voltado à identificação e automação de atividades operacionais que envolviam tarefas repetitivas, consultas em sistemas corporativos, tratamento de dados e atualização de controles.

O objetivo não era simplesmente desenvolver aplicações em Python, mas utilizar programação e automação para transformar processos manuais em fluxos mais rápidos, padronizados e confiáveis.

> **Flow foi um projeto interno da Suzano.** Esta versão pública foi reorganizada e higienizada exclusivamente para demonstrar os conhecimentos técnicos e as soluções desenvolvidas, sem expor informações, dados ou estruturas pertencentes à empresa.

---

## O que era o Flow?

O Flow surgiu a partir da identificação de oportunidades de automação em processos utilizados no dia a dia.

Em vez de tratar cada atividade como uma automação isolada, o projeto buscava estruturar soluções capazes de atuar sobre diferentes etapas de um processo, como:

- Extração de informações de sistemas corporativos.
- Consulta automatizada de dados.
- Tratamento e transformação de informações.
- Integração entre SAP, sistemas web e Excel.
- Processamento de grandes volumes de registros.
- Execução paralela de tarefas.
- Geração e atualização de relatórios.
- Redução de atividades manuais e repetitivas.

A abordagem utilizada seguia, de forma geral, o fluxo:

**Processo → identificação do trabalho manual → estruturação dos dados → automação → processamento → resultado**

Essa experiência também permitiu trabalhar com diferentes estratégias de automação, escolhendo a abordagem de acordo com o problema em vez de aplicar uma única tecnologia a todos os casos.

---

## Soluções desenvolvidas

A pasta reúne diferentes projetos originados desse contexto, cada um direcionado a uma necessidade específica.

| Projeto | Objetivo principal |
|---|---|
| **PGR Orchestrator** | Automatizar o preenchimento de indicadores em sistema web |
| **Lot Miner** | Automatizar a coleta e o processamento de informações de lotes |
| **Prod Forge** | Automatizar etapas de controle e acompanhamento da produção |
| **SAP Bridge** | Automatizar e consolidar consultas realizadas no SAP |
| **Techlyst** | Automatizar a extração e tratamento de listas de materiais (BOM) |
| **VISCONF Engine** | Automatizar análises e cálculos relacionados às necessidades de insumos |

Apesar de possuírem objetivos diferentes, os projetos compartilham uma mesma abordagem: **extrair dados, processá-los e transformar uma sequência de tarefas manuais em um fluxo automatizado.**

---

## Tecnologias e conceitos

O desenvolvimento das soluções envolveu principalmente:

- **Python**
- **SAP GUI Scripting**
- **Selenium**
- **pandas**
- **openpyxl**
- **xlwings**
- **Tkinter**
- **threading**
- **concurrent.futures**
- **Automação de Excel**
- **Processamento paralelo**
- **Tratamento e transformação de dados**
- **Integração entre sistemas**

A escolha das tecnologias variava de acordo com o processo. Em alguns casos, a automação dependia diretamente do SAP; em outros, de sistemas web, arquivos Excel ou da combinação entre diferentes fontes.

---

## Resultados

O desenvolvimento das automações gerou ganhos principalmente relacionados a **tempo de execução, redução de tarefas manuais, padronização e capacidade de processamento**.

Alguns resultados obtidos durante o desenvolvimento incluem:

### Lot Miner

Redução do tempo médio de processamento de aproximadamente **6–7 segundos para cerca de 2 segundos por lote**, utilizando processamento paralelo.

Em processos que poderiam levar até aproximadamente **1 hora**, a execução passou a ocorrer em **menos de 10 minutos**, dependendo do volume processado.

### Techlyst

Processamento de **mais de 180 SKUs em aproximadamente 1 hora**, utilizando 4 threads.

A utilização do SAP como fonte direta dos dados também reduziu problemas associados a etapas intermediárias de cópia e tratamento manual.

### VISCONF Engine

Redução do tempo de execução de aproximadamente **40 minutos para cerca de 2 minutos**, representando uma redução aproximada de **95%** no tempo necessário para executar o processo.

---

## Arquitetura geral

Embora cada aplicação possua sua própria implementação, os projetos seguem uma estrutura conceitual semelhante:

```text
┌──────────────────────┐
│      Fonte de dados  │
│ SAP / Web / Excel    │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│       Extração       │
│ Coleta automatizada  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│       Tratamento     │
│ Validação / limpeza  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│      Processamento   │
│ Regras / cálculos    │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│        Saída         │
│ Excel / relatório    │
└──────────────────────┘
```

Quando necessário, o processamento também pode ser distribuído entre múltiplas threads ou sessões para aumentar a capacidade de execução.

---

## Relação com a experiência profissional

O Flow representou uma experiência prática de desenvolvimento de software aplicada a problemas reais de ambiente corporativo.

Mais do que desenvolver scripts isolados, o trabalho envolveu compreender o processo existente, identificar pontos de intervenção, lidar com diferentes fontes de dados e sistemas e transformar essas etapas em soluções automatizadas.

Essa experiência contribuiu para o desenvolvimento de conhecimentos em:

- Desenvolvimento em Python.
- Automação de processos.
- Integração de sistemas.
- Manipulação e tratamento de dados.
- Automação de SAP e Excel.
- Automação de sistemas web.
- Processamento paralelo.
- Desenvolvimento de interfaces desktop.
- Estruturação e manutenção de aplicações.

---

## Sobre esta versão

Os projetos presentes neste diretório foram **higienizados e adaptados para portfólio**.

Foram removidos ou generalizados elementos específicos do ambiente corporativo, incluindo:

- Dados de produção.
- Informações de materiais.
- Nomes de pessoas e áreas.
- Servidores e caminhos de rede.
- URLs e endereços internos.
- Identificadores corporativos.
- Arquivos e configurações proprietárias.
- Credenciais, tokens e outras informações sensíveis.

A implementação apresentada tem como objetivo demonstrar **a abordagem técnica, a estrutura das soluções e os conhecimentos aplicados**, e não reproduzir o ambiente interno da Suzano.

---

## Estrutura

```text
python_automation_process/
│
├── PGR Orchestrator/
├── Lot Miner/
├── Prod Forge/
├── SAP Bridge/
├── Techlyst/
├── VISCONF Engine/
│
└── README.md
```

Cada diretório possui seu próprio README com informações específicas sobre o problema, funcionamento, tecnologias e resultados da respectiva solução.
