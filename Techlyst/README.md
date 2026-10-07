# Techlyst (TecList Forge)

**Status:** Código Higienizado para Portfólio

## Visão Geral
O **Techlyst** (internamente chamado de *TecList Forge*) é uma aplicação desenvolvida em Python para automatizar a extração de Listas Técnicas (Bill of Materials - BOM) do sistema SAP, utilizando as transações C203 e MM03. 

Esta ferramenta foi projetada para lidar com processos de extração de alto volume e otimizar o tempo através de sessões simultâneas do SAP, processando dados complexos de produção e convertendo-os em planilhas Excel formatadas para análise e tomada de decisão.

## Principais Recursos
- **Extração Multi-threaded:** Utiliza o módulo `threading` para realizar a extração simultânea de dados em várias sessões do SAP, com sincronização por *barrier* (worker barrier synchronization), aumentando drasticamente a velocidade de coleta de dados.
- **Banco de Dados Híbrido Inovador (Memória + JSON em AppData):** A aplicação possui um sistema robusto de banco de dados híbrido. Ele utiliza um dicionário de classificação de materiais em memória e o sincroniza com arquivos JSON salvos na pasta AppData do usuário. Isso permite controle de versão, rápida recuperação de dados e fácil atualização sem necessidade de reimplantação do script.
- **Automação de Interface SAP:** Integração direta com a SAP GUI via COM (Component Object Model) utilizando `win32com`, automatizando cliques, inserção de dados e leitura de campos (Scraping).
- **Processamento de Dados:** Tratamento e limpeza das informações extraídas usando `pandas`.
- **Exportação Profissional:** Geração de arquivos `.xlsx` formatados utilizando `openpyxl`, com tabelas estilizadas, filtros e validação visual de dados.
- **Interface Gráfica Amigável:** Desenvolvido com `Tkinter` (e `ttk`), proporcionando um painel intuitivo para configuração de threads, entrada de SKUs, seleção de centros e acompanhamento em tempo real do log e status.

## Tecnologias Utilizadas
- **Linguagem:** Python
- **Automação SAP:** SAP GUI Scripting API (COM/win32com)
- **Manipulação de Dados:** pandas
- **Geração de Relatórios Excel:** openpyxl
- **Interface Gráfica (GUI):** Tkinter
- **Concorrência:** threading (Barrier, Lock, Event)

## Nota sobre a Higienização do Código
O código-fonte deste repositório foi rigorosamente **higienizado** para remover quaisquer dados sensíveis, informações corporativas confidenciais e credenciais, permitindo sua exibição pública como parte de um portfólio de engenharia de software e automação.

As alterações incluem:
- Nomes de servidores, caminhos de rede e conexões SAP substituídos por variáveis e placeholders genéricos (`SAP_CONNECTION_NAME`, etc.).
- O extenso banco de dados interno de materiais foi resumido a dados de exemplo para demonstrar a estrutura de chave-valor.
- Nomes de plantas, depósitos, layouts e códigos de máquinas (`DEPARA_LINHAS`) foram substituídos por códigos genéricos (`Plant-A`, `MACHINE-01`, etc.).
- Nomes de empresas, portais internos e usuários reais foram completamente anonimizados.

*Este projeto demonstra capacidades avançadas em automação de sistemas corporativos (RPA), arquitetura de dados híbrida e programação concorrente em Python.*
