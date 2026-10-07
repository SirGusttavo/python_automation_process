# Prod Forge

**Prod Forge** é uma automação em Python desenvolvida para a criação e manutenção de planilhas de acompanhamento de produção. A ferramenta utiliza scripts de SAP GUI (através de COM/win32com nas transações MB51, COR3, MM03) para extrair dados de produção e estruturar um arquivo Excel consolidado com tabelas dinâmicas.

## Funcionalidades Principais

- **Registro Automático de Ordens:** Consulta e valida as ordens de produção no SAP.
- **Busca de Pesos:** Utiliza a transação MM03 para localizar os pesos dos materiais, alimentando os dados das ordens.
- **Correção de Cache XML (Tabelas Dinâmicas):** Edita o cache XML das planilhas para garantir o funcionamento correto e atualização das tabelas dinâmicas.
- **Gerenciamento de Segmentadores (Slicers):** Automatiza a manipulação dos filtros de tabelas dinâmicas para a interface gerada.
- **Interface Gráfica (GUI):** Interface construída em Tkinter que oferece visualização em pipeline (etapas do processo) e um terminal de logs integrado.

## Tecnologias Utilizadas

- **Python**
- **SAP GUI Scripting** (via `win32com.client`)
- **openpyxl** (Manipulação de arquivos Excel)
- **pandas** (Análise e estruturação de dados)
- **Tkinter** (Interface gráfica do usuário)
- **threading** (Execução assíncrona de processos)
- **zipfile** (Edição e manutenção de arquivos internos do Excel, como o cache de XML das tabelas dinâmicas)

## Configuração e Instalação

1. Clone ou baixe este repositório.
2. Certifique-se de que o Python 3.x esteja instalado e configurado no seu PATH.
3. Instale as dependências via `pip`:
   ```bash
   pip install pandas openpyxl pypiwin32
   ```
4. Configure sua conexão com o SAP Logon (certifique-se de que a conexão padrão e o SAP GUI Scripting estejam habilitados tanto no servidor quanto no cliente).

## Execução

Execute o arquivo principal do Python (ex: `Prod Forge (v7.2.2) - codigo higienizado.py`). A interface gráfica será aberta e as instruções de progresso estarão visíveis no terminal de logs da aplicação.

## Estrutura do Projeto

O código passou por um processo de higienização de informações corporativas e restritas. Nomes de servidores, caminhos de rede absolutos, chaves de layout específicas de usuário, centros e depósitos sensíveis foram substituídos por valores genéricos (ex: `Plant-A`, `MACHINE-01`, etc). O comportamento da lógica principal permanece intacto.
