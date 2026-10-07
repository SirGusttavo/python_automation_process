# Lot Miner

**Lot Miner** é uma ferramenta de análise em lote desenvolvida em Python para extrair e processar dados de lotes de produtos. O sistema utiliza automação web e integração com ERP para consolidar informações em relatórios detalhados, permitindo um acompanhamento preciso da produção.

## Principais Funcionalidades

- **Extração de Dados via SSRS**: Utiliza Selenium WebDriver (Edge) para navegar, interagir e extrair dados tabulares de relatórios web internos (SSRS) armazenados em iframes.
- **Processamento Paralelo**: Implementa extração paralela com múltiplas instâncias de navegadores (número configurável de janelas do Edge × abas), acelerando significativamente a coleta de dados de grandes volumes de lotes.
- **Integração Opcional com SAP (MB52)**: Permite a extração de dados de estoque através de SAP GUI Scripting, integrando os dados de chão de fábrica com o ERP.
- **Saída em Excel Formatado**: Utiliza `pandas` e `openpyxl` para estruturar, limpar e exportar os dados em um arquivo Excel completo com planilhas de análise e formatação condicional.
- **Interface Gráfica (GUI)**: Apresenta uma interface amigável desenvolvida em `Tkinter` (estilo wizard), incluindo uma janela de acompanhamento do progresso da extração em tempo real.

## Tecnologias Utilizadas

- **Python**: Linguagem principal.
- **Selenium WebDriver (Edge)**: Automação web e web scraping.
- **SAP GUI Scripting (COM)**: Automação e interação com SAP ERP.
- **pandas**: Manipulação e análise de dados tabulares.
- **BeautifulSoup / pd.read_html**: Parsing de HTML.
- **openpyxl / xlwings**: Geração e formatação avançada de relatórios Excel.
- **Tkinter**: Construção da interface gráfica e controle em tempo real.
- **threading & queue**: Gerenciamento do processamento paralelo de tarefas web.

## Estrutura do Código

- **JanelaPrincipal**: Interface de configuração inicial dividida em seções (wizard).
- **Estado (Bridge)**: Gerenciamento thread-safe que permite comunicação contínua entre os workers de extração e a interface de usuário.
- **Worker de Extração**: Rotinas em background que paralelizam as consultas web utilizando Selenium.
- **JanelaProgresso**: Janela para monitoramento dinâmico (polling em milissegundos) do status atual da extração e processamento.
- **Exportação e Formatação**: Funções que padronizam e formatam o relatório de saída, aplicando alertas de qualidade visualmente (células coloridas) quando identificados desvios nos lotes analisados.

## Notas

*Nota: Este projeto foi higienizado para remover informações corporativas sensíveis, tornando-o adequado para fins de demonstração em portfólio.*
