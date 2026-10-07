# PGR Orchestrator

O **PGR Orchestrator** é uma automação desenvolvida em Python para preenchimento de indicadores no portal corporativo (PGR). Ele utiliza Selenium WebDriver (Edge) e xlwings para integração com planilhas Excel, automatizando o lançamento de dados de Waste e Consumo com suporte a execução paralela (dual-driver).

## Funcionalidades

- **Automação Web com Selenium**: Interação com o portal corporativo via Microsoft Edge.
- **Processamento de Dados**: Extração de dados diretamente do Excel via `xlwings` usando automação COM.
- **Execução Paralela**: Suporte a execução de dois drivers (um para Waste e outro para Consumo) simultaneamente usando `concurrent.futures`.
- **Interface Gráfica Avançada (GUI)**: Criada em Tkinter com tema escuro, contendo barras de progresso duplas, cálculo dinâmico de ETA e separação em cards.
- **Auto-login SSO**: Preenchimento automático de e-mail e tratamento de fluxo de login corporativo.
- **Notificações do Windows**: Alertas via balões do sistema ("toast notifications").

## Tecnologias Utilizadas

- **Python**
- **Selenium WebDriver (Edge)**
- **xlwings**
- **Tkinter / ttk**
- **threading** e **concurrent.futures**

## Estrutura do Código

- A interface principal está isolada na classe `ConfiguracaoGUI` (dark theme).
- O progresso em tempo real é exibido pela classe `JanelaProgresso`, com atualização na main thread.
- O mapeamento de indicadores e máquinas foi generalizado para proteção de dados sensíveis da empresa.
- As interações do Excel são gerenciadas pela `ExcelEngine` com segurança contra fechamento acidental de planilhas de terceiros.
- O automador web `WebAutomator` controla a espera explícita (Explicit Waits) e retentativas exponenciais de injeção de dados.

## Instruções

1. Selecione a **Planilha Base**.
2. Preencha o login (se solicitado pelo fluxo SSO).
3. Determine o período e as categorias que deseja automatizar.
4. Escolha entre execução sequencial ou em paralelo.
5. Inicie e acompanhe o preenchimento pelas barras de progresso.

---

> **Aviso de Privacidade:** Este código foi higienizado. Nomes de equipamentos específicos, servidores, domínios e caminhos absolutos foram anonimizados para publicação em portfólio de acordo com boas práticas de segurança da informação.
