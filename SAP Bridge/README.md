# SAP Bridge

O SAP Bridge é uma ferramenta de automação multi-thread para o SAP GUI Scripting. Ela extrai dados de transações do SAP (ZCO1, MB51, Produção) e consolida os resultados em uma planilha Excel mestre utilizando as bibliotecas `pandas` e automação COM (`win32com`).

## Recursos

* **Interface Gráfica**: Interface Tkinter em tema escuro (dark-theme) com acompanhamento de progresso.
* **Multithreading**: Extração em paralelo (1 a 4 janelas SAP) para otimizar a extração da transação ZCO1.
* **Notificações**: Notificações nativas do Windows ao concluir ou encontrar erros.
* **Persistência de Configuração**: Salva as últimas configurações do usuário em um arquivo JSON no diretório AppData.
* **Robustez**: Utiliza travas (`sap_lock`) para evitar conflitos na interface do SAP entre múltiplas threads.

## Tecnologias Utilizadas

* Python 3
* SAP GUI Scripting (COM/win32com, pythoncom)
* pandas (processamento e consolidação de dados)
* openpyxl (suporte ao Excel)
* Tkinter e tkcalendar (GUI)
* threading (paralelismo)

## Requisitos

* SAP GUI instalado (C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe)
* Acesso ativo ao ambiente SAP (SAP_CONNECTION_NAME)
* Microsoft Excel instalado

## Como Utilizar

1. Execute o script Python `SAP Bridge (v4.0.0) - codigo higienizado.py`.
2. Selecione a pasta de destino para os arquivos do SAP.
3. Selecione a planilha Excel mestre onde os dados serão consolidados.
4. Escolha o período de extração e o número de threads desejado.
5. Selecione quais relatórios adicionais deseja extrair (MB51, Produção).
6. Clique em **Executar** e aguarde a finalização com acompanhamento na barra de progresso.
