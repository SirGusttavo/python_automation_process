# VISCONF Engine

VISCONF Engine é um motor de gerenciamento de cadeia de suprimentos e estoques para planejamento de produção. A aplicação lê os planos de produção do Excel, calcula as necessidades de estoque através da explosão da lista de materiais (BOM), determina os requisitos de pallets e gera relatórios de reservas.

## Tecnologias Utilizadas

- **Python**
- **xlwings**: para automação do Excel via COM.
- **openpyxl**: para geração de relatórios estruturados e formatados.
- **SAP GUI Scripting (COM)**: para extração automatizada de dados diretamente do SAP (ex: MB51 para produção, LX02 para estoques).
- **Tkinter**: para a interface gráfica do usuário.
- **Threading**: para processamento em background sem travar a interface.

## Funcionalidades Principais

- **Pipeline Completo**: Interface gráfica (GUI) simples e configurável em múltiplos estágios.
- **Integração SAP**: Automação via SAP GUI Scripting para extração direta e em tempo real.
- **Cálculo de Necessidades**: Explosão de BOM para determinar exatamente as necessidades de insumos e alocação de pallets.
- **Relatórios**: Geração automatizada de relatórios em Excel, com formatação avançada via `openpyxl`.

*Nota: Este projeto foi higienizado e anonimizado para fins de publicação em portfólio. Dados sensíveis, nomes de máquinas, endereços de servidores e conexões SAP foram substituídos por valores genéricos.*
