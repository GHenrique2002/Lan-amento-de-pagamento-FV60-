# Automação de Lançamentos SAP via Excel
Este projeto automatiza o processo de lançamentos de pagamentos no SAP utilizando uma planilha Excel como entrada. A automação é feita via SAP GUI Scripting, com interface gráfica para facilitar o uso por usuários administrativos.

## Funcionalidades
- Conexão automática com o SAP.
- Interface gráfica para envio da planilha modelo.
- Suporte a arquivos `.xlsx` e `.zip`.
- Lançamento automatizado de pagamentos via transação FV60.
- Geração de log de execução.
- Exportação de planilha atualizada com número de documento SAP.

## Requisitos
- Windows com SAP GUI instalado e SAP Scripting habilitado.
- Python 3.8 ou superior.
- Bibliotecas listadas em `requirements.txt`.

## Execução
- Execute o script principal `Sap_Lançamento.py` e siga as instruções na interface gráfica.
- Em caso de dúvidas, consultar Documentação `Manual_fv60.md`