# Manual de Uso - Automação de Lançamentos SAP

## 1. Instalação
Instale as dependências com:

```bash
pip install -r requirements.txt 
```

## 2. Estrutura do Projeto
- Sap_Lançamento.py: Script principal que coordena o processo de automação.
- conectar_sap.py: Responsável por conectar ao SAP via SAP GUI Scripting.
- importar_planilha.py: Interface gráfica para importar a planilha modelo.
- lancamento.py: Contém funções que realizam o lançamento no SAP e formatam dados como CNPJ.

## 3. Execução
- Execute Sap_Lançamento.py.
- Uma janela será aberta para envio da planilha Planilha_modelo.xlsx ou .zip contendo ela.
- O sistema tentará conectar ao SAP. Se falhar, será oferecida a opção de tentar novamente.
- Após conexão, os dados da planilha serão lidos e os lançamentos realizados via transação FV60.
- Ao final, será gerada a planilha "planilha_atualizada.xlsx" com os números de documentos SAP.

## 4. Detalhes Técnicos
A função criar_documento(row, session) realiza o lançamento no SAP preenchendo campos como:

- Empresa (obrigatório)
- Data da fatura (obrigatório)
- CNPJ do fornecedor (obrigatório)
- Referência (obrigatório)
- Montante (obrigatório)
- Texto descritivo (obrigatório)
- Conta razão (obrigatório)
- Centro de custo, Ordem, Elemento PEP (se aplicáveis)
- Data de pagamento (obrigatório)
- *O número do documento SAP é capturado e armazenado.*

## 5. Logs
- Um arquivo sap_execucao.log registra todos os eventos e erros.

## 6. Tratamento de Erros
- Erros de leitura da planilha ou falhas no SAP são tratados com mensagens na interface.
- O sistema continua o processamento mesmo após falhas em lançamentos individuais.

## 7. Observações
- O SAP GUI deve estar aberto e com sessão ativa.
- O SAP Scripting precisa estar habilitado nas configurações do SAP GUI.
- O nome da planilha deve ser exatamente "Planilha_modelo.xlsx".
