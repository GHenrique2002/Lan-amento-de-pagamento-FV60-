# Automação de Lançamentos SAP (FV60) com Python

![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

Um script em Python para automatizar o lançamento de faturas de fornecedores (transação FV60) no SAP, a partir de uma planilha Excel padronizada.

## 🚀 Visão Geral

O lançamento manual de pagamentos e faturas na transação FV60 do SAP é um processo repetitivo, demorado e sujeito a erros de digitação, especialmente quando há um grande volume de documentos.

Este projeto oferece uma solução para automatizar essa tarefa. Utilizando um script em Python, ele lê os dados de uma planilha Excel, abre o SAP GUI e simula a entrada de dados para cada lançamento. O objetivo principal é otimizar o tempo das equipes financeiras e fiscais, aumentar a produtividade e reduzir erros.

## ✨ Funcionalidades

-   **Leitura de Dados via Excel**: Processa lançamentos a partir de um arquivo `.xlsx` estruturado.
-   **Interface Simples**: Pede ao usuário para apenas arrastar e soltar (drag-and-drop) o arquivo a ser processado.
-   **Automação da Digitação**: Preenche automaticamente os campos necessários na transação FV60.
-   **Otimização de Tempo**: Reduz drasticamente o tempo gasto em tarefas manuais e repetitivas.
-   **Redução de Erros**: Minimiza falhas humanas de digitação.

## 🔧 Como Funciona

1.  **Preparação**: O usuário preenche a planilha Excel modelo com os dados dos pagamentos (fornecedor, valor, datas, etc.).
2.  **Execução**: O script Python é executado.
3.  **Seleção do Arquivo**: Uma janela surge, solicitando que o usuário arraste o arquivo Excel preenchido.
4.  **Automação**: O script abre o SAP, acessa a transação FV60 e começa a lançar cada linha da planilha como um novo documento.
5.  **Finalização**: Ao final do processo, os lançamentos estarão pré-lançados no SAP, prontos para verificação e contabilização.

## 📄 Uso

1.  Certifique-se de que o **SAP Logon esteja aberto**.
2.  Preencha a planilha com os dados corretos.
3.  Execute o script principal pelo terminal:
    ```bash
    python main.py
    ```
4.  Arraste o arquivo Excel para a janela que aparecerá na tela.
5.  Aguarde enquanto o script realiza os lançamentos no SAP.

### A Planilha Modelo

O script depende de uma planilha Excel com uma estrutura específica. As colunas devem seguir o padrão para que a leitura dos dados funcione corretamente.


**Importante**: Os nomes das colunas na sua planilha devem corresponder exatamente aos que o script espera.

## 📜 Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
