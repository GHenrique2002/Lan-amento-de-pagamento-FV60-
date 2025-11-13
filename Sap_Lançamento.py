import pandas as pd
import os
import logging
from datetime import datetime
import tkinter as tk
import tkinter.messagebox as messagebox
from openpyxl.styles import numbers
from Scripts.conectar_sap import conectar_sap
from Scripts.importar_planilha import importar_planilha_modelo
from Scripts.lancamento import criar_documento

hora_inicio = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
pasta_atual = os.path.dirname(os.path.abspath(__file__)) # Caminho da pasta atual onde está o script
pasta_saida = os.path.join(pasta_atual, "Saída") # Caminho completo da subpasta "Saída"
os.makedirs(pasta_saida, exist_ok=True) # Cria a pasta se não existir

docnums = []

caminho_log = os.path.join(pasta_saida, f"sap_execucao_{hora_inicio}.log")
caminho_planilha = os.path.join(pasta_saida, f"planilha_atualizada_{hora_inicio}.xlsx")

logging.basicConfig (
    filename=caminho_log,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
 
def campo_valido(valor):
    return pd.notna(valor) and str(valor).strip().lower() != "nan" and str(valor).strip() != ""
 
def main():
    """ Função principal que lê o Excel e lança os pagamentos no SAP. """
    logging.info("Iniciando processo de lançamento SAP.")
    logging.info("Tentando conectar ao SAP...")
    session = conectar_sap()    # 1. Conectar à sessão SAP
    if not session:
        logging.error("Falha ao conectar ao SAP. Processo encerrado.")
        return
    logging.info("Conexão SAP estabelecida com sucesso.")
 
     # 2. Importação e leitura dos dados da planilha Excel
   
    logging.info("Aguardando importação da planilha modelo...")
    planilha = importar_planilha_modelo("Planilha_modelo.xlsx")
    if not planilha:
        messagebox.showerror("Nenhuma planilha foi importada")
        logging.warning("Nenhuma planilha foi importada. Processo encerrado.")
        return
    logging.info(f"Planilha importada: {planilha}")
 
    if planilha:
        try:
            df = pd.read_excel(planilha)
            df = df.astype(str) # Converte colunas para string (texto) para evitar problemas de formatação
            nLinhas = len(df)
            logging.info(f"Planilha carregada com {len(df)} linhas.")
        except Exception as e:
            messagebox.showerror("Erro ao ler a planilha", str(e))
            logging.error(f"Erro ao ler a planilha: {e}", exc_info=True)
            return
 
    # 3. Loop através de cada linha da planilha para fazer o lançamento
    print(f"Iniciando lançamentos de {len(df)} pagamentos...")
    logging.info("Iniciando lançamentos...")
    for index, row in df.iterrows(): # Para percorrer cada linha uma por uma de cima para baixo
        try:
            logging.info(f"Iniciando lançamento para fornecedor: {row['Fornecedor']}")
            print("-" * 30)
            print(f"Lançando pagamento para Fornecedor: {row['Fornecedor']}")
            data_fatura = pd.to_datetime(row['Data fatura'], errors='coerce') #garantir que a data está preenchida com ponto
            data_pagamento = pd.to_datetime(row['Data pgto.'], errors='coerce') #garantir que a data está preenchida com ponto
 
            session.findById("wnd[0]").maximize()
            docnum = criar_documento(row, session)        
            docnums.append(docnum)
            print(docnum)
            print(docnums)

            if docnum == "Erro ao processar lançamento. Verifique e tente novamente.":
                logging.error(f"Erro ao lançar para fornecedor: {row['Fornecedor']}")
            else:
                logging.info(f"Lançamento concluído para fornecedor: {row['Fornecedor']}")

        except Exception as e:
           print(f"Erro sistêmico. Reinicie o processo!")
           continue # Pula para a próxima linha da planilha

    print("-" * 30)
    print("Processo de lançamento em massa finalizado!")
    logging.info("Processo de lançamento finalizado.")
    while len(docnums) < len(df):
        docnums.append("Erro ou não processado")
    df["Docnum"] = docnums
    df.to_excel(caminho_planilha, index=False)
    logging.info("Planilha salva com sucesso!")
    messagebox.showinfo("Sucesso","Processo de lançamento finalizado.")
main()