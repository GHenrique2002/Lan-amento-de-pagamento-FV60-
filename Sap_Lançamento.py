import pandas as pd
import win32com.client
import subprocess
import time
import tkinter as tk
from openpyxl.styles import numbers
from tkinterdnd2 import DND_FILES, TkinterDnD

def importar_arquivo():
    global planilha_modelo
    planilha = None
    def drop(event):
        nonlocal planilha
        planilha = event.data.strip('{}')
        root.destroy()
    root = TkinterDnD.Tk()
    root.title("Arraste a planilha aqui")
    root.geometry("400x200")
    label = tk.Label(root, text="Arraste a planilha aqui", width=40, height=10, bg="lightgray")
    label.pack(pady=40)
    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', drop)
    root.mainloop()
    planilha_modelo = planilha
    print("Planilha recebida com sucesso! ")
    return planilha

def conectar_sap():
   """ Função para se conectar a uma sessão SAP já aberta. """
   try:
       # Tenta se conectar à aplicação SAP GUI
       sap_gui_auto = win32com.client.GetObject("SAPGUI")
       if not isinstance(sap_gui_auto, win32com.client.CDispatch):
           return None
       application = sap_gui_auto.GetScriptingEngine
       if not isinstance(application, win32com.client.CDispatch): # O SAP pode ter múltiplas conexões abertas, vamos pegar a primeira
           return None
       connection = application.Children(0)
       if not isinstance(connection, win32com.client.CDispatch): # O SAP pode ter múltiplas sessões (janelas) abertas, vamos pegar a primeira
           return None
       session = connection.Children(0)
       if not isinstance(session, win32com.client.CDispatch):
           return None
       print(" Conexão com SAP estabelecida com sucesso! ")
       return session
   except Exception as e:
       print(f"Erro ao conectar com o SAP: {e}")
       return None 

def campo_valido(valor):
    return pd.notna(valor) and str(valor).strip().lower() != "nan" and str(valor).strip() != ""

def main():
   importar_arquivo() #capturar planilha
   """ Função principal que lê o Excel e lança os pagamentos no SAP. """
   session = conectar_sap()    # 1. Conectar à sessão SAP
   if not session:
       print("Não foi possível encontrar uma sessão SAP ativa. Verifique se o SAP Logon está aberto. ")
       return
   # 2. Ler os dados da planilha Excel
   try:
       df = pd.read_excel(planilha_modelo)
       df = df.astype(str) # Converte colunas para string (texto) para evitar problemas de formatação
   except FileNotFoundError:
       print("Erro: Planilha não encontrada. Verifique e envie novamente.")
       return
   # 3. Loop através de cada linha da planilha para fazer o lançamento
   print(f"Iniciando lançamentos de {len(df)} pagamentos...")
   for index, row in df.iterrows(): # Para percorrer cada linha uma por uma de cima para baixo
       try:
            print("-" * 30)
            print(f"Lançando pagamento para Fornecedor: {row['Fornecedor']}")
            data_fatura = pd.to_datetime(row['Data da fatura'], errors='coerce') #garantir que a data ta preenchida com ponto
            data_pagamento = pd.to_datetime(row['Data lançamento(data pagamento)'], errors='coerce')#garantir que a data ta preenchida com ponto
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nfv60"   # Inicia a transação (ex: FV60).
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]").sendVKey (7)
            session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = row["Empresa"]
            session.findById("wnd[1]").sendVKey (0)  
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BLDAT").text = data_fatura.strftime('%d.%m.%Y')
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-ACCNT").text = row['Fornecedor']
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-XBLNR").text = row['Referência']
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BUDAT").text = data_pagamento.strftime('%d.%m.%Y')
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-WRBTR").text = row['Montante']
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").text = row["Texto descritivo"]
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").text = row["Conta razão"]
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").text = row["Montante"]
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-SGTXT[11,0]").text = row["Texto descritivo"]
            if campo_valido(row['Centro de Custo']):    
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[17,0]").text = row['Centro de Custo']
            elif campo_valido(row['Ordem']):
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-AUFNR[18,0]").text = row['Ordem']
            elif campo_valido(row['Elemento PEP']):
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,0]").text = row['Elemento PEP']
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,0]").setFocus
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,0]").caretPosition = 9
            time.sleep(1)             # Esperar um pouco para lançar novamente
       except Exception as e:
           print(f"ERRO ao lançar para o fornecedor {row['Fornecedor']}: {e}")
           continue # Pula para a próxima linha da planilha
   print("-" * 30)
   print("Processo de lançamento em massa finalizado!")

main()