import pandas as pd
import win32com.client
import time
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD

def importar_arquivo():
    global planilha_modelo
    planilha = None
    def drop(event):
        nonlocal planilha
        planilha = event.data.strip('{}')
        root.destroy()
    root = TkinterDnD.Tk()
    root.title("Arraste a planilha modelo aqui")
    root.geometry("400x200")
    label = tk.Label(root, text="Arraste a planilha modelo aqui", width=40, height=10, bg="lightgray")
    label.pack(pady=40)
    label.drop_target_register(DND_FILES)
    label.dnd_bind('<<Drop>>', drop)
    root.mainloop()
    planilha_modelo = planilha
    print("Arquivo recebido com sucesso:")
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
   
def main():
   """ Função principal que lê o Excel e lança os pagamentos no SAP. """
   importar_arquivo()
   session = conectar_sap()    # 1. Conectar à sessão SAP
   if not session:
       print("Não foi possível encontrar uma sessão SAP ativa. Verifique se o SAP Logon está aberto. ")
       return
   # 2. Ler os dados da planilha Excel
   try:
       df = pd.read_excel(planilha_modelo)
       df = df.astype(str) # Converte colunas para string (texto) para evitar problemas de formatação
   except FileNotFoundError:
       print("Erro: Planilha 'nome do arquivo em excel' não encontrada. Verifique o nome e o local do arquivo.")
       return
   # 3. Loop através de cada linha da planilha para fazer o lançamento
   print(f"Iniciando lançamentos de {len(df)} pagamentos...")
   for index, row in df.iterrows(): # Para percorrer cada linha uma por uma de cima para baixo
       try:
           print("-" * 30)
           print(f"Lançando pagamento para Fornecedor: {row['Fornecedor']}")
           # Inicia a transação (ex: FB60). Adapte para a sua transação!
           session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB60"
           session.findById("wnd[0]").sendVKey(0) # Pressiona Enter
           # --- PREENCHIMENTO DOS DADOS ---
           # Use os IDs que você gravou no Passo 2. Este é apenas um EXEMPLO.
           # Cabeçalho
           session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR01" # Exemplo: Empresa
           session.findById("wnd[0]/usr/subSUB_HEADER:SAPLFDCB:0051/ctxtINVFO-ACCNT").text = row['Fornecedor']
           session.findById("wnd[0]/usr/txtINVFO-BLDAT").text = row['DataDocumento']
           session.findById("wnd[0]/usr/txtINVFO-XBLNR").text = row['NotaFiscal']
           session.findById("wnd[0]/usr/txtINVFO-WRBTR").text = row['ValorBruto']
           session.findById("wnd[0]/usr/subSUB_HEADER:SAPLFDCB:0051/txtINVFO-SGTXT").text = row['TextoHeader']
           # Item (pode precisar de mais lógica se tiver vários itens)
           # Exemplo: preenchendo a primeira linha do grid de itens
           session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabpTAB_ITENS/ssubSUB_TAB_ITENS:SAPLFSKB:0100/tblSAPLFSKBTABLE_CONTROL/ctxtACGL_ITEM-HKONT[0,0]").text = "40010020" # Exemplo: Conta Contábil
           session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabpTAB_ITENS/ssubSUB_TAB_ITENS:SAPLFSKB:0100/tblSAPLFSKBTABLE_CONTROL/txtACGL_ITEM-WRBTR[1,0]").text = row['ValorBruto']
           # Simular ou Salvar
           # Para simular, pressione o botão de simular
           # session.findById("wnd[0]/tbar[1]/btn[21]").press()
           # Para salvar/lançar diretamente
           session.findById("wnd[0]/tbar[0]/btn[11]").press() # Botão Salvar/Lançar
           # Capturar mensagem da barra de status
           status_message = session.findById("wnd[0]/sbar").text
           print(f"Status: {status_message}")
           # Esperar um pouco para a próxima transação
           time.sleep(1)
       except Exception as e:
           print(f"ERRO ao lançar para o fornecedor {row['Fornecedor']}: {e}")
           continue # Pula para a próxima linha da planilha
   print("-" * 30)
   print("Processo de lançamento em massa finalizado!")

main()