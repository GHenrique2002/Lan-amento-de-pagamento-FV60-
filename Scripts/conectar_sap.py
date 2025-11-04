import win32com.client
import tkinter as tk
import time
import tkinter.messagebox as messagebox

def tentar_conectar_sap(status_label, tentativas=3, intervalo=10):
    """Tenta conectar ao SAP e atualiza o status na interface."""
    session = None
    for tentativa in range(1, tentativas + 1):
        try:
            status_label.config(text=f"Tentativa {tentativa}/{tentativas}: Conectando ao SAP...")
            status_label.update()

            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if not isinstance(sap_gui_auto, win32com.client.CDispatch):
                raise Exception("SAPGUI não disponível")
            application = sap_gui_auto.GetScriptingEngine
            if not isinstance(application, win32com.client.CDispatch):
                raise Exception("Engine de scripting não disponível")
            connection = application.Children(0)
            if not isinstance(connection, win32com.client.CDispatch):
                raise Exception("Conexão SAP não disponível")
            session = connection.Children(0)
            if not isinstance(session, win32com.client.CDispatch):
                raise Exception("Sessão SAP não disponível")

            status_label.config(text="Conexão com SAP estabelecida com sucesso!")
            status_label.update()
            time.sleep(3)
            return session
        except Exception:
            status_label.config(text=f"Falha na tentativa {tentativa}. \nTentando novamente em {intervalo}s...")
            status_label.update()
            time.sleep(intervalo)
    return None

def conectar_sap():

    """Interface principal para conectar ao SAP."""
    while True:
        root = tk.Tk()
        root.title("Status de conexão com o SAP")
        root.geometry("400x200")

        status_label = tk.Label(root, text="Iniciando conexão com SAP...", font=("Arial", 12), fg="blue")
        status_label.pack(pady=40)
        root.update()

        session = tentar_conectar_sap(status_label)

        if session:
            tk.Label(root, text="Conexão estabelecida. Você pode fechar esta janela.", font=("Arial", 10)).pack(pady=10)
            tk.Button(root, text="Fechar", command=root.destroy).pack(pady=20)
            root.mainloop()
            return session
        else:
            root.destroy()  # Fecha a janela antes de abrir o messagebox
            tentar_novamente = messagebox.askyesno("Nenhuma conexão SAP encontrada", "Deseja tentar entrar novamente no SAP?")
            if not tentar_novamente:
                messagebox.showinfo("Encerrado", "Processo de lançamento interrompido.")
                return None
