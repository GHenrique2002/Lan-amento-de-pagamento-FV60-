import tkinter as tk
import tkinter.messagebox as messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import zipfile
import os
import tempfile

def importar_planilha_modelo(nome_planilha_modelo="Planilha_modelo.xlsx"):

    """
    Abre uma janela para o usuário arrastar e soltar uma planilha ou um .zip.
    Se for .zip, verifica se há uma planilha com nome "Planilha_modelo.xlsx" dentro.
    Permite tentativas até que a planilha correta seja enviada ou o usuário desista.
    Retorna o caminho da planilha encontrada ou None.
    """
    while True:
        planilha = None

        def drop(event):
            nonlocal planilha
            caminho = event.data.strip('{}')

            if caminho.lower().endswith('.zip'):
                temp_dir = tempfile.mkdtemp()
                with zipfile.ZipFile(caminho, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                arquivos_extraidos = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir)]

                for f in arquivos_extraidos:
                    if os.path.basename(f).lower() == nome_planilha_modelo.lower():
                        planilha = f

                if not planilha:
                    messagebox.showerror("Erro", f"A planilha '{nome_planilha_modelo}' não foi encontrada no ZIP.")
            elif os.path.basename(caminho).lower() == nome_planilha_modelo.lower():
                planilha = caminho
            else:
                messagebox.showerror("Erro", f"Arquivo inválido. Envie a planilha '{nome_planilha_modelo}' ou um ZIP contendo ela.")

            root.destroy()

        root = TkinterDnD.Tk()
        root.title("Arraste a planilha ou ZIP aqui")
        root.geometry("400x200")
        label = tk.Label(root, text=f"Arraste aqui \na planilha '{nome_planilha_modelo}' \nou ZIP contendo ela", width=40, height=10, bg="lightgray")
        label.pack(pady=40)
        label.drop_target_register(DND_FILES)
        label.dnd_bind('<<Drop>>', drop)
        root.mainloop()

        if planilha:
            messagebox.showinfo("Sucesso", f"Arquivos e '{nome_planilha_modelo}' recebidos com sucesso.")
            return planilha
        else:
            tentar_novamente = messagebox.askyesno("Nenhuma planilha recebida", "Deseja tentar enviar novamente?")
            if not tentar_novamente:
                messagebox.showinfo("Encerrado", "Processo de lançamento interrompido.")
                return None
