import os
import logging
from datetime import datetime
import tkinter as tk
import tkinter.messagebox as messagebox
import re
from tkinterdnd2 import DND_FILES, TkinterDnD
import zipfile
import tempfile
import win32com.client
import time
import pandas as pd

# ==========================================
# CONFIGURAÇÕES & CONSTANTES DO SAP
# ==========================================
DEFAULT_ZTERM = "0001"          # Condição de pagamento padrão
SAP_TRANSACTION = "/nFV60"       # Transação para lançamento prévio de faturas

# Diretório de Saída e Configuração de Logs
hora_inicio = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
pasta_atual = os.path.dirname(os.path.abspath(__file__))
pasta_saida = os.path.join(pasta_atual, "Saída")
os.makedirs(pasta_saida, exist_ok=True)

caminho_log = os.path.join(pasta_saida, f"sap_execucao_{hora_inicio}.log")
caminho_planilha = os.path.join(pasta_saida, f"planilha_atualizada_{hora_inicio}.xlsx")

logging.basicConfig(
    filename=caminho_log,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

docnums = []


def campo_valido(valor):
    """Verifica se o valor do campo não é nulo/vazio/nan."""
    return pd.notna(valor) and str(valor).strip().lower() != "nan" and str(valor).strip() != ""


def formatar_cnpj(cnpj):
    """Remove caracteres especiais e garante 14 dígitos."""
    cnpj_limpo = re.sub(r'\D', '', str(cnpj))
    return cnpj_limpo.zfill(14)


def format_sap_value(valor):
    """Formata valores numéricos para o padrão visual do SAP (duas casas decimais e vírgula)."""
    return f"{valor:.2f}".replace(".", ",")


def to_float(valor):
    """Converte com segurança valores para float."""
    try:
        if pd.notna(valor) and str(valor).lower() != "nan":
            return float(str(valor).replace("R$", "").strip())
        return 0.0
    except Exception:
        return 0.0


def preencher_imputacao_custo(session, objeto_custo):
    """
    Realiza o preenchimento dos objetos de custo utilizando uma lógica
    de fallback sequencial (Centro de Custo -> Ordem -> Elemento PEP -> Diagrama/Operação).
    """
    if not objeto_custo or str(objeto_custo).lower() == "nan":
        return

    val_custo = str(objeto_custo)

    try:
        # 1. Tenta Centro de Custo (KOSTL)
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[17,0]").text = val_custo
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)

        sbar_text = session.findById("wnd[0]/sbar").text.lower() if session.findById("wnd[0]/sbar").text else ""

        if "não" in sbar_text and "existe" in sbar_text:
            # 2. Tenta Ordem Interna (AUFNR)
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[17,0]").text = ""
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-AUFNR[18,0]").text = val_custo
            session.findById("wnd[0]").sendVKey(0)

            sbar_text_2 = session.findById("wnd[0]/sbar").text.lower() if session.findById("wnd[0]/sbar").text else ""

            if "não" in sbar_text_2 and "existe" in sbar_text_2:
                # 3. Tenta Elemento PEP (PROJK)
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-AUFNR[18,0]").text = ""
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,0]").text = val_custo
                session.findById("wnd[0]").sendVKey(0)

                sbar_text_3 = session.findById("wnd[0]/sbar").text.lower() if session.findById("wnd[0]/sbar").text else ""

                if "não" in sbar_text_3 and "existe" in sbar_text_3:
                    # 4. Tenta Diagrama de Rede + Operação (NPLNR + VORNR)
                    session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,0]").text = ""
                    session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-NPLNR[37,0]").text = val_custo[:-4]
                    session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-VORNR[38,0]").text = val_custo[-4:]
                    session.findById("wnd[0]").sendVKey(0)
    except Exception as e:
        print(f"Alerta na atribuição de custo: {e}")


def criar_documento(row, session):
    """Executa a rotina de preenchimento dos campos na transação SAP FV60."""
    try:
        valor_real = to_float(row['VALOR'])
        valor_correto = format_sap_value(valor_real)

        # Acessar a transação parametrizada
        session.findById("wnd[0]/tbar[0]/okcd").text = SAP_TRANSACTION
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        cnpj = formatar_cnpj(str(row["NÚMERO RECEBEDOR"]))

        # Preenchimento do Código da Empresa e Tipo de Documento
        try:
            session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = row["EMPRESA"]
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/cmbINVFO-BLART").key = row["TIPO DOC"]
        except Exception:
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/cmbINVFO-BLART").key = row["TIPO DOC"]
            session.findById("wnd[0]").sendVKey(7)
            session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = row["EMPRESA"]
            session.findById("wnd[1]").sendVKey(0)

            data_fatura = session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BUDAT").text
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BLDAT").text = data_fatura
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)

        # Preenchimento do Fornecedor/Recebedor
        try:
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-ACCNT").text = str(row['NÚMERO RECEBEDOR'])
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").text = str(row['CLASSE DE CUSTOS'])
            time.sleep(0.5)
        except Exception:
            # Fallback de busca por CNPJ
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(4)
            time.sleep(0.5)
            session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
            session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = cnpj
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(0.5)
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").text = str(row['CLASSE DE CUSTOS'])
            time.sleep(0.5)

        # Campos de Cabeçalho e Item
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-XBLNR").text = str(row['REFERENCIA'])
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-WRBTR").text = valor_correto
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").text = str(row['DESCRIÇÃO DOS GASTOS'])

        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").text = valor_correto
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-SGTXT[11,0]").text = str(row['DESCRIÇÃO DOS GASTOS'])

        # Preenchimento de Objetos de Custo
        preencher_imputacao_custo(session, row.get('CC / ODI / DIAGRAMA'))

        # Data do Lançamento
        tipo_doc = str(row['TIPO DOC'])
        if tipo_doc not in ["KB", "KL"] and campo_valido(row.get('DATA')):
            try:
                session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BUDAT").text = pd.to_datetime(row['DATA']).strftime("%d.%m.%Y")
                time.sleep(1)
            except Exception as e:
                print(f"Aviso ao definir data da fatura: {e}")

        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)

        # Aba de Pagamento
        session.findById("wnd[0]/usr/tabsTS/tabpPAYM").select()
        session.findById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZTERM").text = DEFAULT_ZTERM
        session.findById("wnd[0]").sendVKey(0)

        if tipo_doc not in ["KB", "KL"] and campo_valido(row.get('DATA')):
            session.findById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZFBDT").text = pd.to_datetime(row['DATA']).strftime("%d.%m.%Y")

        if campo_valido(row.get("FORMA PAGAMENTO")):
            session.findById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZLSCH").text = str(row["FORMA PAGAMENTO"])

        session.findById("wnd[0]").sendVKey(9)

        # Salvar / Gravar documento
        session.findById("wnd[0]/tbar[1]/btn[42]").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()

        docnum = session.findById("wnd[0]/sbar").text
        return docnum

    except Exception as e:
        erro_sbar = session.findById("wnd[0]/sbar").text if hasattr(session, "findById") else str(e)
        return f"Erro: {erro_sbar}"


def importar_planilha_modelo():
    """Interface gráfica (Drag & Drop) para importar planilha de dados (.xlsx ou .zip)."""
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
                    if f.lower().endswith('.xlsx'):
                        planilha = f
                        break

                if not planilha:
                    messagebox.showerror("Erro", "Nenhuma planilha .xlsx encontrada no arquivo ZIP.")
            elif caminho.lower().endswith('.xlsx'):
                planilha = caminho
            else:
                messagebox.showerror("Erro", "Arquivo inválido. Envie uma planilha .xlsx ou um arquivo .zip.")

            root.destroy()

        root = TkinterDnD.Tk()
        root.title("Importar Planilha de Lançamentos")
        root.geometry("420x220")
        label = tk.Label(root, text="Arraste aqui a planilha (.xlsx) ou arquivo ZIP contendo a planilha", width=42, height=10, bg="#f0f0f0")
        label.pack(pady=30)
        label.drop_target_register(DND_FILES)
        label.dnd_bind('<<Drop>>', drop)
        root.mainloop()

        if planilha:
            messagebox.showinfo("Sucesso", f"Arquivo selecionado: {os.path.basename(planilha)}")
            return planilha
        else:
            tentar_novamente = messagebox.askyesno("Nenhum arquivo recebido", "Deseja tentar importar novamente?")
            if not tentar_novamente:
                messagebox.showinfo("Encerrado", "Operação cancelada pelo usuário.")
                return None


def tentar_conectar_sap(status_label, tentativas=3, intervalo=5):
    """Realiza tentativas de conexão com a instância ativa do SAP GUI Scripting."""
    for tentativa in range(1, tentativas + 1):
        try:
            status_label.config(text=f"Tentativa {tentativa}/{tentativas}: Conectando ao SAP...")
            status_label.update()

            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if not isinstance(sap_gui_auto, win32com.client.CDispatch):
                raise Exception("SAPGUI não disponível.")
            application = sap_gui_auto.GetScriptingEngine
            if not isinstance(application, win32com.client.CDispatch):
                raise Exception("Engine de scripting não disponível.")
            connection = application.Children(0)
            if not isinstance(connection, win32com.client.CDispatch):
                raise Exception("Conexão SAP não disponível.")
            session = connection.Children(0)
            if not isinstance(session, win32com.client.CDispatch):
                raise Exception("Sessão SAP não disponível.")

            status_label.config(text="Conexão com o SAP estabelecida com sucesso!")
            status_label.update()
            time.sleep(2)
            return session
        except Exception:
            status_label.config(text=f"Tentativa {tentativa} falhou. Aguardando {intervalo}s...")
            status_label.update()
            time.sleep(intervalo)
    return None


def conectar_sap():
    """Abre janela de status para efetuar a conexão com o SAP."""
    while True:
        root = tk.Tk()
        root.title("Status da Conexão SAP")
        root.geometry("380x180")

        status_label = tk.Label(root, text="Iniciando conexão com o SAP...", font=("Arial", 11), fg="#0056b3")
        status_label.pack(pady=30)
        root.update()

        session = tentar_conectar_sap(status_label)

        if session:
            tk.Label(root, text="Sessão vinculada. Você pode fechar esta janela.", font=("Arial", 9)).pack(pady=5)
            tk.Button(root, text="Continuar", command=root.destroy, width=15).pack(pady=15)
            root.mainloop()
            return session
        else:
            root.destroy()
            tentar_novamente = messagebox.askyesno("Falha na Conexão", "Não foi possível conectar ao SAP. Deseja tentar novamente?")
            if not tentar_novamente:
                messagebox.showinfo("Encerrado", "Lançamento interrompido.")
                return None


def main():
    """Fluxo principal do programa."""
    logging.info("Iniciando processo de lançamento automatizado no SAP.")
    
    session = conectar_sap()
    if not session:
        logging.error("Falha ao conectar ao SAP. Encerrando execução.")
        return

    logging.info("Conexão SAP estabelecida com sucesso.")

    planilha = importar_planilha_modelo()
    if not planilha:
        logging.warning("Nenhuma planilha selecionada. Encerrando execução.")
        return

    logging.info(f"Planilha importada com sucesso: {planilha}")

    try:
        df = pd.read_excel(planilha, header=3)
        df = df.astype(str)
        df.columns = df.columns.str.strip().str.upper()
        logging.info(f"Planilha carregada contendo {len(df)} registros.")
    except Exception as e:
        messagebox.showerror("Erro de Leitura", f"Erro ao ler a planilha: {e}")
        logging.error(f"Erro ao ler a planilha: {e}", exc_info=True)
        return

    print(f"Iniciando lançamentos de {len(df)} registros no SAP...")
    logging.info("Iniciando lote de lançamentos...")

    for index, row in df.iterrows():
        try:
            recebedor_anonimizado = f"***{str(row['NÚMERO RECEBEDOR'])[-4:]}" if len(str(row['NÚMERO RECEBEDOR'])) > 4 else "FORNECEDOR"
            logging.info(f"Iniciando lançamento para recebedor ID: {recebedor_anonimizado}")
            print("-" * 30)
            print(f"Processando linha {index + 1}/{len(df)}")

            session.findById("wnd[0]").maximize()
            docnum = criar_documento(row, session)
            docnums.append(docnum)

            if str(docnum).startswith("Erro"):
                logging.error(f"Falha ao lançar registro {index + 1}: {docnum}")
            else:
                logging.info(f"Lançamento concluído com sucesso. Retorno: {docnum}")

        except Exception as e:
            print(f"Erro ao processar linha {index + 1}: {e}")
            docnums.append("Erro na execução")
            continue

    print("-" * 30)
    print("Processo finalizado com sucesso!")
    logging.info("Lote de lançamentos concluído.")

    while len(docnums) < len(df):
        docnums.append("Não processado")

    df.replace("nan", "", inplace=True)
    df.fillna("", inplace=True)
    df["DOCNUM_RETORNO"] = docnums
    
    df.to_excel(caminho_planilha, index=False)
    logging.info(f"Relatório final salvo em: {caminho_planilha}")
    messagebox.showinfo("Sucesso", "Processo de lançamento concluído e relatório atualizado.")


if __name__ == "__main__":
    main()
