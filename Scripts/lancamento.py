import re
import time
import pandas as pd

def formatar_cnpj(cnpj):
    """Remove caracteres especiais e garante 14 dígitos."""
    cnpj_limpo = re.sub(r'\D', '', str(cnpj))
    return cnpj_limpo.zfill(14)

def criar_documento(row, session):
    try:
        # Abrir FV60
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFV60"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        print("Transação FV60 aberta.")

        # Empresa
        try:
            session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = row["Empresa"]
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        except:
            session.findById("wnd[0]").sendVKey (7)
            session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = row["Empresa"]
            session.findById("wnd[1]").sendVKey(0)
        
        if pd.notna(row['Data fatura']):
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BLDAT").text = pd.to_datetime(row['Data fatura']).strftime("%d.%m.%Y")
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)

        
        # Abrir pesquisa do fornecedor
        session.findById("wnd[0]").sendVKey (4)
        time.sleep(1)

        # Selecionar aba de busca por CNPJ
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()

        # Inserir CNPJ formatado
        cnpj_formatado = formatar_cnpj(row['CNPJ'])
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = cnpj_formatado
        session.findById("wnd[1]").sendVKey (0)

        # Executar pesquisa e selecionar fornecedor
        session.findById("wnd[1]").sendVKey (0)
        time.sleep(1)
        session.findById("wnd[0]").sendVKey (0)


        # Preencher campos obrigatórios


        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-XBLNR").text = str(row['Referência'])

        valor = str(row['Montante']).replace("R$", "").strip()
        valor_corrigido = float(valor)
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-WRBTR").text = valor_corrigido

        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").text = str(row['Texto descritivo'])

        # Aba Itens
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").text = str(row['Conta razão'])
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").text = valor_corrigido
        time.sleep(1)

        try:
            if (row['Centro Custo']) !="nan":
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[17,0]").text = str(row['Centro Custo'])
        except Exception as e:
            print("Erro ao preencher Centro Custo:", e)

        try:
            if (row['Ordem']) != "nan":
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-AUFNR[18,0]").text = str(row['Ordem'])
        except Exception as e:
            print("Erro ao preencher Ordem:", e)

        try:
            if (row['Elemento PEP']) != "nan":
                session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,0]").text = str(row['Elemento PEP'])
        except Exception as e:
            print("Erro ao preencher Elemento PEP:", e)

        print(row['Data pgto.'])

        try:
            if (row['Data pgto.']) != "nan":
                session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BUDAT").text = pd.to_datetime(row['Data pgto.']).strftime("%d.%m.%Y")
                time.sleep(1)
        except: 
            print(session.findById("wnd[0]/usr/tabsTS/tabpPAYM").text)
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]").sendVKey (0)

        # Aba Pagamento
        session.findById("wnd[0]/usr/tabsTS/tabpPAYM").select()
        session.findById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZFBDT").text = pd.to_datetime(row['Data pgto.']).strftime("%d.%m.%Y")
        session.findById("wnd[0]").sendVKey (9)

        # Salvar documento
        session.findById("wnd[0]/tbar[1]/btn[42]").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()

        # Capturar número do documento
        docnum = session.findById("wnd[0]/sbar").text
        print(docnum)
        return docnum
        
    except Exception as e:
        mensagem_erro = "Erro ao processar lançamento. Verifique e tente novamente."
        return mensagem_erro

