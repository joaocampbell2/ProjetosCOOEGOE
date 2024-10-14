import time
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from openpyxl import load_workbook
from marinetteSEFAZ import loginSEI, obterProcessosDeBloco,procurarArquivos, buscarInformacaoEmDocumento, escreverAnotacao, salvarPlanilha
    
def encontrarFormaDePagamento():
    despachos = procurarArquivos(navegador,"Despacho sobre Autorização de Despesa")
    regexBeneficiario = r"Beneficiário:(.*)\n"
    regexFormaDePagamento = r"Forma de Pagamento:(.*)\n"

    for despacho in reversed(despachos):
        try: 
            beneficiario = buscarInformacaoEmDocumento(navegador,despacho,regexBeneficiario,"Rio de Janeiro").group(1).upper()
            formaPagamentoDespacho = buscarInformacaoEmDocumento(navegador,despacho,regexFormaDePagamento,"Rio de Janeiro").group(1).upper()

        
            if "PERDIMENTO" in formaPagamentoDespacho:
                return "PERDIMENTO"

            if "BRADESCO" in formaPagamentoDespacho:
                return "Depósito Bradesco"
            
            if "CNPJ" in beneficiario:                      
                if any(palavra in formaPagamentoDespacho for palavra in ["GUIA", "GRERJ"]):
                    return "Guia"
                if any(palavra in formaPagamentoDespacho for palavra in ["GRU", "FUNAD"]):
                    return "Guia GRU"

            if "DEPÓSITO JUDICIAL" in formaPagamentoDespacho:
                return "Guia"

            regexBanco = r"(ITAÚ|NUBANK|SANTANDER|Santander|C6 Bank|BANCO DO BRASIL|SICOOB|C6 BANK|PICPAY|CAIXA|CEF|SICRED|BANCO C6)"
            banco = buscarInformacaoEmDocumento(navegador,despacho,regexBanco,"Rio de Janeiro")
            if banco == None:
                banco = ""
            else:
                banco = banco.group(1)
            
    

            if "DEPÓSITO" in formaPagamentoDespacho:
                return "Depósito " + banco
            if  "CPF" in beneficiario:
                return "Depósito " + banco
            if "AGÊNCIA" in formaPagamentoDespacho:
                return "Depósito " + banco
        except:
            traceback.print_exc()
            pass
    
    try:
        indebitos = procurarArquivos(navegador, "Correspondência Interna")
        formaDePagamento = buscarInformacaoEmDocumento(navegador,indebitos[0],r"(INDÉBITO)").group(1)
        if formaDePagamento:
            return "Indébito"
    except:
        pass
    
    print("Não encontrado")
    return "ERRO"




def encontrarValidade():
    
    regexBanco = r"(BANCO DO BRASIL|GRERJ)"
    validade = "SEM VALIDADE"
    
    if tipo == "FIANÇA":
        docs = procurarArquivos(navegador, ["Guia", "GRERJ"])
        for doc in reversed(docs):
            banco = buscarInformacaoEmDocumento(navegador,doc,regexBanco)
            if banco != None:
                if banco.group(1) == "BANCO DO BRASIL":
                    regexValidade = r"(\d{2}\/\d{2}\/\d{4})\n"
                    
                elif banco.group(1) == "GRERJ":
                    regexValidade = r"(\d{2}\/\d{2}\/\d{4})"
                else:
                    continue
                
                validade = buscarInformacaoEmDocumento(navegador,doc,regexValidade)
                print(validade)

                return validade.group(1)
                
        
        return validade

    if tipo =="EXECUÇÃO FISCAL":
        
        docs = procurarArquivos(navegador, "Despacho sobre Autorização de Despesa")
        regexValidade = r"\bvalidade de (.*?)\b do referido"
        
        validade = buscarInformacaoEmDocumento(navegador,docs[-1],regexValidade)
               
        return validade
    




marinette = r"C:\Users\jcampbell1\OneDrive - SEFAZ-RJ\CONTROLE GERENCIAL - PAGAMENTOS\Planilha Gerencial - Marinette.xlsx"


bloco = input("Digite o número do bloco: ")

wb = load_workbook(marinette)

if bloco not in wb.sheetnames:
    wb.create_sheet(bloco,0)
    wb.save(marinette)

df = pd.read_excel(marinette, sheet_name=bloco, dtype={'VALIDADE': str})

tipo = int(input("Qual o tipo de bloco?\n1) Fiança\n2) Execução Fiscal\n3) Caução\n"))


try:
    print(df["PROCESSO"].values)

except:

    if tipo == 1:
        colunas = ["PROCESSO", "FORMA DE PAGAMENTO", "VALIDADE",'PRAZO',"ACOMPANHAMENTO ESPECIAL","VALOR EXTRAORÇAMENTÁRIA",
                   "VALOR ORÇAMENTÁRIA",  "OB EXTRAORÇAMENTÁRIA", "OB ORÇAMENTÁRIA", "UPLOAD EXTRAORÇAMENTÁRIA",
                   "UPLOAD ORÇAMENTÁRIA","COMPROVANTES NOTIFICADOS", "DESPACHO","NOTIFICADO PARA ASSINATURA"]
    if tipo == 2 or tipo ==3:
        colunas = ["PROCESSO", "FORMA DE PAGAMENTO", "VALIDADE",'PRAZO',"ACOMPANHAMENTO ESPECIAL", "VALOR EXTRAORÇAMENTÁRIA", "OB EXTRAORÇAMENTÁRIA"]
    df = pd.DataFrame(columns=colunas, index=None)


match tipo:
    case 1:
        tipo = "FIANÇA"
    case 2:
        tipo = "EXECUÇÃO FISCAL"
    case 3:
        tipo = "CAUÇÃO"


salvarPlanilha(df, marinette, bloco)

navegador = webdriver.Firefox()

loginSEI(navegador, os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")

processos = obterProcessosDeBloco(navegador, bloco)

for i in range(1,len(processos)):
        
        WebDriverWait(navegador,20).until(EC.invisibility_of_element_located(((By.XPATH, "//div[@class = 'sparkling-modal-close']"))))
        WebDriverWait(navegador,20).until(EC.presence_of_element_located(((By.XPATH, "//tbody//tr"))))
        processo = navegador.find_elements(By.XPATH, "//tbody//tr")[i]
        nProcesso = processo.find_element(By.XPATH, './/td[3]//a').text
        
        if nProcesso not in df['PROCESSO'].values or pd.isna(df.loc[df[df["PROCESSO"] == nProcesso].index[0], "FORMA DE PAGAMENTO"]):                          
            
            if tipo == "CAUÇÃO":
                df.loc[len(df)] = {"PROCESSO":nProcesso, "VALIDADE": '-'}
                salvarPlanilha(df,marinette,bloco)
                continue
            
            validade = "-"
            
            if tipo == "EXECUÇÃO FISCAL":
                formaDePagamento = "DARJ"
            
            try:
                WebDriverWait(processo,20).until(EC.element_to_be_clickable(((By.XPATH, './/td[3]//a')))).click()
                navegador.switch_to.window(navegador.window_handles[1])
                
            
                if tipo == "FIANÇA":
                    formaDePagamento = encontrarFormaDePagamento()
    
                if "Depósito" not in formaDePagamento:
                    validade = encontrarValidade()
                
                df.loc[len(df)] = {"PROCESSO":nProcesso,"FORMA DE PAGAMENTO": formaDePagamento,"VALIDADE": validade}
                
                texto = ["Forma de Pagamento: " + formaDePagamento]
                if validade != "-":
                    texto.append('Data de Validade da Guia: ' + validade)
            
            except:
                traceback.print_exc()
                continue
     
            finally:
                navegador.close()
                navegador.switch_to.window(navegador.window_handles[0])
            
            try:
                #escreverAnotacao(navegador,texto,nProcesso)    
                salvarPlanilha(df,marinette,bloco)
            except:
                traceback.print_exc()
                continue


navegador.quit()
