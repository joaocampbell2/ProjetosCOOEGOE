import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from marinetteSEFAZ import loginSEI,incluirDespacho,buscarNumeroDocumento,inserirHyperlinkSEI, obterProcessosDeBloco,procurarArquivos, buscarInformacaoEmDocumento, escreverAnotacao, incluirEmBlocoDeAssinatura, buscarProcessoEmBloco
from selenium.webdriver.common.keys import  Keys


def descobrirTipoProcesso():
    docs = procurarArquivos(nav,"Despacho de Encaminhamento de Processo")
    for doc in reversed(docs):
        try:
            modelo = buscarInformacaoEmDocumento(nav,doc,"DARJ","Estado")
            if modelo:
                return "68111827"
            modelo = buscarInformacaoEmDocumento(nav,doc,"GRE","Estado")
            if modelo:
                return "68112290"
        except:
            pass
        
def preencherDespacho(tipo, comprovantes,modelo= None):
    nav.switch_to.window(nav.window_handles[2])
    nav.maximize_window()
    nav.switch_to.default_content()
    btnSalvar = WebDriverWait(nav,6).until(EC.presence_of_element_located((By.XPATH, '//a[@id = "cke_142"]')))
    try:
    
        iframes = nav.find_elements(By.TAG_NAME, 'iframe')


        nav.switch_to.frame(iframes[2])
        
        corpoTexto = nav.find_element(By.TAG_NAME, 'body')
        corpoTexto.click()
        
        if tipo == "FIANÇA":
        
            if modelo == "68112290":
                
                corpoTexto.send_keys(Keys.PAGE_UP)
                corpoTexto.send_keys(Keys.ARROW_DOWN)
                corpoTexto.send_keys(Keys.ARROW_DOWN)

                for i in range(3):
                    corpoTexto.send_keys(Keys.ARROW_LEFT)
                for i in range(12):
                    corpoTexto.send_keys(Keys.BACKSPACE)
                    
            elif modelo == "68111827":

                corpoTexto.send_keys(Keys.PAGE_UP)
                corpoTexto.send_keys(Keys.END)
                for i in range(41):
                    corpoTexto.send_keys(Keys.ARROW_LEFT)
                for i in range(7):
                    corpoTexto.send_keys(Keys.BACKSPACE)
        
        if tipo =="EXECUÇÃO FISCAL":
            corpoTexto.send_keys(Keys.PAGE_UP)
            corpoTexto.send_keys(Keys.END)
            for i in range(8):
                corpoTexto.send_keys(Keys.CONTROL + Keys.ARROW_LEFT)
            for i in range(4):
                corpoTexto.send_keys(Keys.ARROW_RIGHT)
            for i in range(4):
                corpoTexto.send_keys(Keys.BACKSPACE)

                
        for comprovante in comprovantes:
            inserirHyperlinkSEI(nav,comprovante)
            nav.switch_to.frame(iframes[2])
            corpoTexto.send_keys(Keys.SPACE)
   
        corpoTexto.send_keys(Keys.BACKSPACE)
        
    except:
        traceback.print_exc()
    
    finally:
                
        nav.switch_to.default_content()
        btnSalvar.click()
        nav.close()
        
        nav.switch_to.window(nav.window_handles[1])
        

bloco = input("Digite o número do bloco: ")
tipo = input("Selecione o tipo de processo\n1) Execução\n2) Fiança\n")
nav = webdriver.Firefox()

match tipo:
    case "1":
        tipo = "EXECUÇÃO FISCAL"
    case "2":
        tipo = "FIANÇA"

loginSEI(nav, os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")

processos = obterProcessosDeBloco(nav, bloco)

for i in range(1,len(processos)):
    processo = nav.find_elements(By.XPATH, "//tbody//tr")[i]
    textoProcesso = processo.text
    
    linkProcesso = buscarProcessoEmBloco(nav,i)    
    nProcesso = linkProcesso.text
    
    if "COMPROVANTES OK" in textoProcesso.upper() and "DESPACHO OK" not in textoProcesso.upper():
        print(nProcesso)
        linkProcesso.click()
        nav.switch_to.window(nav.window_handles[1])
        try:
            if tipo == "FIANÇA":
                modelo = descobrirTipoProcesso()
            else:
                modelo = "68113483"
            comprovantes = buscarNumeroDocumento(nav,"Comprovante",lista=True)
            incluirDespacho(nav,"Despacho de Encaminhamento de Processo","Documento Modelo",modelo=modelo,hipotese='Controle Interno (Art. 26, § 3º, da Lei nº 10.180/2001)' )
            try:
            
                preencherDespacho(tipo,comprovantes,modelo = modelo)
            except:
                nav.close()
                nav.switch_to.window(nav.window_handles[1])
                traceback.print_exc()
                pass
            if modelo == "68112290":
                assinatura = "506041 - Fiança GREs - COOCR/COOAJUR"
            if modelo == "68111827":
                assinatura = "135319 - Fiança DARJs E PAs COM PROBLEMAS - COOAJUR"
            if modelo == "68113483":
                assinatura = "407903 - Execução Fiscal"
            if "Guia" not in textoProcesso or tipo == "EXECUÇÃO FISCAL":
                incluirEmBlocoDeAssinatura(nav,assinatura)
            
        except:
            traceback.print_exc()
            continue
        finally:
            nav.close()
            nav.switch_to.window(nav.window_handles[0])
            
        escreverAnotacao(nav,["Despacho Ok"],nProcesso)

nav.quit()