import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from openpyxl import load_workbook
from marinetteSEFAZ import loginSEI,incluirDocumento,buscarNumeroDocumento,inserirHyperlinkSEI, obterProcessosDeBloco,procurarArquivos, buscarInformacaoEmDocumento, escreverAnotacao, incluirEmBlocoDeAssinatura, buscarProcessoEmBloco
from selenium.webdriver.common.keys import  Keys


def descobrirTipoProcesso():
    docs = procurarArquivos(nav,"Despacho de Encaminhamento de Processo")
    for doc in reversed(docs):
        filtro = buscarInformacaoEmDocumento(nav,doc,"DARJ","Estado")
        if filtro:
            return "68111827"
        filtro = buscarInformacaoEmDocumento(nav,doc,"GRE","Estado")
        if filtro:
            return "68112290"
        
def preencherDespacho(filtro, comprovantes):
    nav.switch_to.window(nav.window_handles[2])
    nav.maximize_window()
    nav.switch_to.default_content()
    btnSalvar = WebDriverWait(nav,6).until(EC.presence_of_element_located((By.XPATH, '//a[@id = "cke_142"]')))
    try:
    
        iframes = nav.find_elements(By.TAG_NAME, 'iframe')


        nav.switch_to.frame(iframes[1])
        
        corpoTexto = nav.find_element(By.TAG_NAME, 'body')
        corpoTexto.click()
        
        
        if filtro == "68112290":
            
            corpoTexto.send_keys(Keys.ARROW_DOWN)

            for i in range(2):
                corpoTexto.send_keys(Keys.ARROW_LEFT)
            for i in range(12):
                corpoTexto.send_keys(Keys.BACKSPACE)
        elif filtro == "68111827":

            corpoTexto.send_keys(Keys.PAGE_UP)
            corpoTexto.send_keys(Keys.END)
            for i in range(41):
                corpoTexto.send_keys(Keys.ARROW_LEFT)
            for i in range(7):
                corpoTexto.send_keys(Keys.BACKSPACE)
                
                
        for comprovante in comprovantes:
            inserirHyperlinkSEI(nav,comprovante)
            nav.switch_to.frame(iframes[2])
            corpoTexto.send_keys(Keys.SPACE)
   
         
    except:
        traceback.print_exc()
    
    finally:
                
        nav.switch_to.default_content()
        btnSalvar.click()
        nav.close()
        
        nav.switch_to.window(nav.window_handles[1])
        

bloco = input("Digite o número do bloco: ")
nav = webdriver.Firefox()

loginSEI(nav, os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")

processos = obterProcessosDeBloco(nav, bloco)

for i in range(1,len(processos)):
    processo = nav.find_elements(By.XPATH, "//tbody//tr")[i]
    textoProcesso = processo.text
    
    linkProcesso = buscarProcessoEmBloco(nav,i)    
    nProcesso = linkProcesso.text
    
    if "Comprovantes Ok" in textoProcesso and "Despacho O" not in textoProcesso:
        print(nProcesso)
        linkProcesso.click()
        nav.switch_to.window(nav.window_handles[1])
        try:
        
            filtro = descobrirTipoProcesso()
            comprovantes = buscarNumeroDocumento(nav,"Comprovante",lista=True)
            if comprovantes == []:
                comprovantes = buscarNumeroDocumento(nav,"OB",lista=True)
            print(filtro)
            incluirDocumento(nav,"Despacho de Encaminhamento de Processo","Documento Modelo",modelo=filtro,hipotese='Controle Interno (Art. 26, § 3º, da Lei nº 10.180/2001)' )
            preencherDespacho(filtro,comprovantes)
            
            if filtro == "68112290":
                assinatura = "506041 - Fiança GREs - COOCR/COOAJUR"
            if filtro == "68111827":
                assinatura = "135319 - Fiança DARJs E PAs COM PROBLEMAS - COOAJUR"
            
            if "Guia" not in textoProcesso:
                incluirEmBlocoDeAssinatura(nav,assinatura)
            
        except:
            traceback.print_exc()
            continue
        finally:
            nav.close()
            nav.switch_to.window(nav.window_handles[0])
            
        escreverAnotacao(nav,["Despacho Ok"],nProcesso)
