from marinetteSEFAZ import loginSEI,buscarNumeroDocumento,incluirEmBlocoDeAssinatura, obterProcessosDeBloco,inserirHyperlinkSEI,escreverAnotacao,procurarArquivos, buscarInformacaoEmDocumento,buscarProcessoEmBloco ,incluirProcessoEmBloco,removerProcessoDoBloco,incluirDocumento
import time
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import  Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import os
import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm
from datetime import datetime
from num2words import num2words
    
def preencherDespacho(cda,ef,despacho,guia,valor,darj,validade):
    nav.switch_to.window(nav.window_handles[2])
    nav.maximize_window()
    nav.switch_to.default_content()
    btnSalvar = WebDriverWait(nav,6).until(EC.presence_of_element_located((By.XPATH, '//a[@id = "cke_142"]')))
    try:
    
        iframes = nav.find_elements(By.TAG_NAME, 'iframe')


        nav.switch_to.frame(iframes[1])
        
        corpoTexto = nav.find_element(By.TAG_NAME, 'body')
        corpoTexto.click()
        corpoTexto.send_keys(Keys.CONTROL + "a")
        corpoTexto.send_keys(Keys.BACKSPACE)
        
        nav.switch_to.default_content()
        nav.switch_to.frame(iframes[2])
        
        corpoTexto = nav.find_element(By.TAG_NAME, 'body')
        corpoTexto.click()
        
        corpoTexto.send_keys(Keys.PAGE_UP)
        corpoTexto.send_keys(Keys.ARROW_DOWN)
        corpoTexto.send_keys(Keys.END)
        
        for i in range(5):
            corpoTexto.send_keys(Keys.BACKSPACE)
        
        corpoTexto.send_keys(cda)
        corpoTexto.send_keys(Keys.ARROW_DOWN)
        corpoTexto.send_keys(ef)

        for i in range(3):
            corpoTexto.send_keys(Keys.ARROW_DOWN)
        
        for i in range(5):
            corpoTexto.send_keys(Keys.CONTROL + Keys.ARROW_LEFT)
            
        inserirHyperlinkSEI(nav,despacho)
        
        
        nav.switch_to.frame(iframes[2])



        for i in range(10):
            corpoTexto.send_keys(Keys.CONTROL + Keys.ARROW_LEFT)
        for i in range(7):
            corpoTexto.send_keys(Keys.ARROW_RIGHT)
        
                   
        inserirHyperlinkSEI(nav,guia)
        
        nav.switch_to.frame(iframes[2])

        
        corpoTexto.send_keys(Keys.ARROW_DOWN)

        corpoTexto.send_keys(Keys.HOME)
        
        for i in range(11):
            corpoTexto.send_keys(Keys.CONTROL + Keys.ARROW_RIGHT)
            
        corpoTexto.send_keys(Keys.ARROW_LEFT)

        for i in range(4):
            corpoTexto.send_keys(Keys.BACKSPACE)
        
        
        
        valorExtenso = valor.replace(".","")
        valorExtenso = valorExtenso.replace(",",".")

        reais,centavos = valorExtenso.split(".")
        reais = num2words(reais, lang="pt-BR")

        if centavos != "0":
            centavos = num2words(centavos, lang="pt-BR")
            corpoTexto.send_keys(valor + " (" + reais + " reais e " + centavos + " centavos" + ")")

        else:
            corpoTexto.send_keys(valor + " (" + reais + " reais" + ")")

 
        for i in range(7):
            corpoTexto.send_keys(Keys.CONTROL + Keys.ARROW_RIGHT)
        
        corpoTexto.send_keys(Keys.ARROW_RIGHT)
        inserirHyperlinkSEI(nav,darj)
        nav.switch_to.frame(iframes[2])


        corpoTexto.send_keys(Keys.ARROW_DOWN)
        corpoTexto.send_keys(Keys.ARROW_DOWN)


        corpoTexto.send_keys(Keys.HOME)
        
        for i in range(10):
            corpoTexto.send_keys(Keys.CONTROL + Keys.ARROW_RIGHT)

        corpoTexto.send_keys(Keys.ARROW_LEFT)

            
        for i in range(24):
            corpoTexto.send_keys(Keys.BACKSPACE)
            
        corpoTexto.send_keys(validade)
    except:
        traceback.print_exc()
    
    finally:
                
        nav.switch_to.default_content()
        btnSalvar.click()
        nav.close()
        
        nav.switch_to.window(nav.window_handles[1])
        



    
controle = r"C:\Users\jcampbell1\Downloads\CONTROLE GERENCIAL - EXECUÇÃO FISCAL.xlsx"
df = pd.read_excel(controle,sheet_name="EXECUÇÃO FISCAL",header=3)    
    
    

nav = webdriver.Firefox()
loginSEI(nav,os.environ['login_sefaz'], os.environ['senha_sefaz'], "SEFAZ/COOAJUR")

processos = obterProcessosDeBloco(nav,"938324")
for i in tqdm(range(1,len(processos[1:]) + 1)):
    
    processo = nav.find_elements(By.XPATH, "//tbody//tr")[i].text
    if "Despacho Ok" not in processo:
        try:
            linkProcesso = buscarProcessoEmBloco(nav,i)
            nProcesso = linkProcesso.text
            print(nProcesso)
            
            index = df.index[df['PROCESSO ADMINISTRATIVO'] == nProcesso]
            cda = df.loc[index]["CDA"]
            ef = df.loc[index]["PROCESSO JUDICIAL"]
            montante = df.loc[index]["MONTANTE"].values[0]
            linkProcesso.click()
            nav.switch_to.window(nav.window_handles[1])
        except:
            traceback.print_exc()
            continue
        try:
            numGuia = buscarNumeroDocumento(nav,"Guia")
            numDarj = buscarNumeroDocumento(nav,"DARJ")
            numDespacho = buscarNumeroDocumento(nav,"Despacho de Encaminhamento de Processo")
            
            darjs= procurarArquivos(nav, "DARJ")

            for darj in reversed(darjs):

                validade = buscarInformacaoEmDocumento(nav,darjs[-1],"(VENCIMENTO)\n\n([\n*\w*\s\(\)\.,-\/]*)\n\n01 ",verificador="DARJ")
                if validade:
                    validade = validade.group(2)
        
        
            incluirDocumento(nav,"Despacho sobre Autorização de Despesa","Texto Padrão",modelo="EF - À COOEGOE PARA PGTO",hipotese="Comprometer Atividades (Art.23º, VIII da Lei nº 12.527/2011)" )
            preencherDespacho(cda,ef,numDespacho,numGuia,montante,numDarj,validade)
            incluirEmBlocoDeAssinatura(nav,"232228 - EXECUÇÃO FISCAL - À COOEGOE (superintendente)", "Despacho de Autorização de Despesa")
        except:
            traceback.print_exc()
            continue
        finally:
            nav.close()
            nav.switch_to.window(nav.window_handles[0])
            
        escreverAnotacao(nav,["Despacho Ok"],nProcesso)
        
    
    