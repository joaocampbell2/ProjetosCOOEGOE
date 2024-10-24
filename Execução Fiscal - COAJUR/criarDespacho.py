from marinetteSEFAZ import loginSEI, obterProcessosDeBloco,inserirHyperlinkSEI,  escreverAnotacao,procurarArquivos, buscarInformacaoEmDocumento,buscarProcessoEmBloco ,incluirProcessoEmBloco,removerProcessoDoBloco,incluirDocumento
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
        
    
def buscarNumeroDocumento(nav,nome):
    nav.switch_to.default_content()

    WebDriverWait(nav,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
    # abrirPastas(nav)
    nav.switch_to.default_content()
    WebDriverWait(nav,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
    listaDocs = WebDriverWait(nav,5).until(EC.presence_of_element_located((By.ID, "divArvore")))
    pastas = listaDocs.find_elements(By.XPATH, '//a[contains(@id, "joinPASTA")]//img[contains(@title, "Abrir")]')
    
    for doc in pastas:
        doc.click() 
        WebDriverWait(nav,5).until(EC.presence_of_element_located((By.XPATH, "//*[text() = 'Aguarde...']")))
        WebDriverWait(nav,5).until(EC.invisibility_of_element((By.XPATH, "//*[text() = 'Aguarde...']")))
    docs = nav.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
    
    #fechaabre pastas
    
    for doc in reversed(docs):
        if nome in doc.text:
            return  re.search(r"(\d+)\)?$", doc.text).group(1)



def incluirEmBlocoDeAssinatura(nav,documento,blocoAssinatura):
    print("Incluindo no novo bloco de assinatura...")
    nav.switch_to.default_content()

    iframeBotoes = nav.find_element(By.ID, "ifrVisualizacao")
    nav.switch_to.frame(iframeBotoes)

    arvoreBotoes = nav.find_element(By.ID, "divInfraAreaTela")
    botoesSei = arvoreBotoes.find_element(By.CLASS_NAME, "barraBotoesSEI")
    opcoesBotoesSei = botoesSei.find_elements(By.TAG_NAME, "a")
    for opcaoBotaoSei in opcoesBotoesSei:
        infoBotao = opcaoBotaoSei.find_element(By.TAG_NAME, "img")
        attrTitle = infoBotao.get_attribute("title")
        if attrTitle == "Incluir em Bloco de Assinatura":
            opcaoBotaoSei.click()
            break

    WebDriverWait(nav, 20).until(EC.element_to_be_clickable((By.ID, "selBloco")))
    # Clicar para abrir a aba de blocos
    nav.find_element(By.ID, "selBloco").click()
    # Se for fianca, clicar na opcao 506041 - Fiança GREs - COOCR/COOAJUR
    selecaoBloco = nav.find_element(By.ID, "selBloco")
    optionsBloco = selecaoBloco.find_elements(By.TAG_NAME, "option")
   

    for optionBloco in optionsBloco:
        if optionBloco.text == blocoAssinatura:
            optionBloco.click()
            break
            
    
    # Incluir no bloco de assinatura
    nav.find_element(By.ID, "sbmIncluir").click()

    time.sleep(5)
    nav.switch_to.default_content()

    print("Incluido com sucesso.")
    # Funcao que retira o processo do bloco de assinatura:
    #excluirDoBloco(nav, processo)  
    
controle = r"C:\Users\jcampbell1\Downloads\CONTROLE GERENCIAL - EXECUÇÃO FISCAL.xlsx"
df = pd.read_excel(controle,sheet_name="EXECUÇÃO FISCAL",header=3)    
    
    

nav = webdriver.Firefox()
loginSEI(nav,os.environ['login_sefaz'], os.environ['senha_sefaz'], "SEFAZ/COOAJUR")

processos = obterProcessosDeBloco(nav,"938324")
try:
    for i in tqdm(range(1,len(processos[1:]) + 1)):
        nProcesso = buscarProcessoEmBloco(nav,i)
        print(nProcesso.text)
        
        index = df.index[df['PROCESSO ADMINISTRATIVO'] == nProcesso.text]
        
        cda = df.loc[index]["CDA"]
        ef = df.loc[index]["PROCESSO JUDICIAL"]
        montante = df.loc[index]["MONTANTE"].values[0]
        nProcesso.click()
        nav.switch_to.window(nav.window_handles[1])

        numGuia = buscarNumeroDocumento(nav,"Guia")
        numDarj = buscarNumeroDocumento(nav,"DARJ")
        numDespacho = buscarNumeroDocumento(nav,"Despacho de Encaminhamento de Processo")
        
        darjs= procurarArquivos(nav, "DARJ")

        for darj in reversed(darjs):

            validade = buscarInformacaoEmDocumento(nav,darjs[-1],"(VENCIMENTO)\n\n([\n*\w*\s\(\)\.,-\/]*)\n\n01 ",verificador="DARJ")
            if validade:
                validade = validade.group(2)
        
        try:
            incluirDocumento(nav,"Despacho sobre Autorização de Despesa","Texto Padrão",modelo="EF - À COOEGOE PARA PGTO",hipotese="Comprometer Atividades (Art.23º, VIII da Lei nº 12.527/2011)" )
            preencherDespacho(cda,ef,numDespacho,numGuia,montante,numDarj,validade)
            incluirEmBlocoDeAssinatura(nav,"Despacho de Autorização de Despesa","232228 - EXECUÇÃO FISCAL - À COOEGOE (superintendente)")
        except:
            traceback.print_exc()
        
        finally:
            nav.close()
            nav.switch_to.window(nav.window_handles[0])
            break
        
    
except:
    traceback.print_exc()
    