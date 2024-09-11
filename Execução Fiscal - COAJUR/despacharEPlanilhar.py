import time
from time import sleep
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import  Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import re
from openpyxl import load_workbook
from datetime import date, datetime
from glob import glob
from shutil import move
import tabula
from PyPDF2 import PdfReader


def loginSEI():
    navegador.get("https://sei.rj.gov.br/sip/login.php?sigla_orgao_sistema=ERJ&sigla_sistema=SEI")

    usuario = navegador.find_element(By.XPATH, value='//*[@id="txtUsuario"]')
    usuario.send_keys(os.environ['login_sefaz'])

    senha = navegador.find_element(By.XPATH, value='//*[@id="pwdSenha"]')
    senha.send_keys(os.environ['senha_sefaz'])

    exercicio = Select(navegador.find_element(By.XPATH, value='//*[@id="selOrgao"]'))
    exercicio.select_by_visible_text('SEFAZ')

    btnLogin = navegador.find_element(By.XPATH, value='//*[@id="Acessar"]')
    btnLogin.click()


    navegador.maximize_window()
    time.sleep(5)

    navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    time.sleep(1)
    trocarCoordenacao()

def trocarCoordenacao():
    coordenacao = navegador.find_elements(By.XPATH, "//a[@id = 'lnkInfraUnidade']")[1]
    if coordenacao.get_attribute("innerHTML") == 'SEFAZ/COOEGOE':
        coordenacao.click()
        WebDriverWait(navegador,5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Trocar Unidade')]")))
        navegador.find_element(By.XPATH, "//td[text() = 'SEFAZ/COOAJUR' ]").click() 
        
def abrirPastas():
    navegador.switch_to.default_content()
    WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
    listaDocs = WebDriverWait(navegador,5).until(EC.presence_of_element_located((By.ID, "divArvore")))
    pastas = listaDocs.find_elements(By.XPATH, '//a[contains(@id, "joinPASTA")]')
    for doc in pastas[:-1]:
        doc.click() 
        sleep(4)
        
        
def preencherPlanilha():
    dados = {"data":datetime.now().strftime("%d/%m/%Y") ,"processo" : processoSEI, "matéria": "Execução Fiscal", "status": "PENDENTE DE PGTO", "destinação": "À COOEGOE PARA PAGAMENTO"  }
    
    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
    quantDocs = len(docs) 
    for doc in reversed(range(quantDocs)):
        docTexto = docs[doc].text
        if "DARJ CDA" in docTexto:
            docs[doc].click()
            navegador.switch_to.default_content()            
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            WebDriverWait(navegador,5).until(EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "DARJ")]')))
            body = navegador.find_element(By.XPATH, '//body').text
            cda = re.search(r"(CERTIDÃO)\n\n([\n*\w*\s\(\)\.,-\/]*)\n\n06",body).group(2) #DARJ
            executado = re.search(r"(NOME)\n\n([\n*\w*\s\(\)\.,-]*)\n\n08",body).group(2).replace("\n", "")  #DARJ
            montante = re.search(r"(TOTAL A PAGAR)\n\n([\n*\w*\s\(\)\.,-]*)\n\n14",body).group(2)  #DARJ
            dados["cda"] = cda
            dados["executado"] = executado
            dados["montante"] = montante
            break
    
    dados["processoJudicial"] = encontrarProcessoJudicial()
    
        
def encontrarProcessoJudicial():
    navegador.switch_to.default_content()
    WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
    quantDocs = len(docs)
    for doc in (range(quantDocs)):
        docTexto = docs[doc].text
        if "Documento" in docTexto or "Ofício" in docTexto:   
            docs[doc].click()
            navegador.switch_to.default_content()            
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            time.sleep(2)
            body = navegador.find_element(By.XPATH, '//body').text
            processoJudicial = re.search(r"Executivo Fiscal \(Nº CNJ\):\n (\d[\n*\w*\s\(\)\.,-\/]*)\n\n",body)
            if processoJudicial:
                return processoJudicial.group(1)
            processoJudicial = re.search(r"Executivo Fiscal: (\d[\n*\w*\s\(\)\.,-\/]*)\n",body)
            if processoJudicial:
                return processoJudicial.group(1).replace("\n", "")
            processoJudicial = re.search(r"Ref: (\d[\n*\w*\s\(\)\.,-\/]*)\n\n",body)#DARJ
            if processoJudicial:
                return processoJudicial.group(1)



            navegador.switch_to.default_content()
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
            docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']") 
            
navegador = webdriver.Firefox()
processoSEI = "SEI-140001/050926/2024"

loginSEI()

barraPesquisa = navegador.find_element(By.ID, "txtPesquisaRapida")

barraPesquisa.send_keys(processoSEI)
barraPesquisa.send_keys(Keys.ENTER)


abrirPastas()

preencherPlanilha()