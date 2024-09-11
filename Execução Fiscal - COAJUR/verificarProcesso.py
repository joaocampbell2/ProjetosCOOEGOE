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
        
def verificarCompetencia():
    navegador.switch_to.default_content()
    WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
    quantDocs = len(docs) 
    for doc in reversed(range(quantDocs)):
            docTexto = docs[doc].text
            if "Despacho de Encaminhamento de Processo" in docTexto:
                docs[doc].click()
                navegador.switch_to.default_content()            
                WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
                WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
                WebDriverWait(navegador,5).until(EC.presence_of_element_located((By.XPATH, '//p[contains(text(), "Rio de Janeiro")]')))

                body = navegador.find_element(By.TAG_NAME, 'body').text
                if "processamento do DARJ" in body:
                    if "AUDITORA FISCAL" in body.upper() or "AUDITOR FISCAL" in body.upper():
                        return "Competencia ok"
                    else:
                        return "Competencia inválida"
                else:
                    navegador.switch_to.default_content()
                    WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
                    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
            
                        
    return "Competencia inválida"

def preencherPlanilha():
    processo = processoSEI
    
    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
    quantDocs = len(docs) 
    for doc in reversed(range(quantDocs)):
        docTexto = docs[doc].text
        if "DARJ CDA" in docTexto:
            docs[doc].click()
            navegador.switch_to.default_content()            
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            WebDriverWait(navegador,5).until(EC.presence_of_element_located((By.XPATH, '//p[contains(text(), "Rio de Janeiro")]')))
            body = navegador.find_element(By.XPATH, '//body').text
            cda = re.search(r"(CERTIDÃO)\n\n([\n*\w*\s\(\)\.,-\/]*)\n\n06",body).group(2) #DARJ
            executado = re.search(r"(NOME)\n\n([\n*\w*\s\(\)\.,-]*)\n\n08",body).group(2)  #DARJ
            montante = re.search(r"(TOTAL A PAGAR)\n\n([\n*\w*\s\(\)\.,-]*)\n\n14",body).group(2)  #DARJ
            print(cda)
            print(executado)
            print(montante)
            break
    

    processoJudicial = None #1 OU SEGUNDO DOCUMENTO

def verificarValorEValidade():
    navegador.switch_to.default_content()
    WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))

    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")
    quantDocs = len(docs) 
    for doc in reversed(range(quantDocs)):
        docTexto = docs[doc].text
        if "DARJ CDA" in docTexto:
            docs[doc].click()
            navegador.switch_to.default_content()            
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            time.sleep(2)
            body = navegador.find_element(By.XPATH, '//body').text
            print(body)
            if "DARJ" not in body:
                return "Impossível verificar DARJ", "Impossível verificar DARJ" 
            validade = re.search(r"(VENCIMENTO)\n\n([\n*\w*\s\(\)\.,-\/]*)\n\n01 ",body).group(2)  #DARJ
            validadeData = datetime.strptime(validade, '%d/%m/%Y')

            dias = (validadeData - datetime.now()).days
            if dias < 0:
                validade = "Guia fora de validade"
            else:
                validade = "Validade ok"
            montanteDARJ = re.search(r"(TOTAL A PAGAR)\n\n([\n*\w*\s\(\)\.,-]*)\n\n14 ",body).group(2)  #DARJ
            
            navegador.switch_to.default_content()
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))

            break
    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")

    for doc in reversed(range(quantDocs)):
        docTexto = docs[doc].text
        if "Guia" in docTexto:
            docs[doc].click()
            navegador.switch_to.default_content()            
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, "//span[text() = 'Guia de Recolhimento']")))
            body = navegador.find_element(By.XPATH, '//body').text 
            montanteGuia = re.search(r"Não\n.?(\d[\n*\w*\s\(\)\.,-\/]*,\d\d)\n\nU", body).group(1)
            break      
    
    if montanteDARJ != montanteGuia:
        montante  = "Montante Guia diferente de Montante DARJ"            
    else:
        montante = "Montante ok"
        
    return validade, montante
                    
navegador = webdriver.Firefox()
processoSEI = "SEI-140001/050926/2024"

loginSEI()

barraPesquisa = navegador.find_element(By.ID, "txtPesquisaRapida")

barraPesquisa.send_keys(processoSEI)
barraPesquisa.send_keys(Keys.ENTER)


abrirPastas()

validade, montante = verificarValorEValidade()

competencia = verificarCompetencia()

print("Validade: " + validade)
print("Montante: " + montante)
print("Competencia: " + competencia)

#navegador.quit()