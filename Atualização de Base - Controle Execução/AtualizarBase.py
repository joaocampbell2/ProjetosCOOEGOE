from datetime import date
from glob import glob
import os
from pathlib import Path
import time
import re
import traceback
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from shutil import move
import pandas as pd
import logging
from os.path import getmtime

from tqdm import tqdm
def login(navegador):
    navegador.get("https://siafe2.fazenda.rj.gov.br/Siafe/faces/login.jsp")

    user = os.environ['cpf']
    password = os.environ['senha_sefaz']

    usuario = navegador.find_element(By.XPATH, value='//*[@id="loginBox:itxUsuario::content"]')
    usuario.send_keys(user)

    senha = navegador.find_element(By.XPATH, value='//*[@id="loginBox:itxSenhaAtual::content"]')
    senha.send_keys(password)
    
    btnLogin = navegador.find_element(By.XPATH, value='//*[@id="loginBox:btnConfirmar"]')
    btnLogin.click()

    try:
        WebDriverWait(navegador,2).until(EC.element_to_be_clickable((By.XPATH, "//a[@class = 'x12k']"))).click()        
    except:
        pass

    navegador.maximize_window()
    navegador.get("https://siafe2.fazenda.rj.gov.br/Siafe/faces/flexvision/flexvisionMain.jsp")
    WebDriverWait(nav,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pt1:pt_np3:1:pt_cni4::disclosureAnchor"]'))).click()
    WebDriverWait(nav,10).until(EC.frame_to_be_available_and_switch_to_it((By.ID,'flexFrame')))
def mes():
    meses = ['Janeiro', 'Fevereiro', "Marco", 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

    mes = int(hoje.month)

    for i in range(0,len(meses)):
        if mes - 1 == i:
            x = str(mes) +' - ' +  meses[i]

    return x   
def fazerConsulta(codigo):
    WebDriverWait(nav,180).until(EC.presence_of_element_located((By.XPATH,"//span[text() = 'Consultas de outros usuários']"))).click()
    time.sleep(5)
    botoes = nav.find_elements(By.XPATH,"//div[@class = 'v-splitpanel-second-container v-scrollable']//span[@class = 'v-icon Vaadin-Icons']")
    botoes[0].click()

    #Filtrar consulta
      
    nav.find_element(By.XPATH,"//input[@placeholder = 'Digite para filtrar']").send_keys(codigo)
   
    #Aguardar a consulta certa aparecer
    
    WebDriverWait(nav,15).until(EC.presence_of_element_located((By.XPATH, "//td[text() = '"+ codigo + "']")))
    
    #Executar consulta
    
    consultas = nav.find_elements(By.XPATH, "//tbody[@class = 'v-grid-body']//tr//td")
    consultas[0].click()
    botoes[1].click()
    
    if codigo ==codigoAnna:
        pass
    if codigo == codigoRose:
        inputAno = WebDriverWait(nav,4).until(EC.presence_of_element_located((By.XPATH,'//input[contains(@id,"gwt-uid")]')))
        inputAno.send_keys(hoje.year)
        time.sleep(1)
        inputAno.send_keys(Keys.ENTER)
        
        time.sleep(4)
        
    if codigo == codigoMarcao:
       
        
        WebDriverWait(nav,4).until(EC.presence_of_element_located((By.XPATH,'//input[contains(@id,"gwt-uid")]')))
        inputs = nav.find_elements(By.XPATH,'//input[contains(@id,"gwt-uid")]')
        
        inputs[0].send_keys(hoje.year)
        time.sleep(2)
        inputs[0].send_keys(Keys.ENTER)
        
        inputs[1].send_keys(mes())
        time.sleep(2)
        inputs[1].send_keys(Keys.ENTER)
        time.sleep(5)
        
        nav.find_element(By.XPATH, "//div[@class = 'v-slot v-slot-friendly']").click()
        time.sleep(2)

    nav.find_element(By.XPATH, "//div[@class = 'v-slot v-slot-friendly']").click()

    #Baixar Planilha
    WebDriverWait(nav,180).until(EC.presence_of_element_located((By.XPATH, "//span[@class = 'v-icon v-icon-download']"))).click()
    WebDriverWait(nav,10).until(EC.presence_of_element_located((By.XPATH,  "//span[text() = 'Excel']"))).click()
    arquivo = aguardarDownload()
    #Sair
    nav.find_element(By.TAG_NAME,"body").send_keys(Keys.ESCAPE)
    return arquivo
def aguardarDownload():
    arquivo = ''
    while not os.path.isfile(arquivo):
        file_list = glob(r"C:\Users\\"+os.getlogin()+r"\Downloads\*.xls")
        for file in file_list:
            if (int(time.time()) - int(os.stat(file).st_mtime) < 2):
                arquivo = file
                time.sleep(1)
                return arquivo
                

def encontrarArquivo():
        
    files = Path(r"C:\Users\\"+os.getlogin()+r"\Downloads\\").glob('*.xls')
    arquivo_mais_recente = max(files, key=getmtime)
    newFile = Path(str(arquivo_mais_recente).replace(" ", "_"))

    move(arquivo_mais_recente, newFile)
    return str(newFile)

def salvarPlanilha(df,tabela):
    #SALVA A TABELA SEM APAGAR AS OUTRAS
    writer = pd.ExcelWriter(r"C:\Users\jcampbell1\Downloads\2024 - CONTROLE EXECUÇÃO - COPIA.xlsx", engine='openpyxl', mode='a', if_sheet_exists="replace")
    df.to_excel(writer, sheet_name=tabela, index=False, header=False)
    writer.close()


hoje = date.today()
codigoAnna = "056759"
codigoMarcao = "061581"
codigoRose = "055092"

codigos = [codigoMarcao,codigoAnna,codigoRose]
caminhos=  ["Base 1 - EGE MARCAO","Base 2 - EGE ANNA","Base 3 - EGE ROSE",]
nav = webdriver.Firefox()

login(nav)


try:
    
    for codigo, caminho in zip(codigos,caminhos):
        
        arquivo =fazerConsulta(codigo)
        planilha = pd.read_excel(str(arquivo), header=None, engine= "xlrd")
        salvarPlanilha(planilha,caminho)
        os.remove(arquivo)
except:
    traceback.print_exc()
    
finally:
    nav.quit()
