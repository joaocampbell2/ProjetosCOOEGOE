import time
from time import sleep
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import  Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
import pyautogui
from datetime import date
from glob import glob
from shutil import move

meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", 
         "Dezembro"]

processos = ["PGOV0822P","PGOV0832P","TGRJ0801P","TGRJ0802P","TGRJ0807P", "TGRJ0808P"]

def login(navegador):
    navegador.get("https://sei.rj.gov.br/sip/login.php?sigla_orgao_sistema=ERJ&sigla_sistema=SEI")

    usuario = navegador.find_element(By.XPATH, value='//*[@id="txtUsuario"]')
    usuario.send_keys(os.environ['login_sefaz'])

    senha = navegador.find_element(By.XPATH, value='//*[@id="pwdSenha"]')
    senha.send_keys(os.environ['senha_sefaz'])

    exercicio = Select(navegador.find_element(By.XPATH, value='//*[@id="selOrgao"]'))
    exercicio.select_by_visible_text('SEFAZ')

    btnLogin = navegador.find_element(By.XPATH, value='//*[@id="Acessar"]')
    btnLogin.click()

    time.sleep(5)

    navegador.maximize_window()

    navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    
def baixarRelatorios(navegador):
    
    arvore = WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.ID, "ifrArvore")))    
    visualizacao = navegador.find_element(By.XPATH, "//iframe[@id = 'ifrVisualizacao']")
    navegador.switch_to.frame(arvore)

    listaDocs =  WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.ID, "divArvore")))  
    docs = listaDocs.find_elements(By.TAG_NAME, "a")

    for doc in docs:
        if "FOLHA" in doc.text.upper(): 
            doc.click()
            navegador.switch_to.default_content()
            WebDriverWait(navegador,3).until(EC.frame_to_be_available_and_switch_to_it(visualizacao))
            WebDriverWait(navegador,3).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            time.sleep(0.3)
            
            
            navegador.find_element(By.XPATH, '//button[@id ="secondaryToolbarToggle" ]').click()
            time.sleep(0.3)

            navegador.find_element(By.XPATH, '//button[@id ="secondaryDownload"]').click()
            time.sleep(2)

            pyautogui.press("enter")
            alterarNomeArquivo()
            navegador.switch_to.default_content()
            navegador.switch_to.frame(arvore)

def alterarNomeArquivo():
    arquivo = ""
    while not os.path.isfile(arquivo):
        file_list = glob(r"C:\Users\\"+os.getlogin()+r"\Downloads\*.pdf")
        for file in file_list:
            if (int(time.time()) - int(os.stat(file).st_mtime) < 4):
                arquivo = file
                time.sleep(1)
                break
    
    for processo in processos:
        if processo in arquivo.upper():
            nome = processo + "_" + meses[date.today().month - 1]
            newFile = r"C:\Users\\"+os.getlogin()+r"\Downloads\\" + nome + ".pdf"
            move(arquivo, newFile)
            return
    os.remove(arquivo)

navegador = webdriver.Firefox()
login(navegador)

processo = "SEI-040002/002623/2024"

barraPesquisa = navegador.find_element(By.ID, "txtPesquisaRapida")

barraPesquisa.send_keys(processo)
barraPesquisa.send_keys(Keys.ENTER)

baixarRelatorios(navegador)
