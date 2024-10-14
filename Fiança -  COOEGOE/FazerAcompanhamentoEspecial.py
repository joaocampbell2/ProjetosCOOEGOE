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

from tqdm import tqdm
from marinetteSEFAZ import loginSEI, salvarPlanilha, obterProcessosDeBloco, procurarArquivos, buscarInformacaoEmDocumento

bloco = input("Digite o número do bloco: ")
    
grupo = int(input("Digite o grupo do acompanhamento fiscal\n 1- FIANÇA E VALOR APREENDIDO\n 2- EXECUÇÃO FISCAL\n 3- CAUÇÃO\n"))

match grupo:
    case 1: 
        grupo = "FIANÇA E VALOR APREENDIDO"
    case 2: 
        grupo = "EXECUÇÃO FISCAL"
    case 3:
        grupo = "CAUÇÃO"
    case _:
        ""
marinette = r"C:\Users\jcampbell1\OneDrive - SEFAZ-RJ\CONTROLE GERENCIAL - PAGAMENTOS\Planilha Gerencial - Marinette.xlsx"
df = pd.read_excel(marinette, sheet_name=bloco)


navegador = webdriver.Firefox()
loginSEI(navegador,os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")
processos = obterProcessosDeBloco(navegador,bloco)

for processo in tqdm(processos, total = len(processos)):
    
    WebDriverWait(navegador,20).until(EC.invisibility_of_element_located(((By.XPATH, "//div[@class = 'sparkling-modal-close']"))))
    WebDriverWait(navegador,20).until(EC.presence_of_element_located(((By.XPATH, "//tbody//tr"))))
    nProcesso = processo.find_element(By.XPATH, './/td[3]//a').text
    
    
    if df.loc[df[df["PROCESSO"] == nProcesso].index[0], "ACOMPANHAMENTO ESPECIAL"] != "Ok":
        WebDriverWait(processo,20).until(EC.element_to_be_clickable(((By.XPATH, './/td[3]//a')))).click()
        navegador.switch_to.window(navegador.window_handles[1])
        
        
        if grupo == "FIANÇA E VALOR APREENDIDO":
            docs = procurarArquivos(navegador, "Despacho sobre Autorização de Despesa")
            for doc in reversed(docs):
                buscarInformacaoEmDocumento(navegador,doc,)
        
           