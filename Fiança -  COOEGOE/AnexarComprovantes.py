from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os
import glob
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import  Keys
from marinetteSEFAZ import loginSEI, salvarPlanilha, obterProcessosDeBloco,buscarProcessoEmBloco,escreverAnotacao,incluirDocumentoExterno
from selenium.webdriver.support.ui import Select
import re
def verificaArquivosPasta(processo):
    caminhoPasta = r"C:\Users\\"+os.getlogin()+r"\Downloads\Arquivos-Processos\\"
    arquivosPasta = glob.glob(os.path.join(caminhoPasta, '*'))

    # Manipulacao de string para verificar se determinado arquivo com o nome do processo está na pasta
    listaArquivos = []
    for arquivo in arquivosPasta:
        nomeProcessoFormatado = processo.replace("-", "_")
        nomeProcessoFormatado = nomeProcessoFormatado.replace("/", "_")

        if nomeProcessoFormatado in arquivo:
            listaArquivos.append(arquivo)

    # Retorna a lista de arquivos
    return listaArquivos


bloco = input("Digite o número do bloco: ")
tipoProcesso = "FIANÇA"
nav = webdriver.Firefox()

loginSEI(nav,os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")

processos = obterProcessosDeBloco(nav,bloco)

for i in range(1,len(processos)):
    processo = nav.find_elements(By.XPATH, "//tbody//tr")[i]
    linkProcesso = buscarProcessoEmBloco(nav,i)
    nProcesso = linkProcesso.text
    print(nProcesso)  
    if  "Comprovantes Ok" not in processo.text:
        arquivosProcesso = verificaArquivosPasta(nProcesso)
        nProcesso.click()

        for arquivo in arquivosProcesso:
            ob = re.search(r"(\d{4}OB\d{5})",arquivo).group(1)  
            
            nav.switch_to.window(nav.window_handles[1])
            incluirDocumentoExterno(nav,"Comprovante",arquivo,nome=ob)

        nav.quit()
        nav.switch_to.window(nav.window_handles[0])
        escreverAnotacao(nav,"Comprovantes Ok", nProcesso)
        
    
       

    
    
