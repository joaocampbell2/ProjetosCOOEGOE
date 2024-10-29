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
from marinetteSEFAZ import loginSEI, salvarPlanilha, obterProcessosDeBloco, procurarArquivos, buscarInformacaoEmDocumento,buscarProcessoEmBloco

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
        
nav = webdriver.Firefox()
loginSEI(nav,os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")
processos = obterProcessosDeBloco(nav,bloco)

for i in tqdm(range(1,len(processos[1:]) + 1)):
    #try:
        processo = nav.find_elements(By.XPATH, "//tbody//tr")[i]

    
        linkProcesso = buscarProcessoEmBloco(nav,i)
        
        nProcesso = linkProcesso.text
    
        if "Acompanhamento Especial preenchido" not in processo.text:
            linkProcesso.click()
            nav.switch_to.window(nav.window_handles[1])
            
            
            valoresExtra = []
            valoresOrcamentario = []
            beneficiarios = []
            pagamentos = []
            
            
            despachos = procurarArquivos(nav, "Despacho sobre Autorização de Despesa")

            
            regexFiança = ""
            
            
            if grupo == "FIANÇA E VALOR APREENDIDO":
                for doc in reversed(despachos):
                    buscarInformacaoEmDocumento(nav,doc,"",)
                    
                    #Buscar oq??
            if grupo == "EXECUÇÃO FISCAL":
                darjs = procurarArquivos(nav, "DARJ")
                   
                regexExecucao = r"(CDA \d{4}\/\d{3}\.\d{3}\-\d ?|CDA \d*)[\s\S]*?no valor de R\$ ([^(]*)"
                resultado = buscarInformacaoEmDocumento(nav,despachos[-1],regexExecucao,"Coordenador")
                
                cda = resultado.group(1)
                valor = resultado.group(2)
                    
                regexDARJ = r"\bNOME\n?\n?([\n*\w*\s\(\)\.,-]*)08 - CNPJ\/CPF\n?\n?([\n\d\.\-\/]*)02 - ENDEREÇO COMPLETO"
                
                resultado = buscarInformacaoEmDocumento(nav,darjs[-1],regexDARJ,"ESTADO")
                
                nome = resultado.group(1).strip()
                cpf = resultado.group(2).strip()
                
                texto = [cda + " /","Nome:" + nome + " /", "CPF-CNPJ: " + cpf + " /", "Valor: R$" + valor + " /"]
                                
            if grupo == "CAUÇÃO":
                pass
                
    #except:
        
