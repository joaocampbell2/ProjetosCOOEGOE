import traceback
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import os
import glob
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import  Keys
import re
from marinetteSEFAZ import loginSEI, salvarPlanilha, obterProcessosDeBloco,buscarProcessoEmBloco,escreverAnotacao

def verificaArquivosPasta(processo):
    # Abrir pasta com os arquivos
    caminhoPasta = r"C:\Users\jcampbell1\Downloads\Arquivos-Processos\\"
    #caminhoPasta = "C:\\Users\\Gabriel Ferraz\\Documents\\SEFAZ\\Scripts-SEFAZ\\Arquivos-Anexar"
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


def anexarComprovante(nav,):
    WebDriverWait(nav,5).until(EC.presence_of_element_located(((By.ID, "divInfraAreaTela"))))
    conteudoTelaProcesso = nav.find_element(By.ID, "divInfraAreaTela")
    iframe = conteudoTelaProcesso.find_element(By.ID, "ifrVisualizacao")
    dicio = {}
    # Muda para o iframe, clica em anexar documento e seleciona um a partir do caminho do mesmo
    # try:  # O Try/Except evita que em caso de erro o codigo fique no frame antigo e atrapalhe a execucao
    #     nav.switch_to.frame(iframe)




marinette = r"C:\Users\jcampbell1\OneDrive - SEFAZ-RJ\CONTROLE GERENCIAL - PAGAMENTOS\Planilha Gerencial - Marinette.xlsx"
bloco = input("Digite o número do bloco: ")
df = pd.read_excel(marinette,sheet_name=bloco)
tipoProcesso = "FIANÇA"
nav = webdriver.Firefox()

loginSEI(nav,os.environ['login_sefaz'],os.environ['senha_sefaz'],"SEFAZ/COOEGOE")

processos = obterProcessosDeBloco(nav,bloco)

for i in range(1,len(processos)):
    
    nProcesso = buscarProcessoEmBloco(nav,i)
    nComprovantes = 2
    index = df.index[df['PROCESSO'] == nProcesso.text]
    orcamentaria = df.loc[index]['OB ORÇAMENTÁRIA']
    uploadOrcamentaria = df.loc[index]["UPLOAD ORÇAMENTÁRIA"]
    uploadExtraOrcamentaria = df.loc[index]["UPLOAD EXTRAORÇAMENTÁRIA"]

    comprovantes = df.loc[index]['COMPROVANTES NOTIFICADOS']
    
    if orcamentaria.values[0] == "Erro no Pagamento" or df.loc[index]['OB EXTRAORÇAMENTÁRIA'].values[0] == "Erro no Pagamento" and comprovantes.values[0] != 'Ok':
        escreverAnotacao(nav,df.loc[index]["OB EXTRAORÇAMENTÁRIA"].values[0],nProcesso.text)
        comprovantes = "Ok"
        salvarPlanilha(df,marinette,bloco)
        continue 
    
    if orcamentaria.values[0] == 'Processo sem OB':
            nComprovantes = 1
            uploadOrcamentaria = "Processo sem OB"
            
    if uploadExtraOrcamentaria.values[0] != "Ok" and uploadOrcamentaria.values[0] != "Ok":
        arquivosProcesso = verificaArquivosPasta(nProcesso)
        if len(arquivosProcesso) >= nComprovantes:
            nProcesso.click()
            nav.switch_to.window(nav.window_handles[1])
            anexarComprovante()

    
       

    
    
