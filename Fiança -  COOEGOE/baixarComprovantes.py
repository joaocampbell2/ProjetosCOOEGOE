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
from marinetteSEFAZ import loginSEI,loginSIAFE, obterProcessosDeBloco,buscarProcessoEmBloco,escreverAnotacao
import traceback
from selenium.webdriver.support.ui import Select
import re
from pathlib import Path
from os.path import getmtime
import shutil

def extrairOB(processo,tipoOB):
    if tipoOB == "Extra":
        index = 12
    if tipoOB == "Orcamentaria":
        index =13

    numeroOB = None
    
    WebDriverWait(nav, 20).until(EC.element_to_be_clickable((By.XPATH, '// *[ @id = "pt1:tblOB' + tipoOB + ':sdtFilter::btn"]')))
    try:
        btnLimpar = nav.find_element(By.XPATH, value= '//*[@id="pt1:tblOB' + tipoOB + ':btnClearFilter"]')
        btnLimpar.click()
    except:
        btnFiltro = nav.find_element(By.XPATH, value='// *[ @ id = "pt1:tblOB' + tipoOB + ':sdtFilter::disAcr"]')
        btnFiltro.click()
        try:
            btnLimpar = nav.find_element(By.XPATH, value='//*[@id="pt1:tblOB' + tipoOB + ':btnClearFilter"]')
            btnLimpar.click()
        except:
            None
    
    WebDriverWait(nav,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="pt1:tblOB' + tipoOB + ':table_rtfFilter:0:in_value_rtfFilter::content"]')))

    
    acessoRapido = nav.find_element(By.ID, value='pt1:iTxtCad::content')

    acessoRapido.click()

    filtroProcessoSEI = Select(nav.find_element(By.XPATH,
                                                        value='//*[@id="pt1:tblOB' + tipoOB + ':table_rtfFilter:0:cbx_col_sel_rtfFilter::content"]'))
    filtroProcessoSEI.select_by_visible_text('Processo')

    filtroContemProcesso = Select(nav.find_element(By.XPATH,
                                                            value='//*[@id="pt1:tblOB' + tipoOB + ':table_rtfFilter:0:cbx_op_sel_rtfFilter::content"]'))
    filtroContemProcesso.select_by_visible_text('contém')
    valorFiltroProcesso = nav.find_element(By.XPATH,
                                                    value='//*[@id="pt1:tblOB' + tipoOB + ':table_rtfFilter:0:in_value_rtfFilter::content"]')    
    valorFiltroProcesso.clear()
    valorFiltroProcesso.send_keys(processo)

    acessoRapido.click()
    
    WebDriverWait(nav,8).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pt1:tblOB' + tipoOB + ':table_rtfFilter:1:cbx_col_sel_rtfFilter::content"]')))
    
    try:
        WebDriverWait(nav, 8).until(EC.element_to_be_clickable((By.XPATH, "//span[text()= '"+ processo + "']")))
    except:
        return "OB não encontrada!"
    
    tabelaDataResultado = nav.find_element(By.XPATH, value='//*[@id="pt1:tblOB' + tipoOB + ':tabViewerDec::db"]')
    rows = tabelaDataResultado.find_elements(By.TAG_NAME, value="tr")
    if len(rows) > 0:
        for i in range(len(rows)):
            tabelaDataResultado = nav.find_element(By.XPATH, value='//*[@id="pt1:tblOB' + tipoOB + ':tabViewerDec::db"]')
            rows = tabelaDataResultado.find_elements(By.TAG_NAME, value="tr")
            col = rows[i].find_elements(By.TAG_NAME, value="td")
            status = col[index].text
            if status == "Processado e Pago":
                col[index].click()
                btnVisualizar = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt1:tblOB' + tipoOB + ':btnView"]')))
                btnVisualizar.click()
                try:
                    btnImpComprovante = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt1:btnImpComprov"]')))
                except:
                    btnImpComprovante = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tplSip:btnImpComprov"]')))

                btnImpComprovante.click()
                btnImpPDF = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '// *[@id = "pt1:btnPDF"]')))
                btnImpPDF.click()
                WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pt1:popPageConfig::content"]'))).click()
                btnImpPDFOK = nav.find_element(By.XPATH, value='// *[ @ id = "pt1:dlgPageConfig::ok"]')
                btnImpPDFOK.click()
                
                
                body = nav.find_element(By.XPATH,"//body").text
                
                numeroOB = re.search(r"(\d{4}OB\d{5})",body).group(1)            
                
                nav.switch_to.window(nav.window_handles[1])
                nav.close()
                nav.switch_to.window(nav.window_handles[0])
                try:
                    btnSair = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt1:btnCancelar"]')))
                except:
                    btnSair = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tplSip:btnCancelar"]')))
                btnSair.click()
                
                time.sleep(2)
                try:
                    btnSair = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt1:btnCancelar"]')))
                except:
                    btnSair = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tplSip:btnCancelar"]')))
                    
                btnSair.click()
                
                alterarNomeDownload(processo,tipoOB,numeroOB)
                
            elif status == "Erro no Pagamento":
                col[index].click()
                buscarErroNoPagamento(tipoOB,processo)
            else:
                print(col[index].text)
              
    else:
        print("Não há OB disponível")

def alterarNomeDownload(processo,tipoOB,numeroOB):
    caminho = "C:/users/" + os.getlogin() + "/" + "Downloads/"
    processo = re.sub("[/-]","_",processo)
    novoNome = processo + "_OB_" + tipoOB + "_Numero_" + numeroOB + ".pdf"
    files = Path(caminho).glob('*.pdf')
    arquivo_mais_recente = max(files, key=getmtime)
    newFile = Path(caminho+"Arquivos-Processos/"+novoNome)
    shutil.move(arquivo_mais_recente, newFile)
    
    return print(f"Arquivo {novoNome} criado!")

def buscarErroNoPagamento(tipoOB,processo):
    WebDriverWait(nav,8).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt1:tblOB'+tipoOB+':btnView"]'))).click()

    try:
        WebDriverWait(nav,5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tplSip:slcGuiaDevolucao::disAcr"]'))).click()
    except:
        nav.find_element(By.XPATH, value='//*[@id="pt1:slcGuiaDevolucao::disAcr"]').click()
    WebDriverWait(nav,8).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Visualizar observação"]'))).click()
    
    try:
        observacao = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//textarea[@id = "tplSip:tblSucesso:itxDescricaoObservacao::content"]')))
        
    except:
        observacao = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//textarea[@id = "pt1:tblSucesso:itxDescricaoObservacao::content"]')))

    erro = "ERRO NO PAGAMENTO DA OB "+tipoOB+": " + observacao.text.split('"')[1]
    
    try:
        btnOk = nav.find_element(By.XPATH,"//*[@id = 'pt1:tblSucesso:cmdOkVisualizarLog']")
    except:
        btnOk = nav.find_element(By.XPATH,"//*[@id = 'tplSip:tblSucesso:cmdOkVisualizarLog']")
        
    btnOk.click()
        
    try:
        btnSair = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pt1:btnCancelar"]')))
    except:
        btnSair = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tplSip:btnCancelar"]')))
                
    btnSair.click()
    
    anotarErro(erro,processo)
    
def anotarErro(erro,processo):
    nav.switch_to.new_window()
    loginSEI(nav,os.environ["login_sefaz"],os.environ["senha_sefaz"],"SEFAZ/COOEGOE")
    obterProcessosDeBloco(nav,bloco)
    processoEncontrado = buscarProcessoEmBloco(nav,processo).text
    escreverAnotacao(nav,erro,processoEncontrado)
    nav.close()
    nav.switch_to.window(nav.window_handles[0])
    print("Erro de pagamento alertado")

bloco = input("Digite o número do bloco: ")
bateria = input("Digite a bateria: ")
forma = input("Digite qual forma de pagamento é pra extrair as OB's:\n1) Depósito Bradesco\n2) Depósito outros bancos\n3) Guia\n4) Todos\n")
nav = webdriver.Firefox()

match forma:
    case "1":
        forma = "Depósito Bradesco"
    case "2":
        forma = "Depósito"
    case "3":
        forma = "Guia"
    case "4":
        forma = ""  


loginSEI(nav,os.environ["login_sefaz"], os.environ["senha_sefaz"],"SEFAZ/COOEGOE")

processos = obterProcessosDeBloco(nav,bloco)
processosParaBaixar = []
for i in range(1,len(processos)):
    
    processo = nav.find_elements(By.XPATH, "//tbody//tr")[i].text
    if forma in processo and bateria in processo:
        if forma == "Depósito" and "Bradesco" in processo:
            continue
        linkProcesso = buscarProcessoEmBloco(nav,i)
        processosParaBaixar.append(linkProcesso.text) 

print(processosParaBaixar)

loginSIAFE(nav,os.environ['cpf'],os.environ['senha_siafe'])

link = 'https://siafe2.fazenda.rj.gov.br/Siafe/faces/execucao/financeira/ordemBancariaExtraOrcamentariaCad.jsp'
tipoOB = "Extra"
erros = []

for i in range(0,2):
    nav.get(link) 
    for processo in processosParaBaixar:
        print(processo + " " + tipoOB)
        try:
            extrairOB(processo,tipoOB)
        except:
            print('Não Foi possível extrair a OB ' + tipoOB )
            erros.append({"processo": processo,"tipo": tipoOB, "link": link})
            traceback.print_exc()
            continue
            
    link = 'https://siafe2.fazenda.rj.gov.br/Siafe/faces/execucao/financeira/ordemBancariaOrcamentariaCad.jsp' 
    tipoOB = "Orcamentaria"

while len(erros) > 0:
    erro = erros[0]   
    print(erro["processo"])
    try:
        nav.get(erro["link"])
    except:
        break
    try:
        extrairOB(erro["processo"],erro["tipo"])
        erros.pop(0)
    except:
        try:
            nav.find_element(By.XPATH,"//*[contains(text(), 'Ocorreu um erro interno.')]")
            nav.quit()
            nav = webdriver.Firefox()
            loginSIAFE(nav,os.environ['cpf'],os.environ['senha_siafe'])
            continue
        except:
            pass
        traceback.print_exc()
        time.sleep(2) 
        
    

nav.quit()
