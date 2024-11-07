import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from datetime import  datetime
from marinetteSEFAZ import loginSEI, obterProcessosDeBloco, escreverAnotacao, buscarInformacaoEmDocumento, procurarArquivos,incluirProcessoEmBloco,removerProcessoDoBloco
from tqdm import tqdm
import time
def verificarCompetencia():
    lista = procurarArquivos(navegador, "Despacho de Encaminhamento de Processo")
    pattern1 = r"(?i)(Para o processamento dos? DARJs?|solicitamos o processamento dos? DARJs?|Para processamento dos? DARJs?|Solicitando processar os? Darjs?)"
    pattern2 = r"(?i)(AUDITORA FISCAL|AUDITOR FISCAL|Procurador|Procuradora|Ramon S N de Souza)"

    for item in reversed(lista):
        processamento = buscarInformacaoEmDocumento(navegador,item,[pattern1],"Rio de Janeiro")
        if processamento != None:
            auditor = buscarInformacaoEmDocumento(navegador,item,[pattern2],"Rio de Janeiro")
            if auditor == None:
                return "Inválida"
            else:
                return "Ok"
            
    return "Não encontrada"


def verificarValorEValidade():
    montanteDARJ = None

    darj = procurarArquivos(navegador, "DARJ")
    if not darj:
        return "DARJ não encontrado", "DARJ não encontrado"
    
    
    regex = [r"(VENCIMENTO)\n\n([\n*\w*\s\(\)\.,-\/]*)\n\n01 ", r"(TOTAL A PAGAR)\n\n([\n*\w*\s\(\)\.,-]*)\n\n14 "]
    regexSUAR = [r"(\d\d\/\d\d\/\d\d\d\d)",r"\n([\n*\w*\s\(\)\.,-]*)\n\nTelefone"]
    for doc in reversed(darj):
        lista = buscarInformacaoEmDocumento(navegador,doc,regex,"DARJ")
        if lista:
            validade = lista[0].group(2)
            montanteDARJ = lista[1].group(2)
            break
        lista = buscarInformacaoEmDocumento(navegador,doc,regexSUAR,"SUAR")
        if lista:
            validade = lista[0].group(1)
            montanteDARJ = lista[1].group(1)
            break
        if not lista:
            return "Impossível verificar DARJ", "Impossível verificar DARJ"
    
    validadeData = datetime.strptime(validade, '%d/%m/%Y')
    dias = (validadeData - datetime.now()).days
    if dias < 0:
        validade = "Fora de validade"
    elif dias <= 15:
        validade = "Data de Validade muito próxima"
    else:
        validade = "Ok"
    
    guias = procurarArquivos(navegador, "Guia")
    if not guias:
        return validade, "Guia não encontrada"
    
    regexGuia = [r"Não\n.?(\d[\n*\w*\s\(\)\.,-\/]*,\d\d)\n\nU"]

    for guia in reversed(guias):
        montanteGuia = buscarInformacaoEmDocumento(navegador,guia,regexGuia,"Guia de Recolhimento")
        if montanteGuia != None:
            montanteGuia = montanteGuia[0].group(1)
            break
    else:
        return validade, "Impossível verificar Guia"
    
    if montanteDARJ == None:
        return "DARJ não encontrado", "DARJ não encontrado"

    if montanteDARJ != montanteGuia:
        montante  = "Montante Guia diferente de Montante DARJ"            
    else:
        montante = "Ok"
        
    return validade, montante
   


navegador = webdriver.Firefox()

blocoSolicitado = "616986"

loginSEI(navegador,os.environ['login_sefaz'],os.environ['senha_sefaz'],'SEFAZ/COOAJUR' )

processos = obterProcessosDeBloco(navegador, blocoSolicitado)

numProcessos = len(processos)
i = 1

while i != numProcessos:
            
    processo = navegador.find_elements(By.XPATH, "//tbody//tr")[i]
    if "Montante" not in processo.text:


        linkProcesso = WebDriverWait(processo,3).until(EC.presence_of_element_located((By.XPATH, './/td[3]//a')))

        nProcesso = linkProcesso.text
        linkProcesso.click()
        print(nProcesso)
        navegador.switch_to.window(navegador.window_handles[1])
            
        try:
            
            validade, montante = verificarValorEValidade()

            competencia = verificarCompetencia()
            texto = ["Validade DARJ: " + validade, "Montante: " + montante, "Competência: " + competencia]

            if validade == "Ok" and montante == "Ok" and competencia == "Ok":
                try:  
                    incluirProcessoEmBloco(navegador,nProcesso,"938324")
                except:    
                    traceback.print_exc()

        except:
            traceback.print_exc()
            continue

        finally:
            navegador.close()
            navegador.switch_to.window(navegador.window_handles[0])

        try:
            if validade == "Ok" and montante == "Ok" and competencia == "Ok":
                removerProcessoDoBloco(navegador, nProcesso)
                navegador.find_elements(By.XPATH, "//tbody//tr")
                numProcessos -= 1
            else:
                escreverAnotacao(navegador,texto,nProcesso)
                i += 1
        except:
            traceback.print_exc()
    else:
        i+=1

navegador.quit()