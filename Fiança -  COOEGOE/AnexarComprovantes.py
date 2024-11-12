from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import glob
from marinetteSEFAZ import loginSEI, obterProcessosDeBloco,buscarProcessoEmBloco,escreverAnotacao,incluirDocumentoExterno
import re
import traceback
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
        if arquivosProcesso:
            linkProcesso.click()
            nav.switch_to.window(nav.window_handles[1])
            print(arquivosProcesso)
        
            for arquivo in arquivosProcesso:
                try:
                    ob = re.search(r"(\d{4}OB\d{5})",arquivo).group(1)  
                    incluirDocumentoExterno(nav,"Comprovante",arquivo,nome=ob)
                except:
                    traceback.print_exc()
                    continue
            nav.close()
            nav.switch_to.window(nav.window_handles[0])
            escreverAnotacao(nav,"Comprovantes Ok", nProcesso)
              
nav.quit()
    
    
