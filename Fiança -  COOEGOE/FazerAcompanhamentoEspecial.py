import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
from tqdm import tqdm
from marinetteSEFAZ import loginSEI, obterProcessosDeBloco, procurarArquivos, buscarInformacaoEmDocumento,buscarProcessoEmBloco,escreverAcompanhamentoEspecial

bloco = input("Digite o número do bloco: ")
    
grupo = int(input("Digite o grupo do acompanhamento fiscal\n\n 1- FIANÇA E VALOR APREENDIDO\n 2- EXECUÇÃO FISCAL\n 3- CAUÇÃO\n\n"))

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
    try:
        processo = nav.find_elements(By.XPATH, "//tbody//tr")[i]
        linkProcesso = buscarProcessoEmBloco(nav,i)    
        nProcesso = linkProcesso.text
        print(nProcesso)
        
        linkProcesso.click()
        nav.switch_to.window(nav.window_handles[1])
    

        despachos = procurarArquivos(nav, "Despacho sobre Autorização de Despesa")    
        
        if grupo == "FIANÇA E VALOR APREENDIDO":
            regexFiança = r"(Beneficiário|Credor): ?(.*)\nForma de Pagamento: ?(.*)(\(\d*\))?"
            regexValores = r"(R\$ ?[\d.]+,\d\d)\b[\s\S]*?(R\$ ?[\d.]+,\d\d)?\b"
            texto = []
            for despacho in despachos:
                
                resultado = buscarInformacaoEmDocumento(nav,despacho,[regexFiança,regexValores],"Estado")
                texto.append("Beneficiário: " + resultado[0].group(2).strip())
                texto.append("Forma de Pagamento: " + resultado[0].group(3).strip())
                texto.append("Valor Extraorçamentário: " +resultado[1].group(1))
                try:
                    texto.append("Valor Orçamentário: " +resultado[1].group(2))
                except:
                      pass  
        if grupo == "EXECUÇÃO FISCAL":
            
            darjs = procurarArquivos(nav, "DARJ")
                
            regexExecucao = r"(CDA \d{4}\/\d{3}\.\d{3}\-\d ?|CDA \d*)[\s\S]*?no valor de R\$ ?([^(]*)"
            resultado = buscarInformacaoEmDocumento(nav,despachos[-1],regexExecucao,"Coordenador")
            
            cda = resultado.group(1).strip()
            valor = resultado.group(2).strip()
                
            regexDARJ = r"\bNOME\n?\n?([\n*\w*\s\(\)\.,-\/]*)08 - CNPJ\/CPF\n?\n?([\n\d\.\-\/]*)02 - ENDEREÇO COMPLETO"
            for darj in darjs:
                resultado = buscarInformacaoEmDocumento(nav,darjs[-1],regexDARJ,"ESTADO")
                if resultado:
                    nome = resultado.group(1).replace('\n',"")
                    cpf = resultado.group(2).replace('\n',"")
                    break
            else:
                documento = procurarArquivos(nav,["Documento","Petição","Certidão"])[0]
                resultado = buscarInformacaoEmDocumento(nav,documento,"(CNPJ|CPF) ?:\n? ?([\n\d\.\-\/]*)\n?\n?Nome:\n? ?(.*)\n?\n?","PROCURADORIA")
                if resultado:
                    nome = resultado.group(3).replace('\n',"")
                    cpf = resultado.group(2).replace('\n',"")
                else:
                    nome = ""
                    cpf = ""
                    print("NOME E CPF NÃO ENCONTRADOS")
            texto = [cda,"Nome: " + nome, "CPF-CNPJ: " + cpf, "Valor: R$ " + valor]
                            
        if grupo == "CAUÇÃO":
            pass
        
        escreverAcompanhamentoEspecial(nav,nProcesso,texto,grupo)   
    except:
        traceback.print_exc()
    finally:
        nav.close()
        nav.switch_to.window(nav.window_handles[0])


nav.quit()
