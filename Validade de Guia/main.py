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
from selenium.webdriver.common.action_chains import ActionChains
import os
import sys


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

def encontrarProcessos(navegador,blocoSolicitado):
    navegador.find_element(By.XPATH, "//span[text() = 'Blocos']").click()
    WebDriverWait(navegador,20).until(EC.element_to_be_clickable((By.XPATH, "//span[text() = 'Internos']"))).click()
    blocos = navegador.find_elements(By.XPATH, "//tbody//tr")[1:-1]

    for bloco in blocos:    
        nBloco = bloco.find_elements(By.XPATH,".//td")[1]
        if nBloco.text == blocoSolicitado:
            nBloco.find_element(By.XPATH, './/a').click()
            break
    processos = navegador.find_elements(By.XPATH, "//tbody//tr")[1:]

    for i in range(1,len(processos)):
        processo = navegador.find_elements(By.XPATH, "//tbody//tr")[i]
        WebDriverWait(processo,20).until(EC.element_to_be_clickable(((By.XPATH, './/td[3]//a'))))
        processo.find_element(By.XPATH, './/td[3]//a').click()
        formaDePagamento = encontrarFormaDePagamento(navegador)          
        anotarFormaDePagamento(processo, formaDePagamento,navegador)
        
def encontrarFormaDePagamento(navegador):

    navegador.switch_to.window(navegador.window_handles[1])
    WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore")))
    docs = navegador.find_elements(By.XPATH, "//div[@id = 'divArvore']//div//a[@class = 'infraArvoreNo']")

    for doc in docs:
        docTexto = doc.text
        if "Despacho sobre Autorização de Despesa" in docTexto:
            
            
            doc.click()
            time.sleep(2)
            
            navegador.switch_to.default_content()            
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
            WebDriverWait(navegador,20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvoreHtml")))
            try:
                beneficiario = navegador.find_element(By.XPATH, "//p[@class = 'Tabela_Texto_Alinhado_Esquerda' ][4]" )
                print(beneficiario.text)
                if  "CPF" in beneficiario.text:
                    formaPagamento = "Depósito"
                    navegador.close()
                    navegador.switch_to.window(navegador.window_handles[0])
                    
                    
                    return formaPagamento

                if "CNPJ" in beneficiario.text:
            
                    formaPagamentoDespacho = navegador.find_element(By.XPATH, "//p[@class = 'Tabela_Texto_Alinhado_Esquerda' ][5]" )
                    print(formaPagamentoDespacho.text)
                    
                    if "GUIA" in formaPagamentoDespacho.text:
                        formaPagamento = "Guia"
                    if "DEPÓSITO JUDICIAL" in formaPagamentoDespacho.text:
                        formaPagamento = "Guia"
                    elif "DEPÓSITO" in formaPagamentoDespacho.text:
                        formaPagamento = "Depósito"


                    navegador.close()
                    navegador.switch_to.window(navegador.window_handles[0])


                    return formaPagamento
            

                return ""
            except:
                traceback.print_exc()

def anotarFormaDePagamento(processo,formaPagamento,navegador):

    processo.find_element(By.XPATH,".//td//a//img[@title='Anotações']").click()
                        
    time.sleep(2)
    try:
        iframe = navegador.find_element(By.TAG_NAME, 'iframe')
        navegador.switch_to.frame(iframe)

        txtarea = navegador.find_element(By.XPATH, '//textarea[@id = "txtAnotacao"]')

        txtarea.send_keys(Keys.PAGE_DOWN)
        txtarea.send_keys(Keys.END)
        txtarea.send_keys(Keys.ENTER)
        txtarea.send_keys("Forma de Pagamento: " + formaPagamento)

        time.sleep(1)
        salvar = navegador.find_element(By.XPATH, '//button[@value = "Salvar"]')

        salvar.click()
        
    except:
       traceback.print_exc()
       time.sleep(1)
       navegador.find_element(By.XPATH, "//div[@class = 'sparkling-modal-close']").click()
    finally:
        navegador.switch_to.default_content()


processo = input("Digite o número do bloco: ")

navegador = webdriver.Edge()
login(navegador)

encontrarProcessos(navegador,processo)
