import time
from time import sleep
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import  Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import re
from openpyxl import load_workbook
import pyautogui
from datetime import date
from glob import glob
from shutil import move
import tabula
from PyPDF2 import PdfReader
import pandas as pd

def login(navegador, user, password):
    navegador.get("https://siafe2.fazenda.rj.gov.br/Siafe/faces/login.jsp")
    usuario = WebDriverWait(navegador,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="loginBox:itxUsuario::content"]')))
    usuario.send_keys(user)

    senha = navegador.find_element(By.XPATH, value='//*[@id="loginBox:itxSenhaAtual::content"]')
    senha.send_keys(password)

    exercicio = Select(navegador.find_element(By.XPATH, value='//*[@id="loginBox:cbxExercicio::content"]'))
    exercicio.select_by_visible_text('2024')

    btnLogin = navegador.find_element(By.XPATH, value='//*[@id="loginBox:btnConfirmar"]')
    btnLogin.click()
    navegador.maximize_window()

def verificarSeO(navegador,processo,tipoOB):
    if tipoOB == "Extra":
        index = 12
    if tipoOB == "Orcamentaria":
        index =13

    
    WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.XPATH, '// *[ @id = "pt1:tblOBExtra:sdtFilter::btn"]')))
    try:
        btnLimpar = navegador.find_element(By.XPATH, value= '//*[@id="pt1:tblOBExtra:btnClearFilter"]')
        btnLimpar.click()
    except:
        btnFiltro = navegador.find_element(By.XPATH, value='// *[ @ id = "pt1:tblOBExtra:sdtFilter::disAcr"]')
        btnFiltro.click()
        try:
            btnLimpar = navegador.find_element(By.XPATH, value='//*[@id="pt1:tblOBExtra:btnClearFilter"]')
            btnLimpar.click()
        except:
            None
    
    WebDriverWait(navegador,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="pt1:tblOBExtra:table_rtfFilter:0:in_value_rtfFilter::content"]')))

    
    acessoRapido = navegador.find_element(By.ID, value='pt1:iTxtCad::content')

    acessoRapido.click()

    popUp(navegador)

    filtroProcessoSEI = Select(navegador.find_element(By.XPATH,
                                                        value='//*[@id="pt1:tblOBExtra:table_rtfFilter:0:cbx_col_sel_rtfFilter::content"]'))
    filtroProcessoSEI.select_by_visible_text('Processo')

    filtroContemProcesso = Select(navegador.find_element(By.XPATH,
                                                            value='//*[@id="pt1:tblOBExtra:table_rtfFilter:0:cbx_op_sel_rtfFilter::content"]'))
    filtroContemProcesso.select_by_visible_text('contém')
    valorFiltroProcesso = navegador.find_element(By.XPATH,
                                                    value='//*[@id="pt1:tblOBExtra:table_rtfFilter:0:in_value_rtfFilter::content"]')    
    valorFiltroProcesso.clear()
    valorFiltroProcesso.send_keys(processo)

    acessoRapido.click()
    
    WebDriverWait(navegador,15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pt1:tblOBExtra:table_rtfFilter:1:cbx_col_sel_rtfFilter::content"]')))
    time.sleep(3)
    
    try:
        WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[text()= '"+ processo + "']")))
        return True
    except:
        return False


def preencherTabelaPrazos():
    planilha = load_workbook(r"C:\Users\jcampbell1\Downloads\Planilha Gerencial - Marinette.xlsx")
    prazos = planilha["PRAZOS"]


    prazos.delete_rows(2,prazos.max_row)
    tabelas = planilha.sheetnames


    tabela = planilha["895785"]
    x= 1

    celulasComPrazo = []

    for tabela in tabelas:
        tabela = planilha[tabela]
        x = 1
        for linha in tabela:
            for cell in linha:
                print(cell.value)
            if tabela[f"H{x}"].value =="ERRO" or tabela[f"H{x}"].value == None or tabela[f"H{x}"].value == "OB não encontrada!": 
                try:
                    linhaAtual = []
                    prazo = tabela[f"D{x}"].value
                    
                    if ("PRAZO") not in prazo:
                        for cell in linha:
                            if cell.value != None:
                                celula = cell.value
                                linhaAtual.append(celula)
                                
                        linhaAtual.append(tabela.title) 
                        celulasComPrazo.append(linhaAtual)
                except:
                    traceback.print_exc
                x += 1
        
    for linha in celulasComPrazo:
        try:
            numero = prazos.max_row + 1
            linha[3] = re.sub(r"\d+", str(numero), linha[3])    
            prazos.append(linha)
            planilha.save(marinette)
        except:
            traceback.print_exc()


    planilha._sheets.remove(prazos)
    planilha._sheets.insert(0,prazos)
    planilha.save(marinette)

marinette = r"C:\Users\jcampbell1\Downloads\Planilha Gerencial - Marinette.xlsx"
planilha = load_workbook(marinette)
#tabelas = 
for tabela in planilha:
    tabela = pd.read_excel(marinette, sheet_name= tabela)
    
    for processo in tabela:
        print(processo)



#preencherTabelaPrazos()