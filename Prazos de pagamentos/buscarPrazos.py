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
            planilha.save(r"C:\Users\jcampbell1\Downloads\Planilha Gerencial - Marinette.xlsx")
        except:
            traceback.print_exc()


    planilha._sheets.remove(prazos)
    planilha._sheets.insert(0,prazos)
    planilha.save(r"C:\Users\jcampbell1\Downloads\Planilha Gerencial - Marinette.xlsx")





preencherTabelaPrazos()