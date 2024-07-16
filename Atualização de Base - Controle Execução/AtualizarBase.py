import os
import time
import re
import traceback
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from shutil import move
import pandas as pd
import logging

from tqdm import tqdm

def login(navegador):
    navegador.get("https://siafe2.fazenda.rj.gov.br/Siafe/faces/login.jsp")

    user = os.environ['cpf']
    password = os.environ['senha_sefaz']

    usuario = navegador.find_element(By.XPATH, value='//*[@id="loginBox:itxUsuario::content"]')
    usuario.send_keys(user)

    senha = navegador.find_element(By.XPATH, value='//*[@id="loginBox:itxSenhaAtual::content"]')
    senha.send_keys(password)
    
    btnLogin = navegador.find_element(By.XPATH, value='//*[@id="loginBox:btnConfirmar"]')
    btnLogin.click()

    try:
        WebDriverWait(navegador,2).until(EC.element_to_be_clickable((By.XPATH, "//a[@class = 'x12k']"))).click()        
    except:
        pass

    navegador.maximize_window()
    navegador.get("https://siafe2.fazenda.rj.gov.br/Siafe/faces/flexvision/flexvisionMain.jsp")
    WebDriverWait(navegador,180).until(EC.presence_of_element_located((By.XPATH,"//span[text() = 'Consultas de outros usuários']")))
    

nav = webdriver.Firefox()

login(nav)
nav.quit()