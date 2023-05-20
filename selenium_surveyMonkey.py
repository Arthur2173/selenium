from openpyxl import load_workbook
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as pg

nomeCaminhoArquivo = r'C:\Users\Ana Paula\Desktop\selenium\DadosFormulario.xlsx'
planilha_aberta = load_workbook(filename=nomeCaminhoArquivo)

#Seleciona a sheet que tem as informações a serem passadas para o formulario on-line
sheet_selecionada = planilha_aberta['Dados']


for linha in range(2, len(sheet_selecionada['A']) + 1):
    
    nome = sheet_selecionada['A%s' % linha].value
    email = sheet_selecionada['B%s' % linha].value
    telefone = sheet_selecionada['C%s' % linha].value
    sexo = sheet_selecionada['D%s' % linha].value
    sobre = sheet_selecionada['E%s' % linha].value
    
    #Aguardar para o computador processar as informações
    pg.sleep(2)
    browser = webdriver.Chrome()
    browser.get("https://pt.surveymonkey.com/r/Y9Y6FFR")

    #Aguardar para o computador processar as informações
    pg.sleep(6)

    #Preenche Nome
    browser.find_element(By.NAME,"683928983").send_keys(nome)

    #Preenche Email
    browser.find_element(By.NAME,"683932318").send_keys(email)

    #Preenche Telefone
    browser.find_element(By.NAME,"683930688").send_keys(telefone)

    #Preenche Sobre
    browser.find_element(By.NAME,"683932969").send_keys(sobre)

    if sexo == "Masculino":
        
        #Preenche Radio Button Feminino
        browser.find_element(By.ID,"683931881_4497366118_label").click()
        
        
    
    else:
        
        #Preenche Radio Button Feminino
        browser.find_element(By.ID,"683931881_4497366119_label").click()

    #Aguardar para o computador processar as informações
    pg.sleep(6)
    browser.find_element(By.XPATH,'//*[@id="patas"]/main/article/section/form/div[2]/button').click()
    print(sexo)