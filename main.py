from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from datetime import date
import pandas as pd
import win32com.client as win32
import time

def sendEmailWithRefreshSheet():

    hoje = date.today().strftime("%d/%m/%Y")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = 'pedrodopython@outlook.com'
    mail.Subject = 'Atualização Dos Preços'
    mail.Body = '''
    Olá, tudo bem ? Boa tarde ! 
    Segue acima em anexo os valores das novas planilhas de acordo com a cotação atual das moedas
    Planilha Atualizada No dia {}
    
    '''.format(hoje)

    arquivo = r'C:\Users\pedro sousa\Documents\Repositório Github\Back-end\Projetos Python\Cotação Tempo Real\Produtos Atualizados.xlsx'
    mail.Attachments.Add(arquivo)
    mail.Send()


def atualizaPlanilhaProdutos(dolar_atualizado, euro_atualizado, yuan_atualizado, ouro_atualizado_formatado):
    tables = pd.read_excel('./Produtos.xlsx')

    tables.loc[tables["Moeda"] == 'Dólar', 'Cotação'] == float(dolar_atualizado)
    tables.loc[tables["Moeda"] == 'Euro', 'Cotação'] == float(euro_atualizado)
    tables.loc[tables["Moeda"] == 'Yuan', 'Cotação'] == float(yuan_atualizado)
    tables.loc[tables["Moeda"] == 'Ouro', 'Cotação'] == float(ouro_atualizado_formatado)
    
    tables['Preço de Compra'] = tables['Cotação'] * tables['Preço Original']
    tables['Preço de Venda'] = tables['Preço de Compra'] * tables['Margem']

    tables.to_excel('Produtos Atualizados.xlsx', index=False)
    sendEmailWithRefreshSheet()
      
while True:
    driver = webdriver.Chrome()
    driver.get("http://www.google.com")

    # pegando a cotacao do dolar
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').click()
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dolar")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
    dolar_atualizado = driver.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

    # pegando a cotacao do euro
    driver.get("http://www.google.com")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').click()
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
    euro_atualizado = driver.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

    # pegando a cotacao do Yuan
    driver.get("http://www.google.com")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').click()
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação Yuan")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
    yuan_atualizado = driver.find_element("xpath", '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

    # pegando a cotacao do ouro
    driver.get("http://www.google.com")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').click()
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação ouro")
    driver.find_element("xpath", '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
    driver.find_element("xpath", '//*[@id="rso"]/div[1]/div/div/div/div[1]/div/a/h3').click()
    ouro_atualizado = driver.find_element("xpath", '//*[@id="comercial"]').get_attribute('value')
    ouro_atualizado_formatado = ouro_atualizado.replace(",", ".")

    atualizaPlanilhaProdutos(dolar_atualizado, euro_atualizado, yuan_atualizado, ouro_atualizado_formatado)
    time.sleep(60 * 60)