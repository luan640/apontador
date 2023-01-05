from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime
import os
import numpy as np
import glob
import os.path
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from sys import exit
import gspread
from selenium.webdriver.common.action_chains import ActionChains

#link1 = "http://192.168.3.141/"
link1 = 'http://devcemag.innovaro.com.br:81/sistema'
nav = webdriver.Chrome()
time.sleep(2)
nav.get(link1)

########### LOGIN ###########

#logando 
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("luan araujo")
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("luanaraujo")

time.sleep(2)

WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

time.sleep(2)

#abrindo menu

WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()

time.sleep(2)

#clicando em produção
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[2]/div[7]/span[2]'))).click()

time.sleep(2)

#clicando em SFC
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[2]/div[12]/span[2]'))).click()

time.sleep(2)

#clicando em apontamento
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[2]/div[14]/span[2]'))).click()

time.sleep(3)

########### ACESSANDO PLANILHAS ###########

sheet = 'Apontamento automático'
worksheet1 = 'Usinagem'
worksheet2 = 'Estamparia'
worksheet3 = 'Apontadas'

filename = "service_account.json"

sa = gspread.service_account(filename)
sh = sa.open(sheet)

wks1 = sh.worksheet(worksheet1)
wks2 = sh.worksheet(worksheet2)
wks3 = sh.worksheet(worksheet3)

usinagem = wks1.get_all_records()
estamparia = wks2.get_all_records()
apontadas = wks3.get_all_records()

usinagem = pd.DataFrame(usinagem)
estamparia = pd.DataFrame(estamparia)
apontadas = pd.DataFrame(apontadas)

########### Tratando planilhas ###########

#ajustando datas
for i in range(len(usinagem)):
    try:
        if usinagem['Data'][i] == "":
            usinagem['Data'][i] = usinagem['Data'][i-1]
    except:
        pass

i = None

#filtrando pecas que não foram apontadas
usinagem_filter = usinagem.loc[(usinagem.Status) == '']

#inserindo 0 antes do código da peca
usinagem_filter['Peca'] = usinagem_filter['Peca'].astype(str)

for i in range(len(usinagem)):
    try:
        if len(usinagem_filter['Peca'][i]) == 5:
            usinagem_filter['Peca'][i] = "0" + usinagem_filter['Peca'][i] 
    except:
        pass

i = None

#data de hoje
data_realizado = datetime.now()
ts = pd.Timestamp(data_realizado)
data_realizado = data_realizado.strftime('%d/%m/%Y')

########### PREENCHIMENTO ###########

def preenchimento(peca,data,qtde,i,wks1,pessoa,c):

    #mudando iframe
    iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe1)

    #Insert
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
    time.sleep(2)

    #classe
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[3]/div/input"))).send_keys("Produção por máquina")
    time.sleep(2)

    #data
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/div"))).click()
    time.sleep(2)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[5]/div/input"))).send_keys(data)

    #pessoa
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/div"))).click()
    time.sleep(2)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[8]/div/input"))).send_keys(pessoa)

    #peça
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/div"))).click()
    time.sleep(2)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[10]/div/input"))).send_keys(peca)
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/div"))).click()
    time.sleep(5)

    #processo
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[12]/div/input"))).send_keys('S Usinagem')
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/div"))).click()
    time.sleep(2)

    #quantidade
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/div"))).click()
    time.sleep(2)
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[" + str(c) + "]/td[18]/div/input"))).send_keys(qtde)
    time.sleep(5)

    ActionChains(nav).key_down(Keys.CONTROL).send_keys('m').key_up(Keys.CONTROL).perform()
    time.sleep(2)
    ActionChains(nav).key_down(Keys.CONTROL).send_keys('m').key_up(Keys.CONTROL).perform()
    time.sleep(2)
    
    #innovaro real
    #WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
    #time.sleep(2)

    time.sleep(2)

    # volta p janela principal (fora do iframe)
    nav.switch_to.default_content()

    try:
        texto_erro = nav.find_element("xpath", '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]').text
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
        wks1.update('D' + str(i+2), texto_erro)
    except:
        wks1.update('D' + str(i+2), 'Apontada!')


c = 1

for i in range(0,3):
    try:
        peca = usinagem_filter['Peca'][i]
        qtde = str(usinagem_filter['Qtde'][i])
        data = usinagem_filter['Data'][i]
        pessoa = '4054'
        c = c+2
        preenchimento(peca,data,qtde,i,wks1,pessoa,c)

    except:
        pass