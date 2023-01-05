from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime
import os
import numpy as np
import pyautogui
import glob
import os.path
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from sys import exit

link1 = "http://192.168.3.141/"
nav = webdriver.Chrome()
time.sleep(2)
nav.get(link1)

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

#mudando iframe
iframe1 = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
nav.switch_to.frame(iframe1)

#preenchendo campos
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

time.sleep(2)

#classe
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[3]/div/input'))).send_keys("Produção por máquina")
time.sleep(2)

#pessoa
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[8]/div/div'))).click()
time.sleep(2)
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[8]/div/div'))).click()
time.sleep(2)
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[8]/div/input'))).send_keys('4357')

WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/div'))).click()
time.sleep(2)
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[10]/div/input'))).send_keys('444698')

WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[12]/div/div'))).click()
time.sleep(2)
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[12]/div/input'))).send_keys('S Usinagem')

WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[18]/div/div'))).click()
time.sleep(2)
WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[18]/div/div'))).click()

time.sleep(2)
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[18]/div/input'))).send_keys('82')

time.sleep(2)
WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[18]/div/input'))).send_keys(Keys.INSERT)

WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
time.sleep(2)

# volta p janela principal (fora do iframe)
nav.switch_to.default_content()

try:
    texto_erro = nav.find_element("xpath", '/html/body/div[10]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]').text
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))).click()
    exit(0)
except:
    print("Passou")  