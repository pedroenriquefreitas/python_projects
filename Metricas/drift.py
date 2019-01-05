from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from tabulate import tabulate
import time
import pyautogui
import openpyxl

print('Começando a parte do Chat')

def mes_dia_extenso(no):
    month_list = ['null', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    return month_list[no]

print('Para começar, preciso saber qual o numero do mês que será avaliado (Jan-1, Fev-2, ...)')
mes_drift = input('Qual o mês que sera avaliado?\n')
mes_drift = int(mes_drift)

driver = webdriver.Chrome('/Users/pedroenriqueandrade/Desktop/python_projects/drivers/chromedriver')
time.sleep(1)
driver.get("https://start.drift.com/login")
time.sleep(1)

log_host = driver.find_element_by_xpath('//*[@id="username"]')
log_host.send_keys('faleconosco@empresajunior.com.br')

driver.find_element_by_xpath('//*[@id="root"]/div/div/div[1]/div/div[1]/form/button').click()

time.sleep(1.8)

pswd_host = driver.find_element_by_xpath('//*[@id="password"]')
pswd_host.send_keys('gestorcomercial')


driver.find_element_by_xpath('//*[@id="root"]/div/div/div[1]/div/div/div/form/button').click()

time.sleep(1)

driver.get("https://app.drift.com/inboxes")

time.sleep(1.5)

conversas_chat = driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div')

p = 0
for a in conversas_chat: #antes a lista conversas_chat continha apeans WebElements, nesse 'for' a gente faz com que ela agora seja uma lista de strings com as datas de cada conversa
    conversas_chat[p] = a.get_attribute('title')
    p += 1

print(conversas_chat)

while not any(mes_dia_extenso(mes_drift) in s for s in conversas_chat):
    print('Curva #1')
    driver.find_element_by_xpath('//*[@id="conversation-list"]/div/div/button').click()
    conversas_chat = driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div')
    p = 0
    for a in conversas_chat: #antes a lista conversas_chat continha apeans WebElements, nesse 'for' a gente faz com que ela agora seja uma lista de strings com as datas de cada conversa
        conversas_chat[p] = a.get_attribute('title')
        p += 1

while mes_dia_extenso(mes_drift) in conversas_chat[len(conversas_chat)-1]: #se o ultimo elemento da lista ainda é do mes desejado, devemos carregar mais pra ver se ainda ter algum chat sobrando
    print('Curva #2')
    driver.find_element_by_xpath('//*[@id="conversation-list"]/div/div/button').click()
    conversas_chat = driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div')
    p = 0
    for a in conversas_chat: #antes a lista conversas_chat continha apeans WebElements, nesse 'for' a gente faz com que ela agora seja uma lista de strings com as datas de cada conversa
        conversas_chat[p] = a.get_attribute('title')
        p += 1

matching = [s for s in conversas_chat if mes_dia_extenso(mes_drift) in s] #lista com os chats do periodos analisado

#agora colocar na planilha do excel
workbook = openpyxl.load_workbook('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/template.xlsx')
sheet = workbook.get_sheet_by_name('Sheet1')
print('Planilha aberta com sucesso')
sheet['D12'].value = len(matching)
print('Valores inputados com sucesso com sucesso')
workbook.save('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/tempfsdgflate.xlsx')
print('Planilha salva com sucesso')
