from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
import time
import pyautogui
#import openpyxl

#check if a certain xpath exists
def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

#variavel que armazena a quantidade total de likes
total_likes = 0

#variavel que armazena a quantidade total de comentarios
total_comments = 0

#variavel que armazena a quantidade de videos postados
qtd_vids = 0

#variavel que armazena a quantidade de imagens postadas
qtd_imgs = 0

#variavel que fala se é video ou nao
is_video = False

#precisamos saber qual o mes em questão que esta sendo avaliado
print('Para começar, preciso saber qual o numero do mês que será avaliado (Jan-1, Fev-2, ...)')
mes = input('Qual o mês que sera avaliado?\n')

driver = webdriver.Chrome('/Users/pedroenriqueandrade/Desktop/python_projects/drivers/chromedriver')
time.sleep(1)
driver.get("https://instagram.com/empresajunior")
time.sleep(1)

#fecha a caixa do insta que pedi login
tyu = driver.find_element_by_xpath("//*[@id='react-root']/section/nav/div[2]/div/div/div[3]/div/div/section/div/button")
tyu.click()

#clica na primeira foto
time.sleep(0.6)
pyautogui.click(294, 757)
time.sleep(3)

#checar se a midia esta no periodo analisado
m_time = driver.find_element_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[2]/a/time')
midia_time = m_time.get_attribute("datetime")[:7]
midia_time = midia_time[-2:]

while midia_time != mes: #checa se a midia em questão se insere no mês que sera analisado
    driver.find_element_by_xpath("/html/body/div[3]/div/div[1]/div/div/a[contains(@class, 'HBoOv')]").click() #passa para a proxima midia
    print('Passa para a proxima midia')
    time.sleep(1.5)
    m_time = driver.find_element_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[2]/a/time')
    midia_time = m_time.get_attribute("datetime")[:7]
    midia_time = midia_time[-2:]

#se for um video, o processo para conseguir os likes é diferente
vid_likes = 0
if check_exists_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/span"):
    is_video = True
    pyautogui.click(798, 615)
    time.sleep(0.6)
    vid_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/div[4]/span")
    likes = vid_likes.text
else:
    #acha a quantidade de likes
    num_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/button/span")
    likes = num_likes.text

#para o numero de comentarios, analisamos quantos sao exibidos na tela e diminuimos 1 (a legenda da foto)
if is_video: #se é um video tem que clicar na tela para desmarcar a visualizacao dos likes
    pyautogui.click(870, 403)
while check_exists_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[1]/ul/li[2]/button'):
    load_more_btn = driver.find_element_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[1]/ul/li[2]/button')
    load_more_btn.click()
    time.sleep(1)
comments = driver.find_elements_by_class_name("gElp9")
comments_tot = len(comments) - 1

if is_video:
    print('O elemento é um video')
else:
    print('O elemento é uma foto estática')
print('Quantidade de Likes:')
print(likes)
print('Quantidade de Comentários:')
print(comments_tot)
