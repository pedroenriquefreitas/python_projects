from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pyautogui
import openpyxl

#variavel que armazena a quantidade total de likes
total_likes = 0

#variavel que armazena a quantidade total de comentarios
total_comments = 0

#variavel que armazena a quantidade de videos postados
qtd_vids = 0

#variavel que armazena a quantidade de imagens postadas
qtd_imgs = 0

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

#se for um video, o processo para conseguir os likes é diferente
vid = ''
vid_likes = 0
vid = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/span")
if 'views' in vid.text:
    print('yaaaaaaaa')
    pyautogui.click(798, 615)
    time.sleep(0.6)
    vid_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/div[4]/span")
    likes = vid_likes.text
else:
    #acha a quantidade de likes
    num_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/div[4]/span")
    likes = num_likes.text

print('Quantidade de Likes:')
print(likes, '\n')
