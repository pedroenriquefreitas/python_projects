from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from tabulate import tabulate
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

def check_midia_month(): #função que retorna qual o mês que a midia em questão foi postada
    m_time = driver.find_element_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[2]/a/time')
    midia_t = m_time.get_attribute("datetime")[:7]
    return midia_t[-2:]

def next_insta_midia(): #função que passa para a proxima midia no insta (aperta a seta da direita na tela)
    driver.find_element_by_xpath("/html/body/div[3]/div/div[1]/div/div/a[contains(@class, 'HBoOv')]").click()

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

i = 0

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

#checar se a midia esta no periodo analisado e guarda nessa variavel para a while
media_month = check_midia_month()

while media_month != mes: #checa se a midia em questão se insere no mês que sera analisado e passa para a proxima ate chegar na do mes que queremos analisar
    next_insta_midia()
    time.sleep(1.5)
    media_month = check_midia_month()

while check_midia_month() == mes:
    is_video = False
    i += 1
    print('Analisando a midia ' + str(i)) #log da midia em questão qeu esta sendo avaliada (comeca contato no 1)

    #se for um video, o processo para conseguir os likes é diferente
    vid_likes = 0
    if check_exists_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/span"): #entra aqui se a midia for um video
        is_video = True
        qtd_vids += 1
        pyautogui.click(798, 615)
        time.sleep(0.6)
        vid_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/div[4]/span")
        likes = vid_likes.text
        total_likes += int(likes) #adiciona a qtd de likes no total
    else: #entra aqui se a midia for uma foto (ou um não-video)
        qtd_imgs += 1
        #acha a quantidade de likes
        num_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/button/span")
        likes = num_likes.text
        total_likes += int(likes) #adiciona a qtd de likes no total

    #para o numero de comentarios, analisamos quantos sao exibidos na tela e diminuimos 1 (a legenda da foto)
    if is_video: #se é um video tem que clicar na tela para desmarcar a visualizacao dos likes
        pyautogui.click(870, 403)
    while check_exists_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[1]/ul/li[2]/button'):
        load_more_btn = driver.find_element_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[1]/ul/li[2]/button')
        load_more_btn.click()
        time.sleep(1.3)
    comments = driver.find_elements_by_class_name("gElp9")
    comments_tot = len(comments) - 1
    total_comments += comments_tot
    if is_video:
        print('Video | ' + str(likes) + ' likes | ' + str(comments_tot) + ' comentarios')
    else:
        print('Foto | ' + str(likes) + ' likes | ' + str(comments_tot) + ' comentarios')
    next_insta_midia()
    time.sleep(1.5)

#depois de tudo print os resultados gerais para o meso
print('\n\n\n')
print('Resultados do Instagram para o mês ' + mes + '\n')
print(tabulate([['Videos', qtd_vids], ['Fotos', qtd_imgs], ['Likes', total_likes], ['Comentários', total_comments]], headers=['Campo', 'Valor'], tablefmt='orgtbl'))
print('\n\n')
