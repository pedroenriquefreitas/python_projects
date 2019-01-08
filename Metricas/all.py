from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from tabulate import tabulate
import time
import pyautogui
import openpyxl
import csv
import sys
import random

typing_speed = 10 #wpm
def slow_type(t):
    for l in t:
        sys.stdout.write(l)
        sys.stdout.flush()
        time.sleep(random.random()*10.0/typing_speed)
    print ('')

print('\nAntes de começar, precisamos fazer algumas perguntas')
print('E assim, todo o resto do trabalho você deixa comigo')

print('Para começar, preciso saber qual o numero do mês que será avaliado (Jan-1, Fev-2, ...)')
mes = input('Qual o mês que será avaliado ?\n')
fllwrs = int(input('Quantos seguidores a página tinha no final do mês ' + mes + '?\n'))
print('Também preciso saber quanto foi investido em AdWords no mês ' + mes + ' (coloque no formato xx.xx)')
ad_mney = input('Quanto foi investido em AdWords ?\n')
ad_mney = round(float(ad_mney), 2) #arrendodar já que é quantia em dinheiro e só tem 2 casas decimais
print('Agora preciso saber quanto foi investido em propagandas no Facebook/Instagram no mês ' + mes + ' (coloque no formato xx.xx)')
fbinsta_mney = input('Quanto foi investido em Facebook/Instagram ?\n')
fbinsta_mney = round(float(fbinsta_mney), 2)
print('Tudo certo, o resto você deixa comigo')

slow_type('....')

# BEGINNING OF GOOGLE PART ########################################################
print('Começando a parte do Google')

#pesquisas do termo no mes
pesq_emp = 0

pesquisas_diretas = 0

pesquisas_descoberta = 0

tot_visualizacs_emp = 0

#precisamos saber quanto foi investido em adwords no mes em questão

with open('/Users/pedroenriqueandrade/Desktop/g1.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)

    z = 0 #criado apenas para no for pular as duas primeiras partes
    for line in csv_reader: #guardar numero de novos likes na pagina
        if (z == 0 or z == 1): #as duas primeiras partes da line sao os headers
            z += 1
            continue
        pesq_emp = int(line[4])
        pesquisas_diretas = int(line[5])
        pesquisas_descoberta = int(line[6])
        tot_visualizacs_emp = int(line[7])
        z += 1

print('Parte do Google Finalizada')

# END OF GOOGLE PART ########################################################

# BEGINNING OF FACEBOOK PART ########################################################

print('Começando a parte do Facebook')
#variavel que armazena o numero de curtidas que a pagina tem no ultimo dia do mês
curtidas = 0

#variavel que armazena o numero de novas curtidas que a pagina recebeu
novas_curtidas = 0

#variavel que armazena o numero de visualizações que a pagina recebeu
visualizcs = 0

#variavel que armazena o alcance da pagina no mês
alcance = 0

#alcance organico
alcance_org = 0

#alcance pago
alcance_pago = 0

#alcance viral
alcance_viral = 0

envolvimento = 0

qtd_posts_mes = 0

qtd_compart_mes = 0

qtd_comentarios_mes = 0

qtd_reacoes_mes = 0

qtd_interacoes_mes = 0

qtd_links_mes = 0
alcance_links = 0
envolvimento_links = 0
consumo_links = 0

qtd_photos_mes = 0
alcance_photos = 0
envolvimento_photos = 0
consumo_photos = 0

qtd_videos_mes = 0
alcance_videos = 0
envolvimento_videos = 0
consumo_videos = 0

with open('/Users/pedroenriqueandrade/Desktop/f1.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)

    i = 0 #criado apenas para no for pular as duas primeiras partes
    for line in csv_reader: #guardar numero de novos likes na pagina
        if (i == 0 or i == 1): #as duas primeiras partes da line sao os headers
            i += 1
            continue
        curtidas = int(line[1])
        if line[2] != '':
            novas_curtidas += int(line[2])
        if line[3] != '':
            visualizcs += int(line[3])
        if line[4] != '':
            alcance += int(line[4])
        if line[5] != '':
            alcance_pago += int(line[5])
        if line[6] != '':
            alcance_org += int(line[6])
        if line[7] != '':
            alcance_viral += int(line[7])
        if line[8] != '':
            envolvimento += int(line[8])
        qtd_reacoes_mes += (int(line[9]) + int(line[10]) + int(line[11]) + int(line[12]) + int(line[13]) + int(line[14]))
        i += 1

with open('/Users/pedroenriqueandrade/Desktop/f2.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)

    y = 0
    for line in csv_reader:
        if (y == 0 or y == 1): #as duas primeiras partes da line sao os headers
            y += 1
            continue
        qtd_posts_mes += 1 #como aqui é apenas a quantidade de posts, a qtd de vezes que o for rodar representa o numero de postagens
        if (line[3] == 'Photo'):
            qtd_photos_mes += 1
            alcance_photos += int(line[8])
            envolvimento_photos += int(line[9])
            consumo_photos += int(line[10])
        elif (line[3] == 'Video'):
            qtd_videos_mes += 1
            alcance_videos += int(line[8])
            envolvimento_videos += int(line[9])
            consumo_videos += int(line[10])
        else:
            qtd_links_mes += 1
            alcance_links += int(line[8])
            envolvimento_links += int(line[9])
            consumo_links += int(line[10])
        if line[12] != '':
            qtd_comentarios_mes += int(line[12])
        if line[13] != '':
            qtd_compart_mes += int(line[13])
        y += 1

qtd_interacoes_mes = envolvimento_links + envolvimento_photos + envolvimento_videos
print('Parte do Facebook Finalizada')
# END OF FACEBOOK PART ########################################################

# BEGINNING OF INSTAGRAM PART ########################################################
print('Começando a parte do Instagram')
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

#o mês que sera analizado vai ser reservado na memoria no inicio do codigo
#na parte do instagram ele é interpretado como string

#o numero de seguidores é armazenado no comeco do codigo nas perguntas iniciais\

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

#variavel que armazena a quantidade de visualizações dos videos postados
vid_vws = 0

i = 0

driver = webdriver.Chrome('/Users/pedroenriqueandrade/Desktop/python_projects/drivers/chromedriver')
time.sleep(1)
driver.get("https://instagram.com/empresajunior")
time.sleep(1)

#fecha a caixa do insta que pede login
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
    time.sleep(1.9)
    media_month = check_midia_month()

while check_midia_month() == mes:
    is_video = False
    i += 1
    #se for um video, o processo para conseguir os likes é diferente
    vid_likes = 0
    if check_exists_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/span"): #entra aqui se a midia for um video
        is_video = True
        qtd_vids += 1
        driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/span").click() #clica no numero de views para revelar o numero de likes (so acontece com videos)
        time.sleep(0.6) #espera um tempinho para o react carregar o numero de likes
        views_temp = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/span/span")
        views_temp = views_temp.text
        vid_vws += int(views_temp)
        vid_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/div[4]/span")
        likes = vid_likes.text
        total_likes += int(likes) #adiciona a qtd de likes no total
        pyautogui.click(870, 403) #se é um video tem que clicar na tela para desmarcar a visualizacao dos likes
    else: #entra aqui se a midia for uma foto (ou um não-video)
        qtd_imgs += 1
        #acha a quantidade de likes
        num_likes = driver.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/article/div[2]/section[2]/div/div/button/span")
        likes = num_likes.text
        total_likes += int(likes) #adiciona a qtd de likes no total

    #para o numero de comentarios, analisamos quantos sao exibidos na tela e diminuimos 1 (a legenda da foto)
    while check_exists_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[1]/ul/li[2]/button'):
        load_more_btn = driver.find_element_by_xpath('/html/body/div[3]/div/div[2]/div/article/div[2]/div[1]/ul/li[2]/button')
        load_more_btn.click()
        time.sleep(1.8)
    comments = driver.find_elements_by_class_name("gElp9")
    comments_tot = len(comments) - 1
    total_comments += comments_tot
    next_insta_midia()
    time.sleep(1.9)
driver.close()
print('Parte do Instagram Finalizada')
# END OF INSTAGRAM PART ########################################################

# BEGINNING OF CHAT PART ########################################################

print('Começando a parte do Chat')

def mes_dia_extenso(no):
    month_list = ['null', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    return month_list[no]

mes = int(mes) #na parte do chat o mes é interpretado como int

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

time.sleep(2.5)

driver.get("https://app.drift.com/inboxes")

time.sleep(1.5)

conversas_chat = driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div')

p = 0
for a in conversas_chat: #antes a lista conversas_chat continha apeans WebElements, nesse 'for' a gente faz com que ela agora seja uma lista de strings com as datas de cada conversa
    conversas_chat[p] = a.get_attribute('title')
    p += 1

while not any(mes_dia_extenso(mes) in s for s in conversas_chat):
    time.sleep(0.4)
    driver.find_element_by_xpath('//*[@id="conversation-list"]/div/div/button').click()
    conversas_chat = driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div')
    p = 0
    for a in conversas_chat: #antes a lista conversas_chat continha apeans WebElements, nesse 'for' a gente faz com que ela agora seja uma lista de strings com as datas de cada conversa
        conversas_chat[p] = a.get_attribute('title')
        p += 1

while mes_dia_extenso(mes) in conversas_chat[len(conversas_chat)-1]: #se o ultimo elemento da lista ainda é do mes desejado, devemos carregar mais pra ver se ainda ter algum chat sobrando
    time.sleep(0.4)
    driver.find_element_by_xpath('//*[@id="conversation-list"]/div/div/button').click()
    conversas_chat = driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div')
    p = 0
    for a in conversas_chat: #antes a lista conversas_chat continha apeans WebElements, nesse 'for' a gente faz com que ela agora seja uma lista de strings com as datas de cada conversa
        conversas_chat[p] = a.get_attribute('title')
        p += 1

matching = [s for s in conversas_chat if mes_dia_extenso(mes) in s] #lista com os chats do periodos analisado
for temp in driver.find_elements_by_xpath('//*[@id="conversation-list"]/div/div/div[2]/div[1]/div[3]/div/div'):
    if (temp.get_attribute('title') == matching[len(matching)-1]):
        temp.click()

print('Parte do Chat Finalizada')
# END OF CHAT PART ########################################################

#agora colocar na planilha do excel
workbook = openpyxl.load_workbook('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/template.xlsx')
sheet = workbook.get_sheet_by_name('Sheet1')
print('Planilha aberta com sucesso')
sheet['D12'].value = len(matching)
sheet['D19'].value = ad_mney
sheet['D20'].value = pesq_emp
sheet['D21'].value = pesquisas_diretas
sheet['D22'].value = pesquisas_descoberta
sheet['D23'].value = tot_visualizacs_emp
sheet['D25'].value = fbinsta_mney
sheet['D27'].value = curtidas
sheet['D28'].value = novas_curtidas
sheet['D29'].value = visualizcs
sheet['D30'].value = alcance
sheet['D31'].value = alcance_org
sheet['D32'].value = alcance_viral
sheet['D33'].value = alcance_pago
sheet['D34'].value = envolvimento
sheet['D35'].value = qtd_posts_mes
sheet['D36'].value = qtd_compart_mes
sheet['D38'].value = qtd_comentarios_mes
sheet['D40'].value = qtd_reacoes_mes
sheet['D42'].value = qtd_interacoes_mes
sheet['D44'].value = qtd_links_mes
sheet['D45'].value = alcance_links
sheet['D47'].value = envolvimento_links
sheet['D48'].value = consumo_links
sheet['D49'].value = qtd_photos_mes
sheet['D50'].value = alcance_photos
sheet['D52'].value = envolvimento_photos
sheet['D53'].value = consumo_photos
sheet['D54'].value = qtd_videos_mes
sheet['D55'].value = alcance_videos
sheet['D57'].value = envolvimento_videos
sheet['D58'].value = consumo_videos
sheet['D60'].value = fllwrs
sheet['D62'].value = qtd_vids + qtd_imgs
sheet['D63'].value = total_likes
sheet['D65'].value = total_comments
sheet['D67'].value = qtd_vids
sheet['D68'].value = vid_vws
print('Valores inputados com sucesso com sucesso')
workbook.save('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/Dezembro.xlsx')
print('Planilha salva com sucesso')
