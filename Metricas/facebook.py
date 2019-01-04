import csv
import openpyxl

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

#agora colocar na planilha do excel
workbook = openpyxl.load_workbook('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/Dezembro.xlsx')
sheet = workbook.get_sheet_by_name('Sheet1')
print('Planilha aberta com sucesso')
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
print('Valores inputados com sucesso com sucesso')
workbook.save('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/Dezembro.xlsx')
print('Planilha salva com sucesso')
