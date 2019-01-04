import csv
import openpyxl

print('Começando a parte do Google')

#pesquisas do termo no mes
pesq_emp = 0

pesquisas_diretas = 0

pesquisas_descoberta = 0

tot_visualizacs_emp = 0

#precisamos saber quanto foi investido em adwords no mes em questão
print('Para começar, preciso saber quanto foi investido em AdWords no mes em questão (coloque no formato xx.xx)')
ad_mney = input('Quanto foi investido em AdWords neste mês?\n')
ad_mney = round(float(ad_mney), 2)

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
        print(line[7])
        z += 1


print(ad_mney)
print(pesq_emp)
print(pesquisas_diretas)
print(pesquisas_descoberta)
print(tot_visualizacs_emp)

#agora colocar na planilha do excel
workbook = openpyxl.load_workbook('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/Dezembro.xlsx')
sheet = workbook.get_sheet_by_name('Sheet1')
print('Planilha aberta com sucesso')
sheet['D19'].value = ad_mney
sheet['D20'].value = pesq_emp
sheet['D21'].value = pesquisas_diretas
sheet['D22'].value = pesquisas_descoberta
sheet['D23'].value = tot_visualizacs_emp
print('Valores inputados com sucesso com sucesso')
workbook.save('/Users/pedroenriqueandrade/Desktop/python_projects/Metricas/Dezembro.xlsx')
print('Planilha salva com sucesso')
