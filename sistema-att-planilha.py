#necessário instalar gspread

import gspread
import time

#------Caminho até a planilha e chave para altera-la
CODE = '12OG8eIFIv3cDwLWRaybtXgsgMXiTO3OJW6X9LstZcYA'

gc = gspread.service_account(filename = 'key.json')
sh = gc.open_by_key(CODE)
ws = sh.worksheet('engenharia_de_software')

page = sh.sheet1
#-----------

#Definindo um contador que também representará qual aluno será selecionado
y = 4
while y < 28:

    #variaveis com as faltas, notas e calculo de média
    f = int(page.cell(y, 3).value)

    x1 = int(page.cell(y, 4).value)
    x2 = int(page.cell(y, 5).value)
    x3 = int(page.cell(y, 6).value)

    m = round(int((x1+x2+x3)/30))

    #condicional para faltas
    if f > 15:

        #atualizando dados na tabela
        page.update_cell(y, 7, 'Reprovado por falta')
        page.update_cell(y, 8, '0')
        print('Reprovado por faltas')

    else:

        #condicionais para média
        if m >= 7:
            page.update_cell(y, 7, 'Aprovado')
            page.update_cell(y, 8, '0')
            print('Aprovado! Média{}'.format(m))
        elif m >= 5  and m < 7:
            page.update_cell(y, 7, 'Exame Final')
            #simplificação da equação proposta '5 <= (m+naf)/2' onde 'ef' é a nota o exame final
            ef = 10 - m
            page.update_cell(y, 8, '{}'.format(ef))
            print('Exame Final! Média{}, nota necessária {}'.format(m, ef))
        else:
            page.update_cell(y, 7, 'Reprovado')
            page.update_cell(y, 8, '0')
            print('Reprovado! Média {}'.format(m))


    #adicionado para contornar o quote limit do google cloud
    time.sleep(3)
    y = y+1



 