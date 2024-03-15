import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha-certificado.xlsx')
planilha_alunos = workbook_alunos['Planilha1']

for indice, linha in enumerate(planilha_alunos.iter_rows(min_row=2)):

    nome_participante = linha[0].value
    nome_curso = linha[1].value
    carga_horaria = linha[2].value
    dta_inicio = linha[3].value
    dta_final = linha[4].value
    dta_emissao = linha[5].value


    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 50)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 45)
    fonte_dta = ImageFont.truetype('./tahoma.ttf', 40)

    image = Image.open('Certificado.jpg')
    escrever = ImageDraw.Draw(image)

    escrever.text((370,262), nome_participante, fill= 'black', font= fonte_nome)
    escrever.text((365,395), nome_curso, fill= 'black', font= fonte_geral)
    escrever.text((680,525), str (carga_horaria), fill= 'black', font= fonte_dta)


    if dta_inicio:
        data_inicio_formatada = dta_inicio.strftime('%Y-%m-%d')
        data_final_formatada =  dta_final.strftime('%Y-%m-%d')
        data_emissao_formatada = dta_emissao.strftime('%Y-%m-%d')

    escrever.text((230, 1000), data_inicio_formatada, fill= 'blue', font= fonte_dta)
    escrever.text((230, 1165),data_final_formatada, fill= 'blue', font= fonte_dta)
    escrever.text((1365, 1165), data_emissao_formatada, fill= 'blue', font= fonte_dta)

    image.save(f'{indice} {nome_participante} Certificado.png')

