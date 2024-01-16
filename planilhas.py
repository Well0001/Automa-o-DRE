import openpyxl
def loadValueTotal (arquivoXlsx):
    workbook = openpyxl.load_workbook(arquivoXlsx)
    sheet = workbook['Sheet1']
    valor_celula = sheet['G13'].value
    return valor_celula


def preencherDRE(valor, inicio=9):
    workbook = openpyxl.load_workbook('DRE.xlsx')
    sheet = workbook['LANÇAMENTO']
    celula = sheet[f'G{inicio}']
    celula.value = valor
    workbook.save('DRE.xlsx')

