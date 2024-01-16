import os
import pyautogui
from time import sleep
from pdfToCsv import pdf_to_csv
from gerarXlsx import gerarXlsx
from planilhas import loadValueTotal, preencherDRE

dias_do_mes = pyautogui.prompt(title='Informações : ',
                               text='DIGITE O NÚMERO DE DIAS EXISTENTES NO MÊS ')
mes = pyautogui.prompt(title='Informações : ',
                       text='DIGITE O NÚMERO CORRESPONDENTE AO MÊS UTILIZANDO 2 DIGITOS ')
ano = pyautogui.prompt(title='Informações : ',
                       text='DIGITE O ANO UTILIZANDO 4 DIGITOS ')
desligar = pyautogui.confirm(text='Deseja que o computador desligue após a automação terminar? ',
                             title='Atenção',buttons=['Sim', 'Não'])

confirmacao = pyautogui.confirm(text='Dados Coletados Com Sucesso!\n'
                                     'Certifique que o Ourofarma esteja aberto em'
                                     ' Relatorios-> Vendas -> Vendas Por classe.\n'
                                     'O Programa inicializara em 05 Segundos após a confirmação.',
                  title='ATENÇÃO!!!!', buttons=['Iniciar', 'Fechar'])

dia = int(dias_do_mes)
inicio = 9

if confirmacao == 'Iniciar':
    try:
        sleep(5)
        pyautogui.click(905,317)
        pyautogui.click(421,283)
        pyautogui.typewrite("MANIPULADOS",interval=0.04)
        sleep(0.6)
        pyautogui.hotkey('enter')
        sleep(0.3)
        pyautogui.hotkey('enter')
        pyautogui.hotkey('F10')
        sleep(0.7)
        pyautogui.click(647, 482, interval=1)
        for i in range(dia):
            if i + 1 < 10:
                data = str(f'0{i+1}{mes}{ano}')
            else:
                data = str(f'{i+1}{mes}{ano}')
            nomedoarquivo = f'\{data}.pdf'
            nomedoarquivocsv =f'\{data}.csv'
            nomedoarquivoxlsx = f'\{data}.xlsx'
            caminho_completo = r'C:\Users\User\Desktop\Manip\Temp'+nomedoarquivo
            arquivo_csv = r'C:\Users\User\Desktop\Manip'+nomedoarquivocsv
            arquivo_final = r'C:\Users\User\Desktop\Manip\Xlsx'+nomedoarquivoxlsx
            # 4 Cliques em 504,197
            pyautogui.click(504, 197, clicks=4, interval=0.15)
            # Preencher com a data Inicial
            pyautogui.hotkey('backspace')
            sleep(0.05)
            pyautogui.typewrite(data, interval=0.02)
            sleep(0.1)
            # 4 Cliques em 608,197
            pyautogui.click(608, 197, clicks=4, interval=0.15)
            # Preencher com a data Final
            pyautogui.hotkey('backspace')
            sleep(0.05)
            pyautogui.typewrite(data, interval=0.02)
            sleep(1)
            pyautogui.click(842,591,interval=0.06)
            sleep(1.2)
            pyautogui.click(69,31,interval=0.06)
            sleep(0.9)
            sleep(0.9)
            pyautogui.click(850,311,interval=0.06)
            sleep(0.9)
            pyautogui.typewrite(caminho_completo)
            pyautogui.click(1193,682,interval=0.9)
            sleep(1)
            pyautogui.click(754,456,interval=0.9)
            sleep(1)
            pyautogui.click(1342,8,interval=1)
            sleep(2)
            pdf_to_csv(caminho_completo, arquivo_csv)
            sleep(2.5)
            gerarXlsx(arquivo_csv, arquivo_final)
            sleep(2)
            valorDoDia = loadValueTotal(arquivo_final)
            sleep(2)
            valorDoDia = float(valorDoDia)
            preencherDRE(valor=valorDoDia, inicio=inicio)
            inicio += 1

    except Exception as e:
        pyautogui.alert(text='Oops , Algo Saiu errado, tente novamente',
                        button='Ok')
        pyautogui.alert(text=Exception, button="Ok")
if confirmacao == 'Fechar':
    pyautogui.alert(text='Fechando Aplicação!!',
                    title='Fechando',button='Ok')

if desligar == 'Sim':
    os.system('shutdown -s -t 10')
else:
    pass