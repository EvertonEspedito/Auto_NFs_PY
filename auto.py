#https://consultadanfe.com/#

import openpyxl 
import pyperclip
import pyautogui
from time import sleep
from PySimpleGUI import PySimpleGUI as sg

pagina = 2


sg.theme('Reddit')
layout = [ 
    [sg.Text("Nome do arquivo"),sg.Input(key='nomeArquivo')],
    [sg.Text("Nome da Pagina do arquivo"),sg.Input(key='paginaArquivo')],
    [sg.Button('Começar')] ]
#Janela
janela = sg.Window('Automação NFe e CTe', layout)



#Eventos
while True:
    eventos,valores = janela.read()
    nome_planilha = valores ['nomeArquivo']        
    pagina_planilha  = valores['paginaArquivo']
    if eventos == sg.WIN_CLOSED:
        break
    if eventos == 'Começar':
        nome_planilha = valores ['nomeArquivo']        
        pagina_planilha  = valores['paginaArquivo']
        break


planilha = openpyxl.load_workbook(nome_planilha+'.xlsx')
sheet_planilha = planilha[pagina_planilha]


sleep(3)

for linha in sheet_planilha.iter_rows(min_row=pagina):#LINHA 1
    

    sleep(1)

    nome_produto = linha[0].value
    #COPIAR
    pyperclip.copy(nome_produto)

    #CLICAR
    if(pagina == 2):
        pyautogui.click(462,564,duration=2)
    if(pagina > 2 ):
        pyautogui.click(464,602,duration=2)
        pyautogui.click(464,602,duration=0.1)

    #COLOAR
    pyautogui.hotkey('ctrl', 'v')

    #Buscar
    if(pagina == 2):
        pyautogui.click(852,557,duration=1)
    if(pagina > 2):
        pyautogui.click(859,599,duration=1)    

    #BAIXAR XML
    pyautogui.click(947,455,duration=1)

    #Fechar pop-up
    pyautogui.click(1177,142,duration=2)

    pagina+=1
