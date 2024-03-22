#  baixar os arquivos através das chaves
# https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=
import openpyxl 
import pyperclip
import pyautogui
from time import sleep
from PySimpleGUI import PySimpleGUI as sg



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

for linha in sheet_planilha.iter_rows(min_row=2):#LINHA 1
    

    sleep(1)

    nome_produto = linha[0].value
    #COPIAR
    pyperclip.copy(nome_produto)
    #CLICAR
    pyautogui.click(753,442,duration=2)
    #COLOAR
    pyautogui.hotkey('ctrl', 'v')
    #CHECAR
    pyautogui.click(712,495,duration=2)
    #COMFIRMAR CHAVE
    pyautogui.click(808,565,duration=4)#GOOGLE
    #pyautogui.click(792,570,duration=7)#OPERAGX
    #DOWNLOAD
    pyautogui.click(712,406,duration=2.5)#GOOGLE
    #pyautogui.click(703,419,duration=2.5)#OPERAGX

    #COMFIRMAR CERTIFICADO
    pyautogui.click(1029,249,duration=2)# google
    #pyautogui.click(1031,251,duration=1)# Egde
    #pyautogui.click(1051,184,duration=1)# OPERAGX
    #NOVA Consulta
    pyautogui.click(552,407,duration=2)# google
    #pyautogui.click(577,421,duration=1.5)# OPERAGX
