import datetime
import locale
import pyperclip as pc
import pyautogui as pg
import pandas as pd
import re
import time
from tkinter import filedialog as tf

locale.setlocale(locale.LC_ALL, "pt_BR")
dt = datetime.datetime.today()
hoje = dt.strftime("%d/%m/%Y")
pausa = datetime.timedelta(hours = 1, minutes = 0)
tempo_envio = datetime.timedelta(years = 0, months = 3, days = 0)

#   Não disparar emails fora do Horário Comercial
if dt.hour >= 9 and dt.hour < 18:
    print("FECHAR ESSA JANELA INTERROMPE A AUTOMAÇÃO")
    #   OFFICIAL
    with open(tf.askopenfilename(defaultextension = '.xlsx'), 'rb') as file:
        df = pd.read_excel(file).drop('Unnamed: 0', axis = 1)

    #   TESTING
    #with open('xlsx/Modelo Envio Portfólio.xlsx', 'rb') as file:
    #    df = pd.read_excel(file, sheet_name='RH', engine = 'openpyxl')

    #   Linhas sem data ou observações não dispararam mensagens
    #   Considerar nome vazio como Recursos Humanos      
    #resp = df[ df["Data Envio"].isnull() ]
    copy = df[ df["Observações"].isnull() ].copy()
    copy.fillna({'Nome': 'RH'}, inplace = True)
    #copy.fillna({'Data Envio': hoje}, inplace = True)

    #   Entrar no outlook
    pg.click(400, 1040, button = 'left', interval = 5)
    pg.hotkey('win', 'left')
    pg.press('esc', presses = 2, interval = .5)
    pg.click(192, 306, button = 'left')

    #   Copiar o template
    pg.dragTo(500, 200, button = 'left')
    pg.hotkey('ctrl', 'c')
    i = 0
    while i < len(copy):
        pg.hotkey('ctrl', 'v')
        i += 1
        if i == len(copy): 
            break


    #   Achar o maior tamanho de grupo inferior a 100 com divisão sem resto
    cap = 99
    if len(copy) > 0:
        while len(copy) % cap != 0 and cap > 0:
            cap -= 1


    sent = 0
    start = 0
    finish = start + cap
    while finish <= len(copy):
        resp = copy.iloc[start:finish, 4:6]
        start += len(resp)
        finish = start + len(resp)
        print('Enviando mensagens')
        for nome, email in resp.to_numpy():
            if nome == 'RH':
                #email = email.lower()
            else:
                nome.split()[0].capitalize()
                #email.lower()
            

            #   Enviar as Mensagens
            if re.match("([a-zA-Z0-9._-]{2,}@[a-zA-Z]{2,}.[a-zA-Z]{2,})", email):        
                pc.copy(nome)
                pg.moveTo(793, 171)
                pg.click()
                pg.write(email)
                pg.moveTo(627, 352)
                pg.click()
                pg.hotkey('ctrl', 'v')            
                pg.click(645, 181, button = 'left', interval = 1.5)
        

        if finish < len(copy):
            pg.moveTo(0, 0, interval = 3600)
            dt += pausa
            if dt.hour < 18:
                print(f'Pausa na automação até {dt.hour}:{dt.minute}')
            else:
                print('Encerramos por hoje')
                time.sleep(5)
                break


else:
    print('Estamos fora do horário comercial')
    time.sleep(5)
