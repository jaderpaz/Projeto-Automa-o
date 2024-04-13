#passo a passo do projeto
# 1 passo, entrar no sistema da empresa
# pip install pyautogui

import pyautogui
import time


pyautogui.PAUSE = 2
#pyautogui.click => clicar em algum lugar da tela
#pyautogui.write => escrever um texto
#pyautogui.press => pressionar uma tecla do teclado

#abrir o navegador

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")


link = "https://nfse.vitoria.es.gov.br/"
pyautogui.write(link)
pyautogui.press("enter")

# # dar uma pausa um pouco maior
time.sleep(5)

# # 2 passo, fazer o login
pyautogui.click(x=1411, y=354)
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.write("ivan.silva@oi.net.br")

pyautogui.press("tab")
pyautogui.write("OISA2020")
pyautogui.press("tab")
pyautogui.press("enter")

time.sleep(5)
pyautogui.click(x=721, y=372)
time.sleep(1)
pyautogui.click(x=1311, y=941)
time.sleep(1)
pyautogui.click(x=783, y=715)
time.sleep(1)
pyautogui.click(x=540, y=447)
time.sleep(1)
pyautogui.click(x=645, y=427)
time.sleep(1)
pyautogui.click(x=617, y=518)
time.sleep(1)
pyautogui.press("tab")
time.sleep(1)
pyautogui.press("tab")
time.sleep(1)
pyautogui.press("tab") 
time.sleep(1) 
pyautogui.press("enter")
pyautogui.click(x=660, y=563)
pyautogui.click(x=661, y=411)
time.sleep(1)

pyautogui.click(x=549, y=461)



time.sleep(3)
import pandas
import datetime
tabela = pandas.read_excel("Vitoria teste.xlsx")

for linha in tabela.index:
#     #numero da nota
    #pyautogui.click(x=672, y=400)
    
    numero = tabela.loc[linha,"Nota - Número Documento"]
    pyautogui.write(str(numero))
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")


    data = tabela.loc[linha,"Imposto - Data Base"]
    data_atual = data.strftime("%d/%m/%Y")
    pyautogui.write(str(data_atual))
    time.sleep(1)
    pyautogui.press("tab")

    pagamento = tabela.loc[linha,"Data de pagamento"]
    data_pagamento = pagamento.strftime("%d/%m/%Y")
    pyautogui.write(str(data_pagamento))
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")

    #Verificar o valor
    valor = tabela.loc[linha,"Imposto - Valor Base de Calculo"]
    pyautogui.write(str(valor))
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.5)

    cnpj = tabela.loc[linha,"Parceiro - CNPJ/CPF"]
    pyautogui.click(x=537, y=754)
    pyautogui.write(str(cnpj))
    time.sleep(3)
    pyautogui.press("tab")
    time.sleep(3)           

    pyautogui.click(x=537, y=456)
    time.sleep(1) 
    pyautogui.click(x=545, y=463)


