#pyautogui 
#pyperclip
#pandas - base de dados
#numpy
#openpyxl

import pyautogui
import pyperclip
import time
import pandas as pd
#as pd da o apelido de pd 

pyautogui.PAUSE = 0.5 #delay em tds as linhas do código

pyautogui.press("win") # abre windows
pyautogui.write("chrome") #digita chrome
pyautogui.press("enter") # da enter
#pyautogui.hotkey("ctrl", "t") #abre nova guia 
time.sleep(1)
pyautogui.click(x=638, y=629)
time.sleep(1)

pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga") #copia o link
#pyautogui.write("link") - caracter especial
pyautogui.hotkey("ctrl", "v") # cola o link
pyautogui.press("enter") # da enter
#demora alguns segundos para carregar ent pause o codigo
time.sleep(5) #só pausa nessa linha

#descobrindo a posição do mouse na tela para clicar na pasta
#time.sleep(5)
#pyautogui.positon()
#colocar o mouse em cima da pastq  quer abrir 
#copiar o valor da posição e botar no .click

#pyautogui.click( x=valor, y=valor, clicks = numero de clicks q quer, button="right")
pyautogui.click(x=517, y=335, clicks=2) #duplo click p abir pasta do drive
time.sleep(2) #navegador carregando a pasta
#pegar posição do arquivo para baixar
pyautogui.click(x=467, y=437) #clicar no arquivo
#pegar posição do mouse
pyautogui.click(x=1663, y=205) #clicar nos 3 pontinhos
#pegar posição do mouse
pyautogui.click(x=1305, y=753) #clicar no fazer download
time.sleep(8)

#importando a base de dados

tabela = pd.read_excel(r"C:\Users\Dell\Downloads\Vendas - Dez.xlsx") # ler um arquivo. tipo do arquivo(excel, word,..) : sheets muda aba do documento
#tabela = pd.read_excel(r"C:\Users\Dell\Downloads\Vendas - Dez.xlsx", sheets) # ler um arquivo. tipo do arquivo(excel, word,..) : sheets muda aba do documento
#esse r diz que é uma roadstring, p python n achar q eh caracter especial
print(tabela)
#no jupyter usar display(tabela)
faturamento = tabela["Valor Final"].sum()#pegando coluna da tabela
quantidade = tabela["Quantidade"].sum() #calculando quantidade

#enviar o email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
#clica em escrever
pyautogui.click(x=37, y=201)
time.sleep(2)

#destinatario
pyautogui.write("fuscaazullistrado@gmail.com")
pyautogui.press("tab")
#assunto
pyautogui.press("tab")
pyperclip.copy("RELATÓRIO HUR DUR")
pyautogui.hotkey("ctrl","v")
#corpo do email
pyautogui.press("tab")
# o f antes das aspas indica que ha uma variavel no texto
texto = f"""ola 
corpo do email
com mais de uma linha
res aspas
faturamento foi: R$ {faturamento:,.2f}
quantidade foi:{quantidade:,}"""
pyperclip.copy(texto) #copiando e colando o texto do email 
pyautogui.hotkey("ctrl","v") # p evitar erros de caracter especial

pyautogui.hotkey("ctrl", "enter")