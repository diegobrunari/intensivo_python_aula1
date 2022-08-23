from typing import TextIO
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# Passo 1: Entrar no site (link do canal)
pyautogui.press ("win")
pyautogui.write("brave")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar no sistema e encontrar a base de dados (pasta Explorar)
pyautogui.click(x=340, y=264, clicks=2)
time.sleep(2)

# Passo 3: Download da base de dados
pyautogui.click(x=389, y=337) #clicou no arquivo
pyautogui.click(x=1713, y=155) #clicou nos detalhes
pyautogui.click(x=1456, y=527) #clicou em download
time.sleep(5)
pyautogui.press("enter")

# Passo 4: Calcular indicadores (faturamento, quantidade de produtos)
tabela = pd.read_excel("Vendas - Dez.xlsx")
print(tabela)

quantidade = tabela["Quantidade"].sum()
faturamento = tabela["Valor Final"].sum()

print(quantidade)
print(faturamento)

# Passo 5: Entrar no e-mail
pyautogui.hotkey("ctrl", "t") #abriu nova aba
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox") #copiando link gmail
pyautogui.hotkey("ctrl", "v") #colando link gmail
pyautogui.press("enter") #entrando no site
time.sleep(5)
pyautogui.click(x=82, y=166) #clicou no botão +
pyautogui.write("diego.skp@hotmail.com") #escrever destinario
pyautogui.press("tab") #selecionar e-mail
pyautogui.press("tab") #assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") #passar para corpo do e-mail

texto = f"""
Prezados, bom dia.

O faturamento foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Att., 
Diego.
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# Passo 6: Enviar e-mail com indicadores
pyautogui.hotkey("ctrl", "enter")