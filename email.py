#clicar no escrever destinatario
#automação criada a partir de clicks automatizados no gmail

import pyautogui
import time
import pyperclip

#Projeto para coletar dados web e enviar por e-mail automatizado
#download do arquivo que será enviado

pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

#demora alguns segundos
time.sleep(5)
#passo 2: Navegar no sistema e encontrar base de dados(entrar na pasta exportar)
pyautogui.click(x=351, y=282, clicks=2)
time.sleep(2)
pyautogui.click(x=422, y=292) #clicar no arquivo
time.sleep(2)
pyautogui.click(x=1155, y=194) #clicar 3 pontos
time.sleep(1)
pyautogui.click(x=914, y=622) #fazer download



#verificar a tabela

#passo 3 : importar base de dados para o python
import pandas as pd
function
tabela = pd.read_excel(r'C:\Users\Jorge\Downloads\Vendas - Dez.xlsx')
print(tabela)


#dados importados da tabela para e-mail

#passo 4 : Calcular indicadores
#faturamento = soma coluna valor final
faturamento = tabela['Valor Final'].sum()
#qtdade produtos vendidos - coluna quantidade
quantidade= tabela['Quantidade'].sum()

print(quantidade)



#passo 5 : enviar e-mail
#Abrir o e-mail https://mail.google.com/mail/u/0/#inbox

#envio do e-mail automatizado
pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)

pyautogui.click(x=76, y=215)
time.sleep(2)
pyautogui.write('junior.ornellas.nf@gmail.com')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
#clicar no assunto
pyautogui.click(x=852, y=313)
time.sleep(1)
pyautogui.write('Relatorio de vendas')
#escreve o corpo do e-mail
pyautogui.press('tab')
time.sleep(2)
texto=('''Olá
Esse e um e-mail de teste
esta automatizado e esta sendo enviado sozinho
estou puxando de uma tabela o faturamento foi R${:.2f} e a quantidade foi {}'''.format(faturamento,quantidade))
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
#anexar o arquivo
pyautogui.click(x=954, y=638)
time.sleep(2)
pyautogui.click(x=74, y=226)
time.sleep(2)
pyautogui.click(x=299, y=174, clicks=2)
time.sleep(10)
#envia e-mail
pyautogui.click(x=847, y=643)