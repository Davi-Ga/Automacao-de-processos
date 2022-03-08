import pyautogui as pa
import pyperclip
import time as t
import pandas as pd

#Passo 1: Entrar no sistema da empresa(Abrir o arquivo)
pa.PAUSE = 1
pa.alert("Aperte OK para iniciar a automação")
pa.press("win")
pa.write("Opera")
pa.press("enter")
pa.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pa.hotkey("ctrl","v")
pa.press("enter")

#Passo 2: Navegar no sistema até encontrar a base de dados e baixá-los

t.sleep(2)

pa.click(x=430, y=266,clicks=2)
t.sleep(2)
#Passo 3: Exportar a base de vendas
pa.click(x=333, y=438,button='right')
t.sleep(2)
pa.click(x=380, y=677)

#Passo 4: Calcular os indicadores(Faturamento e quantidade de produtos vendidos)
bd=pd.read_excel(r'C:/Users/davig/Downloads/Vendas - Dez.xlsx')

quantidade=bd['Quantidade'].sum()
faturamento=bd['Valor Final'].sum()

#Passo 5: Enviar um email para a diretoria com os indicadores
pa.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/1/#inbox")
pa.hotkey("ctrl","v")
pa.press("enter")
t.sleep(5)
pa.click(x=110, y=177)
t.sleep(1)
pa.write("sirkamiih@gmail.com")
pa.press("tab")
pa.press("tab")
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pa.hotkey("ctrl","v")       
pa.press("tab")
texto = f""" Dados de ontem 
         
         Bom dia, este é um relatório com os dados obtidos ontem
         Faturamento:{faturamento:,.2f}
         
         Quantidade:{quantidade:,}
         
         Ass
         DaviGA
         """
pyperclip.copy(texto)
pa.hotkey("ctrl","v")
pa.hotkey("ctrl","enter")
pa.alert("Automação concluída com sucesso")
