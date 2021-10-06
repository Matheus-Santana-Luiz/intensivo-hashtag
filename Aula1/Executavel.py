import pyautogui as ptg
import pyperclip as ppc
import pandas as pd
import time

ptg.PAUSE = 1

ptg.press("win")
ptg.write("brave")
ptg.press("enter")
ppc.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
ptg.hotkey("ctrl", "v")
ptg.press("enter")
time.sleep(2)
ptg.click(x=362, y=335, clicks=2)
time.sleep(2)
ptg.rightClick(x=362, y=335)
ptg.click(x=468, y=655)
time.sleep(2)
ptg.press("enter")

tabela = pd.read_excel(r"D://mathe/Downloads/Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

ptg.hotkey("ctrl", "t")
ppc.copy("https://outlook.live.com/mail/0/inbox")
ptg.hotkey("ctrl", "v")
ptg.press("enter")
time.sleep(5)

ptg.click(x=155, y=140)

ppc.copy("pythonimpressionador+diretoria@gmail.com")
ptg.hotkey("ctrl", "v")
ptg.press("tab")
ppc.copy("Relatório de Vendas")
ptg.hotkey("ctrl", "v")
ptg.press("tab")

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Matheus Santana"""

ppc.copy(texto)
ptg.hotkey("ctrl", "v")
ptg.hotkey("ctrl", "enter")