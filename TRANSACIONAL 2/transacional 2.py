import pyautogui
import pyscreeze
import PIL
import time
import pandas as pd

caminho = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/TRANSACIONAL 2/Transacional.xlsx"
ler = pd.read_excel(caminho)
tabela = pd.DataFrame(ler)
print (tabela)

x = 0
y = 0


while True and x < len (tabela):
    plu = "{:.0f}".format(tabela.iloc [x, y + 5])
    dia_validade = "{:.0f}".format(tabela.iloc[x, y + 14])
    mes_validade = "{:.0f}".format(tabela.iloc[x, y + 15])
    ano_validade = "{:.0f}".format(tabela.iloc[x , y + 16])
    dia_fabricacao = "{:.0f}".format(tabela.iloc[x, y + 18]) 
    mes_fabricacao = "{:.0f}".format(tabela.iloc[x, y + 19])
    ano_fabricacao = "{:.0f}".format(tabela.iloc[x, y + 20])
    ilpn = "{:.0f}".format(tabela.iloc[x , y + 12])
    kg_cs ="{:.0f}".format( tabela.iloc [x, y + 22])
    g_unid ="{:.0f}".format( tabela.iloc[x, y + 24])
    
    
    
    print(plu)
    print(dia_validade)
    print(mes_validade)
    print(ano_validade)
    print(dia_fabricacao)
    print(mes_fabricacao)
    print(ano_fabricacao)
    print(ilpn)
    print(kg_cs)
    print(g_unid)

    pyautogui.doubleClick(104, 172)
    time.sleep(0.5)
    pyautogui.hotkey("backspace")
    time.sleep(0.5)
    pyautogui.write("0", interval=0.111)
    pyautogui.write(str(ilpn),interval=0.111)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.doubleClick(86, 207)
    pyautogui.write("OLPNCANCELADAS",interval=0.111)
    time.sleep(0.7)
    pyautogui.press("enter")    
    for i in range(2):    
        with pyautogui.hold("ctrl"):
            pyautogui.press("a")
            time.sleep(1)
    time.sleep(1)
    pyautogui.write(str(plu))
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(0.5)
    with pyautogui.hold("ctrl"):
        pyautogui.press("a")
    time.sleep(0.5)
    pyautogui.write(str(kg_cs),interval=0.111)
    time.sleep(0.5)
    pyautogui.hotkey("tab")
    time.sleep(0.5)
    pyautogui.write(str(g_unid),interval= 0.111)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)
    with pyautogui.hold("ctrl"):
        pyautogui.press("a")
    time.sleep(1)
    pyautogui.doubleClick(34, 219)
    pyautogui.hotkey("backspace")
    pyautogui.write(str(dia_fabricacao),interval=0.111)
    time.sleep(1)
    pyautogui.leftClick(91, 220)
    pyautogui.write(str(mes_fabricacao),interval=0.111)
    time.sleep(1)
    pyautogui.leftClick(152, 220)
    pyautogui.write(str(ano_fabricacao),interval=0.111)
    time.sleep(1)
    pyautogui.leftClick(36, 253)
    pyautogui.write(str(dia_validade),interval=0.111)
    time.sleep(1)
    pyautogui.leftClick(94, 253)
    pyautogui.write(str(mes_validade),interval=0.111)
    time.sleep(1)
    pyautogui.leftClick(158, 253)
    pyautogui.write(str(ano_validade),interval=0.111)
    time.sleep(2)
    with pyautogui.hold("ctrl"):
        pyautogui.press("a")
        time.sleep(2)
    pyautogui.leftClick(823, 224)
    time.sleep(2)
    x = x + 1 


