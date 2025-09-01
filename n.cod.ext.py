import pyautogui
import time
import pandas as pd
import pyperclip

df = pd.read_excel(r"C:\Users\julia.neris\Downloads\27082025.xlsx")
for index, row in df.iterrows():
    numero = str(row[0])
    texto = str(row[1])

    print("Agora vai")
    time.sleep(5)

    #aqui só abre pn
    pyautogui.click(x=439, y=101) 
    time.sleep(2)

    pyperclip.copy(numero)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)  

    pyautogui.click(x=1225, y=392) 
    pyautogui.dragRel(0, -50, duration=0.5)

    pyautogui.click(x=608, y=100) 
    time.sleep(2)

    #aqui cola o numero cod externo
    pyautogui.click(x=1225, y=392) 
    pyautogui.dragRel(0, 300, duration=0.5)

    pyautogui.click(x=321, y=447)  
    time.sleep(2)
    pyperclip.copy(texto)
    time.sleep(2) 
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 's')
    print("Agora foi")

