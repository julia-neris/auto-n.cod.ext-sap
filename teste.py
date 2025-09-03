import pyautogui
import time
import pandas as pd
import pyperclip

df = pd.read_excel(
    r"C:\Users\julia.neris\Downloads\27082025.xlsx",
    dtype={0: str, 1: str, 2: str, 4: str, 5: str, 6: str}
)

for index, row in df.iterrows():
    numero = str(row[0])
    texto = str(row[1])
    chavebanco = str(row[2])
    agencia = str(row[4])
    conta = str(row[5])
    digitoconta = str(row[6])

    print("Agora vai")
    time.sleep(5)

    #aqui só abre pn
    pyautogui.click(x=439, y=101) 
    time.sleep(5)

    pyperclip.copy(numero)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(3)  

    pyautogui.click(x=47, y=291) 
    time.sleep(5)

    pyautogui.click(x=1225, y=392) 
    pyautogui.dragRel(0, 300, duration=0.5)

    pyautogui.click(x=321, y=447)  
    time.sleep(2)
    pyperclip.copy(texto)
    time.sleep(2) 
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    #pagamento
    pyautogui.click(x=561, y=285) 
    time.sleep(2)

    #país
    pyautogui.click(x=133, y=411) 
    time.sleep(2)
    pyautogui.write("BR")
    pyautogui.press('enter')
    time.sleep(2)

    #chave do banco
    pyautogui.click(x=251, y=413)
    time.sleep(2)
    pyperclip.copy(chavebanco)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    pyperclip.copy(agencia)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    #conta bancaria
    pyautogui.click(x=408, y=409)
    time.sleep(2)
    pyperclip.copy(conta)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    pyautogui.write("-")
    time.sleep(1)

    pyperclip.copy(digitoconta)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(5)