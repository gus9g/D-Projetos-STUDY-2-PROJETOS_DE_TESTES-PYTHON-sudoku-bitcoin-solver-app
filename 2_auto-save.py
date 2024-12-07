import openpyxl
import pyautogui
import time

# https://www.sudoku-solutions.com/
time.sleep(1.1)
pyautogui.click(1598,262)
time.sleep(2.1)
pyautogui.click(1534,501)
time.sleep(0.5)



# Move o mouse para o ponto A
pyautogui.moveTo(1015,244)

# Clica e segura
pyautogui.mouseDown()

# Move para o ponto B enquanto segura o clique
pyautogui.moveTo(1306,523)

# Solta o botão do mouse
pyautogui.mouseUp()

pyautogui.hotkey("ctrl","c")

# sublime text
pyautogui.click(2859,1033)

pyautogui.hotkey("ctrl","v")

pyautogui.hotkey("ctrl","f")
time.sleep(0.1)
pyautogui.hotkey("backspace")
pyautogui.write("\\n")
pyautogui.hotkey("alt", "enter")
pyautogui.write(", ")
time.sleep(0.1)
pyautogui.hotkey("ctrl","f")
time.sleep(0.2)
pyautogui.hotkey("backspace")
pyautogui.write("\\b(\\d+\\D+){8}\\d+\\b")
pyautogui.hotkey("alt", "enter")
pyautogui.hotkey("right")
pyautogui.hotkey("right")
pyautogui.hotkey("enter")
pyautogui.hotkey("delete")
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")

# # planilha Campo A1
time.sleep(1.1)
pyautogui.click(2441,283)
time.sleep(1.1)
pyautogui.click(2441,283)
time.sleep(0.8)
pyautogui.hotkey("esc")
pyautogui.hotkey("ctrl", "v")
time.sleep(0.1)
pyautogui.hotkey("ctrl")
time.sleep(0.3)
pyautogui.hotkey("down")
pyautogui.hotkey("enter")

time.sleep(1.5)
pyautogui.hotkey("enter")

time.sleep(1)
# checkbox virgula
pyautogui.click(1959,166)

pyautogui.hotkey("enter")
pyautogui.hotkey("enter")
pyautogui.hotkey("enter")

time.sleep(1)
pyautogui.hotkey("ctrl", "b")




