import openpyxl
import pyautogui
import time

def map_sheet(file_path, sheet_name):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    sheet_values = []

    for row in range(1, 10):
        for col in range(1, 10):
            cell_value = sheet.cell(row=row, column=col).value
            sheet_values.append(cell_value if cell_value is not None else 0)

    return sheet_values

def inScreen(ArrayMap):
    for value in range(0, len(ArrayMap)):
        if(ArrayMap[value] == 0):
            pyautogui.press('tab')
        else:
            time.sleep(0.2)
            pyautogui.typewrite(str(ArrayMap[value]))



# CODIGO
file_path = 'template.xlsx'
sheet_name = 'Plan1'
mapped_values = map_sheet(file_path, sheet_name)

# clear tabela
pyautogui.click(1508,436)


x, y = 1032,241
time.sleep(0.8)
pyautogui.click(x, y)
inScreen(mapped_values)



# Exibir os resultados
# for item in transformed_data:
#     print(item)

# REGEX PARA QUEBRAR LINHA APOS 8 NUMEROS, VSCODE
# ((?:\d, ){8}\d)