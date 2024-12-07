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

def split_and_convert(value):
    integer_part, decimal_part = str(value).split('.')
    return int(integer_part), int(decimal_part)

def positionsNumber():
    positionIndex = [
        {"mx":4295,"my":644},
        {"mx":4328,"my":644},
        {"mx":4364,"my":644},
        {"mx":4386,"my":644},
        {"mx":4413,"my":644},
        {"mx":4446,"my":644},
        {"mx":4484,"my":644},
        {"mx":4508,"my":644},
        {"mx":4541,"my":644}
    ]
    return positionIndex



file_path = 'template.xlsx'
sheet_name = 'Plan1'
sheet_name_pos = 'positions'
mapped_fill_values = map_sheet(file_path, sheet_name)
mapped_positions_values = map_sheet(file_path, sheet_name_pos)

converted_data_positions = [split_and_convert(value) for value in mapped_positions_values]

positions = positionsNumber()


def convertArrayDictionary(array):
    transformed_data = []
    for x, y in array:
        transformed_data.append({
            "mx": x,
            "my": y,
        })
    return transformed_data

arrayTest = convertArrayDictionary(converted_data_positions)

for position, mFvalue in zip(arrayTest, mapped_fill_values):
    x = position["mx"]
    y = position["my"]
    time.sleep(0.1)
    pyautogui.click(x, y)
    for numberPos in position:
        if(mFvalue == 1):
            xN = positions[0]["mx"]
            yN = positions[0]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 2):
            xN = positions[1]["mx"]
            yN = positions[1]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 3):
            xN = positions[2]["mx"]
            yN = positions[2]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 4):
            xN = positions[3]["mx"]
            yN = positions[3]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 5):
            xN = positions[4]["mx"]
            yN = positions[4]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 6):
            xN = positions[5]["mx"]
            yN = positions[5]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 7):
            xN = positions[6]["mx"]
            yN = positions[6]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 8):
            xN = positions[7]["mx"]
            yN = positions[7]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
        elif(mFvalue == 9):
            xN = positions[8]["mx"]
            yN = positions[8]["my"]
            time.sleep(0.1)
            pyautogui.click(xN, yN)
    
