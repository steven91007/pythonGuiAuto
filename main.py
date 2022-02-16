import pyautogui
import pyperclip
import pynput
import json
import pandas as pd
import time
import os

def find_mouse_click_coordinates():

    f = open("mouse_coordinates.json.md", "w")
    mouse = pynput.mouse.Controller()
    list = []
    with pynput.keyboard.Events() as event:
        for i in event:
            # 迭代用法。
            if isinstance(i, pynput.keyboard.Events.Press):
                # 滑鼠點擊事件。
                if i.key == pynput.keyboard.Key.esc:
                    json_format = json.dumps(list)
                    f.write(json_format)
                    f.close()
                    break
                elif i.key == pynput.keyboard.KeyCode.from_char('r'):
                    print(mouse.position)
                    list.append(mouse.position)
        i = event.get(1)

class Reader:

    def __init__(self):
        self.path = ''

    @staticmethod
    def extract_excel_part_numbers(path):
        df = pd.read_excel(path)
        part_numbers = [str(d) for d in df.loc[:, 'Part No.']]
        return part_numbers

    @staticmethod
    def read_mouse_coordinates(path):
        f = open(path, 'r')
        coordinates = json.load(f)
        f.close()
        return coordinates

def start(part_numbers):
    mouse_coordinates = Reader.read_mouse_coordinates('mouse_coordinates.json.md')

    with pynput.keyboard.Events() as event:
        for i in event:
            # 迭代用法。
            if isinstance(i, pynput.keyboard.Events.Press):
                # 滑鼠點擊事件。
                if i.key == pynput.keyboard.Key.esc:
                    break
                elif i.key == pynput.keyboard.KeyCode.from_char('s'):
                    for part_number in part_numbers:
                        pyperclip.copy(part_number)

                        for index, mouse_coordinate in enumerate(mouse_coordinates):
                            coordinate_x = mouse_coordinate[0]
                            coordinate_y = mouse_coordinate[1]
                            if index == 0:
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.5)
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press('enter')
                            elif index == 1:
                                time.sleep(1.5)
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.7)
                            else:
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.5)
        i = event.get(1)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    excel_files_name = []
    part_numbers = []
    for filename in os.listdir('ExcelFiles'):
        ext = filename.split('.')[-1]
        if ext == 'XLSX':

            p = Reader.extract_excel_part_numbers(os.path.join('ExcelFiles', filename))
            part_numbers.extend(p)

    print(len(part_numbers))

    # screenWidth, screenHeight = pyautogui.size()
    # print(screenWidth, screenHeight)
    #
    # currentMouseX, currentMouseY = pyautogui.position()
    # print(currentMouseX, currentMouseY)
    # pyperclip.copy('The text to be copied to the clipboard.')


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
