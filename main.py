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
    def extract_excel_part_numbers(path, key):
        df = pd.read_excel(path)
        part_numbers = [str(d) for d in df.loc[:, key]]
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
        for count, i in enumerate(event):
            # 迭代用法。
            if isinstance(i, pynput.keyboard.Events.Press):
                # 滑鼠點擊事件。
                if i.key == pynput.keyboard.Key.esc:
                    break
                elif i.key == pynput.keyboard.KeyCode.from_char('s'):
                    for part_index, part_number in enumerate(part_numbers):
                        pyperclip.copy(part_number)

                        for index, mouse_coordinate in enumerate(mouse_coordinates):
                            coordinate_x = mouse_coordinate[0]
                            coordinate_y = mouse_coordinate[1]
                            if index == 0:
                                pyautogui.press('enter')
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.3)
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press('enter')
                            elif index == 1:
                                time.sleep(1.5)
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.5)
                            else:
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.5)

        i = event.get(1)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # find_mouse_click_coordinates()
    # pyautogui.alert('This displays some text with an OK button.')
    # excel_files_name = []
    # part_numbers = []
    # key = 'PN'
    # for filename in os.listdir('ExcelFiles'):
    #     ext = filename.split('.')[-1]
    #     if ext == 'xlsm':
    #         p = Reader.extract_excel_part_numbers(os.path.join('ExcelFiles', filename), key)
    #         part_numbers.extend(p)
    # print((part_numbers))
    # start(part_numbers)

    excel_files_name = []
    part_numbers = []
    for filename in os.listdir('ExcelFiles'):
        ext = filename.split('.')[-1]
        if ext == 'XLSX':
            p = Reader.extract_excel_part_numbers(os.path.join('ExcelFiles', filename), key='Part No.')
            part_numbers.extend(p)

    pdf_path = r'C:\Users\AMO.CY.HSU\Documents\磁性元件_spec'
    filename_list = []
    for filename in os.listdir(pdf_path):
        filename_exclude_ext = filename.split('.')[0]
        filename_list.append(filename_exclude_ext[:10])
    lost_part_number = []

    for part_number in part_numbers:
        if part_number not in filename_list:
            lost_part_number.append(part_number)
    last = '4149020890'
    print("4149020890 索引位置: ", lost_part_number.index('4149020890'))
    lost_part_number = lost_part_number[lost_part_number.index('4149020890'):]
    print(lost_part_number)
    start(lost_part_number)
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
