import json
import os
import time
from dataclasses import dataclass
from typing import List, Tuple, Union, Optional
from dacite import from_dict

import pandas as pd
import pyautogui
import pynput
import pyperclip

from tools.ExcelApplication import ExcelOperator
from tools.PdfApplication import PdfApplication
from tools.SAPApplication import SAPApplication
from tools.RequestAPI import API


@dataclass
class Script:
    action: str
    coordinates: tuple
    time_sleep: float = 0


@dataclass
class IDAndDeltaPN:
    id: int
    deltaPN: str


@dataclass
class TenCodeInfos(IDAndDeltaPN):
    modelName: str
    magName: str
    eeName: str
    samples: List[dict]
    remarks: List[str]
    date: str = None
    rev: str = None


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

def sap_specs_auto_process(part_number, data) -> Tuple[str, Optional[TenCodeInfos], Optional[str]]:

    sap = SAPApplication()
    sap.OpenSAP()
    sap.LogInMenu()
    session = sap.Connect2SAP_API()
    sap.WindowResize(session)
    sap.LogInUser(session)
    error_msg = sap.DownloadBOMForm(session, part_number)
    if error_msg:

        return error_msg, None, None
    error_msg = sap.DownloadSpecs(session, part_number)
    if error_msg:
        return error_msg, None, None

    # 防止文件還未下載完成
    time.sleep(5)
    path = r'D:\DeltaBox\OneDrive - Delta Electronics, Inc\文件\SAP\SAP GUI\{}.xls'.format(part_number)
    boom_pdf_path, infos_dataclass = ExcelOperator(path, part_number, data)

    specs_pdf_path_dict_list = [dict(path=os.path.join(r'C:\Users\amo.cy.hsu\Downloads', path),
                                     time=os.path.getctime(os.path.join(r'C:\Users\amo.cy.hsu\Downloads', path)))
                                for path in os.listdir(r'C:\Users\amo.cy.hsu\Downloads') if
                                path.split('.')[-1] == 'pdf']
    newest_specs_pdf = max(specs_pdf_path_dict_list, key=lambda d: d['time'])
    newest_specs_pdf_path = newest_specs_pdf['path']
    pdf_application = PdfApplication()
    path = pdf_application.ReadPDF(newest_specs_pdf_path)
    start_index, end_index = pdf_application.JudgeSpecs(path)

    pdf_application.split_pdf(newest_specs_pdf_path, start_index, end_index, part_number+'_split.pdf')
    merge_pdf_list = [boom_pdf_path, part_number+'_split.pdf']
    output_pdf = pdf_application.merge_pdf(merge_pdf_list, '{}_{}.pdf'.format(str(data.id), part_number))
    output_pdf = os.path.join(os.getcwd(), output_pdf)
    # pdf_application.PDFPrinter(output_pdf)

    return 'successes', infos_dataclass, output_pdf


def fetch_non_processed_data():

    ten_code_data_list: List[TenCodeInfos] = [from_dict(data_class=TenCodeInfos, data=d) for d in API.GetApplications()]

    for ten_code_data in ten_code_data_list:
        # print(ten_code_data.deltaPN)
        ten_code_data.deltaPN = ten_code_data.deltaPN[:-2]
        msg, infos, output_file_path = sap_specs_auto_process(ten_code_data.deltaPN, ten_code_data)

        if msg == 'successes':
            state = 1
            reason = ''
            multipart_form_data = {
                'file': (output_file_path, open(output_file_path, 'rb')),
                'revision': (None, infos.rev),
                'state': (None, state),
                'reason': (None, reason)
            }

        else:
            state = 0
            reason = msg
            multipart_form_data = {
                'file': (None, None),
                'revision': (None, None),
                'state': (None, state),
                'reason': (None, reason)
            }
        # update data
        API.PostApplicationData(ten_code_data.id, multipart_form_data)

def fetch_non_printed_data():

    non_printed_ten_code_data_list: List[IDAndDeltaPN] = [from_dict(data_class=IDAndDeltaPN, data=d) for d in API.GetCheckedApplications()]

    for non_printed_ten_code_data in non_printed_ten_code_data_list:
        pdf_path = os.path.join(os.getcwd(), 'TenCodeMergePDF',
                                '{}_{}.pdf'.format(non_printed_ten_code_data.id, non_printed_ten_code_data.deltaPN))
        print(pdf_path)
        if os.path.exists(pdf_path):
            PdfApplication.PDFPrinter(pdf_path)
            API.PostApplicationPrinted(non_printed_ten_code_data.id)


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
                        time.sleep(8)
                        pyperclip.copy(part_number)

                        for index, mouse_coordinate in enumerate(mouse_coordinates):
                            coordinate_x = mouse_coordinate[0]
                            coordinate_y = mouse_coordinate[1]
                            if index == 0:

                                pyautogui.click(coordinate_x, coordinate_y, duration=0.3)
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press('enter')
                            elif index == 1:
                                time.sleep(3)
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.5)
                            else:
                                pyautogui.click(coordinate_x, coordinate_y, duration=0.5)
            if count == 3:
                break
        i = event.get(1)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # find_mouse_click_coordinates()
    #
    # import os
    # excel_files_name = []
    # part_numbers = []
    # key = 'Part No.'
    # for filename in os.listdir('ExcelFiles'):
    #     ext = filename.split('.')[-1]
    #     if filename != '123.XLSX':
    #         continue
    #     if ext == 'XLSX':
    #         p = Reader.extract_excel_part_numbers(os.path.join('ExcelFiles', filename), key)
    #         part_numbers.extend(p)
    # print((part_numbers))
    # start(part_numbers)
    #
    # excel_files_name = []
    # part_numbers = []
    # for filename in os.listdir('ExcelFiles'):
    #     ext = filename.split('.')[-1]
    #     if ext == 'xlsx' and 'download' in filename:
    #         p = Reader.extract_excel_part_numbers(os.path.join('ExcelFiles', filename), key='Part No.')
    #         part_numbers.extend(p)
    #

    fetch_non_printed_data()

    # import schedule
    # import time
    #
    #
    # def job():
    #     print("I'm working...")
    #
    # def job2():
    #     print("I'm working2...")
    #
    # schedule.every(10).seconds.do(job)
    # schedule.every(15).seconds.do(job2)
    #
    # while True:
    #     schedule.run_pending()
    #     time.sleep(1)