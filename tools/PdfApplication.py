import os
from tools import table_structure_recognition_lines as tsrl
from utility import pdf2png
import cv2
from pikepdf import Pdf
import shutil
import win32print
import subprocess


class PdfApplication:

    def ReadPDF(self, pdf_path):

        return pdf2png(pdf_path, 4)

    def JudgeSpecs(self, png_path_list):
        start_index = 0
        end_index = 0
        for index, png_path in enumerate(png_path_list):

            img = cv2.imread(png_path)
            _, flag = tsrl.recognize_structure(img, str(index))
            print(flag)

            if not flag:
                if start_index == 0:
                    start_index = index

                if index+1 == len(png_path_list):
                    end_index = index
                    break

                next_image = cv2.imread(png_path_list[index+1])
                _, next_flag = tsrl.recognize_structure(next_image, '')

                if next_flag:
                    end_index = index
                    break
        folder_path = os.path.join(*png_path_list[0].split(os.sep)[:-1])
        shutil.rmtree(folder_path)
        return start_index, end_index

    def split_pdf(self, file_name, start_page, end_page, output_pdf):

        pdf = Pdf.open(file_name)

        dst = Pdf.new()
        for n, page in enumerate(pdf.pages):
            if start_page <= n <= end_page:
                dst.pages.append(page)
        dst.save(output_pdf)

    def merge_pdf(self, merge_pdf_path_list, output_pdf):
        """
        merge_list: 需要合併的pdf列表
        output_pdf：合併之後的pdf名
        """
        # 實例一個 PDF文件編寫器
        new_pdf = Pdf.new()
        for pdf_path in merge_pdf_path_list:

            src = Pdf.open(pdf_path)
            new_pdf.pages.extend(src.pages)
        new_pdf.save(output_pdf)
        return output_pdf

    def PDFPrinter(self, pdf_path):

        acrobat = r'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe'
        name = win32print.GetDefaultPrinter()
        cmd = '"{}" /n /o /t "{}" "{}"'.format(acrobat, pdf_path, name)
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


if __name__ == '__main__':
    pdf_path = r'C:\Users\amo.cy.hsu\PycharmProjects\pythonGuiAuto\tools\287070330208.pdf'
    pdf_application = PdfApplication()
    pdf_application.PDFPrinter('')
    # path = pdf_application.ReadPDF(pdf_path)
    # start_index, end_index = pdf_application.JudgeSpecs(path)
    # print(start_index, end_index)
    # pdf_application.split_pdf(pdf_path, start_index, end_index, 'split.pdf')
