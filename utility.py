import fitz
import os
import cv2
import paddle
from paddleocr import PaddleOCR
from paddleocr.tools.infer.utility import draw_ocr_box_txt
import numpy as np
from PIL import Image
import json
from typing import List, Union, Tuple


def pdf2png(pdf_file_path: str, zoom: int, return_type='path') -> list:
    image_root_name = pdf_file_path.split(os.sep)[-1].split('.')[0]
    doc = fitz.open(pdf_file_path)
    zoom = zoom
    trans = fitz.Matrix(zoom, zoom)
    print(image_root_name)
    print("%s开始转换..." % pdf_file_path)
    png_path_list = []
    if doc.pageCount > 0:  # 获取PDF的页数
        for pg in range(doc.pageCount):
            print('fileName', image_root_name)
            page = doc[pg]  # 获得第pg頁
            pm = page.get_pixmap(matrix=trans, alpha=False)  # 将其转化为光栅文件（位数）
            base_folder = 'PngFile'
            folder_path = os.path.join(base_folder, image_root_name)
            if not os.path.exists(folder_path):  # 判断存放图片的文件夹是否存在
                os.makedirs(folder_path)  # 若图片文件夹不存在就创建

            path = os.path.join(folder_path, image_root_name + "_" + str(pg + 1) + ".png")
            print('path:', path)
            pm.save(path)
            png_path_list.append(path)

    if return_type == 'path':
        return png_path_list
    elif return_type == 'nparray':
        nparray_list = []
        for png_path in png_path_list:
            img = cv2.imdecode(np.fromfile(png_path, dtype=np.uint8), -1)
            nparray_list.append(img)
        return nparray_list


def batch_infer(img_list, args):

    rec_model_name = args['rec_model_name']
    rec_char_dict_path = args['rec_char_dict_path']
    ocr = PaddleOCR(lang='ch',
                    rec_char_type='ch',
                    drop_score=0.2,
                    use_angle_cls=True,
                    use_space_char=True,
                    rec_char_dict_path=rec_char_dict_path,
                    rec_model_dir=os.path.join('OCR_Models', rec_model_name)
                    )

    for i, img_path in enumerate(img_list):
        if i == 5: break
        img = cv2.imdecode(np.fromfile(img_path, dtype=np.uint8), -1)
        img_infer_path = img_path.replace('PngFile', os.path.join('inference_results', rec_model_name))
        img_infer_folder_path = os.path.join(*img_infer_path.split(os.sep)[:3])
        if not os.path.exists(img_infer_folder_path):  # 判断存放图片的文件夹是否存在
            os.makedirs(img_infer_folder_path)  # 若图片文件夹不存在就创建

        result = ocr.ocr(img, cls=True)
        for line in result:
            print(line)

        # 显示结果
        image = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
        boxes = [map(tuple, line[0]) for line in result]
        txts = [line[1][0] for line in result]
        scores = [line[1][1] for line in result]
        array_of_tuples = map(tuple, boxes)
        boxes = tuple(array_of_tuples)
        im_show = draw_ocr_box_txt(image,
                                   tuple(boxes),
                                   txts,
                                   scores,
                                   font_path='./simfang.ttf',
                                   drop_score=0.2)
        # im_show = Image.fromarray(im_show)
        # buffer = BytesIO()
        # im_show.save(buffer, format="JPEG")  # "'./inference_results/'+filename)
        json_path = img_infer_path.replace('png', 'json')
        cv2.imencode('.png', im_show)[1].tofile(img_infer_path)
        f = open(json_path, "w")
        json_format = json.dumps(str(result))
        f.write(json_format)
        f.close()


def read_json(path):

    # Opening JSON file
    f = open(path)

    # returns JSON object as
    # a dictionary
    data = json.load(f)
    return data


def IsInclude(TableBBox: Union[List[int], Tuple[int]], WordBBox: Union[List[int], Tuple[int]], buffer=0):

    if TableBBox[0]-buffer < WordBBox[0] and TableBBox[2]+buffer > WordBBox[2] and TableBBox[1]-buffer < WordBBox[1] and TableBBox[3]+buffer > WordBBox[3]:
        return True

    return False

class DataPurify:

    def OnlyFloat(self, text) -> float:

        digit_text = ''
        for t in text:
            if t.isdigit() or t=='.':
                digit_text += t

        return float(digit_text)

if __name__ == '__main__':
    import pandas as pd

    df = pd.DataFrame()
    dict = {'First Name': 'Vikram', 'Last Name': 'Aruchamy', 'Country': 'India'}

    df = df.append(dict, ignore_index=True)

    df2 = pd.DataFrame({'First Name': ['Kumar'],
                        'Last Name': ['Ram'],})

    df = df.append(df2, ignore_index=False)

    print(df)


