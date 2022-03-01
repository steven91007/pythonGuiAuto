import cv2
import numpy as np
img_path = '414851166100-page-003.jpg'
img = cv2.imdecode(np.fromfile(img_path, dtype=np.uint8), -1)

import paddle
paddle.utils.run_check()

from paddleocr import PaddleOCR

from paddleocr.tools.infer.utility import draw_ocr_box_txt

# Paddleocr目前支持的多语言语种可以通过修改lang参数进行切换
# 例如`ch`, `en`, `fr`, `german`, `korean`, `japan`
ocr = PaddleOCR(use_angle_cls=True, lang="ch")  # need to run only once to download and load model into memory
result = ocr.ocr(img, cls=True)
for line in result:
    print(line)

# 显示结果
from PIL import Image

image = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))

boxes = [map(tuple, line[0]) for line in result]
txts = [line[1][0] for line in result]
scores = [line[1][1] for line in result]

print(boxes)