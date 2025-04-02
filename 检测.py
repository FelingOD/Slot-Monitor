import paddleocr
import paddle
import cv2
import numpy as np
import pandas as pd

print(f"PaddlePaddle 版本: {paddle.__version__}")  # 应输出 3.0.0
print(f"OpenCV 版本: {cv2.__version__}")          # 应输出 4.11.0
paddle.utils.run_check()  # 检查 PaddlePaddle 环境是否正常