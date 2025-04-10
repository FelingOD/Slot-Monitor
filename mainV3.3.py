import cv2
import pytesseract
import numpy as np
import pandas as pd
from datetime import datetime
import time
import os

# 设置Tesseract路径
pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'

# 定义监控区域
BALANCE_REGION = (420, 55, 260, 95)  # 区域A：余额显示区域 (x, y, w, h)
COUNTER_REGION = (1670, 1010, 90, 50)  # 区域B：计数器区域

# 初始化数据存储
data = {
    "时间戳": [],
    "SPIN次数": [],
    "余额": [],
    "损耗": [],
    "剩余SPIN次数": []
}
spin_count = 0
previous_balance = None
previous_counter = None
screenshot_dir = "screenshots"  # 截图保存目录

# 创建截图目录
os.makedirs(screenshot_dir, exist_ok=True)

def extract_number(frame, region):
    """从指定区域提取数字，失败返回None"""
    try:
        x, y, w, h = region
        roi = frame[y:y + h, x:x + w]

        # 图像预处理
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        _, threshold = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # OCR识别
        custom_config = r'--oem 3 --psm 6 outputbase digits'
        number_text = pytesseract.image_to_string(threshold, config=custom_config)

        # 清理结果
        number = ''.join(filter(str.isdigit, number_text))
        return int(number) if number else None
    except Exception as e:
        print(f"识别数字时出错: {str(e)}")
        return None

def save_screenshot(frame, counter_value):
    """保存当前帧截图"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{screenshot_dir}/counter_{counter_value if counter_value is not None else 'null'}_{timestamp}.png"
        cv2.imwrite(filename, frame)
        print(f"已保存截图: {filename}")
    except Exception as e:
        print(f"保存截图失败: {str(e)}")

def save_to_excel(data):
    """保存数据到Excel"""
    try:
        df = pd.DataFrame(data)
        filename = f"D:\\GitHubProj\\Slot-Monitor\\slot数据采集源表{datetime.now().strftime('%Y%m%d')}.xlsx"
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Balance Data', index=False)
        writer.close()
        print(f"数据已保存到 {filename}")
    except Exception as e:
        print(f"保存Excel失败: {str(e)}")

# ...（前面的导入和设置保持不变）...

def main():
    global spin_count, previous_balance, previous_counter

    # 使用视频文件
    video_path = "videos\\4.10,23.49.500z.mp4"
    cap = cv2.VideoCapture(video_path)

    # 获取原始视频帧率
    original_fps = cap.get(cv2.CAP_PROP_FPS)
    # 设置3倍速播放的目标帧率
    target_fps = original_fps * 10
    # 计算每帧应延迟的时间（毫秒）
    frame_delay = int(1000 / target_fps) if target_fps > 0 else 1

    # 创建可调整大小的窗口
    cv2.namedWindow("Monitoring", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Monitoring", 800, 600)  # 初始大小

    # 性能优化：跳帧计数器
    frame_counter = 0
    skip_frames = 0  # 可根据需要调整跳帧数

    while cap.isOpened():
        start_time = time.time()

        ret, frame = cap.read()
        if not ret:
            break

        # 跳帧处理（可选，进一步提升速度）
        frame_counter += 1
        if frame_counter % (skip_frames + 1) != 0:
            continue

        # 提取两个区域的数值（增加None检查）
        current_balance_raw = extract_number(frame, BALANCE_REGION)
        current_counter = extract_number(frame, COUNTER_REGION)

        # ...（中间的处理逻辑保持不变）...

        # 显示监控画面（可调整大小）
        debug_frame = frame.copy()
        cv2.rectangle(debug_frame,
                      (BALANCE_REGION[0], BALANCE_REGION[1]),
                      (BALANCE_REGION[0] + BALANCE_REGION[2], BALANCE_REGION[1] + BALANCE_REGION[3]),
                      (0, 255, 0), 2)
        cv2.rectangle(debug_frame,
                      (COUNTER_REGION[0], COUNTER_REGION[1]),
                      (COUNTER_REGION[0] + COUNTER_REGION[2], COUNTER_REGION[1] + COUNTER_REGION[3]),
                      (0, 0, 255), 2)

        # 按比例缩小显示
        display_frame = cv2.resize(debug_frame, None, fx=0.5, fy=0.5)  # 缩小50%
        cv2.imshow("Monitoring", display_frame)

        # 计算处理耗时，调整延迟时间
        processing_time = (time.time() - start_time) * 1000  # 转为毫秒
        adjusted_delay = max(1, frame_delay - int(processing_time))

        # 添加键盘控制
        key = cv2.waitKey(adjusted_delay) & 0xFF
        if key == ord('q'):
            break
        elif key == ord('+'):  # 按+放大
            cv2.resizeWindow("Monitoring", int(display_frame.shape[1] * 1.1), int(display_frame.shape[0] * 1.1))
        elif key == ord('-'):  # 按-缩小
            cv2.resizeWindow("Monitoring", int(display_frame.shape[1] * 0.9), int(display_frame.shape[0] * 0.9))

    cap.release()
    cv2.destroyAllWindows()
    save_to_excel(data)


if __name__ == "__main__":
    main()