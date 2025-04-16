import cv2
import pytesseract
import numpy as np
import pandas as pd
from datetime import datetime
import time
import os

# 设置Tesseract路径
# pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = r'E:\Tesseract-OCR\tesseract.exe'

# 定义监控区域
BALANCE_REGION = (395, 55, 285, 80)  # 区域A：余额显示区域 (x, y, w, h)
COUNTER_REGION = (1680, 1010, 95, 50)  # 区域B：计数器区域

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

def save_to_excel(data):
    """保存数据到Excel"""
    try:
        df = pd.DataFrame(data)
        # filename = f"D:\\GitHubProj\\Slot-Monitor\\documents\\slot数据采集源表{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        filename = f"E:\\GitHubProj\\Slot-Monitor\\documents\\slot数据采集源表{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Balance Data', index=False)
        writer.close()
        print(f"数据已保存到 {filename}")
    except Exception as e:
        print(f"保存Excel失败: {str(e)}")

def main():
    global spin_count, previous_balance, previous_counter

    # 使用视频文件
    video_path = "videos\\416,1437,500z.mp4"
    cap = cv2.VideoCapture(video_path)

    # 创建可调整大小的窗口
    cv2.namedWindow("Monitoring", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Monitoring", 800, 600)  # 初始大小

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            break

        # 提取两个区域的数值（增加None检查）
        current_balance_raw = extract_number(frame, BALANCE_REGION)
        current_counter = extract_number(frame, COUNTER_REGION)

        # 处理余额数据（如果识别失败则设为None）
        current_balance = current_balance_raw + 600000 if current_balance_raw is not None else None

        if current_counter is not None:
            # 修改点：检测计数器是否有变化（原：仅检测减少1）
            if previous_counter is not None and current_counter != previous_counter:
                print(f"计数器变化: {previous_counter} → {current_counter}")

                # 记录数据（所有字段都支持None值）
                spin_count += 1
                change = (current_balance - previous_balance) if (
                            current_balance is not None and previous_balance is not None) else None

                data["时间戳"].append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                data["SPIN次数"].append(spin_count)
                data["余额"].append(current_balance)
                data["损耗"].append(change)
                data["剩余SPIN次数"].append(current_counter)

                # 增强的输出信息（显示null值）
                balance_display = current_balance if current_balance is not None else "null"
                change_display = change if change is not None else "null"
                counter_display = current_counter if current_counter is not None else "null"
                print(f"Spin #{spin_count}: 余额={balance_display}, 变化={change_display}, 计数器={counter_display}")

            previous_counter = current_counter if current_counter is not None else previous_counter

        # 更新余额历史（仅当识别成功时更新）
        if current_balance is not None:
            previous_balance = current_balance

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

        # 添加键盘控制
        key = cv2.waitKey(1) & 0xFF
        if key == ord('q'):
            break
        elif key == ord('+'):  # 按+放大
            cv2.resizeWindow("Monitoring", int(display_frame.shape[1]*1.1), int(display_frame.shape[0]*1.1))
        elif key == ord('-'):  # 按-缩小
            cv2.resizeWindow("Monitoring", int(display_frame.shape[1]*0.9), int(display_frame.shape[0]*0.9))

    cap.release()
    cv2.destroyAllWindows()
    save_to_excel(data)

if __name__ == "__main__":
    main()