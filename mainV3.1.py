import cv2
import pytesseract
import numpy as np
import pandas as pd
from datetime import datetime
import time
import os

# 设置Tesseract路径
pytesseract.pytesseract.tesseract_cmd = r'E:\Tesseract-OCR\tesseract.exe'

# 定义两个监控区域
BALANCE_REGION = (460, 55, 250, 95)  # 区域A：余额显示区域 (x, y, w, h)
COUNTER_REGION = (1720, 1020, 50, 50)  # 区域B：计数器区域

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
    """从指定区域提取数字"""
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


def save_screenshot(frame, counter_value):
    """保存当前帧截图"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{screenshot_dir}/counter_{counter_value}_{timestamp}.png"
    cv2.imwrite(filename, frame)
    print(f"已保存截图: {filename}")


def save_to_excel(data):
    """保存数据到Excel"""
    df = pd.DataFrame(data)
    filename = f"E:\\GitHubProj\\Slot-Monitor\\slot数据采集原表{datetime.now().strftime('%Y%m%d')}.xlsx"
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Balance Data', index=False)
    writer.close()
    print(f"数据已保存到 {filename}")


def main():
    global spin_count, previous_balance, previous_counter

    # 使用视频文件
    video_path = "videos\\test1.mp4"    # 视频绝对地址
    cap = cv2.VideoCapture(video_path)

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            break

        # 提取两个区域的数值
        current_balance = extract_number(frame, BALANCE_REGION)+30000
        current_counter = extract_number(frame, COUNTER_REGION)

        if current_counter is not None:
            # 检测计数器是否减少了1
            if previous_counter is not None and current_counter == previous_counter - 1:
                print(f"计数器减少1: {previous_counter} → {current_counter}")
                save_screenshot(frame, current_counter)

                # 记录余额变化
                if current_balance is not None:
                    spin_count += 1
                    change = current_balance - previous_balance if previous_balance is not None else 0

                    data["Timestamp"].append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    data["Spin Number"].append(spin_count)
                    data["Balance"].append(current_balance)
                    data["Change"].append(change)
                    data["Counter Value"].append(current_counter)

                    print(f"Spin #{spin_count}: 余额={current_balance}, 变化={change}, 计数器={current_counter}")

            previous_counter = current_counter

        if current_balance is not None:
            previous_balance = current_balance

        # 显示监控画面（调试用）
        debug_frame = frame.copy()
        cv2.rectangle(debug_frame,
                      (BALANCE_REGION[0], BALANCE_REGION[1]),
                      (BALANCE_REGION[0] + BALANCE_REGION[2], BALANCE_REGION[1] + BALANCE_REGION[3]),
                      (0, 255, 0), 2)
        cv2.rectangle(debug_frame,
                      (COUNTER_REGION[0], COUNTER_REGION[1]),
                      (COUNTER_REGION[0] + COUNTER_REGION[2], COUNTER_REGION[1] + COUNTER_REGION[3]),
                      (0, 0, 255), 2)

        cv2.imshow("Monitoring", debug_frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    save_to_excel(data)


if __name__ == "__main__":
    main()