import cv2
import pytesseract
import numpy as np
import pandas as pd
from datetime import datetime
import time

# 设置Tesseract路径(根据你的安装位置调整)
pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'  # Windows示例

# 游戏余额区域坐标 (需要根据你的游戏调整)
BALANCE_REGION = (800, 100, 200, 50)  # (x, y, width, height)

# 初始化数据存储
data = {
    "Timestamp": [],
    "Spin Number": [],
    "Balance": [],
    "Change": []
}
spin_count = 0
previous_balance = None


def extract_balance(frame):

    """从游戏帧中提取余额"""
    # 截取余额区域
    x, y, w, h = BALANCE_REGION
    balance_roi = frame[y:y + h, x:x + w]

    # 图像预处理提高OCR识别率
    gray = cv2.cvtColor(balance_roi, cv2.COLOR_BGR2GRAY)
    _, threshold = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # OCR识别
    custom_config = r'--oem 3 --psm 6 outputbase digits'
    balance_text = pytesseract.image_to_string(threshold, config=custom_config)

    # 清理识别结果
    balance = ''.join(filter(str.isdigit, balance_text))
    return int(balance) if balance else None


def save_to_excel(data):
    """保存数据到Excel"""
    df = pd.DataFrame(data)
    filename = f"slot_balance_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Balance Data', index=False)
    writer.close()
    print(f"数据已保存到 {filename}")


def main():
    global spin_count, previous_balance

    # 初始化摄像头/屏幕捕获
    # 方法1: 使用屏幕捕获 (需要安装pyautogui)
    # import pyautogui
    # while True:
    #     screenshot = pyautogui.screenshot()
    #     frame = np.array(screenshot)
    #     frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)

    # 方法2: 使用视频文件 (用于测试)
    video_path = "slot_gameplay.mp4"  # 替换为你的视频文件路径
    cap = cv2.VideoCapture(video_path)

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            break

        current_balance = extract_balance(frame)

        if current_balance is not None:
            # 检测到余额变化(新Spin)
            if previous_balance is None or current_balance != previous_balance:
                spin_count += 1
                change = current_balance - previous_balance if previous_balance is not None else 0

                # 记录数据
                data["Timestamp"].append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                data["Spin Number"].append(spin_count)
                data["Balance"].append(current_balance)
                data["Change"].append(change)

                print(f"Spin #{spin_count}: 余额={current_balance}, 变化={change}")

                previous_balance = current_balance

        # 按'q'退出
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

    # 保存数据到Excel
    save_to_excel(data)


if __name__ == "__main__":
    main()