import time
import pyautogui
import pytesseract
from PIL import Image
import cv2
import numpy as np

# 配置
CLICK_POSITION = (1600, 1050)  # 点击位置的坐标 (x, y)
SCAN_REGION = (420,60,660,150)  # 数字识别区域 (x, y, width, height)
INTERVAL_SECONDS = 5  # 每次操作的间隔时间(秒)
LOG_FILE = "number_log.txt"  # 记录文件


# 设置pytesseract路径（如果需要）
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def click_position():
    """点击指定位置"""
    pyautogui.click(CLICK_POSITION[0], CLICK_POSITION[1])
    print(f"已点击位置: {CLICK_POSITION}")


def capture_region():
    """截取指定区域"""
    x, y, w, h = SCAN_REGION
    screenshot = pyautogui.screenshot(region=(x, y, w, h))
    return screenshot


def preprocess_image(image):
    """图像预处理以提高OCR识别率"""
    # 转换为灰度图
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)

    # 二值化处理
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # 去噪
    kernel = np.ones((1, 1), np.uint8)
    processed = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

    return Image.fromarray(processed)


def extract_numbers(image):
    """从图像中提取数字"""
    # 预处理图像
    processed_img = preprocess_image(image)

    # 使用OCR识别数字
    custom_config = r'--oem 3 --psm 6 outputdigits'
    text = pytesseract.image_to_string(processed_img, config=custom_config)

    # 清理识别结果，只保留数字
    numbers = ''.join(filter(str.isdigit, text))
    return numbers if numbers else "未识别到数字"


def log_numbers(numbers):
    """记录数字到文件"""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"{timestamp} - 识别到的数字: {numbers}\n"

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(log_entry)

    print(log_entry.strip())


def main():
    print("程序开始运行...")
    print(f"配置: 每{INTERVAL_SECONDS}秒点击一次位置{CLICK_POSITION}")
    print(f"数字识别区域: {SCAN_REGION}")
    print(f"记录文件: {LOG_FILE}")
    print("按Ctrl+C终止程序")

    try:
        while True:
            # 点击指定位置
            click_position()

            # 等待短暂时间让界面更新
            time.sleep(1)

            # 截取并识别数字
            screenshot = capture_region()
            numbers = extract_numbers(screenshot)

            # 记录结果
            log_numbers(numbers)

            # 等待下一次操作
            time.sleep(INTERVAL_SECONDS - 1)  # 减去之前等待的1秒

    except KeyboardInterrupt:
        print("\n程序已终止")


if __name__ == "__main__":
    main()