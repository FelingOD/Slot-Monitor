import configparser
import os
import time
import cv2
import numpy as np
import pyautogui
import pandas as pd
from datetime import datetime
from paddleocr import PaddleOCR
import win32gui
import win32con


class SlotRecorder:
    def __init__(self):
        self.load_config()
        self.ocr = PaddleOCR(
            use_angle_cls=True,
            lang="en",
            det_model_dir='en_PP-OCRv3_det',
            rec_model_dir='en_PP-OCRv3_rec',
            use_gpu=False
        )
        self.setup_directories()

    def load_config(self):
        """加载配置文件"""
        self.config = configparser.ConfigParser()
        if not os.path.exists('config.ini'):
            self.create_default_config()
        self.config.read('config.ini')

        # 读取区域设置
        self.balance_area = eval(self.config.get('REGIONS', 'BalanceArea'))
        self.spin_button_area = eval(self.config.get('REGIONS', 'SpinButtonArea'))
        self.spin_active_color = eval(self.config.get('COLORS', 'SpinActiveColor'))
        self.spin_cooldown_color = eval(self.config.get('COLORS', 'SpinCooldownColor'))
        self.window_title = self.config.get('SETTINGS', 'WindowTitle')

    def create_default_config(self):
        """创建默认配置文件"""
        self.config['REGIONS'] = {
            'BalanceArea': '(100, 200, 300, 250)',
            'SpinButtonArea': '(400, 500, 500, 550)'
        }
        self.config['COLORS'] = {
            'SpinActiveColor': '(70, 70, 255)',
            'SpinCooldownColor': '(100, 100, 100)'
        }
        self.config['SETTINGS'] = {
            'WindowTitle': 'BlueStacks',
            'ScreenshotPath': 'screenshots',
            'RecordPath': 'records'
        }
        with open('config.ini', 'w') as f:
            self.config.write(f)

    def setup_directories(self):
        """创建必要的目录"""
        os.makedirs(self.config.get('SETTINGS', 'ScreenshotPath'), exist_ok=True)
        os.makedirs(self.config.get('SETTINGS', 'RecordPath'), exist_ok=True)

    def focus_window(self):
        """将模拟器窗口置顶"""
        hwnd = win32gui.FindWindow(None, self.window_title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            return True
        return False

    def capture_area(self, area, filename=None):
        """截取指定区域"""
        self.focus_window()
        try:
            screenshot = pyautogui.screenshot(region=area)
            if filename:
                screenshot.save(os.path.join(
                    self.config.get('SETTINGS', 'ScreenshotPath'),
                    filename
                ))
            return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        except Exception as e:
            print(f"截图失败: {e}")
            return None

    def preprocess_image(self, img):
        """图像预处理"""
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
        return binary

    def extract_number(self, img):
        """从图像中提取数字"""
        processed = self.preprocess_image(img)
        result = self.ocr.ocr(processed, cls=True)

        numbers = []
        for line in result:
            for word_info in line:
                text = word_info[1][0]
                filtered = ''.join(c for c in text if c.isdigit())
                if filtered:
                    numbers.append(filtered)

        try:
            return int(''.join(numbers)) if numbers else None
        except:
            return None

    def is_spin_available(self):
        """检测Spin按钮是否可用"""
        spin_img = self.capture_area(self.spin_button_area)
        if spin_img is None:
            return False

        h, w = spin_img.shape[:2]
        center_color = spin_img[h // 2, w // 2]
        center_rgb = (center_color[2], center_color[1], center_color[0])

        def color_distance(c1, c2):
            return sum((a - b) ** 2 for a, b in zip(c1, c2)) ** 0.5

        active_dist = color_distance(center_rgb, self.spin_active_color)
        cooldown_dist = color_distance(center_rgb, self.spin_cooldown_color)

        return active_dist < cooldown_dist

    def perform_spin(self):
        """执行Spin操作"""
        center_x = (self.spin_button_area[0] + self.spin_button_area[2]) // 2
        center_y = (self.spin_button_area[1] + self.spin_button_area[3]) // 2

        # 自然移动鼠标并点击
        pyautogui.moveTo(center_x, center_y, duration=0.5)
        pyautogui.click()
        time.sleep(0.2)

    def wait_for_spin_completion(self, timeout=15):
        """等待Spin完成"""
        start_time = time.time()
        last_balance = None
        stable_count = 0

        while time.time() - start_time < timeout:
            current_img = self.capture_area(self.balance_area)
            current_balance = self.extract_number(current_img)

            if current_balance:
                if last_balance == current_balance:
                    stable_count += 1
                    if stable_count >= 2:  # 连续2次相同认为稳定
                        return current_balance
                else:
                    stable_count = 0

                last_balance = current_balance

            time.sleep(0.5)

        return last_balance

    def record_spin(self, spin_number):
        """记录一次Spin"""
        # 1. 获取Spin前余额
        before_img = self.capture_area(
            self.balance_area,
            f"before_{spin_number}.png"
        )
        balance_before = self.extract_number(before_img)

        if balance_before is None:
            print("无法识别Spin前余额!")
            return None

        print(f"[Spin #{spin_number}] 前余额: {balance_before}")

        # 2. 等待Spin可用并执行
        print("等待Spin按钮可用...")
        while not self.is_spin_available():
            time.sleep(0.5)

        print("执行Spin...")
        self.perform_spin()

        # 3. 等待完成并获取结果
        print("等待Spin完成...")
        balance_after = self.wait_for_spin_completion()

        if balance_after is None:
            print("无法识别Spin后余额!")
            return None

        # 保存Spin后截图
        self.capture_area(
            self.balance_area,
            f"after_{spin_number}.png"
        )

        print(f"Spin后余额: {balance_after}")

        # 计算并记录结果
        change = balance_after - balance_before
        result = "Win" if change > 0 else "Loss" if change < 0 else "Even"

        return {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "spin_number": spin_number,
            "balance_before": balance_before,
            "balance_after": balance_after,
            "change": change,
            "result": result
        }

    def save_record(self, record):
        """保存记录到CSV"""
        filename = os.path.join(
            self.config.get('SETTINGS', 'RecordPath'),
            f"records_{datetime.now().strftime('%Y%m%d')}.csv"
        )

        df = pd.DataFrame([record])
        if not os.path.exists(filename):
            df.to_csv(filename, index=False)
        else:
            df.to_csv(filename, mode='a', header=False, index=False)

        print(f"记录已保存到 {filename}")

    def calibrate(self):
        """校准工具"""
        print("\n=== 校准模式 ===")
        print("1. 确保游戏窗口在最前且可见")

        # 1. 校准余额区域
        input("按Enter键截取全屏(用于设置余额区域)...")
        full_img = pyautogui.screenshot()
        full_img.save("calibration_full.png")
        print("全屏截图已保存为 calibration_full.png")

        print("\n请用图片查看器打开 calibration_full.png")
        print("确定金币余额显示区域的坐标 (left, top, right, bottom)")
        balance_coords = input("请输入余额区域坐标(格式:100,200,300,250): ")
        self.config.set('REGIONS', 'BalanceArea', balance_coords)

        # 2. 校准Spin按钮
        print("\n确定Spin按钮区域的坐标 (left, top, right, bottom)")
        spin_coords = input("请输入Spin按钮区域坐标(格式:400,500,500,550): ")
        self.config.set('REGIONS', 'SpinButtonArea', spin_coords)

        # 3. 校准Spin按钮颜色
        print("\n将鼠标移动到Spin按钮中心位置")
        print("5秒内请将鼠标移动到激活状态的Spin按钮上...")
        time.sleep(5)
        x, y = pyautogui.position()
        active_color = pyautogui.screenshot().getpixel((x, y))
        self.config.set('COLORS', 'SpinActiveColor', str(active_color))
        print(f"Spin按钮激活颜色设置为: {active_color}")

        # 4. 校准冷却颜色
        input("\n请执行一次Spin，等待按钮变灰后按Enter键...")
        x, y = pyautogui.position()
        cooldown_color = pyautogui.screenshot().getpixel((x, y))
        self.config.set('COLORS', 'SpinCooldownColor', str(cooldown_color))
        print(f"Spin按钮冷却颜色设置为: {cooldown_color}")

        # 保存配置
        with open('config.ini', 'w') as f:
            self.config.write(f)

        print("\n校准完成! 配置已保存到 config.ini")

    def run(self, max_spins=100):
        """主运行循环"""
        spin_count = 0
        print(f"\n自动记录模式启动，最多记录 {max_spins} 次Spin")
        print("按下Ctrl+C停止记录")

        try:
            while spin_count < max_spins:
                spin_count += 1
                record = self.record_spin(spin_count)

                if record:
                    self.save_record(record)
                    # 随机延迟防止检测
                    delay = 1 + np.random.uniform(0, 2)
                    print(f"等待 {delay:.1f} 秒...\n")
                    time.sleep(delay)
                else:
                    time.sleep(2)
        except KeyboardInterrupt:
            print("\n记录已停止")
        finally:
            print(f"共完成 {spin_count} 次记录")


if __name__ == "__main__":
    recorder = SlotRecorder()

    print("=== Slot游戏自动记录系统 ===")
    print("1. 校准区域和颜色设置")
    print("2. 开始自动记录")
    choice = input("请选择(1/2): ")

    if choice == "1":
        recorder.calibrate()
    elif choice == "2":
        max_spins = int(input("输入最大记录次数(默认100): ") or 100)
        recorder.run(max_spins)