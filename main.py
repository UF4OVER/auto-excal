import openpyxl
from pynput.mouse import Listener, Button
import pyautogui
import time
import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLabel, QFileDialog

# 确保mian_ui模块和ExcelReaderApp类已经定义并且可用
import mian_ui

# 初始化应用和主窗口
app = QApplication(sys.argv)
ex = mian_ui.ExcelReaderApp()
ex.show()
sys.exit(app.exec_())
# 在这里定义read_excel_data函数，因为它是全局使用的
def selectFile(self):
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
    if file_path:
        file_name = os.path.basename(file_path)  # 获取文件名
        self.file_name = file_name
        print("Selected file:", file_name)  # 打印文件名
        self.file_label.setText(file_name)  # 更新文件名标签
        data_f, data_h = self.readExcelData(file_path)  # 传入完整路径
        # 打印数据
        print("F Column Data:", data_f)
        print("H Column Data:", data_h)

# 鼠标监听器
def on_click(x, y, button, pressed):
    global last_click_time, data_index, data
    if not pressed and button == Button.left:
        current_time = time.time()
        if current_time - last_click_time < double_click_interval:
            print(f'Double click detected at ({x}, {y})')
            if data_index < len(data):
                f_data, h_data = data[data_index]
                # 填入F列数据
                pyautogui.typewrite(str(f_data))
                # 模拟Tab键
                pyautogui.press('tab')
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(0.5)

                pyautogui.typewrite(str(h_data))

                data_index += 1
        else:
            last_click_time = current_time

# 初始化变量
last_click_time = 0
double_click_interval = 0.3  # 双击的时间间隔，单位为秒
data_index = 0
data = read_excel_data()

# 监听鼠标事件
with Listener(on_click=on_click) as listener:
    listener.join()


