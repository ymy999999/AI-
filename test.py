import sys
import os
import time
import ctypes
import base64
import json
import threading
import queue
import requests
import psutil
import win32gui
import win32ui
import win32process
from ctypes import wintypes
from PIL import Image
import pytesseract
import numpy as np
import cv2
import pygetwindow as gw
from pynput import keyboard  # 用于监听键盘事件

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QLabel, QPushButton, QVBoxLayout,
    QWidget, QMessageBox, QFileDialog, QHBoxLayout, QListWidget, QListWidgetItem, QSlider,
    QTabWidget, QFormLayout, QLineEdit, QComboBox, QSpinBox, QGridLayout, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QRect, QTimer, QObject, QPoint
from PyQt5.QtGui import QFont, QTextCursor, QTextCharFormat, QColor, QBrush, QIcon, QPalette, QImage, QPixmap

from bs4 import BeautifulSoup  # 用于解析HTML

# ----------------------------- Window Management ----------------------------- #

def find_all_windows():
    """
    查找所有运行中的应用程序窗口的句柄，无论其是否可见或最小化。
    """
    hwnds = []

    def enum_windows_callback(hwnd, lParam):
        window_title = win32gui.GetWindowText(hwnd)
        if window_title:
            hwnds.append(hwnd)

    win32gui.EnumWindows(enum_windows_callback, None)
    return hwnds

def hide_other_windows(exclude_hwnds):
    """
    隐藏所有窗口，除了在exclude_hwnds中的窗口。
    返回隐藏的窗口列表以便恢复。
    """
    hidden_hwnds = []

    def enum_windows_callback(hwnd, lParam):
        if hwnd in exclude_hwnds:
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        # 避免隐藏关键系统窗口
        if win32gui.GetWindowText(hwnd) == "":
            return
        # 隐藏窗口
        try:
            win32gui.ShowWindow(hwnd, 0)  # 0 = SW_HIDE
            hidden_hwnds.append(hwnd)
        except Exception as e:
            print(f"无法隐藏窗口 {hwnd}: {e}")

    win32gui.EnumWindows(enum_windows_callback, None)
    return hidden_hwnds

def show_windows(hwnds):
    """
    显示之前隐藏的窗口。
    """
    for hwnd in hwnds:
        try:
            win32gui.ShowWindow(hwnd, 5)  # 5 = SW_SHOW
        except Exception as e:
            print(f"无法显示窗口 {hwnd}: {e}")

# ----------------------------- Image Capture ----------------------------- #

def capture_window(hwnd):
    """
    捕获指定窗口的图像并返回OpenCV格式的图像。
    """
    try:
        hwndDC = win32gui.GetWindowDC(hwnd)
        if not hwndDC:
            print("无法获取窗口设备上下文。")
            return None
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        # 获取窗口尺寸
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        left, top = win32gui.ClientToScreen(hwnd, (left, top))
        right, bottom = win32gui.ClientToScreen(hwnd, (right, bottom))
        width = right - left
        height = bottom - top

        if width == 0 or height == 0:
            print("窗口宽度或高度为0，无法捕获。")
            return None

        # 创建位图对象
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        # 使用PrintWindow复制窗口内容
        PW_RENDERFULLCONTENT = 0x00000002
        result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), PW_RENDERFULLCONTENT)
        if not result:
            # 尝试使用标志1
            print("PrintWindow失败，尝试使用标志1。")
            result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
            if not result:
                print("PrintWindow失败，无法捕获窗口内容。")
                return None

        # 获取位图信息
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        # 转换为numpy数组
        img = np.frombuffer(bmpstr, dtype=np.uint8)
        if img.size == 0:
            print("捕获到的图像数据为空。")
            return None
        img.shape = (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4)

        # 转换为BGR格式
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        return img

    except Exception as e:
        print(f"捕获窗口时出错: {e}")
        return None

    finally:
        # 释放资源
        try:
            win32gui.DeleteObject(saveBitMap.GetHandle())
        except:
            pass
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

# ----------------------------- OCR and Web Search ----------------------------- #

def perform_ocr(image, ocr_language='chi_sim'):
    """
    对图像执行OCR并返回提取的文本。
    """
    try:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        pil_img = Image.fromarray(gray)

        # 执行OCR
        custom_config = f'--oem 1 --psm 3 -l {ocr_language}'
        extracted_text = pytesseract.image_to_string(pil_img, config=custom_config)

        if not extracted_text.strip():
            print("OCR未检测到文本。")
            return ""

        # 清理文本
        cleaned_text = clean_extracted_text(extracted_text)
        return cleaned_text

    except Exception as e:
        print(f"OCR识别时出错: {e}")
        return ""

def clean_extracted_text(text):
    """
    清理和格式化OCR提取的文本。
    """
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    lines = text.split('\n')
    cleaned_lines = [line.strip() for line in lines if line.strip()]
    cleaned_text = '\n'.join(cleaned_lines)
    return cleaned_text

def perform_web_search(query, search_url, search_api_key=None, max_results=5):
    """
    使用指定的搜索API进行网页搜索，并返回搜索结果摘要。
    """
    try:
        print(f"正在搜索: {query}")
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        }
        params = {'q': query}
        if search_api_key:
            params['api_key'] = search_api_key
        response = requests.post(search_url, data=params, headers=headers, timeout=10)
        if response.status_code != 200:
            print(f"搜索请求失败，状态码: {response.status_code}")
            return "搜索请求失败。"

        soup = BeautifulSoup(response.text, 'html.parser')
        results = []
        for result in soup.find_all('div', class_='result', limit=max_results):
            title_tag = result.find('a', class_='result__a')
            snippet_tag = result.find('a', class_='result__snippet')
            if title_tag and snippet_tag:
                title = title_tag.get_text()
                snippet = snippet_tag.get_text()
                link = title_tag.get('href')
                # 计算与查询的相关性（示例：包含查询关键词的优先级更高）
                relevance = sum(word.lower() in snippet.lower() for word in query.split())
                results.append({'title': title, 'snippet': snippet, 'link': link, 'relevance': relevance})

        if not results:
            print("未找到搜索结果。")
            return "未找到相关的搜索结果。"

        # 按相关性排序
        sorted_results = sorted(results, key=lambda x: x['relevance'], reverse=True)

        # 构建摘要
        formatted_results = []
        for res in sorted_results:
            formatted_results.append(f"标题: {res['title']}\n摘要: {res['snippet']}\n链接: {res['link']}\n")

        search_summary = "\n".join(formatted_results)
        return search_summary

    except Exception as e:
        print(f"进行网络搜索时出错: {e}")
        return "进行网络搜索时出错。"

def send_to_ai(extracted_text, search_results, ai_api_url, ai_api_key=None, model_name='qwen2.5:14b'):
    """
    将文本和搜索结果发送到AI模型，并返回响应。
    """
    try:
        if not extracted_text.strip():
            print("发送到 AI 的文本为空。")
            return "发送到 AI 时出错：文本为空。"

        # 构建AI的提示，主要聚焦于解题
        prompt = (
            "你是一名智能助理，擅长解答各种类型的问题，特别是帮助用户解决问题。请根据以下内容提供详细、准确且有条理的回答,正确应对识别混乱的语句，抓住关键信息发散。。\n\n"
            "### 题目内容：\n"
            f"{extracted_text}\n\n"
            "### 网络搜索结果：\n"
            f"{search_results}\n\n"
            "### 任务要求：\n"
            "1. 识别题目类型（如选择题、简答题、填空题等）。\n"
            "2. 根据题目类型，以适当的格式提供答案。\n"
            "   - **选择题**：请提供正确选项，并简要解释原因。\n"
            "   - **简答题**：请提供简明扼要的答案。\n"
            "   - **填空题**：请填入正确的答案。\n"
            "3. 确保答案准确且具有逻辑性。\n"
            "4. 保持题目序号与答案序号一致。\n"
            "5. 优化答案的结构和语言表达，使其易于理解。\n"
            "6. 准确地识别单选题和多选题。\n"
            "7. 正确应对识别混乱的语句，抓住关键信息发散。\n\n"
            "### 答案："
        )

        headers = {'Content-Type': 'application/json'}
        if ai_api_key:
            headers['Authorization'] = f'Bearer {ai_api_key}'

        data = {
            'model': model_name,
            'messages': [{
                'role': 'user',
                'content': prompt
            }],
            'stream': False
        }

        response = requests.post(ai_api_url, headers=headers, json=data, timeout=30)  # 增加超时时间

        if response.status_code == 200:
            result = response.json()
            content = result.get('message', {}).get('content', '')
            if content:
                return content
            else:
                print("AI 模型返回的内容为空。")
                return "AI 模型返回的内容为空。"
        else:
            print(f"AI 模型返回错误状态码: {response.status_code}")
            print(f"响应内容: {response.text}")
            return f"AI 模型错误，状态码: {response.status_code}"
    except requests.exceptions.RequestException as e:
        print(f"AI 模型请求异常: {e}")
        return f"AI 模型请求异常: {e}"
    except Exception as e:
        print(f"AI 模型错误: {e}")
        return "发送到 AI 时出错。"

# ----------------------------- Screenshot Processing Thread ----------------------------- #

class ScreenshotThread(QThread):
    ocr_text_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, hwnd, window_title, settings):
        super().__init__()
        self._run_flag = True
        self.hwnd = hwnd
        self.window_title = window_title
        self.settings = settings  # Dictionary containing all settings

    def run(self):
        while self._run_flag:
            # 检查窗口是否仍然存在和可见
            if not win32gui.IsWindow(self.hwnd):
                message = f"{self.window_title}: 窗口不存在。\n"
                print(message)
                self.ocr_text_signal.emit(message)
                break

            # 捕获窗口图像
            frame = capture_window(self.hwnd)
            if frame is None:
                message = f"{self.window_title}: 图像捕获失败。\n"
                print(message)
                self.ocr_text_signal.emit(message)
                break

            # 执行OCR
            extracted_text = perform_ocr(frame, ocr_language=self.settings.get('ocr_language', 'chi_sim'))

            if not extracted_text.strip():
                message = f"{self.window_title}: 未检测到文本。\n"
                print(message)
                self.ocr_text_signal.emit(message)
            else:
                # 进行网络搜索
                search_results = perform_web_search(
                    query=extracted_text,
                    search_url=self.settings.get('search_api_url', 'https://html.duckduckgo.com/html/'),
                    search_api_key=self.settings.get('search_api_key', None),
                    max_results=self.settings.get('search_max_results', 5)
                )

                # 发送到AI
                ai_response = send_to_ai(
                    extracted_text=extracted_text,
                    search_results=search_results,
                    ai_api_url=self.settings.get('ai_api_url', 'http://localhost:11434/api/chat'),
                    ai_api_key=self.settings.get('ai_api_key', None),
                    model_name=self.settings.get('ai_model', 'qwen2.5:14b')
                )

                # 准备分隔符和时间戳
                separator = (
                    "<hr style='border:1px solid gray;'>"
                    f"<b>窗口: {self.window_title}</b> "
                    f"<span style='color: gray;'>{time.strftime('%Y-%m-%d %H:%M:%S')}</span><br>"
                )
                # 发射AI响应，使用HTML格式，并优化显示
                formatted_response = f"{separator}<pre style='font-size: 12pt; white-space: pre-wrap;'>{ai_response}</pre>"
                self.ocr_text_signal.emit(formatted_response)

            # 等待下一次捕获
            time.sleep(self.settings.get('capture_interval', 1))

    def stop(self):
        self._run_flag = False
        self.wait()

# ----------------------------- Overlay Window ----------------------------- #

class OverlayWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI Response Overlay")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.WindowDoesNotAcceptFocus)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        # QTextEdit用于显示AI响应
        self.text_edit = QTextEdit(self)
        self.text_edit.setReadOnly(True)
        self.text_edit.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.text_edit.setStyleSheet("""
            background-color: rgba(0, 0, 0, 180);
            color: white;
            padding: 10px;
            border: none;
            font-size: 12pt;
        """)
        font = QFont("Segoe UI", 12)
        self.text_edit.setFont(font)

        # 启用富文本模式
        self.text_edit.setAcceptRichText(True)

        # 布局设置
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        self.setLayout(layout)

        # 窗口不透明度
        self.setWindowOpacity(0.6)  # 调整为0.6以更好地显示内容

        # 自动隐藏滚动条
        self.text_edit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.text_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        # 将窗口放置在屏幕的右侧，并统一尺寸
        screen_geometry = QApplication.desktop().availableGeometry(self)
        window_width = 500  # 宽度设置为500
        window_height = 1400  # 高度设置为1400
        self.setGeometry(
            screen_geometry.width() - window_width - 10,  # 右侧边缘留10像素
            20,  # 顶部留20像素
            window_width,
            window_height
        )

    def update_text(self, text):
        """
        将新文本附加到QTextEdit中，并自动滚动到最新内容。
        支持HTML格式化。
        """
        self.text_edit.moveCursor(QTextCursor.End)
        self.text_edit.insertHtml(text + "<br>")
        self.text_edit.moveCursor(QTextCursor.End)

# ----------------------------- Keyboard Listener ----------------------------- #

class KeyboardListener(QObject):
    quit_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener.start()

    def on_press(self, key):
        try:
            if key.char.lower() == 'q':
                print("检测到 'Q' 键被按下，准备退出程序。")
                self.quit_signal.emit()
        except AttributeError:
            # 特殊键（如Ctrl, Shift等）会引发AttributeError
            pass

    def stop(self):
        self.listener.stop()

# ----------------------------- Settings Tabs ----------------------------- #

class AISettingsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QFormLayout()

        # AI API URL
        self.ai_api_url_input = QLineEdit()
        self.ai_api_url_input.setText('http://localhost:11434/api/chat')
        layout.addRow(QLabel("<b>AI API URL:</b>"), self.ai_api_url_input)

        # AI API Key (可选)
        self.ai_api_key_input = QLineEdit()
        self.ai_api_key_input.setPlaceholderText("可选")
        layout.addRow(QLabel("<b>AI API Key:</b>"), self.ai_api_key_input)

        # AI 模型名称
        self.ai_model_input = QLineEdit()
        self.ai_model_input.setText('qwen2.5:14b')
        layout.addRow(QLabel("<b>AI 模型名称:</b>"), self.ai_model_input)

        self.setLayout(layout)

class SearchSettingsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QFormLayout()

        # 搜索API URL
        self.search_api_url_input = QLineEdit()
        self.search_api_url_input.setText('https://html.duckduckgo.com/html/')
        layout.addRow(QLabel("<b>搜索 API URL:</b>"), self.search_api_url_input)

        # 搜索API Key (可选)
        self.search_api_key_input = QLineEdit()
        self.search_api_key_input.setPlaceholderText("可选")
        layout.addRow(QLabel("<b>搜索 API Key:</b>"), self.search_api_key_input)

        # 最大搜索结果数量
        self.search_max_results_spin = QSpinBox()
        self.search_max_results_spin.setRange(1, 30)
        self.search_max_results_spin.setValue(6)
        layout.addRow(QLabel("<b>最大搜索结果数:</b>"), self.search_max_results_spin)

        self.setLayout(layout)

class DisplaySettingsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QFormLayout()

        # 显示窗口宽度
        self.display_width_spin = QSpinBox()
        self.display_width_spin.setRange(300, 2000)
        self.display_width_spin.setValue(500)  # 与OverlayWindow的宽度一致
        layout.addRow(QLabel("<b>覆盖窗口宽度:</b>"), self.display_width_spin)

        # 显示窗口高度
        self.display_height_spin = QSpinBox()
        self.display_height_spin.setRange(400, 2000)
        self.display_height_spin.setValue(1400)  # 设置为1400
        layout.addRow(QLabel("<b>覆盖窗口高度:</b>"), self.display_height_spin)

        # 显示窗口位置X
        self.display_pos_x_spin = QSpinBox()
        self.display_pos_x_spin.setRange(0, 3000)
        self.display_pos_x_spin.setValue(0)  # 默认动态计算
        layout.addRow(QLabel("<b>覆盖窗口位置 X:</b>"), self.display_pos_x_spin)

        # 显示窗口位置Y
        self.display_pos_y_spin = QSpinBox()
        self.display_pos_y_spin.setRange(0, 2000)
        self.display_pos_y_spin.setValue(20)
        layout.addRow(QLabel("<b>覆盖窗口位置 Y:</b>"), self.display_pos_y_spin)

        self.setLayout(layout)

class OCRSettingsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QFormLayout()

        # OCR 语言选择
        self.ocr_language_combo = QComboBox()
        self.ocr_language_combo.addItems(['chi_sim', 'eng', 'fra', 'deu', 'spa'])  # 根据需要添加更多
        self.ocr_language_combo.setCurrentText('chi_sim')
        layout.addRow(QLabel("<b>OCR 语言:</b>"), self.ocr_language_combo)

        # Tesseract 安装路径 (可选)
        self.tesseract_path_input = QLineEdit()
        self.tesseract_path_input.setPlaceholderText("可选")
        self.tesseract_path_input.setText(pytesseract.pytesseract.tesseract_cmd)
        layout.addRow(QLabel("<b>Tesseract 路径:</b>"), self.tesseract_path_input)

        # 选择 Tesseract 可执行文件按钮
        self.browse_tesseract_button = QPushButton("浏览")
        self.browse_tesseract_button.clicked.connect(self.browse_tesseract)
        layout.addRow(QLabel(""), self.browse_tesseract_button)

        self.setLayout(layout)

    def browse_tesseract(self):
        """
        浏览并选择Tesseract可执行文件
        """
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        fileName, _ = QFileDialog.getOpenFileName(
            self, "选择 Tesseract 可执行文件", "", "可执行文件 (*.exe);;所有文件 (*)", options=options
        )
        if fileName:
            self.tesseract_path_input.setText(fileName)
            pytesseract.pytesseract.tesseract_cmd = fileName

# ----------------------------- About Tab ----------------------------- #

class AboutTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout()

        disclaimer = QTextEdit()
        disclaimer.setReadOnly(True)
        disclaimer.setHtml("""
            <h2>免责声明</h2>
            <p>本软件仅用作学习和研究用途，如有人用作违规或非法用途，使用本软件的风险由用户自行承担，本软件仅可用于练习题辅助大家学习，切勿用于违规途径，软件开发的本意是辅助大家好好学习，天天向上。</p>
            <h3>软件作者</h3>
            <p>爱划水的小咸鱼</p>
        """)
        layout.addWidget(disclaimer)
        self.setLayout(layout)

# ----------------------------- Settings Tab ----------------------------- #

class SettingsTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        # 创建标签页
        self.tabs = QTabWidget()
        self.ai_settings_tab = AISettingsTab()
        self.search_settings_tab = SearchSettingsTab()
        self.display_settings_tab = DisplaySettingsTab()
        self.ocr_settings_tab = OCRSettingsTab()
        self.about_tab = AboutTab()  # 添加关于标签页
        self.tabs.addTab(self.ai_settings_tab, "AI 设置")
        self.tabs.addTab(self.search_settings_tab, "搜索设置")
        self.tabs.addTab(self.display_settings_tab, "显示设置")
        self.tabs.addTab(self.ocr_settings_tab, "OCR 设置")
        self.tabs.addTab(self.about_tab, "关于")

        # 布局
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)
        self.setLayout(main_layout)

# ----------------------------- Window Selection Tab ----------------------------- #

class WindowSelectionTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        # 布局
        layout = QVBoxLayout()

        # 窗口列表
        self.window_list = QListWidget()
        self.window_list.setSelectionMode(QListWidget.MultiSelection)
        self.window_list.setStyleSheet("""
            QListWidget {
                background-color: #2b2b2b;
                color: white;
                border: 1px solid #444;
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #3c3f41;
            }
        """)
        layout.addWidget(self.window_list)

        # 刷新按钮
        self.refresh_button = QPushButton("刷新窗口列表")
        self.refresh_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(self.refresh_button)

        self.setLayout(layout)

# ----------------------------- Main Application ----------------------------- #

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI 窗口 OCR 工具")
        self.setGeometry(100, 100, 1400, 900)
        self.display_width = 500
        self.display_height = 1400

        # 设置应用程序图标（可选）
        # self.setWindowIcon(QIcon('icon.png'))

        # 隐藏任务栏图标
        self.setWindowFlags(self.windowFlags() | Qt.Tool)

        # 创建标签页
        self.tabs = QTabWidget()
        self.window_selection_tab = WindowSelectionTab()
        self.settings_tab = SettingsTab()
        self.tabs.addTab(self.window_selection_tab, "窗口选择")
        self.tabs.addTab(self.settings_tab, "设置")

        # 创建按钮
        self.start_button = QPushButton("开始检测")
        self.stop_button = QPushButton("停止检测")
        self.stop_button.setEnabled(False)

        # 设置按钮样式
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #008CBA;
                color: white;
                border: none;
                padding: 10px 20px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #007bb5;
            }
        """)
        self.stop_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)

        # 创建状态标签
        self.status_label = QLabel("状态: 未开始")
        self.status_label.setStyleSheet("font-size: 12pt; color: green;")

        # 添加滑块控件来调整捕获区域大小
        self.scale_slider = QSlider(Qt.Horizontal, self)
        self.scale_slider.setMinimum(10)   # 10%
        self.scale_slider.setMaximum(100)
        self.scale_slider.setValue(100)
        self.scale_slider.setTickPosition(QSlider.TicksBelow)
        self.scale_slider.setTickInterval(10)
        self.scale_slider.valueChanged.connect(self.scale_changed)

        self.scale_label = QLabel("捕获区域比例: 100%")
        self.scale_label.setStyleSheet("font-size: 12pt;")

        # 布局
        main_layout = QVBoxLayout()

        # 添加标签页
        main_layout.addWidget(self.tabs)

        # 添加捕获区域滑块
        scale_layout = QHBoxLayout()
        scale_layout.addWidget(self.scale_label)
        scale_layout.addWidget(self.scale_slider)
        main_layout.addLayout(scale_layout)

        # 添加按钮和状态标签
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()  # 添加弹性空间
        button_layout.addWidget(self.status_label)
        main_layout.addLayout(button_layout)

        # 设置中央窗口
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # 初始化变量
        self.threads = []
        self.hidden_windows = []
        self.overlay = OverlayWindow()
        self.overlay.show()
        self.settings = self.load_default_settings()

        # 键盘监听器
        self.keyboard_listener = KeyboardListener()
        self.keyboard_listener.quit_signal.connect(self.handle_quit)

        # 连接刷新按钮
        self.window_selection_tab.refresh_button.clicked.connect(self.refresh_window_list)

        # 连接开始和停止按钮
        self.start_button.clicked.connect(self.start_detection)
        self.stop_button.clicked.connect(self.stop_detection)

        # 启动时刷新窗口列表
        self.refresh_window_list()

    def load_default_settings(self):
        """
        加载默认设置
        """
        screen_geometry = QApplication.desktop().availableGeometry(self)
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()

        default_display_width = 500
        default_display_height = 1400
        default_display_pos_x = screen_width - default_display_width - 10
        default_display_pos_y = 20

        return {
            'ai_api_url': 'http://localhost:11434/api/chat',
            'ai_api_key': None,
            'ai_model': 'qwen2.5:14b',
            'search_api_url': 'https://html.duckduckgo.com/html/',
            'search_api_key': None,
            'search_max_results': 6,  # 保持搜索结果数量不变
            'display_width': default_display_width,
            'display_height': default_display_height,
            'display_pos_x': default_display_pos_x,
            'display_pos_y': default_display_pos_y,
            'ocr_language': 'chi_sim',
            'tesseract_path': pytesseract.pytesseract.tesseract_cmd,
            'capture_interval': 1
        }

    def refresh_window_list(self):
        """
        刷新窗口列表，显示所有运行中的窗口标题。
        """
        self.window_selection_tab.window_list.clear()
        try:
            hwnds = find_all_windows()
            if not hwnds:
                QMessageBox.warning(self, "窗口错误", "未找到任何运行中的应用程序窗口。")
                return

            for hwnd in hwnds:
                window_title = win32gui.GetWindowText(hwnd)
                if not window_title:
                    window_title = "<无标题窗口>"
                item_text = f"{window_title} (HWND: {hwnd})"
                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, hwnd)  # 存储句柄
                self.window_selection_tab.window_list.addItem(item)
        except Exception as e:
            print(f"刷新窗口列表时出错: {e}")

    def handle_tab_change(self, index):
        """
        处理标签页切换
        """
        pass

    def save_settings(self):
        """
        保存设置标签中的所有设置。
        """
        # AI 设置
        ai_settings = self.settings_tab.ai_settings_tab
        self.settings['ai_api_url'] = ai_settings.ai_api_url_input.text()
        self.settings['ai_api_key'] = ai_settings.ai_api_key_input.text() if ai_settings.ai_api_key_input.text() else None
        self.settings['ai_model'] = ai_settings.ai_model_input.text()

        # 搜索设置
        search_settings = self.settings_tab.search_settings_tab
        self.settings['search_api_url'] = search_settings.search_api_url_input.text()
        self.settings['search_api_key'] = search_settings.search_api_key_input.text() if search_settings.search_api_key_input.text() else None
        self.settings['search_max_results'] = search_settings.search_max_results_spin.value()

        # 显示设置
        display_settings = self.settings_tab.display_settings_tab
        self.settings['display_width'] = display_settings.display_width_spin.value()
        self.settings['display_height'] = display_settings.display_height_spin.value()
        # 如果 display_pos_x 和 display_pos_y 被手动调整，使用用户设置，否则使用默认的动态计算位置
        self.settings['display_pos_x'] = display_settings.display_pos_x_spin.value() if display_settings.display_pos_x_spin.value() else self.settings['display_pos_x']
        self.settings['display_pos_y'] = display_settings.display_pos_y_spin.value()

        # OCR 设置
        ocr_settings = self.settings_tab.ocr_settings_tab
        self.settings['ocr_language'] = ocr_settings.ocr_language_combo.currentText()
        self.settings['tesseract_path'] = ocr_settings.tesseract_path_input.text()
        if self.settings['tesseract_path']:
            pytesseract.pytesseract.tesseract_cmd = self.settings['tesseract_path']

        # 保存捕获间隔
        self.settings['capture_interval'] = self.settings.get('capture_interval', 1)

        QMessageBox.information(self, "设置已保存", "设置已成功保存。")

    def start_detection(self):
        """
        开始选定窗口的检测过程。
        """
        self.save_settings()  # 保存当前设置

        selected_items = self.window_selection_tab.window_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "选择错误", "请从窗口列表中选择至少一个窗口。")
            return

        # 解析选定的窗口句柄
        selected_hwnds = []
        for item in selected_items:
            hwnd = item.data(Qt.UserRole)
            selected_hwnds.append(hwnd)

        if not selected_hwnds:
            QMessageBox.warning(self, "解析错误", "无法解析选定的窗口句柄。")
            return

        # 排除覆盖窗口不被隐藏
        overlay_hwnd = self.get_overlay_hwnd()
        exclude_hwnds = set(selected_hwnds)
        exclude_hwnds.add(overlay_hwnd)

        # 隐藏其他窗口
        self.hidden_windows = hide_other_windows(exclude_hwnds)

        # 设置覆盖窗口的位置和大小
        self.overlay.setGeometry(
            self.settings['display_pos_x'],
            self.settings['display_pos_y'],
            self.settings['display_width'],
            self.settings['display_height']
        )

        # 为每个选定的窗口创建并启动截图处理线程
        for hwnd in selected_hwnds:
            window_title = win32gui.GetWindowText(hwnd)
            if not window_title:
                window_title = "<无标题窗口>"
            thread = ScreenshotThread(hwnd=hwnd, window_title=window_title, settings=self.settings)
            thread.ocr_text_signal.connect(self.overlay.update_text)
            thread.error_signal.connect(self.handle_error)
            thread.start()
            self.threads.append(thread)
            print(f"已启动窗口检测线程: {window_title} (HWND: {hwnd})")

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.status_label.setText("状态: 正在检测")
        self.status_label.setStyleSheet("font-size: 12pt; color: orange;")
        QMessageBox.information(self, "检测已开始", "已开始检测选定的窗口。")

    def stop_detection(self):
        """
        停止所有检测线程并恢复隐藏的窗口。
        """
        print("正在停止检测...")
        for thread in self.threads:
            thread.stop()
        self.threads = []

        # 恢复隐藏的窗口
        show_windows(self.hidden_windows)
        self.hidden_windows = []

        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.status_label.setText("状态: 已停止")
        self.status_label.setStyleSheet("font-size: 12pt; color: green;")
        QMessageBox.information(self, "检测已停止", "所有检测已停止，窗口已恢复。")

    def get_overlay_hwnd(self):
        """
        获取覆盖窗口的句柄。
        """
        self.overlay.show()
        time.sleep(0.1)  # 等待窗口显示
        hwnd = win32gui.FindWindow(None, self.overlay.windowTitle())
        if hwnd == 0:
            print("无法找到 Overlay 窗口的句柄。")
        return hwnd

    def handle_quit(self):
        """
        处理退出信号，确保所有线程停止，窗口恢复，并退出应用程序。
        """
        print("正在退出应用程序...")
        self.stop_detection()
        self.keyboard_listener.stop()
        self.overlay.close()
        QApplication.quit()

    def handle_error(self, error_message):
        """
        处理来自线程的错误信号。
        """
        error_html = f"<span style='color: red;'>{error_message}</span>"
        self.overlay.update_text(error_html)

    def closeEvent(self, event):
        """
        重写关闭事件，确保正确清理资源。
        """
        self.handle_quit()
        event.accept()

    def update_image(self, cv_img):
        """
        将OpenCV图像转换为Qt图像并显示
        """
        qt_img = self.convert_cv_qt(cv_img)
        self.label.setPixmap(qt_img)

    def update_fps(self, fps):
        """
        更新FPS显示
        """
        self.fps_label.setText(f"FPS: {fps:.2f}")

    def update_capture_status(self, status):
        """
        更新捕获状态显示
        """
        if status:
            self.capture_status_label.setText("捕获状态: 成功")
        else:
            self.capture_status_label.setText("捕获状态: 失败")

    def convert_cv_qt(self, cv_img):
        """
        将OpenCV图像转换为QPixmap
        """
        try:
            rgb_image = cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGB)
        except Exception as e:
            print(f"图像颜色转换出错: {e}")
            return QPixmap()

        h, w, ch = rgb_image.shape
        bytes_per_line = ch * w
        convert_to_Qt_format = QImage(
            rgb_image.data, w, h, bytes_per_line, QImage.Format_RGB888
        )
        p = convert_to_Qt_format.scaled(
            self.display_width, self.display_height, Qt.KeepAspectRatio
        )
        return QPixmap.fromImage(p)

    def scale_changed(self, value):
        """
        更新捕获区域比例标签并更新线程的 capture_interval
        """
        self.scale_label.setText(f"捕获区域比例: {value}%")
        self.settings['capture_interval'] = value / 100.0  # 根据滑块值调整捕获间隔

# ----------------------------- Entry Point ----------------------------- #

def main():
    # 隐藏控制台窗口
    if sys.platform == "win32":
        import win32con
        import win32gui

        # 获取当前脚本的窗口句柄
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

    app = QApplication(sys.argv)

    # 设置应用程序的整体样式
    app.setStyle("Fusion")
    palette = QPalette()

    # 设置窗口背景颜色
    palette.setColor(QPalette.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.WindowText, Qt.white)

    # 设置按钮颜色
    palette.setColor(QPalette.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ButtonText, Qt.white)

    # 设置文本输入框颜色
    palette.setColor(QPalette.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipText, Qt.white)

    # 设置文本颜色
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.PlaceholderText, QColor(170, 170, 170))

    # 设置边框颜色
    palette.setColor(QPalette.ButtonText, Qt.white)

    app.setPalette(palette)

    application = App()
    application.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
