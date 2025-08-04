import os
import sys
import subprocess
import re
import socket

def check_internet():
    try:
        socket.create_connection(("1.1.1.1", 53), timeout=3)
        return True
    except:
        return False

def auto_install(module_name):
    print(f"⚙️  Đang cài đặt {module_name} ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])
        print(f"✅ Đã cài xong {module_name}\n")
    except Exception as e:
        print(f"❌ Lỗi khi cài {module_name}: {e}")
        sys.exit(1)

def extract_top_level_modules(file_path):
    imported = set()
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip().split("#")[0]
            if line.startswith("import "):
                parts = line.split()
                if len(parts) > 1:
                    imported.add(parts[1].split(".")[0])
            elif line.startswith("from "):
                parts = line.split()
                if len(parts) > 1:
                    imported.add(parts[1].split(".")[0])
    return list(imported)

def check_missing_modules():
    if not check_internet():
        print("⚠️ Không có kết nối Internet. Không thể tự động cài thư viện.")
        sys.exit(1)

    current_file = os.path.abspath(__file__)
    modules = extract_top_level_modules(current_file)

    all_imported = True

    for module in modules:
        try:
            __import__(module)
        except ImportError:
            all_imported = False
            auto_install(module)

    if all_imported:
        print("✅ Tất cả thư viện đã sẵn sàng.\n")

# GỌI KIỂM TRA NGAY KHI KHỞI ĐỘNG
check_missing_modules()

# Tiếp tục các import quan trọng bên dưới
import pygame
import hid
import time
from openpyxl import Workbook
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QScrollArea, QFileDialog, QMessageBox, QGridLayout, QSizePolicy
)
from PySide6.QtCore import Qt, QTimer
from functools import partial

from PySide6.QtCore import QPropertyAnimation, QEasingCurve
from PySide6.QtWidgets import QGraphicsOpacityEffect
from PySide6.QtCore import QAbstractAnimation

from collections import defaultdict

from enum import Enum, auto

import win32pipe
import win32file
import pywintypes

import psutil  # Thư viện kiểm tra tiến trình chạy nhanh hơn tasklist

# === Cấu hình Rumble Server ===
SERVER_EXE = 'XboxRumbleServer.exe'
SERVER_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PORTS = [50500, 50501, 50502, 50503]
DEFAULT_PIPE = r'\\.\pipe\XboxRumblePipe'
SEND_INTERVAL = 0.03

class RumbleMode(Enum):
    OFF = 0
    MOTOR = 1
    TRIGGER = 2

def is_server_running():
    # Kiểm tra nhanh process chạy đúng tên server (cách này nhanh gấp nhiều lần tasklist)
    for proc in psutil.process_iter(attrs=['name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == SERVER_EXE.lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def start_server():
    exe_path = os.path.join(SERVER_DIR, SERVER_EXE)
    if not os.path.exists(exe_path):
        print(f"Không tìm thấy file {exe_path}. Vui lòng để XboxRumbleServer.exe cùng thư mục với Python script.")
        return False
    if not is_server_running():
        print("Đang khởi động XboxRumbleServer...")
        DETACHED_PROCESS = 0x00000008
        subprocess.Popen(
            [exe_path],
            cwd=SERVER_DIR,
            creationflags=DETACHED_PROCESS,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True
        )
        time.sleep(2)
    return True

TOOLBAR_BTN_HEIGHT = 30
TOOLBAR_BTN_FONT_SIZE = 16

HID_SCAN_INTERVAL = 0.5   # Quét lại thiết bị HID mỗi 0.5 giây (tối ưu performance)

DATA_FOLDER = "controller_data"
if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)

# Các hãng controller phổ biến
controller_vids = {0x045e, 0x054c, 0x057e, 0x20d6, 0x0e6f, 0x0f0d, 0x1532, 0x0079}
controller_words = [
    "controller", "gamepad", "joycon", "xbox", "dualshock", "dualsense",
    "pro", "switch", "ps4", "ps5"
]
def is_controller(d):
    prod = (d.get("product_string") or "").lower()
    manu = (d.get("manufacturer_string") or "").lower()
    usage = d.get("usage")
    vid = d.get("vendor_id")
    if vid in controller_vids:
        for kw in controller_words:
            if kw in prod or kw in manu:
                return True
    if d.get("usage_page") == 0x01 and usage in [0x04, 0x05]:
        return True
    if d.get("usage_page") == 0x01 and usage == 0x06 and "controller" in prod:
        return True
    return False

def detect_model(pid_set):
    if pid_set == {0x2ff}:
        return "Jelling Controller"
    elif 0x2ff in pid_set and 0xb02 in pid_set:
        return "Durham Controller"
    return None

MOTOR_FALLBACK = {
     "left_trigger": "left",    # Nếu trigger trái bận, rung sang motor lớn trái
     "left": "left_trigger",    # Nếu motor lớn trái bận, rung sang trigger trái
     "right_trigger": "right",  # Nếu trigger phải bận, rung sang motor nhỏ phải
     "right": "right_trigger",  # Nếu motor nhỏ phải bận, rung sang trigger phải
}

pygame.init()
pygame.joystick.init()

class ControllerTester(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Xbox Controller Tester (PySide6)")
        self.setMinimumSize(430, 740)
        self.resize(430, 740)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)

        self.joystick = None

        # ==== Thêm các biến trạng thái lặp để điều khiển hiển thị =====
        self.connected_state_timer = 0        # Đếm thời gian đã hiển thị
        self.connected_state_stage = 0        # 0: "Đã kết nối", 1: "Chế độ rung: Tắt"
        self.connected_state_last_toggle = time.time()
        self.connected_display_mode = False   # True: đang hiển thị lặp; False: hiện trạng thái rung
        self.was_connected = False            # Lưu trạng thái kết nối trước đó

        # Trạng thái rung
        self.rumble_mode = RumbleMode.OFF   # OFF: Tắt, MOTOR: Motor lớn, TRIGGER: Trigger

        self.nguong_rung = {
            "A": [(10, 0.3)],
            "B": [(10, 0.3)],
            "X": [(10, 0.3)],
            "Y": [(10, 0.3)],
            "Down": [(10, 0.3)],            "Right": [(10, 0.3)],
            "Left": [(10, 0.3)],
            "Up": [(10, 0.3)],
            "Menu": [(10, 0.3)],
            "View": [(10, 0.3)],
            "Share": [(10, 0.3)],
            "LS": [(10, 0.3)],
            "RS": [(10, 0.3)],
            "LB": [(10, 0.3), (20, 0.3), (30, 0.5)],
            "RB": [(10, 0.3), (20, 0.3), (30, 0.5)],
            "Guide": [(5, 0.3), (10, 0.3), (15, 0.3), (20, 0.3), (25, 0.5)],
        }

        self.trigger_rumble_thresholds = {
            "jelling controller": [(5, 0.2), (10, 0.2), (15, 0.5)],
            "durham controller": [(5, 0.2), (10, 0.2), (15, 0.5), (20, 0.2), (25, 0.2), (30, 0.5), (35, 0.2), (40, 0.2), (45, 0.5)],
            # Các loại khác nếu muốn
        }

        self.last_combo_up = False     # Cho tổ hợp test motor
        self.last_combo_down = False   # Cho tổ hợp test trigger

        # Sử dụng defaultdict(set) để tự khởi tạo set cho mỗi nút khi truy cập lần đầu
        self.rumble_alerted = defaultdict(set)

        self.rumble_busy = False  # Chặn rung test khi đang rung cảnh báo

        self.last_rumble_state = None
        self.last_keep_alive = 0
        
        self.last_hid_scan = 0           # Thời điểm quét HID gần nhất
        self.hid_cache = []              # Cache danh sách controller HID
        self.last_pid_set = set()        # Để so sánh Jelling/Durham khi cập nhật

        self.motor_rumble_end_time = {
            "left": 0.0,          # motor lớn
            "right": 0.0,         # motor nhỏ
            "left_trigger": 0.0,  # rung trigger trái
            "right_trigger": 0.0  # rung trigger phải
        }
        self.motor_rumble_intensity = {
            "left": 0.0,
            "right": 0.0,
            "left_trigger": 0.0,
            "right_trigger": 0.0
        }

        self.current_test_rumble_state = {
            "left": 0.0,
            "right": 0.0,
            "left_trigger": 0.0,
            "right_trigger": 0.0
        }
        
        # ===== Toolbar =====
        toolbar = QHBoxLayout()
        btn_min_width = 120
        btn_min_height = TOOLBAR_BTN_HEIGHT
        btn_font = f"{TOOLBAR_BTN_FONT_SIZE}pt"
        # Xuất Kết Quả
        self.btn_export = QPushButton("Xuất Kết Quả")
        self.btn_export.setMinimumWidth(btn_min_width)
        self.btn_export.setMinimumHeight(btn_min_height)
        self.btn_export.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_export.setStyleSheet(f"""
            QPushButton {{
                background: #1877f2; color: white; font-weight: bold; border-radius:5px; font-size: {btn_font};
            }}
            QPushButton:hover {{
                background: #2490fa; color: #fff; border: 2px solid #1877f2;
            }}
            QPushButton:pressed {{
                background: #185fc2; border: 2px solid #185fc2;
            }}
        """)
        # Reset
        self.btn_reset = QPushButton("Reset Tất cả")
        self.btn_reset.setMinimumWidth(btn_min_width)
        self.btn_reset.setMinimumHeight(btn_min_height)
        self.btn_reset.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_reset.setStyleSheet(f"""
            QPushButton {{
                background: #666; color: white; border-radius:5px; font-size: {btn_font};
            }}
            QPushButton:hover {{
                background: #8e8e8e; color: #fff; border: 2px solid #666;
            }}
            QPushButton:pressed {{
                background: #444; border: 2px solid #444;
            }}
        """)
        # Khóa đếm
        self.btn_lock = QPushButton("🔒 Khóa đếm")
        self.btn_lock.setCheckable(True)
        self.btn_lock.setMinimumWidth(btn_min_width)
        self.btn_lock.setMinimumHeight(btn_min_height)
        self.btn_lock.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_lock.setToolTip("Khóa/mở khóa số đếm của trang hiện tại (viền cam)")
        self.btn_lock.setStyleSheet(f"""
            QPushButton {{
                background: #666; color: white; border-radius:5px; font-size: {btn_font};
            }}
            QPushButton:checked, QPushButton:pressed {{
                background: orange; color: black; font-weight:bold; border: 2px solid #e88500;
            }}
            QPushButton:hover {{
                background: #8e8e8e; color: #fff; border: 2px solid orange;
            }}
        """)
        toolbar.addWidget(self.btn_export)
        toolbar.addWidget(self.btn_reset)
        toolbar.addWidget(self.btn_lock)
        toolbar.addStretch(1)
        main_layout.addLayout(toolbar)

        # ==== THAY ĐỔI: label trạng thái chỉ còn 1 label (hiển thị lặp) ====
        self.status_label = QLabel("Đang chờ kết nối...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.status_label.setStyleSheet("font-size: 16pt; font-weight: bold; color: orange; margin:0px 0;")
        main_layout.addWidget(self.status_label)

        self.status_opacity = QGraphicsOpacityEffect()
        self.status_label.setGraphicsEffect(self.status_opacity)
        self.status_opacity.setOpacity(1.0)

        self.status_anim = QPropertyAnimation(self.status_opacity, b"opacity")
        self.status_anim.setDuration(350)  # thời gian chuyển hiệu ứng (ms)
        self.status_anim.setEasingCurve(QEasingCurve.InOutCubic)
        self._pending_status = None  # giá trị sẽ set sau hiệu ứng mờ
        self._fade_text = None
        self._fade_style = None
        self.status_anim.finished.connect(self._do_status_fadein)
        self._fade_in_stage = False  # Thêm biến kiểm soát giai đoạn hiệu ứng
        self._fade_stage = "idle"  # "idle", "fadeout", "fadein"
        self._next_status = None

        # Thông báo cảnh báo nằm dưới trạng thái kết nối/chế độ rung
        self.server_log_label = QLabel("")
        self.server_log_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.server_log_label.setStyleSheet("font-size: 11pt; color: gray; margin: 0px 0;")
        main_layout.addWidget(self.server_log_label)

        self.info_labels = {}
        info_fields = ["VID", "PID", "Serial", "Product"]
        info_vbox = QVBoxLayout()
        for field in info_fields:
            lbl = QLabel(f"{field}: ...")
            lbl.setStyleSheet("font-size: 12pt; margin-top:2px; margin-bottom:2px;")
            lbl.setMinimumHeight(28)
            self.info_labels[field] = lbl
            info_vbox.addWidget(lbl)
        main_layout.addLayout(info_vbox)

        data_widget = QWidget()
        self.data_layout = QGridLayout(data_widget)
        main_layout.addWidget(data_widget, stretch=1)

        self.scroll_area = QScrollArea()
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setWidgetResizable(True)
        self.pagebar_content = QWidget()
        self.pagebar_layout = QHBoxLayout(self.pagebar_content)
        self.pagebar_layout.setContentsMargins(4, 2, 4, 2)
        self.pagebar_layout.setSpacing(6)
        self.scroll_area.setWidget(self.pagebar_content)
        main_layout.addWidget(self.scroll_area)

        self.pages = []
        self.known_keys = set()
        self.current_page_idx = None
        self.page_buttons = []
        self.info_fields = info_fields
        self.locked_pages = set()

        self.label_min_width = 160
        self.label_min_height = 36
        self.fontsize = 13

        self.status_buttons = {}
        
        self.last_button_text = {}
        self.last_button_style = {}
        
        self.status_labels = {}
        self.display_name_map = {
            "LB": "Bumper",
            "RB": "Bumper",
            "LS": "Thumbstick",
            "RS": "Thumbstick",
            "LS_X": "X",
            "LS_Y": "Y",
            "RS_X": "X",
            "RS_Y": "Y",
            "LT": "Trigger",
            "RT": "Trigger",
        }
        self.create_status_items()

        self.btn_export.clicked.connect(self.export_results)
        self.btn_reset.clicked.connect(self.reset_all)
        self.btn_lock.clicked.connect(self.toggle_lock)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_status)
        self.timer.start(33)

        self.scroll_area.horizontalScrollBar().valueChanged.connect(self.adjust_pagebar_width)
        self.resizeEvent = self.on_resize_event

        self.prev_joystick_count = 0

        self.last_hid_scan = 0

        self.update_lock_button()

        self.last_good_method = None
        self.last_good_port = None

    def set_status_with_fade(self, text, style):
        # Nếu đang đổi trạng thái (chưa fade xong), lưu vào hàng đợi
        if self._fade_stage != "idle":
            self._next_status = (text, style)
            return

        # Nếu nội dung không đổi, bỏ qua fade cho mượt
        if text == self.status_label.text() and style == self.status_label.styleSheet():
            return

        # Bắt đầu fade-out
        self._fade_stage = "fadeout"
        self._pending_text = text
        self._pending_style = style
        self.status_anim.stop()
        self.status_anim.setStartValue(1.0)
        self.status_anim.setEndValue(0.0)
        self.status_anim.start()


    def _do_status_fadein(self):
        if self._fade_stage == "fadeout":
            # Kết thúc fade-out: đổi nội dung, bắt đầu fade-in
            self.status_label.setText(self._pending_text)
            self.status_label.setStyleSheet(self._pending_style)
            self._fade_stage = "fadein"
            self.status_anim.stop()
            self.status_anim.setStartValue(0.0)
            self.status_anim.setEndValue(1.0)
            self.status_anim.start()
        elif self._fade_stage == "fadein":
            # Kết thúc fade-in: quay lại trạng thái idle, kiểm tra nếu có status mới đợi
            self._fade_stage = "idle"
            if self._next_status:
                text, style = self._next_status
                self._next_status = None
                self.set_status_with_fade(text, style)

    
    def showEvent(self, event):
        super().showEvent(event)
        screen = QApplication.primaryScreen()
        geo = screen.availableGeometry()
        frame = self.frameGeometry()
        x = geo.right() - frame.width() + 1
        y = geo.bottom() - frame.height() + 1
        self.move(x, y)

    def is_invalid_device_info(self, info):
        return not info or all((str(val).upper() == 'N/A' or val is None or val == "") for val in info.values())

    def create_status_items(self):
        resettable_names = (
            ["Guide", "Menu", "View", "Share",
             "A", "B", "X", "Y",
             "Down", "Right", "Left", "Up",
             "LB", "RB", "LS", "RS", "LT", "RT"]
        )
        button_names = {
            0: "A", 1: "B", 2: "X", 3: "Y",
            4: "LB", 5: "RB", 6: "View", 7: "Menu",
            8: "LS", 9: "RS", 10: "Guide", 11: "Share"
        }
        ordered_dpad = ["Down", "Right", "Left", "Up"]
        abxy_order = ["A", "B", "X", "Y"]
        self.button_names = button_names
        self.ordered_dpad = ordered_dpad
        self.abxy_order = abxy_order
        self.trigger_axes = {"LT": 4, "RT": 5}
        self.trigger_prev_val = {"LT": -1.0, "RT": -1.0}
        self.stick_axes = {
            "LS_X": 0, "LS_Y": 1,
            "RS_X": 2, "RS_Y": 3
        }
        self.press_count = {name: 0 for name in
               list(button_names.values()) + ordered_dpad + ["LT", "RT"]}
        self.button_prev_state = {index: False for index in button_names}
        self.dpad_prev_state = set()

        row = 0
        def additem(name, row, col):
            display_name = self.display_name_map.get(name, name)
            if name in resettable_names:
                btn = QPushButton(f"{display_name}: 0")
                btn.setMinimumSize(self.label_min_width, self.label_min_height)
                btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background: #eee; font-size: {self.fontsize}pt; border: 1px solid #ccc;
                    }}
                    QPushButton:hover {{
                        border: 2px solid orange;
                    }}
                """)
                btn.setToolTip("Bấm vào đây để reset số đếm!")
                btn.clicked.connect(partial(self.reset_single, name))
                self.data_layout.addWidget(btn, row, col)
                self.status_buttons[name] = btn
            else:
                lbl = QLabel(f"{display_name}: 0.000")
                lbl.setMinimumSize(self.label_min_width, self.label_min_height)
                lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                lbl.setStyleSheet(f"background: #fafaff; font-size: {self.fontsize}pt; border: 1px solid #ccc;")
                self.data_layout.addWidget(lbl, row, col)
                self.status_labels[name] = lbl

        additem("Guide", row, 0)
        additem("Share", row, 4)
        row += 1
        additem("View", row, 0)
        additem("Menu", row, 4)
        for i, name in enumerate(self.ordered_dpad):
            additem(name, row + 1 + i, 0)
        for i, name in enumerate(self.abxy_order):
            additem(name, row + 1 + i, 4)
        row += 5
        additem("LB", row + 1, 0)
        additem("RB", row + 1, 4)
        additem("LS", row + 2, 0)
        additem("RS", row + 2, 4)
        additem("LS_X", row + 3, 0)
        additem("RS_X", row + 3, 4)
        additem("LS_Y", row + 4, 0)
        additem("RS_Y", row + 4, 4)
        additem("LT", row + 5, 0)
        additem("RT", row + 5, 4)

    def scroll_to_current_page(self):
        if self.current_page_idx is not None and self.page_buttons:
            btn_widget = self.page_buttons[self.current_page_idx][2]
            scrollBar = self.scroll_area.horizontalScrollBar()
            btn_x = btn_widget.pos().x()
            btn_w = btn_widget.size().width()
            viewport_w = self.scroll_area.viewport().width()
            new_value = btn_x - max(0, (viewport_w - btn_w) // 2)
            scrollBar.setValue(new_value)

    def add_new_page(self, device_key, info_dict):
        data = {name: 0 for name in self.press_count}
        page = {
            "key": device_key,
            "data": data.copy(),
            "info": info_dict.copy()
        }
        self.pages.append(page)
        self.known_keys.add(device_key)
        self.switch_page(len(self.pages)-1)
        self.update_pagebar()
        self.update_lock_button()
        QTimer.singleShot(50, self.scroll_to_current_page)

    def switch_page(self, idx):
        if not (0 <= idx < len(self.pages)): return
        if self.current_page_idx == idx: return
        self.current_page_idx = idx
        page = self.pages[idx]
        for field in self.info_fields:
            self.info_labels[field].setText(f"{field}: {page['info'].get(field,'...')}")
        self.last_button_text.clear()
        self.last_button_style.clear()
        self.update_pagebar()
        self.update_lock_button()
        QTimer.singleShot(50, self.scroll_to_current_page)

    def toggle_lock(self):
        idx = self.current_page_idx
        if idx is None:
            return
        if self.btn_lock.isChecked():
            self.locked_pages.add(idx)
        else:
            self.locked_pages.discard(idx)
        self.update_lock_button()
        self.update_pagebar()

    def update_lock_button(self):
        idx = self.current_page_idx
        self.btn_lock.setMinimumHeight(TOOLBAR_BTN_HEIGHT)
        if idx is not None and idx in self.locked_pages:
            self.btn_lock.setChecked(True)
            self.btn_lock.setText("🔓 Đã khóa")
            self.btn_lock.setToolTip("Trang này đang bị khóa số đếm")
        else:
            self.btn_lock.setChecked(False)
            self.btn_lock.setText("🔒 Khóa đếm")
            self.btn_lock.setToolTip("Bấm để khóa số đếm trang hiện tại")

    def update_button_status(self, name, btn, new_text, new_style):
        """
        Hàm rút gọn cho việc cập nhật text và style cho nút, chỉ update khi giá trị thực sự thay đổi.
        """
        if self.last_button_text.get(name) != new_text:
            btn.setText(new_text)
            self.last_button_text[name] = new_text
        if self.last_button_style.get(name) != new_style:
            btn.setStyleSheet(new_style)
            self.last_button_style[name] = new_style

    def update_pagebar(self):
        for tup in self.page_buttons:
            tup[2].deleteLater()  # xóa widget cha
        self.page_buttons.clear()
        for i, page in enumerate(self.pages):
            btn = QPushButton(f"Trang {i+1}")
            btn.setCheckable(True)
            btn.setChecked(i == self.current_page_idx)
            btn.setMinimumWidth(120)
            btn.setMinimumHeight(TOOLBAR_BTN_HEIGHT)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            if i == self.current_page_idx and i in self.locked_pages:
                border_style = "border:2.5px solid orange;"
            elif i == self.current_page_idx:
                border_style = "border:3px solid #0d47a1;"
            elif i in self.locked_pages:
                border_style = "border:2.5px solid orange;"
            else:
                border_style = "border:1px solid #ccc;"
            base_style = f"""
                QPushButton {{
                    font-size:{TOOLBAR_BTN_FONT_SIZE}pt; min-width:100px; min-height:{TOOLBAR_BTN_HEIGHT}px;
                    background:#1877f2; color:white; border-radius:6px;
                    {border_style}
                }}
                QPushButton:checked {{
                    background:#1877f2; color:white;
                    {border_style}
                }}
                QPushButton:hover {{
                    background: #2490fa; color: #fff;
                    {border_style}
                }}
                QPushButton:pressed {{
                    background: #185fc2; color: #fff;
                    {border_style}
                }}
            """
            btn.setStyleSheet(base_style)
            btn.setToolTip(self.make_page_tooltip(i))
            btn.clicked.connect(lambda _, idx=i: self.switch_page(idx))
            btn_x = QPushButton("✕")
            btn_x.setFixedSize(TOOLBAR_BTN_HEIGHT, TOOLBAR_BTN_HEIGHT)
            btn_x.setStyleSheet(f"""
                QPushButton {{ background: transparent; color: #b00; border:none; font-weight: bold; font-size:{TOOLBAR_BTN_FONT_SIZE}pt; }}
                QPushButton:hover {{ color: #ff3333; border:1px solid #ff9999; }}
            """)
            btn_x.setToolTip("Xóa trang này")
            btn_x.clicked.connect(partial(self.delete_page, i))
            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            btn_layout.setSpacing(0)
            btn_layout.addWidget(btn)
            btn_layout.addWidget(btn_x)
            self.pagebar_layout.addWidget(btn_widget)
            self.page_buttons.append((btn, btn_x, btn_widget))
        self.adjust_pagebar_width()

    def make_page_tooltip(self, idx):
        if not (0 <= idx < len(self.pages)): return ""
        info = self.pages[idx].get('info', {})
        return "\n".join([f"{k}: {v}" for k, v in info.items()])

    def delete_page(self, idx):
        if not (0 <= idx < len(self.pages)): return
        reply = QMessageBox.question(self, "Xác nhận", "Bạn chắc chắn muốn xóa trang này?", QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        del self.pages[idx]
        self.known_keys = set(page["key"] for page in self.pages)
        self.locked_pages = set(i for i in self.locked_pages if i < len(self.pages))
        if self.current_page_idx is not None:
            if self.current_page_idx >= len(self.pages):
                self.current_page_idx = len(self.pages)-1 if self.pages else None
        self.update_pagebar()
        self.update_lock_button()
        QTimer.singleShot(50, self.scroll_to_current_page)

    def reset_single(self, name):
        if self.current_page_idx is not None:
            if name in self.press_count:
                self.pages[self.current_page_idx]["data"][name] = 0
                self.pages[self.current_page_idx]["data"][name] = 0
                self.last_button_text.clear()
                self.last_button_style.clear()

    def reset_all(self):
        if self.current_page_idx is not None:
            for name in self.press_count:
                self.pages[self.current_page_idx]["data"][name] = 0
                self.last_button_text.clear()
                self.last_button_style.clear()

    def export_results(self):
        if self.current_page_idx is None: return
        page = self.pages[self.current_page_idx]
        dialog = QFileDialog(self)
        path, _ = dialog.getSaveFileName(self, "Xuất kết quả ra Excel", filter="Excel Files (*.xlsx)")
        if not path: return
        wb = Workbook()
        ws = wb.active
        ws.title = "Test Results"
        ws.append(["Tên", "Số lần"])
        for name, count in page["data"].items():
            ws.append([self.display_name_map.get(name, name), count])
        wb.save(path)
        QMessageBox.information(self, "Đã lưu", f"Kết quả đã được xuất vào:\n{path}")

    def adjust_pagebar_width(self, *args):
        total_width = sum([btn[2].sizeHint().width() + 10 for btn in self.page_buttons])
        viewport_width = self.scroll_area.viewport().width()
        self.pagebar_content.setMinimumWidth(max(total_width, viewport_width))
        self.pagebar_content.setMaximumWidth(max(total_width, viewport_width))
        self.pagebar_content.adjustSize()

    def on_resize_event(self, event):
        self.adjust_pagebar_width()
        event.accept()
    @staticmethod
    def get_rumble_mode_text(mode):
        if mode == RumbleMode.MOTOR:
            return "Chế độ rung: Motor lớn"
        elif mode == RumbleMode.TRIGGER:
            return "Chế độ rung: Trigger"
        return "Chế độ rung: Tắt"

    def send_rumble(self, left, right, left_trig, right_trig, duration_ms=0, _retry=False):
        """
        Gửi lệnh rung đến XboxRumbleServer. Nếu duration_ms > 0 sẽ rung giữ đúng thời gian (ms).
        Nếu duration_ms = 0 (hoặc không truyền), server sẽ giữ rung cho đến khi nhận lệnh mới.
        """
        data = f"{left},{right},{left_trig},{right_trig},{duration_ms}"
        ports_to_try = DEFAULT_PORTS if self.last_good_port is None else [self.last_good_port] + [p for p in DEFAULT_PORTS if p != self.last_good_port]
        # Thử từng port TCP
        for port in ports_to_try:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(0.2)
                s.connect(('127.0.0.1', port))
                s.sendall(data.encode('utf-8'))
                s.close()
                if self.last_good_method != "tcp" or self.last_good_port != port:
                    print(f"Gửi rung TCP port {port}")
                self.last_good_method = "tcp"
                self.last_good_port = port
                return True
            except Exception:
                pass
        # Nếu lỗi hết thì thử Named Pipe
        try:
            handle = win32file.CreateFile(
                DEFAULT_PIPE,
                win32file.GENERIC_WRITE,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            )
            win32file.WriteFile(handle, data.encode('utf-8'))
            win32file.CloseHandle(handle)
            if self.last_good_method != "pipe":
                print(f"Gửi rung Named Pipe: {DEFAULT_PIPE}")
            self.last_good_method = "pipe"
            self.last_good_port = None
            return True
        except pywintypes.error as e:
            if e.args[0] == 2 and not _retry:
                print("Không kết nối được server. Đang tự khởi động lại XboxRumbleServer...")
                start_server()
                time.sleep(1)
                return self.send_rumble(left, right, left_trig, right_trig, _retry=True)
            else:
                print("Lỗi gửi rung:", e)
        except Exception as e:
            print("Lỗi không xác định:", e)
        return False

    def is_testing_rumble_now(self, motor=None):
        """Trả về True nếu đang test rung motor chỉ định (motor: 'left', 'right', ...).
        Nếu motor=None, kiểm tra tất cả."""
        if not hasattr(self, 'current_test_rumble_state'):
            return False
        if motor is None:
            return any(v > 0 for v in self.current_test_rumble_state.values())
        return self.current_test_rumble_state.get(motor, 0.0) > 0

    def queue_motor_rumble(self, motor, intensity, duration_ms):
        # motor: "left", "right", "left_trigger", "right_trigger"
        # intensity: 0.0 – 1.0
        # duration_ms: int
        end_time = time.time() + duration_ms / 1000.0
        # Nếu đã có rung sẵn và thời điểm mới lớn hơn, thì giữ thời điểm mới
        self.motor_rumble_end_time[motor] = max(self.motor_rumble_end_time[motor], end_time)
        # Nếu muốn intensity cộng dồn, dùng max hoặc + tuỳ bạn. Mặc định dùng intensity mới luôn:
        self.motor_rumble_intensity[motor] = max(self.motor_rumble_intensity[motor], intensity)

    def queue_motor_rumble_flexible(self, motor, intensity, duration_ms):
        test_rumble_now = self.is_testing_rumble_now(motor)
        fallback_motor = MOTOR_FALLBACK.get(motor, None)
        target_motor = motor

        if test_rumble_now and fallback_motor:
            # Nếu motor chính đang bận test rung, thử rung sang motor cùng phía nếu chưa bận
            if not self.is_testing_rumble_now(fallback_motor):
                target_motor = fallback_motor
            else:
                # Nếu cả hai đều bận test rung, bỏ rung cảnh báo hoặc log lại nếu cần
                return

        self.queue_motor_rumble(target_motor, intensity, duration_ms)

    def handle_rumble_alert(self, name, v, display_name, thresholds, alert_func, check_less=True):
        """
        Gom logic cảnh báo rung vào 1 hàm dùng chung.
        - name: tên nút ("LT", "RB", ...)
        - v: số lần hiện tại của nút
        - display_name: tên hiển thị cho người dùng
        - thresholds: danh sách [(nguong, t_rung)] cảnh báo cho nút đó
        - alert_func: hàm thực thi lệnh rung, nhận duration_ms
        - check_less: kiểm tra < nguong để reset cảnh báo (True mặc định)
        """
        for nguong, t_rung in thresholds:
            duration_ms = int(t_rung * 1000)
            # Khi đạt ngưỡng và chưa cảnh báo -> rung + log
            if v == nguong and nguong not in self.rumble_alerted[name]:
                self.server_log_label.setText(
                    f"Rung cảnh báo {display_name} đạt {nguong} lần!"
                )
                self.rumble_busy = True
                alert_func(duration_ms)
                QTimer.singleShot(duration_ms + 100, self.clear_rumble_busy)
                self.rumble_alerted[name].add(nguong)
            # Nếu quay lại < nguong thì cho phép cảnh báo lại ở lần sau
            if check_less and v < nguong and nguong in self.rumble_alerted[name]:
                self.rumble_alerted[name].remove(nguong)

    def safe_axis(self, val, min_val=-1.0, max_val=1.0, default=0.0):
        """
        Đảm bảo giá trị axis nằm trong [min_val, max_val]. Nếu lỗi hoặc None thì trả về default.
        """
        try:
            if val is None or not isinstance(val, (int, float)):
                return default
            return max(min(val, max_val), min_val)
        except Exception:
            return default

    def safe_button(self, val, default=False):
        """
        Đảm bảo giá trị button chỉ nhận True/False. Nếu lỗi thì trả về default.
        """
        try:
            return bool(val)
        except Exception:
            return default


    def update_status(self):

        self.check_joystick_connection()

        pygame.event.pump()

        self.handle_combo_rumble_mode()

        self.handle_rumble_logic()
        
        count = self.refresh_gamepad_count()
        
        pressed_states = {}
        trigger_values = {}

        # Nếu chưa có joystick hoặc chưa init, mà count > 0 thì khởi tạo lại
        if (self.joystick is None or not self.joystick.get_init()) and count > 0:
            try:
                self.joystick = pygame.joystick.Joystick(0)
                self.joystick.init()
            except pygame.error:
                self.joystick = None

        connected = (self.joystick is not None and self.joystick.get_init())

        if connected:
            device_key = self.get_joystick_hid_key() if self.joystick and self.joystick.get_init() else None
            info_dict = self.get_joystick_hid_info() or {field: "N/A" for field in self.info_fields}
            page_keys = [page["key"] for page in self.pages]
            if device_key is not None and device_key not in self.known_keys:
                if not self.is_invalid_device_info(info_dict):
                    self.add_new_page(device_key, info_dict)
            elif device_key in page_keys:
                idx = page_keys.index(device_key)
                if self.current_page_idx != idx:
                    self.switch_page(idx)
            self.update_button_and_trigger_states(pressed_states, trigger_values)

        self.update_ui_display(pressed_states, trigger_values)

        # Cập nhật lại Product nếu phát hiện đổi model
        hid_list = self.hid_cache
        for page in self.pages:
            for d in hid_list:
                if (d["serial"] or "N/A") == page["key"]:
                    # Cập nhật lại Product và PID!
                    page["info"]["Product"] = d.get("product") or "Unknown"
                    page["info"]["PID"] = f"{d['pid']:04X}"
                    page["info"]["VID"] = f"{d['vid']:04X}"
        # Nếu đang có page đang chọn, update giao diện product
        if self.current_page_idx is not None and self.pages:
            page = self.pages[self.current_page_idx]
            for field in self.info_fields:
                val = page['info'].get(field, '...')
                self.info_labels[field].setText(f"{field}: {val}")

    def check_joystick_connection(self):
        """
        Kiểm tra kết nối của joystick.
        Nếu joystick bị rút hoặc chưa cắm, sẽ đặt lại self.joystick = None.
        """
        if self.joystick is not None and not self.joystick.get_init():
            self.joystick = None

    def handle_combo_rumble_mode(self):
        """
        Bắt tổ hợp LB + RB + D-Pad lên/xuống để chuyển đổi chế độ rung (Motor lớn/Trigger/Tắt).
        Cập nhật trạng thái giao diện và tự động khởi động XboxRumbleServer nếu cần.
        """
        try:
            lb = self.joystick.get_button(4) if self.joystick else 0
            rb = self.joystick.get_button(5) if self.joystick else 0
            dpad = self.joystick.get_hat(0) if self.joystick else (0, 0)
        except Exception:
            lb, rb, dpad = 0, 0, (0, 0)

        combo_up = lb and rb and (dpad[1] == 1)
        combo_down = lb and rb and (dpad[1] == -1)

        if combo_up and not self.last_combo_up:
            self.rumble_mode = RumbleMode.MOTOR if self.rumble_mode != RumbleMode.MOTOR else RumbleMode.OFF
            
            self.server_log_label.setText("Đang bật chế độ test rung motor lớn" if self.rumble_mode == RumbleMode.MOTOR else "Đã tắt chế độ rung")
            # Thêm đoạn reset test rung:
            if self.rumble_mode == RumbleMode.OFF:
                for key in self.current_test_rumble_state:
                    self.current_test_rumble_state[key] = 0.0
            if self.rumble_mode == RumbleMode.MOTOR:
                if not is_server_running():
                    self.server_log_label.setText("Đang khởi động XboxRumbleServer...")
                    start_server()
                    self.server_log_label.setText("Đã khởi động XboxRumbleServer!")
        if combo_down and not self.last_combo_down:
            self.rumble_mode = RumbleMode.TRIGGER if self.rumble_mode != RumbleMode.TRIGGER else RumbleMode.OFF
            
            self.server_log_label.setText("Đang bật chế độ test rung trigger" if self.rumble_mode == RumbleMode.TRIGGER else "Đã tắt chế độ rung")
            # Thêm đoạn reset test rung:
            if self.rumble_mode == RumbleMode.OFF:
                for key in self.current_test_rumble_state:
                    self.current_test_rumble_state[key] = 0.0
            if self.rumble_mode == RumbleMode.TRIGGER:
                if not is_server_running():
                    self.server_log_label.setText("Đang khởi động XboxRumbleServer...")
                    start_server()
                    self.server_log_label.setText("Đã khởi động XboxRumbleServer!")
        self.last_combo_up = combo_up
        self.last_combo_down = combo_down

    def handle_rumble_logic(self):
        """
        Xử lý gửi lệnh rung tùy chế độ (Motor lớn/Trigger/Tắt).
        - Lấy giá trị trigger trái/phải
        - Nếu giá trị rung đổi: gửi lệnh rung mới, reset timer giữ rung
        - Nếu không đổi nhưng cần giữ rung: gửi lại lệnh sau 4s/8s
        """
        left = right = left_trigger = right_trigger = 0.0
        # Lấy giá trị trục LT/RT
        lt_raw = self.joystick.get_axis(4) if self.joystick and self.joystick.get_init() else -1.0
        rt_raw = self.joystick.get_axis(5) if self.joystick and self.joystick.get_init() else -1.0

        is_lt_active = lt_raw > -0.99
        is_rt_active = rt_raw > -0.99

        if not self.rumble_busy:
            if self.rumble_mode == RumbleMode.MOTOR:
                left = (lt_raw + 1.0) / 2.0 if is_lt_active else 0.0
                right = (rt_raw + 1.0) / 2.0 if is_rt_active else 0.0
            elif self.rumble_mode == RumbleMode.TRIGGER:
                left_trigger = (lt_raw + 1.0) / 2.0 if is_lt_active else 0.0
                right_trigger = (rt_raw + 1.0) / 2.0 if is_rt_active else 0.0
            # Cập nhật trạng thái test rung
            self.current_test_rumble_state["left"] = left
            self.current_test_rumble_state["right"] = right
            self.current_test_rumble_state["left_trigger"] = left_trigger
            self.current_test_rumble_state["right_trigger"] = right_trigger

    def refresh_gamepad_count(self):
        """
        Kiểm tra số lượng gamepad hiện tại.
        Nếu số lượng thay đổi, khởi tạo lại joystick & quét HID ngay.
        Nếu không, chỉ quét HID theo chu kỳ thời gian để cập nhật cache.
        Trả về số lượng tay cầm hiện tại.
        """
        count = pygame.joystick.get_count()
        if not hasattr(self, "prev_joystick_count"):
            self.prev_joystick_count = 0

        # Nếu số lượng tay cầm đổi, quét lại HID ngay lập tức
        if count != self.prev_joystick_count:
            pygame.joystick.quit()
            pygame.joystick.init()
            self.prev_joystick_count = count
            self.find_gamepads_hid(force=True)
        # Ngược lại, chỉ quét lại HID định kỳ mỗi 1–2s
        elif time.time() - self.last_hid_scan > HID_SCAN_INTERVAL:
            self.find_gamepads_hid()
            self.last_hid_scan = time.time()
        return count

    def update_button_and_trigger_states(self, pressed_states, trigger_values):
        """
        Cập nhật trạng thái các nút, D-Pad, và Trigger.
        Tăng số đếm, cập nhật pressed_states và trigger_values cho page hiện tại.
        """
        if self.current_page_idx is None or not self.pages:
            return

        page = self.pages[self.current_page_idx]
        locked = self.current_page_idx in self.locked_pages

        # Xử lý các nút thường (A, B, X, Y, LB, RB, Menu, View, Share, Guide, LS, RS)
        for i, name in self.button_names.items():
            pressed = False
            if self.joystick and self.joystick.get_init():
                try:
                    pressed = self.safe_button(self.joystick.get_button(i)) if self.joystick and self.joystick.get_init() else False
                except Exception:
                    pressed = False
            pressed_states[name] = pressed
            v = page["data"].get(name, 0)
            if not locked and pressed and not self.button_prev_state[i]:
                page["data"][name] += 1
                v = page["data"][name]
            self.button_prev_state[i] = pressed

        # Xử lý D-Pad
        dpad_map = {(0, -1): "Down", (1, 0): "Right", (-1, 0): "Left", (0, 1): "Up"}
        current_dpad = set()
        num_hats = self.joystick.get_numhats() if self.joystick and self.joystick.get_init() else 0
        for i in range(num_hats):
            hat = self.joystick.get_hat(i)
            if not isinstance(hat, tuple) or len(hat) != 2:
                hat = (0, 0)
            if hat in dpad_map:
                current_dpad.add(dpad_map[hat])
        for name in self.ordered_dpad:
            active = name in current_dpad
            pressed_states[name] = active
            v = page["data"].get(name, 0)
            if not locked and active and name not in self.dpad_prev_state:
                page["data"][name] += 1
                v = page["data"][name]
            if active:
                self.dpad_prev_state.add(name)
            else:
                self.dpad_prev_state.discard(name)

        # Xử lý Trigger (LT/RT)
        for trig in ["LT", "RT"]:
            axis = self.trigger_axes[trig]
            raw_val = -1.0
            val = 0.0
            active = False
            if self.joystick and self.joystick.get_init():
                try:
                    raw_val = self.safe_axis(self.joystick.get_axis(axis))
                    val = round((raw_val + 1.0) / 2.0, 3)
                except Exception:
                    raw_val = -1.0
                    val = 0.0
                active = raw_val > 0.95
            trigger_values[trig] = val
            v = page["data"].get(trig, 0)
            if not locked and active and self.trigger_prev_val[trig] <= 0.95:
                page["data"][trig] += 1
                v = page["data"][trig]
            self.trigger_prev_val[trig] = raw_val

    def update_ui_display(self, pressed_states, trigger_values):
        """
        Cập nhật hiển thị giao diện nút, màu sắc, label, trạng thái, cảnh báo rung, v.v.
        Cập nhật giao diện, bao gồm logic lặp trạng thái 'Đã kết nối' / 'Chế độ rung: Tắt'
        Không được setText hay setStyleSheet cho self.status_label ở ngoài logic phía trên nữa!
        Chỉ duy nhất đoạn điều khiển trạng thái dưới đây mới cập nhật trạng thái giao diện!
        """
        if self.current_page_idx is None or not self.pages:
            # Nếu không có tay cầm
            self.set_status_with_fade(
                "Đang chờ kết nối...",
                "font-size: 16pt; font-weight: bold; color: orange; margin:0px 0;"
            )
            self.connected_display_mode = False
            self.was_connected = False
            return

        page = self.pages[self.current_page_idx]
        cur_key = self.get_joystick_hid_key() if self.joystick and self.joystick.get_init() else None
        connected = self.joystick and self.joystick.get_init() and cur_key == page["key"]

        # ==== ĐIỀU KHIỂN HIỂN THỊ TRẠNG THÁI THEO YÊU CẦU ====

        now = time.time()
        # Nếu vừa mới kết nối thì reset lại timer và bật chế độ hiển thị lặp
        if connected and not self.was_connected:
            self.connected_state_last_toggle = now
            self.connected_state_stage = 0
            self.connected_display_mode = True

        # Nếu vừa ngắt kết nối thì về trạng thái chờ
        if not connected:
            self.set_status_with_fade(
                "Đang chờ kết nối...",
                "font-size: 16pt; font-weight: bold; color: orange; margin:0px 0;"
            )
            self.connected_display_mode = False
            self.was_connected = False
        else:
            # Nếu đang bật chế độ rung (không phải OFF) thì chỉ hiển thị chế độ rung
            if self.rumble_mode == RumbleMode.MOTOR:
                self.set_status_with_fade(
                    "Chế độ rung: Motor",
                    "font-size: 16pt; font-weight: bold; color: deepskyblue; margin:0px 0;"
                )
                self.connected_display_mode = False
            elif self.rumble_mode == RumbleMode.TRIGGER:
                self.set_status_with_fade(
                    "Chế độ rung: Trigger",
                    "font-size: 16pt; font-weight: bold; color: deepskyblue; margin:0px 0;"
                )
                self.connected_display_mode = False
            else:
                # Chỉ hiển thị lặp "Đã kết nối" / "Chế độ rung: Tắt" khi chế độ rung là OFF
                if self.connected_display_mode:
                    if now - self.connected_state_last_toggle > 5:
                        self.connected_state_stage = 1 - self.connected_state_stage
                        self.connected_state_last_toggle = now
                    if self.connected_state_stage == 0:
                        self.set_status_with_fade(
                            "Đã kết nối",
                            "font-size: 16pt; font-weight: bold; color: green; margin:0px 0;"
                        )
                    else:
                        self.set_status_with_fade(
                            "Chế độ rung: Tắt",
                            "font-size: 16pt; font-weight: bold; color: deepskyblue; margin:0px 0;"
                        )
                else:
                    # Nếu vừa tắt chế độ rung, bắt đầu lại vòng lặp
                    self.connected_display_mode = True
                    self.connected_state_stage = 0
                    self.connected_state_last_toggle = now
            self.was_connected = True

        for name, btn in self.status_buttons.items():
            v = page["data"].get(name, 0)
            display_name = self.display_name_map.get(name, name)

            if name in ["LT", "RT"]:
                val = trigger_values.get(name, 0.0)
                new_text = f"{display_name}: {val:.3f} ({v})"
                is_durham = page["info"].get("Product", "").strip().lower() == "durham controller"
                if is_durham:
                    # Durham chia 9 phần
                    if val > 0.000:
                        new_style = f"QPushButton {{ background: orange; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 44:
                        grad = "background: deepskyblue;"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 39:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.888 deepskyblue, stop:0.888 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 34:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.777 deepskyblue, stop:0.777 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 29:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.666 deepskyblue, stop:0.666 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 24:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.555 deepskyblue, stop:0.555 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 19:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.444 deepskyblue, stop:0.444 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 14:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.333 deepskyblue, stop:0.333 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 9:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.222 deepskyblue, stop:0.222 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    elif v > 4:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.111 deepskyblue, stop:0.111 #eee, stop:1 #eee);"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    else:
                        new_style = f"QPushButton {{ background: #eee; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                else:
                    is_pressed = val > 0.000
                    if is_pressed:
                        new_style = f"QPushButton {{ background: orange; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                    else:
                        if v > 14:
                            grad = "background: deepskyblue;"
                        elif v > 9:
                            grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.666 deepskyblue, stop:0.666 #eee, stop:1 #eee);"
                        elif v > 4:
                            grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.333 deepskyblue, stop:0.333 #eee, stop:1 #eee);"
                        else:
                            grad = "background: #eee;"
                        new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"

                self.update_button_status(name, btn, new_text, new_style)

                # Cảnh báo rung đặc biệt cho trigger
                product_name = page["info"].get("Product", "").strip().lower()
                thresholds = self.trigger_rumble_thresholds.get(product_name, [])
                def trigger_alert_func(duration_ms, n=name):
                    if n == "LT":
                        self.queue_motor_rumble_flexible("left_trigger", 1.0, duration_ms)
                    else:
                        self.queue_motor_rumble_flexible("right_trigger", 1.0, duration_ms)
                self.handle_rumble_alert(name, v, display_name, thresholds, trigger_alert_func)

            elif name in ["LB", "RB"]:
                new_text = f"{display_name}: {v}"
                is_pressed = pressed_states.get(name, False)
                if is_pressed:
                    new_style = f"QPushButton {{ background: orange; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                else:
                    if v > 29:
                        grad = "background: deepskyblue;"
                    elif v > 19:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.666 deepskyblue, stop:0.666 #eee, stop:1 #eee);"
                    elif v > 9:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.333 deepskyblue, stop:0.333 #eee, stop:1 #eee);"
                    else:
                        grad = "background: #eee;"
                    new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                self.update_button_status(name, btn, new_text, new_style)
                if name in self.nguong_rung:
                    def button_alert_func(duration_ms, n=name):
                        if n in ["LB", "Down", "Right", "Left", "Up", "View", "LS", "Guide"]:
                            self.queue_motor_rumble_flexible("left", 1.0, duration_ms)
                        else:
                            self.queue_motor_rumble_flexible("right", 1.0, duration_ms)
                    self.handle_rumble_alert(name, v, display_name, self.nguong_rung[name], button_alert_func)

            elif name == "Guide":
                new_text = f"{display_name}: {v}"
                is_pressed = pressed_states.get(name, False)
                if is_pressed:
                    new_style = f"QPushButton {{ background: orange; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                else:
                    if v > 24:
                        grad = "background: deepskyblue;"
                    elif v > 19:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.8 deepskyblue, stop:0.8 #eee, stop:1 #eee);"
                    elif v > 14:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.6 deepskyblue, stop:0.6 #eee, stop:1 #eee);"
                    elif v > 9:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.4 deepskyblue, stop:0.4 #eee, stop:1 #eee);"
                    elif v > 4:
                        grad = "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 deepskyblue, stop:0.2 deepskyblue, stop:0.2 #eee, stop:1 #eee);"
                    else:
                        grad = "background: #eee;"
                    new_style = f"QPushButton {{ {grad} font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                self.update_button_status(name, btn, new_text, new_style)
                if name in self.nguong_rung:
                    def button_alert_func(duration_ms, n=name):
                        if n in ["LB", "Down", "Right", "Left", "Up", "View", "LS", "Guide"]:
                            self.queue_motor_rumble_flexible("left", 1.0, duration_ms)
                        else:
                            self.queue_motor_rumble_flexible("right", 1.0, duration_ms)
                    self.handle_rumble_alert(name, v, display_name, self.nguong_rung[name], button_alert_func)

            else:
                new_text = f"{display_name}: {v}"
                is_pressed = pressed_states.get(name, False)
                if is_pressed:
                    new_style = f"QPushButton {{ background: orange; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                elif v > 9:
                    new_style = f"QPushButton {{ background: deepskyblue; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                else:
                    new_style = f"QPushButton {{ background: #eee; font-size: {self.fontsize}pt; border: 1px solid #ccc; }} QPushButton:hover {{ border: 2px solid orange; }}"
                self.update_button_status(name, btn, new_text, new_style)

                # Cảnh báo rung các nút thường (ngoại trừ trigger)
                if name in self.nguong_rung:
                    def button_alert_func(duration_ms, n=name):
                        if n in ["LB", "Down", "Right", "Left", "Up", "View", "LS", "Guide"]:
                            self.queue_motor_rumble_flexible("left", 1.0, duration_ms)
                        else:
                            self.queue_motor_rumble_flexible("right", 1.0, duration_ms)
                    self.handle_rumble_alert(name, v, display_name, self.nguong_rung[name], button_alert_func)

        # Cập nhật thumbstick
        cur_key = self.get_joystick_hid_key() if self.joystick and self.joystick.get_init() else None
        connected = self.joystick and self.joystick.get_init() and cur_key == page["key"]
        for name, axis in self.stick_axes.items():
            val = 0.0
            if connected:
                try:
                    val = round(self.safe_axis(self.joystick.get_axis(axis)), 3)
                    if name.endswith("_Y"):
                        val *= -1
                except Exception:
                    val = 0.0
            display_name = self.display_name_map.get(name, name)
            if name in self.status_labels:
                self.status_labels[name].setText(f"{display_name}: {val:.3f}")
                if abs(val) > 0.150:
                    self.status_labels[name].setStyleSheet(
                        f"background: orange; font-size: {self.fontsize}pt; border: 1px solid #ccc;"
                    )
                else:
                    self.status_labels[name].setStyleSheet(
                        f"background: #fafaff; font-size: {self.fontsize}pt; border: 1px solid #ccc;"
                    )

        # Cập nhật label trạng thái kết nối và info device
        for field in self.info_fields:
            val = page['info'].get(field, '...')
            self.info_labels[field].setText(f"{field}: {val}")
        self.process_motor_rumble_queue()

    def process_motor_rumble_queue(self):
        now = time.time()
        state = {
            "left": 0.0,
            "right": 0.0,
            "left_trigger": 0.0,
            "right_trigger": 0.0
        }
        for motor in state.keys():
            # Nếu đang test rung motor này thì lấy giá trị từ trạng thái test rung
            if self.is_testing_rumble_now(motor):
                state[motor] = self.current_test_rumble_state.get(motor, 0.0)
            elif self.motor_rumble_end_time[motor] > now:
                state[motor] = self.motor_rumble_intensity[motor]
            else:
                state[motor] = 0.0
                self.motor_rumble_intensity[motor] = 0.0
                self.motor_rumble_end_time[motor] = 0.0

        self.send_rumble(state["left"], state["right"], state["left_trigger"], state["right_trigger"], 0)

    def clear_rumble_busy(self):
        self.rumble_busy = False

    def get_joystick_hid_info(self):
        pygame_name = self.joystick.get_name() if self.joystick else ""
        hid_list = self.hid_cache
        for d in hid_list:
            if pygame_name.strip().lower() in d["product_real"].strip().lower():
                info = {
                    "VID": f"{d['vid']:04X}",
                    "PID": f"{d['pid']:04X}",
                    "Serial": d["serial"] or "N/A",
                    "Product": d["product"] or "N/A",
                }
                return info
        return None

    def get_joystick_hid_key(self):
        pygame_name = self.joystick.get_name() if self.joystick else ""
        hid_list = self.hid_cache
        for d in hid_list:
            if pygame_name.strip().lower() in d["product_real"].strip().lower():
                key = f"{d['serial'] or 'NOSERIAL'}"
                return key
        return pygame_name.replace(" ", "_")

    def find_gamepads_hid(self, force=False):
        now = time.time()
        if force or (now - getattr(self, 'last_hid_scan', 0) > HID_SCAN_INTERVAL):
            hid_list = [d for d in hid.enumerate() if is_controller(d)]
            pid_set = {d['product_id'] for d in hid_list}
            # Định danh theo từng device
            result = []
            for d in hid_list:
                product_real = d.get("product_string") or ""
                # Xác định Product theo từng device
                if d['product_id'] == 0x2ff and pid_set == {0x2ff}:
                    product_display = "Jelling Controller"
                elif 0x2ff in pid_set and 0xb02 in pid_set and d['product_id'] in (0x2ff, 0xb02):
                    product_display = "Durham Controller"
                else:
                    product_display = product_real
                result.append({
                    "vid": d["vendor_id"],
                    "pid": d["product_id"],
                    "serial": d.get("serial_number") or "",
                    "product": product_display,
                    "product_real": product_real,
                    "path": d["path"].decode() if isinstance(d["path"], bytes) else d["path"],
                })
            self.hid_cache = result
            self.last_hid_scan = now
            self.last_pid_set = pid_set
        return self.hid_cache

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ControllerTester()
    win.show()
    sys.exit(app.exec())
