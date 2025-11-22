# -*- coding: utf-8 -*-
"""
AutoClicker 2.4 - 单文件完整版（tkinter + pystray）
功能清单（精简）：
- 多点位管理：添加/删除/编辑/上下移动/重命名
- 每点独立设置：X,Y,名称,点击键(左/右),点内延时(秒),点击次数,点击间隔
- 捕获坐标（全局 F2 / 窗口聚焦 F2）
- 全局快捷键：F2 记录坐标、F3 开始/停止、F4 暂停/继续（需 keyboard）
- 启动前任务预览确认、启动倒计时
- 暂停/继续/停止控制（安全线程化）
- 随机偏移(px ±) 与 随机延时(s ±)
- 托盘最小化（pystray），托盘菜单支持 开/暂停/停止/退出/打开窗口
- 浅/深主题切换（设置中）
- 防误触检测：若鼠标快速移动则自动暂停任务
- 点击统计：总点击数/成功/失败/耗时，实时显示并可导出 click_log.txt
- 自动保存/加载上次配置（click_points.json）
- 导入/导出配置（JSON）
- 自动更新检查（可配置检查 URL；默认为空不会联网）
- 声音提示（Windows winsound）
- 详细日志写入 autoclicker_log.txt
- 点位独立设置：每个点位可独立设置点击次数和点击间隔
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import threading
import time
import pyautogui
import json
import os
import random
import datetime
import math
import sys

# optional libs
try:
    import keyboard

    HAVE_KEYBOARD = True
except Exception:
    HAVE_KEYBOARD = False

try:
    from pystray import Icon, Menu, MenuItem
    from PIL import Image, ImageDraw

    HAVE_PYSTRAY = True
except Exception:
    HAVE_PYSTRAY = False

try:
    import requests

    HAVE_REQUESTS = True
except Exception:
    HAVE_REQUESTS = False

# winsound for Windows notifications
try:
    import winsound

    def play_beep():
        try:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass

except Exception:

    def play_beep():
        pass


# ------------- Files & constants -------------
CONFIG_FILE = "click_points.json"
LOG_FILE = "autoclicker_log.txt"
CLICK_REPORT = "click_log.txt"
VERSION = "2.5.1"  # 更新版本号
# If you have a URL to check updates, set it here.
# For safety, by default it's empty; update_check won't run if empty.
# UPDATE_CHECK_URL = "https://raw.githubusercontent.com/MGHYGitHub/AutoClicker/main/version.json"  # e.g. "https://example.com/autoclicker/version.json"
# 在文件开头的常量定义部分添加备用URL
UPDATE_CHECK_URLS = [
    "https://raw.githubusercontent.com/MGHYGitHub/AutoClicker/main/version.json",
    "https://cdn.jsdelivr.net/gh/MGHYGitHub/AutoClicker@main/version.json",  # jsDelivr CDN
    # 可以添加更多备用源
]


# ------------- Logging & helpers -------------
def log(msg):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}] {msg}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass
    # also print to stdout for debugging
    print(line, end="")


def safe_int(v, default=0):
    try:
        return int(v)
    except Exception:
        return default


def safe_float(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


# ------------- Modern UI Styles -------------
class ModernStyles:
    # Color schemes
    LIGHT_THEME = {
        "bg": "#f8f9fa",
        "card_bg": "#ffffff",
        "accent": "#4361ee",
        "accent_hover": "#3a56d4",
        "secondary": "#6c757d",
        "success": "#28a745",
        "warning": "#ffc107",
        "danger": "#dc3545",
        "text": "#212529",
        "text_light": "#6c757d",
        "border": "#dee2e6",
        "input_bg": "#ffffff",
        "highlight": "#e9ecef",
    }

    DARK_THEME = {
        "bg": "#1a1d23",
        "card_bg": "#252a33",
        "accent": "#5b6bf0",
        "accent_hover": "#4a58d4",
        "secondary": "#6c757d",
        "success": "#2ecc71",
        "warning": "#f39c12",
        "danger": "#e74c3c",
        "text": "#e4e6eb",
        "text_light": "#b0b3b8",
        "border": "#3a3f48",
        "input_bg": "#2d323b",
        "highlight": "#343a46",
    }

    @classmethod
    def get_theme(cls, theme_name):
        return cls.LIGHT_THEME if theme_name == "light" else cls.DARK_THEME

    @classmethod
    def setup_styles(cls):
        style = ttk.Style()

        # Configure modern flat style for buttons
        style.configure(
            "Modern.TButton",
            padding=(15, 8),
            relief="flat",
            borderwidth=0,
            focuscolor="none",
        )

        # Configure modern frame style
        style.configure("Modern.TFrame", relief="flat", borderwidth=0)

        # Configure modern label frame style
        style.configure("Modern.TLabelframe", relief="flat", borderwidth=1)

        style.configure(
            "Modern.TLabelframe.Label", foreground=cls.LIGHT_THEME["accent"]
        )

        # Configure progressbar style
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=cls.LIGHT_THEME["border"],
            background=cls.LIGHT_THEME["accent"],
            borderwidth=0,
            lightcolor=cls.LIGHT_THEME["accent"],
            darkcolor=cls.LIGHT_THEME["accent"],
        )

        # Configure combobox style
        style.configure(
            "Modern.TCombobox",
            fieldbackground=cls.LIGHT_THEME["input_bg"],
            background=cls.LIGHT_THEME["input_bg"],
            borderwidth=1,
            relief="flat",
        )

        # Configure notebook style
        style.configure(
            "Modern.TNotebook", background=cls.LIGHT_THEME["bg"], borderwidth=0
        )

        style.configure(
            "Modern.TNotebook.Tab",
            padding=(20, 8),
            background=cls.LIGHT_THEME["highlight"],
            foreground=cls.LIGHT_THEME["text"],
            borderwidth=0,
        )

        style.map(
            "Modern.TNotebook.Tab",
            background=[
                ("selected", cls.LIGHT_THEME["accent"]),
                ("active", cls.LIGHT_THEME["accent_hover"]),
            ],
            foreground=[("selected", "#ffffff")],
        )


# ------------- Glass Frame Effect -------------
class GlassFrame(tk.Frame):
    def __init__(self, parent, blur_color=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.blur_color = blur_color or "#ffffff"

    def apply_glass_effect(self, theme):
        """Apply glass morphism effect to the frame"""
        if theme == "light":
            self.configure(
                bg=ModernStyles.LIGHT_THEME["card_bg"],
                highlightbackground=ModernStyles.LIGHT_THEME["border"],
                highlightthickness=1,
                relief="flat",
            )
        else:
            self.configure(
                bg=ModernStyles.DARK_THEME["card_bg"],
                highlightbackground=ModernStyles.DARK_THEME["border"],
                highlightthickness=1,
                relief="flat",
            )


# ------------- Modern Button -------------
class ModernButton(tk.Button):
    def __init__(self, parent, *args, **kwargs):
        # Extract custom parameters
        self.button_type = kwargs.pop(
            "button_type", "primary"
        )  # primary, secondary, success, warning, danger
        self.corner_radius = kwargs.pop("corner_radius", 8)

        # Set default styling
        kwargs.update(
            {"relief": "flat", "border": 0, "cursor": "hand2", "font": ("Segoe UI", 10)}
        )

        super().__init__(parent, *args, **kwargs)

        # Bind hover effects
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress>", self._on_press)
        self.bind("<ButtonRelease>", self._on_release)

    def _on_enter(self, event):
        """Mouse enter effect"""
        if self["state"] != "disabled":
            self.configure(bg=self.hover_color)

    def _on_leave(self, event):
        """Mouse leave effect"""
        if self["state"] != "disabled":
            self.configure(bg=self.normal_color)

    def _on_press(self, event):
        """Mouse press effect"""
        if self["state"] != "disabled":
            self.configure(bg=self.press_color)

    def _on_release(self, event):
        """Mouse release effect"""
        if self["state"] != "disabled":
            self.configure(bg=self.normal_color)

    def apply_theme(self, theme):
        """Apply theme colors to the button"""
        colors = ModernStyles.get_theme(theme)

        if self.button_type == "primary":
            self.normal_color = colors["accent"]
            self.hover_color = colors["accent_hover"]
            self.press_color = colors["accent_hover"]
            fg_color = "#ffffff"
        elif self.button_type == "secondary":
            self.normal_color = colors["secondary"]
            self.hover_color = colors["border"]
            self.press_color = colors["border"]
            fg_color = "#ffffff"
        elif self.button_type == "success":
            self.normal_color = colors["success"]
            self.hover_color = "#218838"
            self.press_color = "#218838"
            fg_color = "#ffffff"
        elif self.button_type == "warning":
            self.normal_color = colors["warning"]
            self.hover_color = "#e0a800"
            self.press_color = "#e0a800"
            fg_color = "#212529"
        elif self.button_type == "danger":
            self.normal_color = colors["danger"]
            self.hover_color = "#c82333"
            self.press_color = "#c82333"
            fg_color = "#ffffff"
        else:
            self.normal_color = colors["card_bg"]
            self.hover_color = colors["highlight"]
            self.press_color = colors["highlight"]
            fg_color = colors["text"]

        self.configure(
            bg=self.normal_color,
            fg=fg_color,
            activebackground=self.press_color,
            activeforeground=fg_color,
        )


# ------------- Main App -------------
class AutoClickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"高级自动连点器 {VERSION}")
        self.root.geometry("1200x800")
        self.root.minsize(1100, 700)

        # 设置窗口图标
        self.setup_icons()
        # Setup modern styles
        ModernStyles.setup_styles()

        # Apply initial styling to root
        self.root.configure(bg=ModernStyles.LIGHT_THEME["bg"])

        self.show_confirmation_var = tk.BooleanVar(value=True)
        # 窗口管理相关
        self.window_handle = None
        self.auto_switch_window = tk.BooleanVar(value=False)
        self.window_title_var = tk.StringVar(value="")
        self.window_handles = {}  # 存储窗口句柄映射
        # 添加 HAVE_WIN32 变量
        try:
            import win32gui
            import win32con

            self.HAVE_WIN32 = True
        except ImportError:
            self.HAVE_WIN32 = False
        # 添加调试模式设置
        self.debug_mode_var = tk.BooleanVar(value=False)  # 默认关闭调试模式
        self.enable_safety_check = tk.BooleanVar(value=True)  # 默认开启但调整参数
        self.last_safety_trigger = 0  # 记录上次触发时间
        # 添加快捷键防重复机制
        self.last_capture_time = 0
        self.capture_cooldown = 0.3  # 0.3秒冷却时间

        # 安全检测相关
        self.last_capture_time = 0
        self.capture_cooldown = 0.3
        self.last_click_time = 0
        self.current_target_point = None  # 当前目标点位
        self.last_safety_trigger = 0

        # Program state
        # click_points: list of dicts: {'x':int,'y':int,'name':str,'button':'left'/'right','delay':float, 'click_count':int, 'click_interval':float}
        self.click_points = []
        self.is_capturing = False
        self.is_running = False
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
        self.task_thread = None

        # Settings variables
        self.click_count_var = tk.IntVar(value=1)
        self.base_delay_var = tk.DoubleVar(value=1.0)
        self.loop_var = tk.IntVar(value=1)
        self.random_offset_var = tk.IntVar(value=0)
        self.random_delay_var = tk.DoubleVar(value=0.0)
        self.countdown_var = tk.IntVar(value=3)
        self.auto_action_var = tk.StringVar(value="none")  # none / sound
        self.theme_var = tk.StringVar(value="light")  # light/dark

        # Safety config
        self.mouse_move_threshold_px = 250  # 这个现在可以作为默认值，但实际使用动态阈值
        self.mouse_move_check_interval = 0.25  # seconds

        # 确保这个变量在UI创建之前初始化
        self.safety_threshold_var = tk.IntVar(value=250)

        # Statistics
        self.stats = {
            "total_click_attempts": 0,
            "successful_clicks": 0,
            "failed_clicks": 0,
            "start_time": None,
            "end_time": None,
            "loops_completed": 0,
        }

        # Tray
        self.tray_icon = None
        self.tray_thread = None
        self.tray_running = False

        # Build UI
        self.create_menu()
        self.create_ui()

        # Bind shortcuts
        self.bind_shortcuts()

        # Load saved config if exists
        self.load_points(quiet=True)

        # Start coordinate preview update
        self.last_mouse_pos = pyautogui.position()
        self.update_coord_preview()

        # Auto update check in background (non-blocking)
        if UPDATE_CHECK_URLS:
            # 延迟2秒后检查，避免影响程序启动速度
            self.root.after(
                2000,
                lambda: threading.Thread(target=self.check_update, daemon=True).start(),
            )

        log("程序启动")
        self.add_progress_text("程序已启动。")

    # ---------------- UI: menu ----------------
    def create_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入点位配置...", command=self.import_points)
        file_menu.add_command(label="导出点位配置...", command=self.export_points)
        file_menu.add_separator()
        file_menu.add_command(label="保存当前配置", command=self.save_points)
        file_menu.add_command(label="另存为...", command=self.save_points_as)
        file_menu.add_separator()
        file_menu.add_command(label="导出执行报告", command=self.export_click_report)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_exit)
        menubar.add_cascade(label="文件", menu=file_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(
            label="主题: 浅色", command=lambda: self.set_theme("light")
        )
        settings_menu.add_command(
            label="主题: 深色", command=lambda: self.set_theme("dark")
        )
        menubar.add_cascade(label="设置", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(
            label="GitHub更新帮助", command=self.show_github_update_help
        )  # 新增这一行
        help_menu.add_command(label="关于", command=self.show_about)
        help_menu.add_separator()
        help_menu.add_command(label="打开日志文件", command=self.open_log_file)
        help_menu.add_command(label="检查更新", command=self.manual_check_update)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    # ---------------- UI: main layout ----------------
    def create_ui(self):
        # Main container with modern styling
        main_container = GlassFrame(self.root, blur_color="#ffffff")
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        main_container.apply_glass_effect(self.theme_var.get())

        # Header
        header_frame = GlassFrame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        header_frame.apply_glass_effect(self.theme_var.get())

        title_label = tk.Label(
            header_frame,
            text=f"高级自动连点器 {VERSION}",
            font=("Segoe UI", 18, "bold"),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["accent"],
        )
        title_label.pack(side=tk.LEFT, pady=10)

        # 使用Notebook实现标签页
        self.notebook = ttk.Notebook(main_container, style="Modern.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 主任务标签页
        self.main_task_frame = GlassFrame(self.notebook)
        self.notebook.add(self.main_task_frame, text="主任务")
        self.main_task_frame.apply_glass_effect(self.theme_var.get())

        # 构建主任务标签页的内容
        self.create_main_task_ui()

        # Status bar
        self.status_bar = tk.Label(
            main_container,
            text="就绪",
            anchor=tk.W,
            font=("Segoe UI", 9),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text_light"],
        )
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))

        # Apply initial theme
        self.apply_theme()

        # 在创建完UI后绑定组合框事件
        if hasattr(self, "window_combo"):
            self.window_combo.bind("<<ComboboxSelected>>", self.on_window_combo_select)

    def create_main_task_ui(self):
        """创建主任务界面 - 完整的滚动条解决方案"""
        # 创建主框架
        main_frame = tk.Frame(self.main_task_frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建滚动条
        scrollbar = tk.Scrollbar(main_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建画布
        canvas = tk.Canvas(
            main_frame,
            yscrollcommand=scrollbar.set,
            bg=ModernStyles.get_theme(self.theme_var.get())["bg"],
        )
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=canvas.yview)

        # 创建可滚动的内部框架
        scrollable_frame = GlassFrame(canvas)
        scrollable_frame_id = canvas.create_window(
            (0, 0), window=scrollable_frame, anchor="nw"
        )

        def configure_scroll_region(event):
            # 更新滚动区域
            canvas.configure(scrollregion=canvas.bbox("all"))

        def configure_canvas_width(event):
            # 当画布大小改变时，调整内部框架的宽度
            canvas.itemconfig(scrollable_frame_id, width=event.width)

        scrollable_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_canvas_width)

        # 应用主题
        scrollable_frame.apply_glass_effect(self.theme_var.get())

        # 创建内容（使用水平分栏）
        self.create_scrollable_content(scrollable_frame)

    def create_scrollable_content(self, parent):
        """在可滚动框架中创建内容"""
        # 使用水平分栏
        left_frame = GlassFrame(parent)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        left_frame.apply_glass_effect(self.theme_var.get())

        right_frame = GlassFrame(parent)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        right_frame.apply_glass_effect(self.theme_var.get())

        # 创建内容
        self.create_left_panel(left_frame)
        self.create_right_panel(right_frame)

    def create_left_panel(self, parent):
        # 坐标获取区域
        coord_frame = GlassFrame(parent)
        coord_frame.pack(fill=tk.X, padx=4, pady=4)
        coord_frame.apply_glass_effect(self.theme_var.get())

        # 标题
        coord_title = tk.Label(
            coord_frame,
            text="坐标获取 & 点位管理",
            font=("Segoe UI", 12, "bold"),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        coord_title.pack(anchor="w", pady=(8, 12))

        instr = tk.Label(
            coord_frame,
            text='点击"开始捕获"后移动到目标位置，按 F2 记录坐标（支持全局）。双击列表项可编辑。',
            wraplength=600,
            font=("Segoe UI", 9),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text_light"],
        )
        instr.pack(fill=tk.X, pady=(0, 12))

        # 鼠标键设置和捕获按钮放在同一行
        top_row = tk.Frame(
            coord_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        top_row.pack(fill=tk.X, pady=4)

        # 鼠标键设置
        mb_frame = tk.Frame(
            top_row, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        mb_frame.pack(side=tk.LEFT)
        tk.Label(
            mb_frame,
            text="默认鼠标键:",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).pack(side=tk.LEFT)
        self.default_button_var = tk.StringVar(value="left")
        mb_combo = ttk.Combobox(
            mb_frame,
            textvariable=self.default_button_var,
            values=["left", "right", "middle"],
            state="readonly",
            width=8,
            style="Modern.TCombobox",
        )
        mb_combo.pack(side=tk.LEFT, padx=6)
        mb_combo.bind(
            "<<ComboboxSelected>>",
            lambda e: self.add_progress_text(
                f"默认鼠标键设置为 {self.default_button_var.get()}"
            ),
        )

        # 捕获按钮
        btn_frame = tk.Frame(
            top_row, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        btn_frame.pack(side=tk.RIGHT)
        self.coord_btn = ModernButton(
            btn_frame,
            text="开始获取坐标 (F2)",
            command=self.start_coord_capture,
            button_type="primary",
        )
        self.coord_btn.pack(side=tk.LEFT)
        self.coord_btn.apply_theme(self.theme_var.get())

        self.stop_capture_btn = ModernButton(
            btn_frame,
            text="停止获取坐标",
            state=tk.DISABLED,
            command=self.stop_coord_capture,
            button_type="secondary",
        )
        self.stop_capture_btn.pack(side=tk.LEFT, padx=6)
        self.stop_capture_btn.apply_theme(self.theme_var.get())

        # 坐标预览
        coord_preview_frame = tk.Frame(
            coord_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        coord_preview_frame.pack(fill=tk.X, pady=8)
        self.current_coord_label = tk.Label(
            coord_preview_frame,
            text="当前坐标: (0,0)",
            font=("Segoe UI", 10),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        self.current_coord_label.pack(side=tk.LEFT)
        self.coord_status = tk.Label(
            coord_preview_frame,
            text="捕获未启动",
            fg=ModernStyles.get_theme(self.theme_var.get())["danger"],
            font=("Segoe UI", 10),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
        )
        self.coord_status.pack(side=tk.RIGHT)

        # 点位列表区域
        points_frame = GlassFrame(parent)
        points_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        points_frame.apply_glass_effect(self.theme_var.get())

        # 标题
        points_title = tk.Label(
            points_frame,
            text="点位列表（双击编辑）",
            font=("Segoe UI", 12, "bold"),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        points_title.pack(anchor="w", pady=(8, 12))

        # 主内容区域：列表在左，按钮在右
        content_frame = tk.Frame(
            points_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 列表区域（占据主要空间）
        list_frame = tk.Frame(
            content_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.points_listbox = tk.Listbox(
            list_frame,
            font=("Consolas", 9),
            activestyle="none",
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            selectbackground=ModernStyles.get_theme(self.theme_var.get())["accent"],
            relief="flat",
            borderwidth=1,
        )
        self.points_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.points_listbox.bind("<Double-1>", self.edit_selected_point)

        list_scrollbar = tk.Scrollbar(list_frame, command=self.points_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.points_listbox.config(yscrollcommand=list_scrollbar.set)

        # 按钮区域（右侧竖向排列）
        button_frame = tk.Frame(
            content_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))

        # 按钮分组：操作按钮和移动按钮
        btn_width = 12

        # 主要操作按钮
        op_buttons = [
            ("添加当前坐标", self.add_current_point, "primary"),
            ("删除选中", self.delete_selected_point, "danger"),
            ("清空所有", self.clear_all_points, "warning"),
        ]

        for text, command, btn_type in op_buttons:
            btn = ModernButton(
                button_frame,
                text=text,
                command=command,
                button_type=btn_type,
                width=btn_width,
            )
            btn.pack(pady=4, fill=tk.X)
            btn.apply_theme(self.theme_var.get())

        # 分隔线
        separator = ttk.Separator(button_frame, orient="horizontal")
        separator.pack(fill=tk.X, pady=8)

        # 移动和重命名按钮
        move_buttons = [
            ("上移", lambda: self.move_selected(-1), "success"),
            ("下移", lambda: self.move_selected(1), "success"),
            ("重命名", self.rename_selected_point, "secondary"),
        ]

        for text, command, btn_type in move_buttons:
            btn = ModernButton(
                button_frame,
                text=text,
                command=command,
                button_type=btn_type,
                width=btn_width,
            )
            btn.pack(pady=4, fill=tk.X)
            btn.apply_theme(self.theme_var.get())

        # 窗口管理区域
        self.create_window_management_ui(parent)

    def create_right_panel(self, parent):
        # 设置区域
        settings = GlassFrame(parent)
        settings.pack(fill=tk.X, padx=4, pady=4)
        settings.apply_glass_effect(self.theme_var.get())

        # 标题
        settings_title = tk.Label(
            settings,
            text="任务设置",
            font=("Segoe UI", 12, "bold"),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        settings_title.pack(anchor="w", pady=(8, 12))

        # 使用网格布局更整齐
        settings_grid = tk.Frame(
            settings, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        settings_grid.pack(fill=tk.X, padx=4, pady=4)

        # 第一行：基础设置
        tk.Label(
            settings_grid,
            text="默认点击次数:",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=0, column=0, sticky="w", padx=2, pady=6)
        tk.Spinbox(
            settings_grid,
            from_=1,
            to=999,
            textvariable=self.click_count_var,
            width=8,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).grid(row=0, column=1, padx=2, pady=6)

        tk.Label(
            settings_grid,
            text="默认延时(秒):",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=0, column=2, sticky="w", padx=(10, 2), pady=6)
        tk.Spinbox(
            settings_grid,
            from_=0.0,
            to=60.0,
            increment=0.1,
            textvariable=self.base_delay_var,
            width=8,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).grid(row=0, column=3, padx=2, pady=6)

        tk.Label(
            settings_grid,
            text="循环次数:",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=0, column=4, sticky="w", padx=(10, 2), pady=6)
        tk.Spinbox(
            settings_grid,
            from_=1,
            to=99999,
            textvariable=self.loop_var,
            width=8,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).grid(row=0, column=5, padx=2, pady=6)

        # 第二行：随机设置
        tk.Label(
            settings_grid,
            text="随机偏移(px ±):",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=1, column=0, sticky="w", padx=2, pady=6)
        tk.Spinbox(
            settings_grid,
            from_=0,
            to=500,
            textvariable=self.random_offset_var,
            width=8,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).grid(row=1, column=1, padx=2, pady=6)

        tk.Label(
            settings_grid,
            text="随机延时(s ±):",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=1, column=2, sticky="w", padx=(10, 2), pady=6)
        tk.Spinbox(
            settings_grid,
            from_=0.0,
            to=10.0,
            increment=0.05,
            textvariable=self.random_delay_var,
            width=8,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).grid(row=1, column=3, padx=2, pady=6)

        tk.Label(
            settings_grid,
            text="启动倒计时(秒):",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=1, column=4, sticky="w", padx=(10, 2), pady=6)
        tk.Spinbox(
            settings_grid,
            from_=0,
            to=60,
            textvariable=self.countdown_var,
            width=8,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).grid(row=1, column=5, padx=2, pady=6)

        # 第三行：动作和选项
        tk.Label(
            settings_grid,
            text="任务结束动作:",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).grid(row=2, column=0, sticky="w", padx=2, pady=6)
        ttk.Combobox(
            settings_grid,
            textvariable=self.auto_action_var,
            values=["none", "sound"],
            state="readonly",
            width=8,
            style="Modern.TCombobox",
        ).grid(row=2, column=1, padx=2, pady=6)

        # 选项区域
        options_frame = tk.Frame(
            settings_grid, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        options_frame.grid(row=2, column=2, columnspan=4, sticky="ew", padx=2, pady=6)

        tk.Checkbutton(
            options_frame,
            text="启动前确认",
            variable=self.show_confirmation_var,
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            selectcolor=ModernStyles.get_theme(self.theme_var.get())["accent"],
        ).pack(side=tk.LEFT, padx=8)
        tk.Checkbutton(
            options_frame,
            text="调试模式",
            variable=self.debug_mode_var,
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            selectcolor=ModernStyles.get_theme(self.theme_var.get())["accent"],
        ).pack(side=tk.LEFT, padx=8)
        tk.Checkbutton(
            options_frame,
            text="安全检测",
            variable=self.enable_safety_check,
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            selectcolor=ModernStyles.get_theme(self.theme_var.get())["accent"],
        ).pack(side=tk.LEFT, padx=8)

        # 安全阈值设置
        safety_frame = tk.Frame(
            options_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        safety_frame.pack(side=tk.LEFT, padx=8)
        tk.Label(
            safety_frame,
            text="阈值:",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).pack(side=tk.LEFT)
        tk.Spinbox(
            safety_frame,
            from_=50,
            to=1000,
            textvariable=self.safety_threshold_var,
            width=5,
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
        ).pack(side=tk.LEFT, padx=2)

        # 控制按钮
        ctrl_frame = tk.Frame(
            settings, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        ctrl_frame.pack(fill=tk.X, pady=(12, 0))

        self.start_btn = ModernButton(
            ctrl_frame,
            text="启动任务 (F3)",
            command=self.start_task,
            button_type="success",
            height=2,
            font=("Segoe UI", 11, "bold"),
        )
        self.start_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.start_btn.apply_theme(self.theme_var.get())

        self.pause_btn = ModernButton(
            ctrl_frame,
            text="暂停 (F4)",
            command=self.toggle_pause,
            button_type="warning",
            state=tk.DISABLED,
            width=10,
        )
        self.pause_btn.pack(side=tk.LEFT, padx=2)
        self.pause_btn.apply_theme(self.theme_var.get())

        self.stop_btn = ModernButton(
            ctrl_frame,
            text="停止",
            command=self.stop_task,
            button_type="danger",
            state=tk.DISABLED,
            width=8,
        )
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn.apply_theme(self.theme_var.get())

        self.tray_btn = ModernButton(
            ctrl_frame,
            text="托盘",
            command=self.minimize_to_tray,
            button_type="secondary",
            width=6,
        )
        self.tray_btn.pack(side=tk.LEFT, padx=2)
        self.tray_btn.apply_theme(self.theme_var.get())

        # 进度和日志区域
        progress_frame = GlassFrame(parent)
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        progress_frame.apply_glass_effect(self.theme_var.get())

        # 标题
        progress_title = tk.Label(
            progress_frame,
            text="执行进度与日志",
            font=("Segoe UI", 12, "bold"),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        progress_title.pack(anchor="w", pady=(8, 12))

        # 当前任务信息
        task_frame = tk.Frame(
            progress_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        task_frame.pack(fill=tk.X, pady=(0, 8))
        self.current_task_label = tk.Label(
            task_frame,
            text="当前任务: 无",
            anchor=tk.W,
            justify=tk.LEFT,
            font=("Segoe UI", 9),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        self.current_task_label.pack(fill=tk.X)

        # 进度条
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            style="Modern.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        # 统计信息
        stats_frame = tk.Frame(
            progress_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        stats_frame.pack(fill=tk.X, pady=(0, 8))
        self.stats_label = tk.Label(
            stats_frame,
            text=self._stats_text(),
            anchor=tk.W,
            justify=tk.LEFT,
            fg=ModernStyles.get_theme(self.theme_var.get())["text_light"],
            font=("Segoe UI", 9),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
        )
        self.stats_label.pack(fill=tk.X)

        # 日志文本框
        log_frame = tk.Frame(
            progress_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_details = tk.Text(
            log_frame,
            height=12,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg=ModernStyles.get_theme(self.theme_var.get())["input_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            relief="flat",
            borderwidth=1,
            padx=8,
            pady=8,
        )
        self.progress_details.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        log_scrollbar = tk.Scrollbar(log_frame, command=self.progress_details.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.progress_details.config(yscrollcommand=log_scrollbar.set)

    # ---------------- Theme ----------------
    def set_theme(self, name):
        self.theme_var.set(name)
        self.apply_theme()
        self.add_progress_text(f"主题切换为：{name}")

    def apply_theme(self):
        theme = self.theme_var.get()
        colors = ModernStyles.get_theme(theme)

        # 应用主题到根窗口
        self.root.configure(bg=colors["bg"])

        # 递归应用主题到所有子组件
        def apply_to_children(widget):
            try:
                if isinstance(widget, (GlassFrame, tk.Frame, tk.LabelFrame)):
                    if hasattr(widget, "apply_glass_effect"):
                        widget.apply_glass_effect(theme)
                    else:
                        widget.configure(bg=colors["card_bg"])
                elif isinstance(widget, tk.Label):
                    widget.configure(bg=colors["card_bg"], fg=colors["text"])
                elif isinstance(widget, ModernButton):
                    widget.apply_theme(theme)
                elif isinstance(widget, tk.Button):
                    # 保留特殊按钮的颜色，只修改普通按钮
                    current_bg = widget.cget("bg")
                    if current_bg in ["SystemButtonFace", "#f0f0f0", "#d9d9d9"]:
                        widget.configure(bg=colors["highlight"], fg=colors["text"])
                elif isinstance(widget, tk.Entry):
                    widget.configure(
                        bg=colors["input_bg"],
                        fg=colors["text"],
                        insertbackground=colors["text"],
                        relief="flat",
                    )
                elif isinstance(widget, tk.Listbox):
                    widget.configure(
                        bg=colors["input_bg"],
                        fg=colors["text"],
                        selectbackground=colors["accent"],
                        relief="flat",
                    )
                elif isinstance(widget, tk.Text):
                    widget.configure(
                        bg=colors["input_bg"],
                        fg=colors["text"],
                        insertbackground=colors["text"],
                        relief="flat",
                    )
                elif isinstance(widget, tk.Scrollbar):
                    widget.configure(bg=colors["border"], troughcolor=colors["bg"])
                elif isinstance(widget, tk.Spinbox):
                    widget.configure(
                        bg=colors["input_bg"],
                        fg=colors["text"],
                        relief="flat",
                        borderwidth=1,
                    )
                elif isinstance(widget, ttk.Combobox):
                    style = ttk.Style()
                    style.configure(
                        "Modern.TCombobox",
                        fieldbackground=colors["input_bg"],
                        background=colors["input_bg"],
                        foreground=colors["text"],
                        borderwidth=1,
                        relief="flat",
                    )
                elif isinstance(widget, ttk.Progressbar):
                    style = ttk.Style()
                    style.configure(
                        "Modern.Horizontal.TProgressbar",
                        troughcolor=colors["border"],
                        background=colors["accent"],
                        borderwidth=0,
                    )
                elif isinstance(widget, ttk.Notebook):
                    style = ttk.Style()
                    style.configure(
                        "Modern.TNotebook", background=colors["bg"], borderwidth=0
                    )
                    style.configure(
                        "Modern.TNotebook.Tab",
                        background=colors["highlight"],
                        foreground=colors["text"],
                        borderwidth=0,
                    )
                    style.map(
                        "Modern.TNotebook.Tab",
                        background=[
                            ("selected", colors["accent"]),
                            ("active", colors["accent_hover"]),
                        ],
                        foreground=[("selected", "#ffffff")],
                    )
            except Exception:
                pass

            for child in widget.winfo_children():
                apply_to_children(child)

        for child in self.root.winfo_children():
            apply_to_children(child)

        # 状态栏
        try:
            self.status_bar.configure(bg=colors["card_bg"], fg=colors["text_light"])
        except Exception:
            pass

    # ---------------- Coordinate preview & capture ----------------
    def update_coord_preview(self):
        if not self.root.winfo_exists():
            return
        try:
            x, y = pyautogui.position()
            self.current_coord_label.config(text=f"当前坐标: ({x}, {y})")
        except Exception:
            self.current_coord_label.config(text="当前坐标: (N/A)")
        # schedule mouse move safety check too
        self.check_mouse_movement()
        self.root.after(120, self.update_coord_preview)

    def start_coord_capture(self):
        self.is_capturing = True
        self.coord_btn.config(state=tk.DISABLED)
        self.stop_capture_btn.config(state=tk.NORMAL)
        self.coord_status.config(
            text="捕获模式已启动 - 按 F2 或窗口内'添加当前坐标'记录位置",
            fg=ModernStyles.get_theme(self.theme_var.get())["success"],
        )
        self.add_progress_text("坐标捕获模式已启动")
        log("坐标捕获模式启动")

    def stop_coord_capture(self):
        self.is_capturing = False
        self.coord_btn.config(state=tk.NORMAL)
        self.stop_capture_btn.config(state=tk.DISABLED)
        self.coord_status.config(
            text="捕获模式已停止",
            fg=ModernStyles.get_theme(self.theme_var.get())["danger"],
        )
        self.add_progress_text("坐标捕获模式已停止")
        log("坐标捕获模式停止")

    # ---------------- Points management ----------------
    def add_point(
        self,
        x,
        y,
        name=None,
        button=None,
        delay=None,
        click_count=None,
        click_interval=None,
        keys="",
        action_type="click",
    ):
        if name is None:
            name = f"点位{len(self.click_points)+1}({x},{y})"
        if button is None:
            button = self.default_button_var.get() or "left"
        if delay is None:
            delay = float(self.base_delay_var.get())
        if click_count is None:
            click_count = int(self.click_count_var.get())
        if click_interval is None:
            click_interval = float(self.base_delay_var.get())

        p = {
            "x": int(x),
            "y": int(y),
            "name": str(name),
            "button": button,
            "delay": float(delay),
            "click_count": int(click_count),
            "click_interval": float(click_interval),
            "keys": keys,
            "action_type": action_type,
        }
        self.click_points.append(p)
        self.update_points_list()

        action_desc = (
            f"按键:{keys}"
            if action_type == "keyboard"
            else f"键:{button} 点击:{click_count}次"
        )
        self.add_progress_text(
            f"添加点位: {name} ({x},{y}) {action_desc} 延时:{delay} 间隔:{click_interval}"
        )
        log(f"添加点位 {name} ({x},{y})")

    def add_current_point(self):
        try:
            x, y = pyautogui.position()
            # 弹出对话框选择按钮、延时、点击次数和点击间隔
            dlg = AddPointDialog(
                self.root,
                x,
                y,
                default_button=self.default_button_var.get(),
                default_delay=self.base_delay_var.get(),
                default_click_count=self.click_count_var.get(),
                default_click_interval=self.base_delay_var.get(),
                theme=self.theme_var.get(),
            )
            self.root.wait_window(dlg.top)
            if dlg.result:
                # 现在接收9个返回值
                (
                    nx,
                    ny,
                    name,
                    btn,
                    delay,
                    click_count,
                    click_interval,
                    keys,
                    action_type,
                ) = dlg.result
                self.add_point(
                    nx,
                    ny,
                    name=name,
                    button=btn,
                    delay=delay,
                    click_count=click_count,
                    click_interval=click_interval,
                    keys=keys,
                    action_type=action_type,
                )
        except Exception as e:
            messagebox.showerror("错误", f"无法获取当前鼠标坐标: {e}")

    def update_points_list(self):
        self.points_listbox.delete(0, tk.END)
        for i, p in enumerate(self.click_points):
            # 确保所有必要的字段都存在
            name = p.get("name", f"点位{i+1}")
            x = p.get("x", 0)
            y = p.get("y", 0)
            button = p.get("button", "left")
            delay = p.get("delay", 1.0)
            click_count = p.get("click_count", 1)
            click_interval = p.get("click_interval", 0.1)
            action_type = p.get("action_type", "click")
            keys = p.get("keys", "")

            if action_type == "keyboard":
                display = f"{i+1}. {name} - 键盘操作 [{keys}] 延时:{delay}"
            else:
                display = f"{i+1}. {name} - ({x},{y}) [{button}] 延时:{delay} 点击:{click_count}次 间隔:{click_interval}"
            self.points_listbox.insert(tk.END, display)

    def delete_selected_point(self):
        sel = self.points_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "请先选择点位")
            return
        idx = sel[0]

        # 检查索引是否有效
        if idx < 0 or idx >= len(self.click_points):
            messagebox.showwarning("警告", "选择的点位无效")
            return

        removed = self.click_points.pop(idx)
        self.update_points_list()  # 确保立即更新列表显示
        self.add_progress_text(f"删除点位: {removed['name']}")
        log(f"删除点位 {removed['name']}")

    def clear_all_points(self):
        if not self.click_points:
            messagebox.showinfo("提示", "点位列表已经为空")
            return
        if messagebox.askyesno("确认", "确定要清空所有点位吗？"):
            cnt = len(self.click_points)
            self.click_points.clear()
            self.update_points_list()
            # 清除列表选择
            self.points_listbox.selection_clear(0, tk.END)
            self.add_progress_text(f"已清空 {cnt} 个点位")
            log("清空所有点位")

    def move_selected(self, delta):
        sel = self.points_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "请先选择点位")
            return
        idx = sel[0]

        # 检查索引是否有效
        if idx < 0 or idx >= len(self.click_points):
            messagebox.showwarning("警告", "选择的点位无效")
            return

        new = idx + delta
        if new < 0 or new >= len(self.click_points):
            return

        self.click_points[idx], self.click_points[new] = (
            self.click_points[new],
            self.click_points[idx],
        )
        self.update_points_list()
        self.points_listbox.selection_clear(0, tk.END)
        self.points_listbox.selection_set(new)
        self.add_progress_text(f"点位移动: {idx+1} -> {new+1}")
        log(f"点位移动 {idx+1} -> {new+1}")

    def rename_selected_point(self):
        sel = self.points_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "请先选择点位")
            return
        idx = sel[0]

        # 检查索引是否有效
        if idx < 0 or idx >= len(self.click_points):
            messagebox.showwarning("警告", "选择的点位无效")
            return

        p = self.click_points[idx]
        new_name = simpledialog.askstring(
            "重命名点位",
            "输入新的点位名称：",
            initialvalue=p.get("name", f"点位{idx+1}"),
            parent=self.root,
        )
        if new_name:
            old = p.get("name", f"点位{idx+1}")
            self.click_points[idx]["name"] = new_name
            self.update_points_list()
            self.add_progress_text(f"点位重命名: {old} -> {new_name}")
            log(f"点位重命名 {old} -> {new_name}")

    def edit_selected_point(self, event=None):
        sel = self.points_listbox.curselection()
        if not sel:
            return
        idx = sel[0]

        # 检查索引是否有效
        if idx < 0 or idx >= len(self.click_points):
            messagebox.showwarning("警告", "选择的点位无效")
            return

        p = self.click_points[idx]

        # 确保所有字段都有默认值
        x = p.get("x", 0)
        y = p.get("y", 0)
        name = p.get("name", f"点位{idx+1}")
        button = p.get("button", "left")
        delay = p.get("delay", 1.0)
        click_count = p.get("click_count", 1)
        click_interval = p.get("click_interval", 0.1)
        keys = p.get("keys", "")
        action_type = p.get("action_type", "click")

        # 传递所有必要的参数
        dlg = EditPointDialog(
            self.root,
            x,
            y,
            name,
            button,
            delay,
            click_count,
            click_interval,
            keys,
            action_type,
            theme=self.theme_var.get(),
        )
        self.root.wait_window(dlg.top)
        if dlg.result:
            (
                nx,
                ny,
                nname,
                nbutton,
                ndelay,
                nclick_count,
                nclick_interval,
                nkeys,
                naction_type,
            ) = dlg.result
            self.click_points[idx] = {
                "x": int(nx),
                "y": int(ny),
                "name": nname,
                "button": nbutton,
                "delay": float(ndelay),
                "click_count": int(nclick_count),
                "click_interval": float(nclick_interval),
                "keys": nkeys,
                "action_type": naction_type,
            }
            self.update_points_list()

            action_desc = (
                f"按键:{nkeys}"
                if naction_type == "keyboard"
                else f"键:{nbutton} 点击:{nclick_count}次"
            )
            self.add_progress_text(
                f"点位已编辑: {nname} ({nx},{ny}) {action_desc} 延时:{ndelay} 间隔:{nclick_interval}"
            )
            log(f"点位编辑 -> {nname} ({nx},{ny})")

    # ---------------- Import/Export/Save/Load ----------------
    def import_points(self):
        path = filedialog.askopenfilename(
            title="导入点位配置",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                # support both old list-of-lists and new list-of-dicts
                new_list = []
                for item in data:
                    if isinstance(item, list) and len(item) >= 3:
                        x = int(item[0])
                        y = int(item[1])
                        name = str(item[2])
                        new_list.append(
                            {
                                "x": x,
                                "y": y,
                                "name": name,
                                "button": self.default_button_var.get(),
                                "delay": self.base_delay_var.get(),
                                "click_count": self.click_count_var.get(),
                                "click_interval": self.base_delay_var.get(),
                            }
                        )
                    elif isinstance(item, dict):
                        # ensure keys exist
                        x = int(item.get("x", 0))
                        y = int(item.get("y", 0))
                        name = str(item.get("name", f"点位({x},{y})"))
                        button = item.get("button", self.default_button_var.get())
                        delay = float(item.get("delay", self.base_delay_var.get()))
                        click_count = int(
                            item.get("click_count", self.click_count_var.get())
                        )
                        click_interval = float(
                            item.get("click_interval", self.base_delay_var.get())
                        )
                        new_list.append(
                            {
                                "x": x,
                                "y": y,
                                "name": name,
                                "button": button,
                                "delay": delay,
                                "click_count": click_count,
                                "click_interval": click_interval,
                            }
                        )
                self.click_points = new_list
                self.update_points_list()
                self.add_progress_text(
                    f"已导入 {len(self.click_points)} 个点位 ({os.path.basename(path)})"
                )
                log(f"导入配置 {path}")
            else:
                messagebox.showerror("错误", "不支持的配置格式")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")

    def export_points(self):
        path = filedialog.asksaveasfilename(
            title="导出点位配置",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
            self.add_progress_text(f"已导出到 {os.path.basename(path)}")
            log(f"导出配置 {path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def save_points(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
            messagebox.showinfo(
                "成功", f"已保存 {len(self.click_points)} 个点位到 {CONFIG_FILE}"
            )
            self.add_progress_text("点位配置已保存")
            log(f"保存配置 {CONFIG_FILE}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def save_points_as(self):
        path = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
            self.add_progress_text(f"已另存为 {os.path.basename(path)}")
            log(f"另存为 {path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def load_points(self, quiet=False):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # normalize
                new_list = []
                for item in data:
                    if isinstance(item, dict):
                        x = int(item.get("x", 0))
                        y = int(item.get("y", 0))
                        name = item.get("name", f"点位({x},{y})")
                        button = item.get("button", self.default_button_var.get())
                        delay = float(item.get("delay", self.base_delay_var.get()))
                        click_count = int(
                            item.get("click_count", self.click_count_var.get())
                        )
                        click_interval = float(
                            item.get("click_interval", self.base_delay_var.get())
                        )
                        keys = item.get("keys", "")
                        action_type = item.get("action_type", "click")

                        new_list.append(
                            {
                                "x": x,
                                "y": y,
                                "name": name,
                                "button": button,
                                "delay": delay,
                                "click_count": click_count,
                                "click_interval": click_interval,
                                "keys": keys,
                                "action_type": action_type,
                            }
                        )
                    elif isinstance(item, list):
                        x = int(item[0])
                        y = int(item[1])
                        name = str(item[2])
                        new_list.append(
                            {
                                "x": x,
                                "y": y,
                                "name": name,
                                "button": self.default_button_var.get(),
                                "delay": self.base_delay_var.get(),
                                "click_count": self.click_count_var.get(),
                                "click_interval": self.base_delay_var.get(),
                                "keys": "",
                                "action_type": "click",
                            }
                        )
                self.click_points = new_list
                self.update_points_list()
                if not quiet:
                    self.add_progress_text(f"已加载 {len(self.click_points)} 个点位")
                log(f"加载配置 {CONFIG_FILE}")
            except Exception as e:
                log(f"加载点位配置失败: {e}")

    # ---------------- Task control ----------------
    def start_task(self):
        if self.is_running:
            messagebox.showinfo("提示", "任务已在运行中")
            return
        if not self.click_points:
            messagebox.showwarning("警告", "请先添加至少一个点位！")
            return

        # validate params
        base_delay = safe_float(self.base_delay_var.get(), 1.0)
        loop_count = safe_int(self.loop_var.get(), 1)
        offset_px = safe_int(self.random_offset_var.get(), 0)
        rand_delay = safe_float(self.random_delay_var.get(), 0.0)
        countdown = safe_int(self.countdown_var.get(), 0)

        # 修改这里：根据设置决定是否显示确认窗口
        if self.show_confirmation_var.get():
            # task summary
            summary = (
                f"点位数量: {len(self.click_points)}\n"
                f"基础延时: {base_delay} 秒\n"
                f"随机偏移: ±{offset_px} px\n"
                f"随机延时: ±{rand_delay} 秒\n"
                f"循环次数: {loop_count}\n"
                f"启动倒计时: {countdown} 秒\n"
                f"任务结束动作: {self.auto_action_var.get()}"
            )
            if not messagebox.askokcancel("确认任务", summary):
                return

        # countdown
        if countdown > 0:
            for i in range(countdown, 0, -1):
                self.status_bar.config(text=f"倒计时：{i} 秒 — 准备开始...")
                self.root.update()
                time.sleep(1)

        # prepare events & stats
        self.stop_event.clear()
        self.pause_event.clear()
        self.is_running = True
        self.start_btn.config(state=tk.DISABLED, text="任务运行中...")
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.add_progress_text("任务启动")
        log("任务启动")

        # reset stats
        self.stats = {
            "total_click_attempts": 0,
            "successful_clicks": 0,
            "failed_clicks": 0,
            "start_time": datetime.datetime.now().isoformat(),
            "end_time": None,
            "loops_completed": 0,
        }
        self.update_stats_display()

        # run task thread
        self.task_thread = threading.Thread(
            target=self.run_click_task,
            args=(loop_count, offset_px, rand_delay),
            daemon=True,
        )
        self.task_thread.start()

    def stop_task(self):
        if not self.is_running:
            return
        self.stop_event.set()
        self.pause_event.clear()
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL, text="启动任务 (F3)")
        self.pause_btn.config(state=tk.DISABLED, text="暂停 (F4)")
        self.stop_btn.config(state=tk.DISABLED)
        self.status_bar.config(text="已停止")
        self.add_progress_text("任务已停止")
        log("任务停止")

    def toggle_pause(self):
        if not self.is_running:
            return
        if not self.pause_event.is_set():
            self.pause_event.set()
            self.pause_btn.config(text="继续 (F4)")
            self.add_progress_text("任务已暂停")
            self.status_bar.config(text="已暂停 - 按 F4 继续")
            log("任务暂停")
        else:
            self.pause_event.clear()
            self.pause_btn.config(text="暂停 (F4)")
            self.add_progress_text("任务继续")
            self.status_bar.config(text="继续运行")
            log("任务继续")

    def run_click_task(self, loop_count, offset_px, rand_delay):

        # 确保目标窗口激活
        if not self.ensure_target_window_active():
            self.add_progress_text("错误: 无法切换到目标窗口，任务停止")
            log("任务停止: 无法切换到目标窗口")
            self.stop_event.set()
            return

        total_points = len(self.click_points)
        total_ops = sum(p["click_count"] for p in self.click_points) * loop_count
        ops_done = 0

        try:
            for loop in range(loop_count):
                if self.stop_event.is_set():
                    break
                if not self.is_running:
                    break
                self.add_progress_text(f"开始第 {loop+1}/{loop_count} 次循环")

                for idx, p in enumerate(self.click_points):
                    if self.stop_event.is_set():
                        break

                    # 等待暂停
                    while self.pause_event.is_set() and not self.stop_event.is_set():
                        self.status_bar.config(text="已暂停（按 F4 继续）")
                        time.sleep(0.1)

                    if self.stop_event.is_set():
                        break

                    # 记录当前目标点位
                    self.current_target_point = (p["x"], p["y"])

                    # 更新UI显示
                    self.root.after(0, lambda i=idx: self.highlight_point(i))
                    cur_info = f"当前: 循环 {loop+1}/{loop_count} — 点位 {idx+1}/{total_points} - {p['name']} ({p['x']},{p['y']})"
                    self.root.after(
                        0, lambda t=cur_info: self.current_task_label.config(text=t)
                    )

                    # 调试模式处理
                    if self.debug_mode_var.get():
                        self.add_progress_text(
                            f"调试模式：点位 {idx+1} 准备就绪，按 F4 继续..."
                        )
                        self.status_bar.config(
                            text=f"调试模式：点位 {idx+1} 就绪，按 F4 继续"
                        )
                        self.pause_event.set()

                        while (
                            self.pause_event.is_set() and not self.stop_event.is_set()
                        ):
                            time.sleep(0.1)

                        if self.stop_event.is_set():
                            break

                    # 关键修复：正确的缩进开始
                    click_count = p.get("click_count", 1)
                    action_type = p.get("action_type", "click")

                    for c in range(click_count):
                        if self.stop_event.is_set():
                            break

                        if action_type == "keyboard":
                            # 键盘操作
                            keys = p.get("keys", "")
                            if keys and HAVE_KEYBOARD:
                                try:
                                    keyboard.send(keys)
                                    self.stats["successful_clicks"] += 1
                                    self.add_progress_text(f"执行键盘操作: {keys}")
                                    log(f"键盘操作: {keys}")
                                except Exception as e:
                                    self.stats["failed_clicks"] += 1
                                    self.add_progress_text(f"键盘操作失败: {e}")
                                    log(f"键盘操作异常: {e}")
                        else:
                            # 鼠标点击操作 - 关键修复：确保使用正确的坐标
                            rx = p["x"]
                            ry = p["y"]

                            # 应用随机偏移
                            if offset_px > 0:
                                dx = random.randint(-offset_px, offset_px)
                                dy = random.randint(-offset_px, offset_px)
                                rx = p["x"] + dx
                                ry = p["y"] + dy

                            # 执行点击 - 确保坐标正确传递
                            try:
                                btn = p.get("button", "left")
                                # 调试信息：打印实际点击的坐标
                                debug_msg = f"准备点击: ({rx}, {ry}) - 原始坐标: ({p['x']}, {p['y']})"
                                self.add_progress_text(debug_msg)
                                log(debug_msg)

                                pyautogui.click(rx, ry, button=btn)
                                self.stats["successful_clicks"] += 1
                            except Exception as e:
                                self.stats["failed_clicks"] += 1
                                self.add_progress_text(f"点击失败: {e}，任务停止")
                                log(f"点击异常: {e}")
                                self.stop_event.set()
                                break
                            finally:
                                self.stats["total_click_attempts"] += 1

                            # 记录点击时间
                            self.last_click_time = time.time()

                            ops_done += 1
                            progress = (
                                (ops_done / total_ops) * 100 if total_ops > 0 else 100
                            )
                            self.root.after(
                                0, lambda p=progress: self.progress_var.set(p)
                            )

                            offset_info = (
                                f"偏移({rx-p['x']},{ry-p['y']})"
                                if offset_px > 0
                                else ""
                            )
                            self.add_progress_text(
                                f"循环{loop+1}/{loop_count} - 点位{idx+1}/{total_points} - 点击{c+1}/{click_count} {offset_info}"
                            )
                            log(f"点击: {p['name']} ({rx},{ry}) 按钮:{btn}")

                            # 点击间隔延时
                            click_interval = p.get("click_interval", 0.1)
                            if rand_delay > 0:
                                click_interval = click_interval + random.uniform(
                                    -rand_delay, rand_delay
                                )
                                if click_interval < 0:
                                    click_interval = 0
                            if c < click_count - 1 and not self.stop_event.is_set():
                                time.sleep(click_interval)

                    # 清除目标点位记录
                    self.current_target_point = None

                    # 点位间延时
                    if idx < total_points - 1 and not self.stop_event.is_set():
                        time.sleep(p.get("delay", 1.0))

                # 完成一次循环
                self.stats["loops_completed"] += 1

                # 循环间延时
                if loop < loop_count - 1 and not self.stop_event.is_set():
                    time.sleep(1.0)

            # finished normally or stopped
            if not self.stop_event.is_set():
                self.add_progress_text("任务完成！")
                log("任务完成")
                # post-action
                if self.auto_action_var.get() == "sound":
                    play_beep()
            else:
                self.add_progress_text("任务中断。")
                log("任务中断")
        except Exception as e:
            self.add_progress_text(f"任务异常终止: {e}")
            log(f"任务异常: {e}")
        finally:
            # 清理目标点位记录
            self.current_target_point = None
            self.stats["end_time"] = datetime.datetime.now().isoformat()
            # finalize: update ui in main thread
            self.root.after(0, self._task_finished_cleanup)
            # write click report
            try:
                self.write_click_report(silent=True)
            except Exception as ex:
                log(f"写入报告失败: {ex}")

    def highlight_point(self, index):
        try:
            self.points_listbox.selection_clear(0, tk.END)
            self.points_listbox.selection_set(index)
            self.points_listbox.see(index)
            # also flash the item: we'll change background temporarily if possible (tk Listbox doesn't support per-item bg easily)
            # so we use selection to highlight visually
        except Exception:
            pass

    def _task_finished_cleanup(self):
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL, text="启动任务 (F3)")
        self.pause_btn.config(state=tk.DISABLED, text="暂停 (F4)")
        self.stop_btn.config(state=tk.DISABLED)
        self.status_bar.config(text="就绪")
        self.progress_var.set(0.0)
        try:
            self.points_listbox.selection_clear(0, tk.END)
        except Exception:
            pass
        # update stats label
        self.update_stats_display()

    # ---------------- Shortcuts & global keys ----------------
    def bind_shortcuts(self):
        # local bindings
        self.root.bind("<F2>", lambda e: self.local_capture())
        self.root.bind("<F3>", lambda e: self._hotkey_toggle_start_stop()())
        self.root.bind("<F4>", lambda e: self.toggle_pause())

        # global keyboard hooks via keyboard lib if available
        if HAVE_KEYBOARD:
            try:
                # 添加全局F2快捷键，但使用防重复机制
                keyboard.add_hotkey("f2", self.global_capture_coord)
                keyboard.add_hotkey("f3", lambda: self._hotkey_toggle_start_stop()())
                keyboard.add_hotkey("f4", lambda: self.toggle_pause())
                log("全局快捷键已设置 (F2 记录, F3 开始/停止, F4 暂停/继续)")
            except Exception as e:
                log(f"设置全局快捷键失败: {e}")
                self.add_progress_text("全局快捷键设置失败或权限不足")
        else:
            self.add_progress_text("未检测到 keyboard 库或权限，全局快捷键不可用。")

    def _hotkey_toggle_start_stop(self):
        def _inner():
            if self.is_running:
                self.stop_task()
            else:
                self.start_task()

        return _inner

    def global_capture_coord(self):
        if self.is_capturing:
            # 防重复机制
            current_time = time.time()
            if current_time - self.last_capture_time < self.capture_cooldown:
                return
            self.last_capture_time = current_time

            try:
                x, y = pyautogui.position()
                # 使用默认设置添加点位
                self.add_point(x, y)
                self.coord_status.config(
                    text=f"已记录坐标: ({x}, {y})",
                    fg=ModernStyles.get_theme(self.theme_var.get())["success"],
                )
                # 在主线程中更新状态
                self.root.after(
                    0,
                    lambda: self.coord_status.config(
                        text=f"已记录坐标: ({x}, {y})",
                        fg=ModernStyles.get_theme(self.theme_var.get())["success"],
                    ),
                )
            except Exception as e:
                log(f"全局捕获失败: {e}")

    def local_capture(self, event=None):
        if self.is_capturing:
            # 防重复机制
            current_time = time.time()
            if current_time - self.last_capture_time < self.capture_cooldown:
                return
            self.last_capture_time = current_time

            try:
                x, y = pyautogui.position()
                self.add_point(x, y)
                self.coord_status.config(
                    text=f"已记录坐标: ({x}, {y})",
                    fg=ModernStyles.get_theme(self.theme_var.get())["success"],
                )
            except Exception as e:
                log(f"本地捕获失败: {e}")

    # ---------------- Safety: mouse movement detection ----------------
    def check_mouse_movement(self):
        try:
            cur = pyautogui.position()
            prev = self.last_mouse_pos
            dx = cur[0] - prev[0]
            dy = cur[1] - prev[1]
            dist = math.hypot(dx, dy)

            current_time = time.time()

            # 使用UI中设置的动态阈值
            threshold = self.safety_threshold_var.get()

            # 只有在任务运行中且未暂停时才检查安全
            if (
                self.enable_safety_check.get()
                and self.is_running
                and not self.pause_event.is_set()
            ):

                # 检查是否是任务移动
                is_task_movement = False

                # 方法1：检查是否在点击后的短时间内
                if hasattr(self, "last_click_time"):
                    time_since_last_click = current_time - self.last_click_time
                    if time_since_last_click < 3.0:  # 点击后3秒内认为是任务移动
                        is_task_movement = True

                # 方法2：检查鼠标是否正在向目标点位移动
                if (
                    hasattr(self, "current_target_point")
                    and self.current_target_point is not None
                ):
                    target_x, target_y = self.current_target_point
                    # 计算到目标点的距离
                    to_target_dist = math.hypot(cur[0] - target_x, cur[1] - target_y)
                    # 如果鼠标正在接近目标点，认为是任务移动
                    if to_target_dist < 100:  # 距离目标点100像素内
                        is_task_movement = True

                # 只有当不是任务移动且移动距离超过阈值时才触发安全暂停
                if (
                    dist > threshold
                    and not is_task_movement
                    and current_time - self.last_safety_trigger > 5
                ):  # 5秒冷却

                    self.pause_event.set()
                    self.pause_btn.config(text="继续 (F4)")
                    self.add_progress_text("安全机制：检测到异常鼠标移动，已暂停任务")
                    self.add_progress_text("如果是正常操作，请按F4继续任务")
                    log(f"安全暂停：鼠标移动距离 {dist:.1f}px (阈值: {threshold}px)")
                    self.last_safety_trigger = current_time

            self.last_mouse_pos = cur
        except Exception:
            pass

    # ---------------- Help / Logs / Reports ----------------
    def show_help(self):
        help_text = """
╔═══════════════════════════════════════════════════╗
║                AutoClicker 使用说明               ║
╚═══════════════════════════════════════════════════╝

🎯 基本操作流程：
1. 点击【开始获取坐标】进入捕获模式
2. 移动鼠标到目标位置，按 F2 记录坐标（支持全局捕获）
3. 配置任务参数后点击【启动任务】开始执行

📝 点位管理：
• 双击点位列表可编辑坐标、名称、点击参数
• 支持右键菜单：上移/下移/删除/重命名点位
• 每个点位可独立设置：点击次数、点击间隔、延时
• 支持鼠标左键、右键、中键点击操作
• 新增键盘操作类型：支持快捷键组合输入

⚙️ 任务设置：
• 循环次数：任务重复执行的次数
• 随机偏移：±像素范围，使点击位置更自然
• 随机延时：±秒范围，模拟人类操作间隔
• 启动倒计时：任务开始前的准备时间
• 任务结束动作：可选择播放提示音

🛡️ 安全特性：
• 智能防误触：检测到异常鼠标移动自动暂停
• 可调节安全阈值：50-1000像素灵敏度
• 调试模式：逐步执行，便于调试复杂任务
• 启动前确认：显示任务摘要，避免误操作

⌨️ 快捷键：
• F2 - 记录当前鼠标坐标（全局/窗口内）
• F3 - 开始/停止任务（全局）
• F4 - 暂停/继续任务（全局）

🖼️ 窗口管理：
• 自动窗口切换：执行前自动激活目标窗口
• 窗口列表：显示所有可用窗口供选择
• 窗口捕获：一键获取当前活动窗口信息

🔧 高级功能：
• 主题切换：支持浅色/深色主题
• 托盘运行：最小化到系统托盘后台运行
• 日志记录：详细的操作日志和点击统计
• 配置导入导出：JSON格式保存和分享配置
• 执行报告：生成详细的点击统计报告

📊 统计信息：
• 实时显示：总尝试次数、成功/失败次数
• 循环进度：当前完成的循环次数
• 时间统计：任务开始和结束时间
• 报告导出：可导出为文本文件

💡 使用技巧：
• 对于重复性操作，设置合适的随机偏移和延时
• 使用调试模式逐步验证复杂点击序列
• 合理设置安全阈值，平衡灵敏度和误触防护
• 利用窗口管理功能确保点击在正确的应用窗口

🆘 注意事项：
• 请合理使用，避免用于游戏作弊等违规用途
• 执行期间避免大幅度移动鼠标
• 建议先在调试模式下测试任务流程
• 如遇问题可查看日志文件排查原因

版本：{} - 更智能、更安全的自动点击工具
        """.format(
            VERSION
        )

        # 创建自定义对话框显示帮助信息
        help_window = tk.Toplevel(self.root)
        help_window.title("AutoClicker 详细使用说明")
        help_window.geometry("800x700")
        help_window.resizable(True, True)
        help_window.transient(self.root)
        help_window.grab_set()

        # 应用主题
        colors = ModernStyles.get_theme(self.theme_var.get())
        help_window.configure(bg=colors["bg"])

        # 主框架
        main_frame = GlassFrame(help_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # 在这里设置内边距
        main_frame.apply_glass_effect(self.theme_var.get())

        # 标题
        title_label = tk.Label(
            main_frame,
            text="AutoClicker 使用说明",
            font=("Segoe UI", 16, "bold"),
            fg=colors["accent"],
            pady=10,
            bg=colors["card_bg"],
        )
        title_label.pack()

        # 创建文本框和滚动条
        text_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        text_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 文本框
        help_text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            bg=colors["input_bg"],
            fg=colors["text"],
            padx=15,
            pady=15,
            relief="flat",
            borderwidth=1,
        )
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.config(state=tk.DISABLED)  # 设置为只读

        # 滚动条
        scrollbar = tk.Scrollbar(
            text_frame, orient=tk.VERTICAL, command=help_text_widget.yview
        )
        help_text_widget.configure(yscrollcommand=scrollbar.set)

        help_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮框架
        button_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        button_frame.pack(fill=tk.X, pady=10)

        # 快速入门按钮
        quick_start_btn = ModernButton(
            button_frame,
            text="🎯 快速入门指南",
            command=self.show_quick_start,
            button_type="primary",
            width=15,
        )
        quick_start_btn.pack(side=tk.LEFT, padx=5)
        quick_start_btn.apply_theme(self.theme_var.get())

        # 快捷键参考按钮
        shortcut_btn = ModernButton(
            button_frame,
            text="⌨️ 快捷键参考",
            command=self.show_shortcut_reference,
            button_type="secondary",
            width=15,
        )
        shortcut_btn.pack(side=tk.LEFT, padx=5)
        shortcut_btn.apply_theme(self.theme_var.get())

        # 在关于对话框的按钮部分添加更新检查按钮
        update_btn = ModernButton(
            btn_frame,
            text="🔍 检查更新",  # 可以加个图标更直观
            command=self.manual_check_update,  # 绑定到手动检查更新的函数
            button_type="success",
            width=12,  # 宽度可以稍微调整下
        )
        update_btn.pack(side=tk.LEFT, padx=5)  # 放在关闭按钮旁边
        update_btn.apply_theme(self.theme_var.get())

        # 关闭按钮
        close_btn = ModernButton(
            button_frame,
            text="关闭",
            command=help_window.destroy,
            button_type="secondary",
            width=10,
        )
        close_btn.pack(side=tk.RIGHT, padx=5)
        close_btn.apply_theme(self.theme_var.get())

        # 居中显示
        help_window.update_idletasks()
        x = (self.root.winfo_rootx() + self.root.winfo_width() // 2) - (
            help_window.winfo_width() // 2
        )
        y = (self.root.winfo_rooty() + self.root.winfo_height() // 2) - (
            help_window.winfo_height() // 2
        )
        help_window.geometry(f"+{x}+{y}")

        # 绑定ESC键关闭
        help_window.bind("<Escape>", lambda e: help_window.destroy())
        help_window.bind("<Return>", lambda e: help_window.destroy())

    def show_quick_start(self):
        """显示快速入门指南"""
        quick_start_text = """
╔═══════════════════════════════════════════════════╗
║                  🎯 快速入门指南                 ║
╚═══════════════════════════════════════════════════╝

1. 第一步：获取坐标
   • 点击【开始获取坐标】按钮
   • 移动鼠标到目标位置
   • 按 F2 键记录坐标
   • 重复以上步骤添加多个点位

2. 第二步：配置任务
   • 设置循环次数（如：10次）
   • 设置基础延时（如：1秒）
   • 可选：设置随机偏移（如：5像素）
   • 可选：设置随机延时（如：0.5秒）

3. 第三步：启动任务
   • 点击【启动任务】或按 F3
   • 确认任务参数
   • 等待倒计时结束
   • 任务自动执行

4. 第四步：监控和调整
   • 实时查看执行进度
   • 按 F4 可暂停/继续
   • 按 F3 可停止任务
   • 查看统计信息

💡 提示：首次使用建议先添加1-2个点位测试！
        """

        self._show_info_dialog("快速入门指南", quick_start_text)

    def show_shortcut_reference(self):
        """显示快捷键参考"""
        shortcut_text = """
╔═══════════════════════════════════════════════════╗
║                  ⌨️ 快捷键参考                   ║
╚═══════════════════════════════════════════════════╝

全局快捷键（需要 keyboard 库支持）：
┌──────────┬──────────────────────────────────────┐
│  快捷键  │               功能                   │
├──────────┼──────────────────────────────────────┤
│   F2     │ 记录当前鼠标坐标                     │
│   F3     │ 开始/停止任务                        │
│   F4     │ 暂停/继续任务                        │
└──────────┴──────────────────────────────────────┘

窗口内快捷键：
┌──────────┬──────────────────────────────────────┐
│  快捷键  │               功能                   │
├──────────┼──────────────────────────────────────┤
│  F2      │ 记录坐标（捕获模式）                 │
│  F3      │ 开始/停止任务                        │
│  F4      │ 暂停/继续任务                        │
│  Ctrl+S  │ 保存配置                             │
│  Ctrl+O  │ 导入配置                             │
│  Ctrl+E  │ 导出配置                             │
│  Del     │ 删除选中点位                         │
└──────────┴──────────────────────────────────────┘

💡 提示：如果全局快捷键无效，请以管理员权限运行程序！
        """

        self._show_info_dialog("快捷键参考", shortcut_text)

    def _show_info_dialog(self, title, content):
        """显示信息对话框的通用方法"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x500")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()

        # 应用主题
        colors = ModernStyles.get_theme(self.theme_var.get())
        dialog.configure(bg=colors["bg"])

        # 主框架
        main_frame = GlassFrame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_frame.apply_glass_effect(self.theme_var.get())

        # 文本框
        text_widget = tk.Text(
            main_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg=colors["input_bg"],
            fg=colors["text"],
            padx=15,
            pady=15,
            relief="flat",
        )
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)

        # 滚动条
        scrollbar = tk.Scrollbar(
            main_frame, orient=tk.VERTICAL, command=text_widget.yview
        )
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 关闭按钮
        btn_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        btn_frame.pack(fill=tk.X, pady=10)

        close_btn = ModernButton(
            btn_frame, text="关闭", command=dialog.destroy, button_type="primary"
        )
        close_btn.pack(side=tk.RIGHT)
        close_btn.apply_theme(self.theme_var.get())

        # 居中显示
        dialog.update_idletasks()
        x = (self.root.winfo_rootx() + self.root.winfo_width() // 2) - (
            dialog.winfo_width() // 2
        )
        y = (self.root.winfo_rooty() + self.root.winfo_height() // 2) - (
            dialog.winfo_height() // 2
        )
        dialog.geometry(f"+{x}+{y}")

        # 绑定ESC键关闭
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def open_log_file(self):
        if os.path.exists(LOG_FILE):
            try:
                if os.name == "nt":
                    os.startfile(LOG_FILE)
                elif os.name == "posix":
                    if os.system("which xdg-open >/dev/null 2>&1") == 0:
                        os.system(f"xdg-open {LOG_FILE} >/dev/null 2>&1 &")
                    elif os.system("which open >/dev/null 2>&1") == 0:
                        os.system(f"open {LOG_FILE} >/dev/null 2>&1 &")
                else:
                    messagebox.showinfo(
                        "日志文件", f"日志保存在: {os.path.abspath(LOG_FILE)}"
                    )
            except Exception:
                messagebox.showinfo(
                    "日志文件", f"日志保存在: {os.path.abspath(LOG_FILE)}"
                )
        else:
            messagebox.showinfo("日志文件", "日志文件尚未生成")

    def add_progress_text(self, text):
        now = time.strftime("%H:%M:%S")
        line = f"{now} - {text}\n"
        try:
            self.progress_details.config(state=tk.NORMAL)
            self.progress_details.insert(tk.END, line)
            self.progress_details.see(tk.END)
            self.progress_details.config(state=tk.DISABLED)
        except Exception:
            pass

    # ---------------- Click report (statistics) ----------------
    def _stats_text(self):
        s = self.stats
        if not s or s["start_time"] is None:
            return "统计：尚未开始任务"
        start = s.get("start_time", "N/A")
        end = s.get("end_time", "N/A")
        return (
            f"统计：尝试 {s['total_click_attempts']} 次，成功 {s['successful_clicks']}，失败 {s['failed_clicks']}，"
            f"循环完成 {s['loops_completed']}，开始 {start}，结束 {end}"
        )

    def update_stats_display(self):
        self.stats_label.config(text=self._stats_text())

    def write_click_report(self, silent=False):
        try:
            s = self.stats
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            content = []
            content.append(f"AutoClicker 执行报告 - {now}")
            content.append(f"版本: {VERSION}")
            content.append(f"点位数: {len(self.click_points)}")
            content.append(
                f"参数: base_delay={self.base_delay_var.get()}, loop={self.loop_var.get()}, random_offset={self.random_offset_var.get()}, random_delay={self.random_delay_var.get()}"
            )
            content.append("点位列表:")
            for i, p in enumerate(self.click_points, 1):
                content.append(
                    f"  {i}. {p['name']} ({p['x']},{p['y']}) 按键:{p['button']} 延时:{p['delay']} 点击:{p['click_count']}次 间隔:{p['click_interval']}"
                )
            content.append("统计:")
            content.append(f"  尝试: {s['total_click_attempts']}")
            content.append(f"  成功: {s['successful_clicks']}")
            content.append(f"  失败: {s['failed_clicks']}")
            content.append(f"  循环完成: {s['loops_completed']}")
            content.append(f"  开始时间: {s.get('start_time')}")
            content.append(f"  结束时间: {s.get('end_time')}")
            content.append("日志结束\n\n")
            with open(CLICK_REPORT, "a", encoding="utf-8") as f:
                f.write("\n".join(content) + "\n")
            if not silent:
                messagebox.showinfo("报告已生成", f"执行报告已追加到 {CLICK_REPORT}")
            log(f"写入执行报告到 {CLICK_REPORT}")
        except Exception as e:
            log(f"写入报告失败: {e}")

    def export_click_report(self):
        # save a copy
        if not os.path.exists(CLICK_REPORT):
            messagebox.showinfo("提示", "尚无报告文件，先运行一次任务以生成")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("文本文件", "*.txt")]
        )
        if not path:
            return
        try:
            with open(CLICK_REPORT, "r", encoding="utf-8") as src, open(
                path, "w", encoding="utf-8"
            ) as dst:
                dst.write(src.read())
            messagebox.showinfo("导出成功", f"已导出报告到 {path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    # ---------------- Tray integration ----------------
    def minimize_to_tray(self):
        if not HAVE_PYSTRAY:
            messagebox.showwarning(
                "托盘支持不可用", "未检测到 pystray 或 Pillow，无法最小化到托盘。"
            )
            return
        # hide window and start tray in background thread
        self.hide_window_to_tray()

    def hide_window_to_tray(self):
        """隐藏窗口到托盘 - 使用自定义图标"""
        try:
            # 创建托盘图标图像
            img = None

            # 尝试使用自定义图标
            if hasattr(self, "icon_dir") and self.icon_dir:
                icon_sizes = [64, 32, 48, 128, 256, 16]
                for size in icon_sizes:
                    icon_path = os.path.join(self.icon_dir, f"{size}.png")
                    if os.path.exists(icon_path):
                        try:
                            img = Image.open(icon_path)
                            print(f"使用托盘图标: {icon_path}")
                            break
                        except Exception as e:
                            print(f"加载托盘图标 {icon_path} 失败: {e}")
                            continue

            # 如果自定义图标加载失败，创建默认图标
            if img is None:
                print("使用默认托盘图标")
                img = Image.new("RGB", (64, 64), color=(30, 144, 255))
                d = ImageDraw.Draw(img)
                d.ellipse((10, 10, 54, 54), fill="white")
                d.text((18, 18), "AC", fill="black")

            # 创建托盘菜单
            menu = Menu(
                MenuItem("打开主窗口", lambda icon, item: self.show_from_tray()),
                MenuItem("开始任务", lambda icon, item: self.start_task()),
                MenuItem("暂停/继续", lambda icon, item: self.toggle_pause()),
                MenuItem("停止任务", lambda icon, item: self.stop_task()),
                MenuItem("退出", lambda icon, item: self.tray_exit()),
            )

            self.tray_icon = Icon("AutoClicker", img, f"AutoClicker {VERSION}", menu)

            # 隐藏窗口
            self.root.withdraw()

            # 在单独线程中运行托盘图标
            def tray_run():
                try:
                    self.tray_icon.run()
                except Exception as e:
                    log(f"托盘图标运行异常: {e}")

            self.tray_thread = threading.Thread(target=tray_run, daemon=True)
            self.tray_thread.start()
            self.tray_running = True
            self.add_progress_text("已最小化到托盘（后台运行）")
            log("最小化到托盘")

        except Exception as e:
            print(f"创建托盘图标失败: {e}")
            messagebox.showerror("错误", f"托盘图标创建失败: {e}")

    def show_from_tray(self):
        try:
            if self.tray_icon:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
            self.root.deiconify()
            self.root.lift()
            self.tray_running = False
            self.add_progress_text("从托盘恢复窗口")
        except Exception as e:
            log(f"从托盘恢复失败: {e}")

    def tray_exit(self):
        try:
            if self.tray_icon:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
            # ensure stop running tasks
            self.stop_event.set()
            self.tray_running = False
            self.root.quit()
            log("通过托盘退出程序")
        except Exception as e:
            log(f"托盘退出异常: {e}")

    # ---------------- Update check ----------------
    def check_update(self, auto_check=True):
        """检查更新 - 支持多个备用源和重试机制
        auto_check: True表示自动检查（后台），False表示手动检查（显示对话框）
        """
        if not HAVE_REQUESTS:
            if not auto_check:  # 只有手动检查时才显示警告
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "检查更新",
                        "需要 requests 库才能检查更新。\n请运行: pip install requests",
                    ),
                )
            return

        if not UPDATE_CHECK_URLS:
            return

        # 显示检查中的状态
        self.status_bar.config(text="正在检查更新...")
        if not auto_check:
            self.add_progress_text("开始检查更新，使用多个备用源...")

        def check_with_feedback():
            success = False
            last_error = ""
            update_available = False
            update_info = {}

            for url_index, url in enumerate(UPDATE_CHECK_URLS):
                for attempt in range(2):  # 每个URL尝试2次
                    try:
                        # 显示当前尝试信息
                        source_name = "GitHub" if "github" in url else "CDN"
                        attempt_info = f"尝试从 {source_name} 检查更新"
                        if attempt > 0:
                            attempt_info += f" (重试 {attempt})"

                        if not auto_check:
                            self.add_progress_text(attempt_info)

                        # 重试延迟
                        if attempt > 0:
                            time.sleep(2)  # 重试前等待2秒

                        # 发送请求
                        resp = requests.get(url, timeout=15)

                        if resp.status_code == 200:
                            data = resp.json()
                            remote_version = data.get("version")
                            download_url = data.get("download_url") or data.get("url")
                            changelog = data.get("changelog", "")
                            release_notes = data.get("release_notes", "")

                            if remote_version and self.is_newer_version(
                                remote_version, VERSION
                            ):
                                update_available = True
                                update_info = {
                                    "version": remote_version,
                                    "download_url": download_url,
                                    "changelog": changelog,
                                    "release_notes": release_notes,
                                }
                                if not auto_check:

                                    def notify():
                                        update_msg = f"发现新版本 {remote_version}！\n\n当前版本: {VERSION}"

                                        if changelog:
                                            update_msg += f"\n\n更新内容:\n{changelog}"
                                        if release_notes:
                                            update_msg += (
                                                f"\n\n发布说明:\n{release_notes}"
                                            )

                                        update_msg += "\n\n是否打开下载页面？"

                                        if messagebox.askyesno(
                                            "检测到新版本", update_msg
                                        ):
                                            self.open_download_page(
                                                download_url
                                                or "https://github.com/MGHYGitHub/AutoClicker/releases"
                                            )

                                    self.root.after(0, notify)
                                if auto_check:
                                    self.add_progress_text(
                                        f"发现新版本: {remote_version}"
                                    )
                                else:
                                    self.add_progress_text(
                                        f"发现新版本: {remote_version}"
                                    )
                            else:
                                if not auto_check:

                                    def notify_up_to_date():
                                        messagebox.showinfo(
                                            "检查更新", f"当前已是最新版本 ({VERSION})"
                                        )

                                    self.root.after(0, notify_up_to_date)
                                if auto_check:
                                    self.add_progress_text("当前已是最新版本")
                                else:
                                    self.add_progress_text("当前已是最新版本")

                            success = True
                            if not auto_check:
                                self.add_progress_text(
                                    f"更新检查成功 (来源: {source_name})"
                                )
                            break  # 成功获取，跳出重试循环

                        else:
                            last_error = f"HTTP状态码: {resp.status_code}"
                            log(f"更新检查失败 [{url}]，HTTP状态码: {resp.status_code}")

                    except requests.exceptions.Timeout:
                        last_error = "连接超时"
                        log(f"更新检查超时 [{url}] (尝试 {attempt + 1})")

                    except requests.exceptions.ConnectionError as e:
                        last_error = f"连接错误: {str(e)}"
                        log(f"更新检查连接错误 [{url}] (尝试 {attempt + 1}): {e}")

                    except requests.exceptions.JSONDecodeError as e:
                        last_error = "响应数据格式错误"
                        log(f"更新检查JSON解析错误 [{url}]: {e}")

                    except Exception as e:
                        last_error = f"未知错误: {str(e)}"
                        log(f"更新检查失败 [{url}] (尝试 {attempt + 1}): {e}")

                if success:
                    break  # 成功获取，跳出URL循环

            # 更新状态栏
            self.root.after(0, lambda: self.status_bar.config(text="就绪"))

            # 自动检查时发现新版本，显示通知但不阻塞
            if auto_check and update_available:

                def show_update_notification():
                    update_msg = (
                        f"发现新版本 {update_info['version']}！\n\n当前版本: {VERSION}"
                    )

                    if update_info["changelog"]:
                        update_msg += f"\n\n更新内容:\n{update_info['changelog']}"
                    if update_info["release_notes"]:
                        update_msg += f"\n\n发布说明:\n{update_info['release_notes']}"

                    update_msg += "\n\n是否打开下载页面？"

                    # 使用askyesno会阻塞，但我们希望用户能看到通知
                    if messagebox.askyesno("检测到新版本", update_msg):
                        self.open_download_page(
                            update_info["download_url"]
                            or "https://github.com/MGHYGitHub/AutoClicker/releases"
                        )

                self.root.after(0, show_update_notification)

            # 所有尝试都失败
            if not success and not auto_check:
                self.root.after(0, lambda: self.show_update_failed_message(last_error))
                self.add_progress_text(f"更新检查失败: {last_error}")

        # 在后台线程中执行检查
        threading.Thread(target=check_with_feedback, daemon=True).start()

    def show_update_failed_message(self, error_detail=""):
        """显示更新失败消息"""
        error_msg = (
            "暂时无法连接到更新服务器。\n\n"
            "这可能是因为:\n"
            "• 网络连接问题\n"
            "• GitHub访问限制\n"
            "• 防火墙或代理设置\n"
            "• 服务器暂时不可用\n\n"
        )

        if error_detail:
            error_msg += f"错误详情: {error_detail}\n\n"

        error_msg += (
            "您可以:\n"
            "1. 检查网络连接后重试\n"
            "2. 手动访问项目页面查看更新\n"
            "3. 稍后再试"
        )

        # 提供更多选项
        choice = messagebox.askyesno(
            "检查更新失败", error_msg + "\n\n是否立即手动访问项目页面？"
        )

        if choice:
            self.open_download_page(
                "https://github.com/MGHYGitHub/AutoClicker/releases"
            )

    def is_newer_version(self, remote, current):
        """增强版版本号比较"""
        try:
            # 处理版本号中的非数字字符（如 "v2.5.1" -> "2.5.1"）
            remote = remote.lstrip("vV")
            current = current.lstrip("vV")

            # 分割版本号
            remote_parts = list(map(int, remote.split(".")))
            current_parts = list(map(int, current.split(".")))

            # 确保长度一致
            max_len = max(len(remote_parts), len(current_parts))
            remote_parts.extend([0] * (max_len - len(remote_parts)))
            current_parts.extend([0] * (max_len - len(current_parts)))

            # 逐个比较
            for r, c in zip(remote_parts, current_parts):
                if r > c:
                    return True
                elif r < c:
                    return False
            return False  # 版本相同
        except:
            # 如果解析失败，使用简单的字符串比较
            try:
                return remote > current
            except:
                return False

    def open_download_page(self, url):
        """打开下载页面"""
        try:
            import webbrowser

            webbrowser.open(url)
            self.add_progress_text(f"已打开下载页面: {url}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开浏览器: {e}")

    def manual_check_update(self):
        """手动检查更新 - 提供更多选项"""
        if not UPDATE_CHECK_URLS:
            response = messagebox.askyesno(
                "未配置更新检查", "更新检查未配置。\n\n是否查看配置说明？"
            )
            if response:
                self.show_github_update_help()
            return

        if not HAVE_REQUESTS:
            messagebox.showwarning(
                "缺少依赖",
                "需要 requests 库才能检查更新。\n\n请运行以下命令安装:\npip install requests",
            )
            return

        # 提供选项让用户选择
        choice = messagebox.askyesnocancel(
            "检查更新",
            "即将检查更新，这需要网络连接。\n\n"
            "是 - 立即检查更新（自动尝试多个源）\n"
            "否 - 手动访问GitHub页面\n"
            "取消 - 取消操作",
        )

        if choice is None:  # 取消
            return
        elif not choice:  # 否 - 手动访问
            self.open_download_page(
                "https://github.com/MGHYGitHub/AutoClicker/releases"
            )
            return

        # 是 - 开始检查（手动检查，auto_check=False）
        self.status_bar.config(text="正在检查更新...")
        self.add_progress_text("开始手动检查更新...")

        def check_with_feedback():
            try:
                self.check_update(auto_check=False)
            finally:
                self.root.after(0, lambda: self.status_bar.config(text="就绪"))

        threading.Thread(target=check_with_feedback, daemon=True).start()

    def show_github_update_help(self):
        """显示 GitHub 更新配置帮助"""
        help_text = """
    GitHub 更新检查配置：

    1. 自动配置（已设置）：
    当前已配置多个备用源，包括 GitHub 和 CDN。

    2. 版本文件要求：
    在 GitHub 仓库根目录创建 version.json 文件，内容如下：
    {{
        "version": "2.5.1",
        "download_url": "https://github.com/.../下载链接",
        "changelog": "修复了...",
        "release_notes": "详细更新说明"
    }}

    3. 版本文件位置：
    推荐使用 GitHub Raw 地址：
    https://raw.githubusercontent.com/用户名/仓库名/分支名/version.json

    4. 版本号比较：
    - 支持语义化版本号 (如 2.5.1)
    - 会自动比较远程版本是否比当前版本新

    当前配置状态: 已配置多个备用源
    当前版本: {version}
    主要更新检查URL: {update_url}
    """.format(
            version=VERSION,
            update_url=UPDATE_CHECK_URLS[0] if UPDATE_CHECK_URLS else "未配置",
        )

        messagebox.showinfo("GitHub 更新配置说明", help_text)

    def is_newer_version(self, remote, current):
        """增强版版本号比较"""
        try:
            # 处理版本号中的非数字字符（如 "v2.5.1" -> "2.5.1"）
            remote = remote.lstrip("vV")
            current = current.lstrip("vV")

            # 分割版本号
            remote_parts = list(map(int, remote.split(".")))
            current_parts = list(map(int, current.split(".")))

            # 确保长度一致
            max_len = max(len(remote_parts), len(current_parts))
            remote_parts.extend([0] * (max_len - len(remote_parts)))
            current_parts.extend([0] * (max_len - len(current_parts)))

            # 逐个比较
            for r, c in zip(remote_parts, current_parts):
                if r > c:
                    return True
                elif r < c:
                    return False
            return False  # 版本相同
        except:
            # 如果解析失败，使用简单的字符串比较
            try:
                return remote > current
            except:
                return False

    def open_download_page(self, url):
        """打开下载页面"""
        try:
            import webbrowser

            webbrowser.open(url)
            self.add_progress_text(f"已打开下载页面: {url}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开浏览器: {e}")

    def manual_check_update(self):
        """手动检查更新 - 提供更多选项"""
        if not UPDATE_CHECK_URLS:
            response = messagebox.askyesno(
                "未配置更新检查", "更新检查未配置。\n\n是否查看配置说明？"
            )
            if response:
                self.show_github_update_help()
            return

        if not HAVE_REQUESTS:
            messagebox.showwarning(
                "缺少依赖",
                "需要 requests 库才能检查更新。\n\n请运行以下命令安装:\npip install requests",
            )
            return

        # 提供选项让用户选择
        choice = messagebox.askyesnocancel(
            "检查更新",
            "即将检查更新，这需要网络连接。\n\n"
            "是 - 立即检查更新（自动尝试多个源）\n"
            "否 - 手动访问GitHub页面\n"
            "取消 - 取消操作",
        )

        if choice is None:  # 取消
            return
        elif not choice:  # 否 - 手动访问
            self.open_download_page(
                "https://github.com/MGHYGitHub/AutoClicker/releases"
            )
            return

        # 是 - 开始检查（手动检查，auto_check=False）
        self.status_bar.config(text="正在检查更新...")
        self.add_progress_text("开始手动检查更新...")

        def check_with_feedback():
            try:
                self.check_update(auto_check=False)
            finally:
                self.root.after(0, lambda: self.status_bar.config(text="就绪"))

        threading.Thread(target=check_with_feedback, daemon=True).start()

    # ---------------- Exit ----------------
    def on_exit(self):
        if self.is_running:
            if not messagebox.askyesno(
                "退出确认", "任务正在运行，确定退出并停止任务吗？"
            ):
                return
        # set stop
        self.stop_event.set()
        # save current points
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        log("程序退出")
        try:
            if self.tray_icon:
                try:
                    self.tray_icon.stop()
                except Exception:
                    pass
        except Exception:
            pass
        self.root.destroy()

    # ---------------- Keyboard safe wrapper for start/stop used by tray ----------------
    def start_task_wrapper(self):
        # used by tray menu (non-UI thread) - schedule on main thread
        self.root.after(0, self.start_task)

    def stop_task_wrapper(self):
        self.root.after(0, self.stop_task)

    # ---------------- About dialog ----------------
    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("关于 AutoClicker")
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()

        # 应用主题
        colors = ModernStyles.get_theme(self.theme_var.get())
        about_window.configure(bg=colors["bg"])

        # 居中显示
        about_window.geometry(
            "+%d+%d" % (self.root.winfo_rootx() + 200, self.root.winfo_rooty() + 150)
        )

        main_frame = GlassFrame(about_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        main_frame.apply_glass_effect(self.theme_var.get())

        # 标题
        title_label = tk.Label(
            main_frame,
            text=f"AutoClicker {VERSION}",
            font=("Segoe UI", 16, "bold"),
            fg=colors["accent"],
            bg=colors["card_bg"],
        )
        title_label.pack(pady=(0, 10))

        # 描述
        desc_text = "高级自动连点器\n" "支持多点位管理、独立点击设置、安全检测等功能"
        desc_label = tk.Label(
            main_frame,
            text=desc_text,
            font=("Segoe UI", 11),
            fg=colors["text"],
            bg=colors["card_bg"],
        )
        desc_label.pack(pady=(0, 20))

        # 作者信息
        author_frame = GlassFrame(main_frame, padx=10, pady=10)
        author_frame.pack(fill=tk.X, pady=(0, 15))
        author_frame.apply_glass_effect(self.theme_var.get())

        author_info = [
            ("作者", "Eli_Morgan"),
            ("版本", VERSION),
            ("开发语言", "Python 3"),
            ("界面库", "Tkinter"),
        ]

        for i, (label, value) in enumerate(author_info):
            row_frame = tk.Frame(author_frame, bg=colors["card_bg"])
            row_frame.pack(fill=tk.X, pady=2)
            tk.Label(
                row_frame,
                text=f"{label}:",
                width=8,
                anchor="w",
                font=("Segoe UI", 9, "bold"),
                bg=colors["card_bg"],
                fg=colors["text"],
            ).pack(side=tk.LEFT)
            tk.Label(
                row_frame,
                text=value,
                anchor="w",
                font=("Segoe UI", 9),
                bg=colors["card_bg"],
                fg=colors["text"],
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # GitHub链接 - Markdown风格
        github_frame = tk.Frame(author_frame, bg=colors["card_bg"])
        github_frame.pack(fill=tk.X, pady=2)
        tk.Label(
            github_frame,
            text="项目地址:",
            width=8,
            anchor="w",
            font=("Segoe UI", 9, "bold"),
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)

        # 创建类似Markdown链接的样式
        github_link = tk.Label(
            github_frame,
            text="GitHub",
            font=("Segoe UI", 9),
            fg=colors["accent"],
            cursor="hand2",
            underline=True,
            bg=colors["card_bg"],
        )
        github_link.pack(side=tk.LEFT)
        github_link.bind("<Button-1>", lambda e: self.open_github())

        # 联系方式
        contact_frame = GlassFrame(main_frame, padx=10, pady=10)
        contact_frame.pack(fill=tk.X, pady=(0, 15))
        contact_frame.apply_glass_effect(self.theme_var.get())

        contacts = [("邮箱", "Eli_Morgan2025@outlook.com")]

        for i, (label, value) in enumerate(contacts):
            row_frame = tk.Frame(contact_frame, bg=colors["card_bg"])
            row_frame.pack(fill=tk.X, pady=2)
            tk.Label(
                row_frame,
                text=f"{label}:",
                width=8,
                anchor="w",
                font=("Segoe UI", 9, "bold"),
                bg=colors["card_bg"],
                fg=colors["text"],
            ).pack(side=tk.LEFT)
            tk.Label(
                row_frame,
                text=value,
                anchor="w",
                font=("Segoe UI", 9),
                bg=colors["card_bg"],
                fg=colors["text"],
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 特别感谢
        thanks_frame = GlassFrame(main_frame, padx=10, pady=10)
        thanks_frame.pack(fill=tk.X, pady=(0, 15))
        thanks_frame.apply_glass_effect(self.theme_var.get())

        thanks_text = (
            "• PyAutoGUI - 鼠标键盘自动化\n"
            "• Keyboard - 全局快捷键支持\n"
            "• PyStray - 系统托盘集成\n"
            "• Pillow - 图像处理支持\n"
            "• 所有贡献者和用户"
        )
        thanks_label = tk.Label(
            thanks_frame,
            text=thanks_text,
            justify=tk.LEFT,
            font=("Segoe UI", 9),
            bg=colors["card_bg"],
            fg=colors["text"],
        )
        thanks_label.pack(anchor="w")

        # 版权信息
        copyright_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        copyright_frame.pack(fill=tk.X, pady=(10, 0))

        copyright_text = "© 2025 AutoClicker. 保留所有权利。"
        copyright_label = tk.Label(
            copyright_frame,
            text=copyright_text,
            font=("Segoe UI", 8),
            fg=colors["text_light"],
            bg=colors["card_bg"],
        )
        copyright_label.pack()

        # 按钮框架 - 在这里添加更新检查按钮
        btn_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        # 更新检查按钮
        update_btn = ModernButton(
            btn_frame,
            text="🔍 检查更新",
            command=self.manual_check_update,
            button_type="success",
            width=12,
        )
        update_btn.pack(side=tk.LEFT, padx=5)
        update_btn.apply_theme(self.theme_var.get())

        # 关闭按钮
        close_btn = ModernButton(
            btn_frame,
            text="关闭",
            command=about_window.destroy,
            button_type="primary",
            width=10,
        )
        close_btn.pack(side=tk.RIGHT, padx=5)
        close_btn.apply_theme(self.theme_var.get())

        # 绑定ESC键关闭
        about_window.bind("<Escape>", lambda e: about_window.destroy())

        # 更新窗口以计算正确的大小，然后居中显示
        about_window.update_idletasks()

        # 获取窗口的请求大小
        width = about_window.winfo_reqwidth()
        height = about_window.winfo_reqheight()

        # 计算居中位置
        x = (self.root.winfo_rootx() + self.root.winfo_width() // 2) - (width // 2)
        y = (self.root.winfo_rooty() + self.root.winfo_height() // 2) - (height // 2)

        # 设置窗口位置和最小大小
        about_window.geometry(f"+{x}+{y}")
        about_window.minsize(width, height)

    def open_github(self):
        """打开GitHub项目页面"""
        import webbrowser

        try:
            webbrowser.open("https://github.com/MGHYGitHub/AutoClicker")
            self.add_progress_text("已打开GitHub项目页面")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开浏览器: {e}")

    # ---------------- Window Management ----------------
    def capture_current_window(self):
        """获取当前活动窗口信息"""
        if not self.HAVE_WIN32:  # 添加这行检查
            messagebox.showerror("错误", "需要安装 pywin32 库才能使用窗口管理功能")
            return

        try:
            import win32gui
            import win32con

            # 获取当前活动窗口
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                window_title = win32gui.GetWindowText(hwnd)
                if window_title:
                    self.window_title_var.set(window_title)
                    self.window_handle = hwnd
                    self.window_status.config(
                        text=f"已捕获窗口: {window_title}",
                        fg=ModernStyles.get_theme(self.theme_var.get())["success"],
                    )
                    self.add_progress_text(f"已捕获目标窗口: {window_title}")
                    log(f"捕获窗口: {window_title} (句柄: {hwnd})")
                    return

            self.window_status.config(
                text="未找到有效窗口",
                fg=ModernStyles.get_theme(self.theme_var.get())["danger"],
            )
            messagebox.showwarning("警告", "未检测到有效窗口标题")

        except Exception as e:
            messagebox.showerror("错误", f"获取窗口信息失败: {e}")
            log(f"获取窗口失败: {e}")

    def clear_target_window(self):
        """清除目标窗口设置"""
        self.window_title_var.set("")
        self.window_handle = None
        self.window_status.config(
            text="未选择目标窗口",
            fg=ModernStyles.get_theme(self.theme_var.get())["danger"],
        )
        self.add_progress_text("已清除目标窗口设置")

    def test_window_switch(self):
        """测试窗口切换"""
        if not self.window_handle and not self.window_title_var.get():
            messagebox.showwarning("警告", "请先选择目标窗口")
            return

        try:
            if self.switch_to_target_window():
                messagebox.showinfo("成功", "窗口切换测试成功！")
            else:
                messagebox.showerror("错误", "窗口切换失败，请检查窗口是否存在")
        except Exception as e:
            messagebox.showerror("错误", f"窗口切换测试失败: {e}")

    def switch_to_target_window(self):
        """切换到目标窗口"""
        if not self.HAVE_WIN32:  # 添加这行检查
            log("pywin32 库未安装，无法切换窗口")
            return False

        try:
            import win32gui
            import win32con
            import time

            target_title = self.window_title_var.get().strip()
            if not target_title and not self.window_handle:
                log("未设置目标窗口")
                return False

            def window_enum_handler(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if target_title.lower() in title.lower():
                        ctx.append(hwnd)

            # 如果已有句柄，直接使用
            if self.window_handle and win32gui.IsWindow(self.window_handle):
                hwnd = self.window_handle
            else:
                # 通过标题查找窗口
                windows = []
                win32gui.EnumWindows(window_enum_handler, windows)

                if not windows:
                    log(f"未找到包含 '{target_title}' 的窗口")
                    return False

                hwnd = windows[0]
                self.window_handle = hwnd

            # 激活窗口
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)

            # 等待窗口激活
            time.sleep(0.5)

            log(f"成功切换到窗口: {win32gui.GetWindowText(hwnd)}")
            return True

        except Exception as e:
            log(f"切换窗口失败: {e}")
            return False

        except ImportError:
            log("pywin32 库未安装，无法切换窗口")
            return False
        except Exception as e:
            log(f"切换窗口失败: {e}")
            return False

    def ensure_target_window_active(self):
        """确保目标窗口处于活动状态"""
        if self.auto_switch_window.get() and self.window_title_var.get():
            return self.switch_to_target_window()
        return True

    def create_window_management_ui(self, parent):
        """创建完整的窗口管理界面"""
        window_frame = GlassFrame(parent)
        window_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        window_frame.apply_glass_effect(self.theme_var.get())

        # 标题
        window_title = tk.Label(
            window_frame,
            text="窗口管理",
            font=("Segoe UI", 12, "bold"),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        )
        window_title.pack(anchor="w", pady=(8, 12))

        # 窗口选择区域
        window_select_frame = tk.Frame(
            window_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        window_select_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            window_select_frame,
            text="目标窗口:",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
        ).pack(side=tk.LEFT)

        # 窗口选择组合框
        self.window_combo = ttk.Combobox(
            window_select_frame,
            textvariable=self.window_title_var,
            width=40,
            style="Modern.TCombobox",
        )
        self.window_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 刷新窗口列表按钮 - 修复
        refresh_btn = ModernButton(
            window_select_frame,
            text="刷新窗口列表",
            command=self.refresh_window_list,
            button_type="primary",
        )
        refresh_btn.pack(side=tk.LEFT, padx=2)
        refresh_btn.apply_theme(self.theme_var.get())

        # 窗口操作按钮
        window_btn_frame = tk.Frame(
            window_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        window_btn_frame.pack(fill=tk.X, pady=4)

        # 获取当前窗口按钮 - 修复
        capture_btn = ModernButton(
            window_btn_frame,
            text="获取当前窗口",
            command=self.capture_current_window,
            button_type="primary",
        )
        capture_btn.pack(side=tk.LEFT, padx=2)
        capture_btn.apply_theme(self.theme_var.get())

        # 测试切换按钮 - 修复
        test_btn = ModernButton(
            window_btn_frame,
            text="测试切换",
            command=self.test_window_switch,
            button_type="success",
        )
        test_btn.pack(side=tk.LEFT, padx=2)
        test_btn.apply_theme(self.theme_var.get())

        # 清除窗口按钮 - 修复
        clear_btn = ModernButton(
            window_btn_frame,
            text="清除窗口",
            command=self.clear_target_window,
            button_type="danger",
        )
        clear_btn.pack(side=tk.LEFT, padx=2)
        clear_btn.apply_theme(self.theme_var.get())

        # 显示所有窗口按钮 - 修复
        show_all_btn = ModernButton(
            window_btn_frame,
            text="显示所有窗口",
            command=self.show_all_windows,
            button_type="secondary",
        )
        show_all_btn.pack(side=tk.LEFT, padx=2)
        show_all_btn.apply_theme(self.theme_var.get())

        # 自动切换选项
        auto_switch_frame = tk.Frame(
            window_frame, bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"]
        )
        auto_switch_frame.pack(fill=tk.X, pady=4)

        self.auto_switch_window = tk.BooleanVar(value=False)
        tk.Checkbutton(
            auto_switch_frame,
            text="自动切换窗口",
            variable=self.auto_switch_window,
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text"],
            selectcolor=ModernStyles.get_theme(self.theme_var.get())["accent"],
        ).pack(side=tk.LEFT)

        tk.Label(
            auto_switch_frame,
            text="执行任务前自动切换到目标窗口",
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
            fg=ModernStyles.get_theme(self.theme_var.get())["text_light"],
        ).pack(side=tk.LEFT, padx=10)

        # 窗口状态显示
        self.window_status = tk.Label(
            window_frame,
            text="未选择目标窗口",
            fg=ModernStyles.get_theme(self.theme_var.get())["danger"],
            font=("Segoe UI", 9),
            bg=ModernStyles.get_theme(self.theme_var.get())["card_bg"],
        )
        self.window_status.pack(fill=tk.X, pady=2)

    def on_window_combo_select(self, event):
        """当从组合框选择窗口时的处理"""
        selected = self.window_combo.get()
        if (
            selected
            and hasattr(self, "window_handles")
            and selected in self.window_handles
        ):
            hwnd = self.window_handles[selected]
            self.window_handle = hwnd

            # 提取纯窗口标题（去掉句柄信息）
            title = selected.split(" (句柄: ")[0]
            self.window_title_var.set(title)

            self.window_status.config(
                text=f"已选择窗口: {title}",
                fg=ModernStyles.get_theme(self.theme_var.get())["success"],
            )
            self.add_progress_text(f"已选择窗口: {title} (句柄: {hwnd})")

    def refresh_window_list(self):
        """刷新窗口列表"""
        if not self.HAVE_WIN32:
            messagebox.showerror("错误", "需要安装 pywin32 库才能使用窗口管理功能")
            return

        try:
            import win32gui

            windows = []

            def enum_windows_callback(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if title:  # 只显示有标题的窗口
                        class_name = win32gui.GetClassName(hwnd)
                        # 过滤掉一些系统窗口
                        if class_name not in ["Progman", "WorkerW", "Shell_TrayWnd"]:
                            windows.append((title, hwnd))

            win32gui.EnumWindows(enum_windows_callback, None)

            # 更新组合框
            window_titles = [f"{title} (句柄: {hwnd})" for title, hwnd in windows]
            self.window_combo["values"] = window_titles

            # 保存窗口句柄映射
            self.window_handles = {
                f"{title} (句柄: {hwnd})": hwnd for title, hwnd in windows
            }

            self.add_progress_text(f"刷新窗口列表，找到 {len(windows)} 个窗口")

        except Exception as e:
            messagebox.showerror("错误", f"刷新窗口列表失败: {e}")
            log(f"刷新窗口列表失败: {e}")

    def show_all_windows(self):
        """显示所有窗口的详细信息"""
        if not self.HAVE_WIN32:
            messagebox.showerror("错误", "需要安装 pywin32 库才能使用窗口管理功能")
            return

        try:
            import win32gui

            windows_info = []

            def enum_windows_callback(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    windows_info.append(
                        {
                            "hwnd": hwnd,
                            "title": title,
                            "class": class_name,
                            "rect": rect,
                            "size": f"{rect[2]-rect[0]}x{rect[3]-rect[1]}",
                        }
                    )

            win32gui.EnumWindows(enum_windows_callback, None)

            # 创建窗口列表对话框
            self.show_window_list_dialog(windows_info)

        except Exception as e:
            messagebox.showerror("错误", f"获取窗口列表失败: {e}")

    def show_window_list_dialog(self, windows_info):
        """显示窗口列表对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("窗口列表 - 双击选择窗口")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()

        # 应用主题
        colors = ModernStyles.get_theme(self.theme_var.get())
        dialog.configure(bg=colors["bg"])

        # 主框架
        main_frame = GlassFrame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_frame.apply_glass_effect(self.theme_var.get())

        # 创建树形视图
        tree_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        columns = ("title", "class", "size", "hwnd")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)

        # 设置列
        tree.heading("title", text="窗口标题")
        tree.heading("class", text="窗口类")
        tree.heading("size", text="窗口大小")
        tree.heading("hwnd", text="窗口句柄")

        tree.column("title", width=400)
        tree.column("class", width=150)
        tree.column("size", width=100)
        tree.column("hwnd", width=100)

        # 添加数据
        for info in windows_info:
            tree.insert(
                "",
                tk.END,
                values=(
                    info["title"] or "(无标题)",
                    info["class"],
                    info["size"],
                    info["hwnd"],
                ),
            )

        # 滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮框架
        btn_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        btn_frame.pack(fill=tk.X)

        def on_select():
            selection = tree.selection()
            if selection:
                item = selection[0]
                values = tree.item(item, "values")
                title = values[0]
                hwnd = int(values[3])

                # 更新主界面
                self.window_title_var.set(title)
                self.window_handle = hwnd
                self.window_status.config(
                    text=f"已选择窗口: {title}",
                    fg=ModernStyles.get_theme(self.theme_var.get())["success"],
                )
                self.add_progress_text(f"已选择窗口: {title} (句柄: {hwnd})")

                dialog.destroy()

        def on_refresh():
            dialog.destroy()
            self.show_all_windows()

        select_btn = ModernButton(
            btn_frame, text="选择窗口", command=on_select, button_type="primary"
        )
        select_btn.pack(side=tk.LEFT, padx=(0, 10))
        select_btn.apply_theme(self.theme_var.get())

        refresh_btn = ModernButton(
            btn_frame, text="刷新列表", command=on_refresh, button_type="secondary"
        )
        refresh_btn.pack(side=tk.LEFT, padx=(0, 10))
        refresh_btn.apply_theme(self.theme_var.get())

        cancel_btn = ModernButton(
            btn_frame, text="取消", command=dialog.destroy, button_type="secondary"
        )
        cancel_btn.pack(side=tk.LEFT)
        cancel_btn.apply_theme(self.theme_var.get())

        # 绑定双击事件
        tree.bind("<Double-1>", lambda e: on_select())

        # 绑定回车键
        dialog.bind("<Return>", lambda e: on_select())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def show_github_update_help(self):
        """显示 GitHub 更新配置帮助"""
        help_text = """
    GitHub 更新检查配置：

    1. 自动配置（已设置）：
    当前已配置为: 
    https://raw.githubusercontent.com/MGHYGitHub/AutoClicker/main/version.json

    2. 版本文件要求：
    在 GitHub 仓库根目录创建 version.json 文件，内容如下：
    {{
        "version": "2.5.1",
        "download_url": "https://github.com/.../下载链接",
        "changelog": "修复了...",
        "release_notes": "详细更新说明"
    }}

    3. 版本文件位置：
    推荐使用 GitHub Raw 地址：
    https://raw.githubusercontent.com/用户名/仓库名/分支名/version.json

    4. 版本号比较：
    - 支持语义化版本号 (如 2.5.1)
    - 会自动比较远程版本是否比当前版本新

    当前配置状态: 已配置
    当前版本: {version}
    更新检查URL: {update_url}
    """.format(
            version=VERSION,
            update_url=UPDATE_CHECK_URLS[0] if UPDATE_CHECK_URLS else "未配置",
        )

        messagebox.showinfo("GitHub 更新配置说明", help_text)

    def setup_icons(self):
        """设置图标 - 单文件exe专用"""
        try:
            # 判断运行环境
            if getattr(sys, "frozen", False):
                # 打包后的环境
                base_path = sys._MEIPASS
                icon_dir = os.path.join(base_path, "ICON")
            else:
                # 开发环境
                icon_dir = "ICON"

            print(f"图标目录: {icon_dir}")

            if os.path.exists(icon_dir):
                # 设置窗口图标
                self.set_window_icon(icon_dir)
                self.icon_dir = icon_dir
                print("✅ 图标加载成功")
            else:
                print("⚠ 图标目录不存在，使用默认图标")
                self.icon_dir = None

        except Exception as e:
            print(f"❌ 图标设置失败: {e}")

    def set_window_icon(self, icon_dir):
        """设置窗口图标"""
        try:
            # 按优先级尝试不同尺寸
            sizes = [256, 128, 64, 48, 32, 16]

            for size in sizes:
                icon_path = os.path.join(icon_dir, f"{size}.png")
                if os.path.exists(icon_path):
                    try:
                        from PIL import Image, ImageTk

                        img = Image.open(icon_path)
                        photo = ImageTk.PhotoImage(img)

                        # 设置窗口图标
                        self.root.iconphoto(True, photo)

                        # 保存引用
                        if not hasattr(self, "icon_images"):
                            self.icon_images = []
                        self.icon_images.append(photo)

                        print(f"✅ 使用图标: {size}x{size}")
                        break
                    except Exception as e:
                        print(f"⚠ 加载图标失败 {size}x{size}: {e}")
                        continue
        except Exception as e:
            print(f"❌ 窗口图标设置失败: {e}")

    def set_window_icon(self, icon_dir):
        """设置窗口图标"""
        try:
            # 尝试不同尺寸的图标
            icon_sizes = [64, 48, 32, 128, 256, 16]

            for size in icon_sizes:
                icon_path = os.path.join(icon_dir, f"{size}.png")
                print(f"尝试加载图标: {icon_path}")
                if os.path.exists(icon_path):
                    try:
                        # 使用PIL加载图像
                        from PIL import Image, ImageTk

                        img = Image.open(icon_path)
                        # 转换为PhotoImage
                        photo = ImageTk.PhotoImage(img)
                        # 设置窗口图标
                        self.root.iconphoto(True, photo)
                        # 保存引用防止垃圾回收
                        if not hasattr(self, "icon_images"):
                            self.icon_images = []
                        self.icon_images.append(photo)
                        print(f"成功设置窗口图标: {icon_path}")
                        return
                    except Exception as e:
                        print(f"加载窗口图标 {icon_path} 失败: {e}")
                        continue

            print("未找到可用的图标文件")

        except Exception as e:
            print(f"设置窗口图标失败: {e}")


# ---------------- Dialogs for adding/editing points ----------------
class AddPointDialog:
    def __init__(
        self,
        parent,
        x,
        y,
        default_button="left",
        default_delay=1.0,
        default_click_count=1,
        default_click_interval=0.1,
        theme="light",
    ):
        self.top = tk.Toplevel(parent)
        self.top.title("添加点位")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.result = None
        self.theme = theme
        colors = ModernStyles.get_theme(theme)

        # 应用主题
        self.top.configure(bg=colors["bg"])

        # 居中显示
        self.top.geometry(
            "+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50)
        )

        main_frame = GlassFrame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        main_frame.apply_glass_effect(theme)

        tk.Label(
            main_frame,
            text=f"当前位置：({x}, {y})",
            font=("Segoe UI", 9),
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(anchor="w", pady=(0, 8))

        # 名称
        name_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        name_frame.pack(fill=tk.X, pady=4)
        tk.Label(
            name_frame,
            text="名称:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.name_entry = tk.Entry(
            name_frame, bg=colors["input_bg"], fg=colors["text"], relief="flat"
        )
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.name_entry.insert(0, f"点位({x},{y})")

        # 坐标设置
        coord_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        coord_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            coord_frame,
            text="X坐标:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.x_entry = tk.Entry(
            coord_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.x_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.x_entry.insert(0, str(x))

        tk.Label(
            coord_frame,
            text="Y坐标:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.y_entry = tk.Entry(
            coord_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.y_entry.pack(side=tk.LEFT)
        self.y_entry.insert(0, str(y))

        # 操作类型和按键设置
        action_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        action_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            action_frame,
            text="操作类型:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.action_type_var = tk.StringVar(value="click")
        action_combo = ttk.Combobox(
            action_frame,
            textvariable=self.action_type_var,
            values=["click", "keyboard"],
            state="readonly",
            width=10,
            style="Modern.TCombobox",
        )
        action_combo.pack(side=tk.LEFT, padx=(0, 10))
        action_combo.bind("<<ComboboxSelected>>", self.on_action_type_change)

        # 点击键设置
        self.click_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        self.click_frame.pack(fill=tk.X, pady=4)

        # 窗口管理相关
        self.target_window = None
        self.window_handle = None
        self.auto_switch_window = tk.BooleanVar(value=False)  # 自动切换窗口
        self.window_title_var = tk.StringVar(value="")  # 目标窗口标题

        tk.Label(
            self.click_frame,
            text="点击键:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.button_var = tk.StringVar(value=default_button)
        ttk.Combobox(
            self.click_frame,
            textvariable=self.button_var,
            values=["left", "right", "middle"],
            state="readonly",
            width=8,
            style="Modern.TCombobox",
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(
            self.click_frame,
            text="点击次数:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.click_count_entry = tk.Entry(
            self.click_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.click_count_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.click_count_entry.insert(0, str(default_click_count))

        # 键盘操作设置（初始隐藏）
        self.keyboard_frame = tk.Frame(main_frame, bg=colors["card_bg"])

        tk.Label(
            self.keyboard_frame,
            text="按键组合:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.keys_var = tk.StringVar(value="")
        keys_entry = ttk.Entry(
            self.keyboard_frame,
            textvariable=self.keys_var,
            width=20,
            style="Modern.TCombobox",
        )
        keys_entry.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(
            self.keyboard_frame,
            text="例如: ctrl+a, ctrl+c, ctrl+v",
            font=("Segoe UI", 8),
            bg=colors["card_bg"],
            fg=colors["text_light"],
        ).pack(side=tk.LEFT)

        # 延时设置
        delay_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        delay_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            delay_frame,
            text="点位延时(秒):",
            width=12,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.delay_entry = tk.Entry(
            delay_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.delay_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.delay_entry.insert(0, str(default_delay))

        tk.Label(
            delay_frame,
            text="点击间隔(秒):",
            width=12,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.click_interval_entry = tk.Entry(
            delay_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.click_interval_entry.pack(side=tk.LEFT)
        self.click_interval_entry.insert(0, str(default_click_interval))

        # 按钮
        btn_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        btn_frame.pack(fill=tk.X, pady=(8, 0))

        ok_btn = ModernButton(
            btn_frame,
            text="确认",
            command=lambda: self.on_ok(x, y),
            button_type="primary",
            width=8,
        )
        ok_btn.pack(side=tk.RIGHT, padx=(6, 0))
        ok_btn.apply_theme(theme)

        cancel_btn = ModernButton(
            btn_frame,
            text="取消",
            command=self.on_cancel,
            button_type="secondary",
            width=8,
        )
        cancel_btn.pack(side=tk.RIGHT)
        cancel_btn.apply_theme(theme)

        # 初始显示正确的操作类型界面
        self.on_action_type_change()

        self.name_entry.focus()
        self.name_entry.select_range(0, tk.END)

    def on_action_type_change(self, event=None):
        """操作类型改变时更新界面"""
        action_type = self.action_type_var.get()

        if action_type == "click":
            self.click_frame.pack(fill=tk.X, pady=4)
            self.keyboard_frame.pack_forget()
        else:  # keyboard
            self.click_frame.pack_forget()
            self.keyboard_frame.pack(fill=tk.X, pady=4)

    def on_ok(self, x, y):
        name = self.name_entry.get().strip() or f"点位({x},{y})"
        action_type = self.action_type_var.get()

        try:
            delay = float(self.delay_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效的延时")
            return

        try:
            click_interval = float(self.click_interval_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效的点击间隔")
            return

        if action_type == "click":
            btn = self.button_var.get()
            try:
                click_count = int(self.click_count_entry.get())
            except Exception:
                messagebox.showerror("错误", "请输入有效的点击次数")
                return
            keys = ""
        else:  # keyboard
            btn = "keyboard"
            click_count = 1
            keys = self.keys_var.get().strip()
            if not keys:
                messagebox.showerror("错误", "请输入按键组合")
                return

        # 统一返回9个值，与EditPointDialog保持一致
        self.result = (
            x,
            y,
            name,
            btn,
            delay,
            click_count,
            click_interval,
            keys,
            action_type,
        )
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()


class EditPointDialog:
    def __init__(
        self,
        parent,
        x,
        y,
        name,
        button,
        delay,
        click_count,
        click_interval,
        keys="",
        action_type="click",
        theme="light",
    ):
        self.top = tk.Toplevel(parent)
        self.top.title("编辑点位")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.result = None
        self.theme = theme
        colors = ModernStyles.get_theme(theme)

        # 使用传入的参数
        self.keys = keys
        self.action_type = action_type

        # 应用主题
        self.top.configure(bg=colors["bg"])

        # 居中显示
        self.top.geometry(
            "+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50)
        )

        main_frame = GlassFrame(self.top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        main_frame.apply_glass_effect(theme)

        # 名称
        name_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        name_frame.pack(fill=tk.X, pady=4)
        tk.Label(
            name_frame,
            text="名称:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.name_entry = tk.Entry(
            name_frame, bg=colors["input_bg"], fg=colors["text"], relief="flat"
        )
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.name_entry.insert(0, name)

        # 坐标设置
        coord_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        coord_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            coord_frame,
            text="X坐标:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.x_entry = tk.Entry(
            coord_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.x_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.x_entry.insert(0, str(x))

        tk.Label(
            coord_frame,
            text="Y坐标:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.y_entry = tk.Entry(
            coord_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.y_entry.pack(side=tk.LEFT)
        self.y_entry.insert(0, str(y))

        # 操作类型和按键设置
        action_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        action_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            action_frame,
            text="操作类型:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.action_type_var = tk.StringVar(value=action_type)
        action_combo = ttk.Combobox(
            action_frame,
            textvariable=self.action_type_var,
            values=["click", "keyboard"],
            state="readonly",
            width=10,
            style="Modern.TCombobox",
        )
        action_combo.pack(side=tk.LEFT, padx=(0, 10))
        action_combo.bind("<<ComboboxSelected>>", self.on_action_type_change)

        # 点击键设置
        self.click_frame = tk.Frame(main_frame, bg=colors["card_bg"])

        tk.Label(
            self.click_frame,
            text="点击键:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.button_var = tk.StringVar(value=button if button != "keyboard" else "left")
        ttk.Combobox(
            self.click_frame,
            textvariable=self.button_var,
            values=["left", "right", "middle"],
            state="readonly",
            width=8,
            style="Modern.TCombobox",
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(
            self.click_frame,
            text="点击次数:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.click_count_entry = tk.Entry(
            self.click_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.click_count_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.click_count_entry.insert(0, str(click_count))

        # 键盘操作设置
        self.keyboard_frame = tk.Frame(main_frame, bg=colors["card_bg"])

        tk.Label(
            self.keyboard_frame,
            text="按键组合:",
            width=8,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.keys_var = tk.StringVar(value=keys)
        keys_entry = ttk.Entry(
            self.keyboard_frame,
            textvariable=self.keys_var,
            width=20,
            style="Modern.TCombobox",
        )
        keys_entry.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(
            self.keyboard_frame,
            text="例如: ctrl+a, ctrl+c, ctrl+v",
            font=("Segoe UI", 8),
            bg=colors["card_bg"],
            fg=colors["text_light"],
        ).pack(side=tk.LEFT)

        # 延时设置
        delay_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        delay_frame.pack(fill=tk.X, pady=4)

        tk.Label(
            delay_frame,
            text="点位延时(秒):",
            width=12,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.delay_entry = tk.Entry(
            delay_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.delay_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.delay_entry.insert(0, str(delay))

        tk.Label(
            delay_frame,
            text="点击间隔(秒):",
            width=12,
            anchor="w",
            bg=colors["card_bg"],
            fg=colors["text"],
        ).pack(side=tk.LEFT)
        self.click_interval_entry = tk.Entry(
            delay_frame,
            width=8,
            bg=colors["input_bg"],
            fg=colors["text"],
            relief="flat",
        )
        self.click_interval_entry.pack(side=tk.LEFT)
        self.click_interval_entry.insert(0, str(click_interval))

        # 按钮
        btn_frame = tk.Frame(main_frame, bg=colors["card_bg"])
        btn_frame.pack(fill=tk.X, pady=(8, 0))

        ok_btn = ModernButton(
            btn_frame,
            text="确定",
            command=self.on_ok,
            button_type="primary",
            width=8,
        )
        ok_btn.pack(side=tk.RIGHT, padx=(6, 0))
        ok_btn.apply_theme(theme)

        cancel_btn = ModernButton(
            btn_frame,
            text="取消",
            command=self.on_cancel,
            button_type="secondary",
            width=8,
        )
        cancel_btn.pack(side=tk.RIGHT)
        cancel_btn.apply_theme(theme)

        # 初始显示正确的操作类型界面
        self.on_action_type_change()

        self.name_entry.focus()
        self.name_entry.select_range(0, tk.END)

    def on_action_type_change(self, event=None):
        """操作类型改变时更新界面"""
        action_type = self.action_type_var.get()

        if action_type == "click":
            self.click_frame.pack(fill=tk.X, pady=4)
            self.keyboard_frame.pack_forget()
        else:  # keyboard
            self.click_frame.pack_forget()
            self.keyboard_frame.pack(fill=tk.X, pady=4)

    def on_ok(self):
        try:
            nx = int(self.x_entry.get())
            ny = int(self.y_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效整数坐标")
            return

        nname = self.name_entry.get().strip() or f"点位({nx},{ny})"
        action_type = self.action_type_var.get()

        try:
            ndelay = float(self.delay_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效延时")
            return

        try:
            nclick_interval = float(self.click_interval_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效点击间隔")
            return

        if action_type == "click":
            nbutton = self.button_var.get()
            try:
                nclick_count = int(self.click_count_entry.get())
            except Exception:
                messagebox.showerror("错误", "请输入有效点击次数")
                return
            nkeys = ""
        else:  # keyboard
            nbutton = "keyboard"
            nclick_count = 1
            nkeys = self.keys_var.get().strip()
            if not nkeys:
                messagebox.showerror("错误", "请输入按键组合")
                return

        self.result = (
            nx,
            ny,
            nname,
            nbutton,
            ndelay,
            nclick_count,
            nclick_interval,
            nkeys,
            action_type,
        )
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()


# ---------------- Main run ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = AutoClickerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_exit)
    root.mainloop()
