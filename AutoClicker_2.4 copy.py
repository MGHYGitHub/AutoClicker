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
VERSION = "2.4"
# If you have a URL to check updates, set it here.
# For safety, by default it's empty; update_check won't run if empty.
UPDATE_CHECK_URL = ""  # e.g. "https://example.com/autoclicker/version.json"

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
    print(line, end='')

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

# ------------- Main App -------------
class AutoClickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"高级自动连点器 {VERSION}")
        self.root.geometry("1000x700")
        self.root.minsize(950, 600)
        # 添加这个设置变量
        self.show_confirmation_var = tk.BooleanVar(value=True)  # 默认显示确认窗口
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
            'total_click_attempts': 0,
            'successful_clicks': 0,
            'failed_clicks': 0,
            'start_time': None,
            'end_time': None,
            'loops_completed': 0
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
        if UPDATE_CHECK_URL:
            threading.Thread(target=self.check_update, daemon=True).start()

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
        settings_menu.add_command(label="主题: 浅色", command=lambda: self.set_theme("light"))
        settings_menu.add_command(label="主题: 深色", command=lambda: self.set_theme("dark"))
        menubar.add_cascade(label="设置", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)  # 新增关于菜单
        help_menu.add_separator()
        help_menu.add_command(label="打开日志文件", command=self.open_log_file)
        help_menu.add_command(label="检查更新", command=self.manual_check_update)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    # ---------------- UI: main layout ----------------
    def create_ui(self):
        pad = 10
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)

        # Left and right area
        left = tk.Frame(main_frame, bg='#ffffff', bd=1, relief=tk.RIDGE)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,8))
        right = tk.Frame(main_frame, bg='#ffffff', bd=1, relief=tk.RIDGE)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(8,0))

        # --- Left: capture & points ---
        self.create_left_panel(left)

        # --- Right: task settings & logs ---
        self.create_right_panel(right)

        # Status bar
        self.status_bar = tk.Label(self.root, text="就绪", anchor=tk.W, font=('Arial', 10), bg='#f0f0f0')
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # Apply initial theme
        self.apply_theme()

    def create_left_panel(self, parent):
        coord_frame = tk.LabelFrame(parent, text="坐标获取 & 点位管理", padx=8, pady=8)
        coord_frame.pack(fill=tk.BOTH, expand=False, padx=8, pady=8)

        instr = tk.Label(coord_frame, text="点击\"开始捕获\"后移动到目标位置，按 F2 记录坐标（支持全局）。双击列表项可编辑。", 
                        wraplength=600)
        instr.pack(fill=tk.X, pady=(0,6))

        # 鼠标键设置和捕获按钮放在同一行
        top_row = tk.Frame(coord_frame)
        top_row.pack(fill=tk.X, pady=4)
        
        # 鼠标键设置
        mb_frame = tk.Frame(top_row)
        mb_frame.pack(side=tk.LEFT)
        tk.Label(mb_frame, text="默认鼠标键:").pack(side=tk.LEFT)
        self.default_button_var = tk.StringVar(value="left")
        mb_combo = ttk.Combobox(mb_frame, textvariable=self.default_button_var, 
                               values=["left","right","middle"], state='readonly', width=8)
        mb_combo.pack(side=tk.LEFT, padx=6)
        mb_combo.bind('<<ComboboxSelected>>', lambda e: self.add_progress_text(f"默认鼠标键设置为 {self.default_button_var.get()}"))

        # 捕获按钮
        btn_frame = tk.Frame(top_row)
        btn_frame.pack(side=tk.RIGHT)
        self.coord_btn = tk.Button(btn_frame, text="开始获取坐标 (F2)", bg='#3498db', fg='white', 
                                  command=self.start_coord_capture, width=15)
        self.coord_btn.pack(side=tk.LEFT)
        self.stop_capture_btn = tk.Button(btn_frame, text="停止获取坐标", bg='#95a5a6', fg='white', 
                                         state=tk.DISABLED, command=self.stop_coord_capture, width=12)
        self.stop_capture_btn.pack(side=tk.LEFT, padx=6)

        # 坐标预览
        coord_preview_frame = tk.Frame(coord_frame)
        coord_preview_frame.pack(fill=tk.X, pady=4)
        self.current_coord_label = tk.Label(coord_preview_frame, text="当前坐标: (0,0)", font=('Arial', 10))
        self.current_coord_label.pack(side=tk.LEFT)
        self.coord_status = tk.Label(coord_preview_frame, text="捕获未启动", fg='#e74c3c', font=('Arial', 10))
        self.coord_status.pack(side=tk.RIGHT)

        # 点位列表区域
        points_frame = tk.LabelFrame(parent, text="点位列表（双击编辑）", padx=6, pady=6)
        points_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        
        # 主内容区域：列表在左，按钮在右
        content_frame = tk.Frame(points_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 列表区域（占据主要空间）
        list_frame = tk.Frame(content_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.points_listbox = tk.Listbox(list_frame, font=('Consolas', 9), activestyle='none')
        self.points_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.points_listbox.bind('<Double-1>', self.edit_selected_point)

        list_scrollbar = tk.Scrollbar(list_frame, command=self.points_listbox.yview)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.points_listbox.config(yscrollcommand=list_scrollbar.set)

        # 按钮区域（右侧竖向排列）
        button_frame = tk.Frame(content_frame)
        button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))
        
        # 按钮分组：操作按钮和移动按钮
        btn_width = 12
        
        # 主要操作按钮
        op_buttons = [
            ("添加当前坐标", self.add_current_point, '#3498db'),
            ("删除选中", self.delete_selected_point, '#e74c3c'),
            ("清空所有", self.clear_all_points, '#e67e22'),
        ]
        
        for text, command, color in op_buttons:
            btn = tk.Button(button_frame, text=text, command=command, 
                           bg=color, fg='white', width=btn_width)
            btn.pack(pady=2, fill=tk.X)
        
        # 分隔线
        separator = ttk.Separator(button_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=6)
        
        # 移动和重命名按钮
        move_buttons = [
            ("上移", lambda: self.move_selected(-1), '#2ecc71'),
            ("下移", lambda: self.move_selected(1), '#2ecc71'),
            ("重命名", self.rename_selected_point, '#9b59b6'),
        ]
        
        for text, command, color in move_buttons:
            btn = tk.Button(button_frame, text=text, command=command, 
                           bg=color, fg='white', width=btn_width)
            btn.pack(pady=2, fill=tk.X)

    def create_right_panel(self, parent):
        settings = tk.LabelFrame(parent, text="任务设置", padx=8, pady=8)
        settings.pack(fill=tk.X, padx=8, pady=8)

        # 使用网格布局更整齐
        settings_grid = tk.Frame(settings)
        settings_grid.pack(fill=tk.X, padx=4, pady=4)
        
        # 第一行：基础设置
        tk.Label(settings_grid, text="默认点击次数:").grid(row=0, column=0, sticky='w', padx=2, pady=2)
        tk.Spinbox(settings_grid, from_=1, to=999, textvariable=self.click_count_var, width=8).grid(row=0, column=1, padx=2, pady=2)
        
        tk.Label(settings_grid, text="默认延时(秒):").grid(row=0, column=2, sticky='w', padx=(10,2), pady=2)
        tk.Spinbox(settings_grid, from_=0.0, to=60.0, increment=0.1, textvariable=self.base_delay_var, width=8).grid(row=0, column=3, padx=2, pady=2)
        
        tk.Label(settings_grid, text="循环次数:").grid(row=0, column=4, sticky='w', padx=(10,2), pady=2)
        tk.Spinbox(settings_grid, from_=1, to=99999, textvariable=self.loop_var, width=8).grid(row=0, column=5, padx=2, pady=2)

        # 第二行：随机设置
        tk.Label(settings_grid, text="随机偏移(px ±):").grid(row=1, column=0, sticky='w', padx=2, pady=2)
        tk.Spinbox(settings_grid, from_=0, to=500, textvariable=self.random_offset_var, width=8).grid(row=1, column=1, padx=2, pady=2)
        
        tk.Label(settings_grid, text="随机延时(s ±):").grid(row=1, column=2, sticky='w', padx=(10,2), pady=2)
        tk.Spinbox(settings_grid, from_=0.0, to=10.0, increment=0.05, textvariable=self.random_delay_var, width=8).grid(row=1, column=3, padx=2, pady=2)
        
        tk.Label(settings_grid, text="启动倒计时(秒):").grid(row=1, column=4, sticky='w', padx=(10,2), pady=2)
        tk.Spinbox(settings_grid, from_=0, to=60, textvariable=self.countdown_var, width=8).grid(row=1, column=5, padx=2, pady=2)

        # 第三行：动作和选项
        tk.Label(settings_grid, text="任务结束动作:").grid(row=2, column=0, sticky='w', padx=2, pady=2)
        ttk.Combobox(settings_grid, textvariable=self.auto_action_var, values=["none", "sound"], 
                    state='readonly', width=8).grid(row=2, column=1, padx=2, pady=2)
        
        # 选项区域
        options_frame = tk.Frame(settings_grid)
        options_frame.grid(row=2, column=2, columnspan=4, sticky='ew', padx=2, pady=2)
        
        tk.Checkbutton(options_frame, text="启动前确认", variable=self.show_confirmation_var).pack(side=tk.LEFT, padx=8)
        tk.Checkbutton(options_frame, text="调试模式", variable=self.debug_mode_var).pack(side=tk.LEFT, padx=8)
        tk.Checkbutton(options_frame, text="安全检测", variable=self.enable_safety_check).pack(side=tk.LEFT, padx=8)
        
        # 安全阈值设置
        safety_frame = tk.Frame(options_frame)
        safety_frame.pack(side=tk.LEFT, padx=8)
        tk.Label(safety_frame, text="阈值:").pack(side=tk.LEFT)
        tk.Spinbox(safety_frame, from_=50, to=1000, textvariable=self.safety_threshold_var, width=5).pack(side=tk.LEFT, padx=2)

        # 控制按钮
        ctrl_frame = tk.Frame(settings)
        ctrl_frame.pack(fill=tk.X, pady=(8,0))
        
        self.start_btn = tk.Button(ctrl_frame, text="启动任务 (F3)", bg='#2ecc71', fg='white', 
                                  command=self.start_task, height=2, font=('Arial', 10, 'bold'))
        self.start_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,4))
        
        self.pause_btn = tk.Button(ctrl_frame, text="暂停 (F4)", bg='#f39c12', fg='white', 
                                  command=self.toggle_pause, state=tk.DISABLED, width=10)
        self.pause_btn.pack(side=tk.LEFT, padx=2)
        
        self.stop_btn = tk.Button(ctrl_frame, text="停止", bg='#e74c3c', fg='white', 
                                 command=self.stop_task, state=tk.DISABLED, width=8)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        
        self.tray_btn = tk.Button(ctrl_frame, text="托盘", command=self.minimize_to_tray, width=6)
        self.tray_btn.pack(side=tk.LEFT, padx=2)

        # 进度和日志区域
        progress_frame = tk.LabelFrame(parent, text="执行进度与日志", padx=8, pady=8)
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # 当前任务信息
        task_frame = tk.Frame(progress_frame)
        task_frame.pack(fill=tk.X, pady=(0,6))
        self.current_task_label = tk.Label(task_frame, text="当前任务: 无", anchor=tk.W, justify=tk.LEFT, font=('Arial', 9))
        self.current_task_label.pack(fill=tk.X)

        # 进度条
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0,6))

        # 统计信息
        stats_frame = tk.Frame(progress_frame)
        stats_frame.pack(fill=tk.X, pady=(0,6))
        self.stats_label = tk.Label(stats_frame, text=self._stats_text(), anchor=tk.W, justify=tk.LEFT, 
                                   fg="#333", font=('Arial', 9))
        self.stats_label.pack(fill=tk.X)

        # 日志文本框
        log_frame = tk.Frame(progress_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_details = tk.Text(log_frame, height=12, state=tk.DISABLED, font=('Consolas', 9))
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
        if theme == "dark":
            # 深色主题配色
            bg = "#1e1e1e"
            panel = "#2d2d2d"
            fg = "#eaeaea"
            entry_bg = "#3d3d3d"
            highlight = "#4a4a4a"
            button_bg = "#404040"
            button_fg = "#ffffff"
        else:
            # 浅色主题配色
            bg = "#f5f5f5"
            panel = "#ffffff"
            fg = "#222222"
            entry_bg = "#ffffff"
            highlight = "#e0e0e0"
            button_bg = "#f0f0f0"
            button_fg = "#000000"

        # 应用主题到所有组件
        self.root.configure(bg=bg)
        
        # 状态栏
        try:
            self.status_bar.configure(bg=bg, fg=fg)
        except Exception:
            pass
        
        # 递归应用主题到所有子组件
        def apply_to_children(widget):
            try:
                if isinstance(widget, (tk.Frame, tk.LabelFrame)):
                    widget.configure(bg=panel)
                elif isinstance(widget, tk.Label):
                    widget.configure(bg=panel, fg=fg)
                elif isinstance(widget, tk.Button):
                    # 保留特殊按钮的颜色，只修改普通按钮
                    current_bg = widget.cget('bg')
                    if current_bg in ['SystemButtonFace', '#f0f0f0', '#d9d9d9']:
                        widget.configure(bg=button_bg, fg=button_fg)
                elif isinstance(widget, tk.Entry):
                    widget.configure(bg=entry_bg, fg=fg, insertbackground=fg)
                elif isinstance(widget, tk.Listbox):
                    widget.configure(bg=entry_bg, fg=fg, selectbackground=highlight)
                elif isinstance(widget, tk.Text):
                    widget.configure(bg=entry_bg, fg=fg, insertbackground=fg)
                elif isinstance(widget, ttk.Combobox):
                    # ttk 组件需要特殊处理
                    style = ttk.Style()
                    if theme == "dark":
                        style.configure("TCombobox", fieldbackground=entry_bg, background=entry_bg, foreground=fg)
                    else:
                        style.configure("TCombobox", fieldbackground=entry_bg, background=entry_bg, foreground=fg)
            except Exception:
                pass
            
            for child in widget.winfo_children():
                apply_to_children(child)
        
        for child in self.root.winfo_children():
            apply_to_children(child)

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
        self.coord_status.config(text="捕获模式已启动 - 按 F2 或窗口内'添加当前坐标'记录位置", fg='#2ecc71')
        self.add_progress_text("坐标捕获模式已启动")
        log("坐标捕获模式启动")

    def stop_coord_capture(self):
        self.is_capturing = False
        self.coord_btn.config(state=tk.NORMAL)
        self.stop_capture_btn.config(state=tk.DISABLED)
        self.coord_status.config(text="捕获模式已停止", fg='#e74c3c')
        self.add_progress_text("坐标捕获模式已停止")
        log("坐标捕获模式停止")

    # ---------------- Points management ----------------
    def add_point(self, x, y, name=None, button=None, delay=None, click_count=None, click_interval=None):
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
            'x': int(x), 
            'y': int(y), 
            'name': str(name), 
            'button': button, 
            'delay': float(delay),
            'click_count': int(click_count),
            'click_interval': float(click_interval)
        }
        self.click_points.append(p)
        self.update_points_list()
        self.add_progress_text(f"添加点位: {name} ({x},{y}) 键:{button} 延时:{delay} 点击:{click_count}次 间隔:{click_interval}")
        log(f"添加点位 {name} ({x},{y})")

    # 添加缺失的方法 - 在 create_left_panel 之前
    def add_current_point(self):
        try:
            x, y = pyautogui.position()
            # 弹出对话框选择按钮、延时、点击次数和点击间隔
            dlg = AddPointDialog(self.root, x, y, 
                            default_button=self.default_button_var.get(), 
                            default_delay=self.base_delay_var.get(),
                            default_click_count=self.click_count_var.get(),
                            default_click_interval=self.base_delay_var.get())
            self.root.wait_window(dlg.top)
            if dlg.result:
                name, btn, delay, click_count, click_interval = dlg.result
                self.add_point(x, y, name=name, button=btn, delay=delay, 
                            click_count=click_count, click_interval=click_interval)
        except Exception as e:
            messagebox.showerror("错误", f"无法获取当前鼠标坐标: {e}")

    def update_points_list(self):
        self.points_listbox.delete(0, tk.END)
        for i, p in enumerate(self.click_points):
            display = f"{i+1}. {p['name']} - ({p['x']},{p['y']}) [{p['button']}] 延时:{p['delay']} 点击:{p['click_count']}次 间隔:{p['click_interval']}"
            self.points_listbox.insert(tk.END, display)

    def update_points_list(self):
        self.points_listbox.delete(0, tk.END)
        for i, p in enumerate(self.click_points):
            display = f"{i+1}. {p['name']} - ({p['x']},{p['y']}) [{p['button']}] 延时:{p['delay']} 点击:{p['click_count']}次 间隔:{p['click_interval']}"
            self.points_listbox.insert(tk.END, display)

    def delete_selected_point(self):
        sel = self.points_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "请先选择点位")
            return
        idx = sel[0]
        removed = self.click_points.pop(idx)
        self.update_points_list()
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
            self.add_progress_text(f"已清空 {cnt} 个点位")
            log("清空所有点位")

    def move_selected(self, delta):
        sel = self.points_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "请先选择点位")
            return
        idx = sel[0]
        new = idx + delta
        if new < 0 or new >= len(self.click_points):
            return
        self.click_points[idx], self.click_points[new] = self.click_points[new], self.click_points[idx]
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
        p = self.click_points[idx]
        new_name = simpledialog.askstring("重命名点位", "输入新的点位名称：", initialvalue=p['name'], parent=self.root)
        if new_name:
            old = p['name']
            self.click_points[idx]['name'] = new_name
            self.update_points_list()
            self.add_progress_text(f"点位重命名: {old} -> {new_name}")
            log(f"点位重命名 {old} -> {new_name}")

    def edit_selected_point(self, event=None):
        sel = self.points_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        p = self.click_points[idx]
        dlg = EditPointDialog(self.root, p['x'], p['y'], p['name'], p['button'], p['delay'], p['click_count'], p['click_interval'])
        self.root.wait_window(dlg.top)
        if dlg.result:
            nx, ny, nname, nbutton, ndelay, nclick_count, nclick_interval = dlg.result
            self.click_points[idx] = {
                'x': int(nx), 
                'y': int(ny), 
                'name': nname, 
                'button': nbutton, 
                'delay': float(ndelay),
                'click_count': int(nclick_count),
                'click_interval': float(nclick_interval)
            }
            self.update_points_list()
            self.add_progress_text(f"点位已编辑: {nname} ({nx},{ny}) 键:{nbutton} 延时:{ndelay} 点击:{nclick_count}次 间隔:{nclick_interval}")
            log(f"点位编辑 -> {nname} ({nx},{ny})")

    # ---------------- Import/Export/Save/Load ----------------
    def import_points(self):
        path = filedialog.askopenfilename(title="导入点位配置", filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                # support both old list-of-lists and new list-of-dicts
                new_list = []
                for item in data:
                    if isinstance(item, list) and len(item) >= 3:
                        x = int(item[0]); y = int(item[1]); name = str(item[2])
                        new_list.append({
                            'x':x,'y':y,'name':name,'button':self.default_button_var.get(),
                            'delay':self.base_delay_var.get(),'click_count':self.click_count_var.get(),
                            'click_interval':self.base_delay_var.get()
                        })
                    elif isinstance(item, dict):
                        # ensure keys exist
                        x = int(item.get('x',0)); y = int(item.get('y',0))
                        name = str(item.get('name', f"点位({x},{y})"))
                        button = item.get('button', self.default_button_var.get())
                        delay = float(item.get('delay', self.base_delay_var.get()))
                        click_count = int(item.get('click_count', self.click_count_var.get()))
                        click_interval = float(item.get('click_interval', self.base_delay_var.get()))
                        new_list.append({
                            'x':x,'y':y,'name':name,'button':button,'delay':delay,
                            'click_count':click_count,'click_interval':click_interval
                        })
                self.click_points = new_list
                self.update_points_list()
                self.add_progress_text(f"已导入 {len(self.click_points)} 个点位 ({os.path.basename(path)})")
                log(f"导入配置 {path}")
            else:
                messagebox.showerror("错误", "不支持的配置格式")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")

    def export_points(self):
        path = filedialog.asksaveasfilename(title="导出点位配置", defaultextension=".json", filetypes=[("JSON 文件", "*.json")])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
            self.add_progress_text(f"已导出到 {os.path.basename(path)}")
            log(f"导出配置 {path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def save_points(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", f"已保存 {len(self.click_points)} 个点位到 {CONFIG_FILE}")
            self.add_progress_text("点位配置已保存")
            log(f"保存配置 {CONFIG_FILE}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def save_points_as(self):
        path = filedialog.asksaveasfilename(title="另存为", defaultextension=".json", filetypes=[("JSON 文件", "*.json")])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.click_points, f, ensure_ascii=False, indent=2)
            self.add_progress_text(f"已另存为 {os.path.basename(path)}")
            log(f"另存为 {path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def load_points(self, quiet=False):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # normalize
                new_list = []
                for item in data:
                    if isinstance(item, dict):
                        x = int(item.get('x',0)); y = int(item.get('y',0))
                        name = item.get('name', f"点位({x},{y})")
                        button = item.get('button', self.default_button_var.get())
                        delay = float(item.get('delay', self.base_delay_var.get()))
                        click_count = int(item.get('click_count', self.click_count_var.get()))
                        click_interval = float(item.get('click_interval', self.base_delay_var.get()))
                        new_list.append({
                            'x':x,'y':y,'name':name,'button':button,'delay':delay,
                            'click_count':click_count,'click_interval':click_interval
                        })
                    elif isinstance(item, list):
                        x = int(item[0]); y = int(item[1]); name = str(item[2])
                        new_list.append({
                            'x':x,'y':y,'name':name,'button':self.default_button_var.get(),
                            'delay':self.base_delay_var.get(),'click_count':self.click_count_var.get(),
                            'click_interval':self.base_delay_var.get()
                        })
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
            summary = (f"点位数量: {len(self.click_points)}\n"
                    f"基础延时: {base_delay} 秒\n"
                    f"随机偏移: ±{offset_px} px\n"
                    f"随机延时: ±{rand_delay} 秒\n"
                    f"循环次数: {loop_count}\n"
                    f"启动倒计时: {countdown} 秒\n"
                    f"任务结束动作: {self.auto_action_var.get()}")
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
        self.start_btn.config(state=tk.DISABLED, text="任务运行中...", bg='#a67c00')
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.add_progress_text("任务启动")
        log("任务启动")

        # reset stats
        self.stats = {
            'total_click_attempts': 0,
            'successful_clicks': 0,
            'failed_clicks': 0,
            'start_time': datetime.datetime.now().isoformat(),
            'end_time': None,
            'loops_completed': 0
        }
        self.update_stats_display()

        # run task thread
        self.task_thread = threading.Thread(target=self.run_click_task, args=(loop_count, offset_px, rand_delay), daemon=True)
        self.task_thread.start()

    def stop_task(self):
        if not self.is_running:
            return
        self.stop_event.set()
        self.pause_event.clear()
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL, text="启动任务", bg='#2ecc71')
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
        total_points = len(self.click_points)
        total_ops = sum(p['click_count'] for p in self.click_points) * loop_count
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

                    # 等待暂停（常规暂停）
                    while self.pause_event.is_set() and not self.stop_event.is_set():
                        self.status_bar.config(text="已暂停（按 F4 继续）")
                        time.sleep(0.1)
                    
                    if self.stop_event.is_set():
                        break

                    # 记录当前目标点位，用于安全检测
                    self.current_target_point = (p['x'], p['y'])
                    
                    # highlight
                    self.root.after(0, lambda i=idx: self.highlight_point(i))
                    cur_info = f"当前: 循环 {loop+1}/{loop_count} — 点位 {idx+1}/{total_points} - {p['name']} ({p['x']},{p['y']}) 键:{p['button']} 点击:{p['click_count']}次 间隔:{p['click_interval']}"
                    self.root.after(0, lambda t=cur_info: self.current_task_label.config(text=t))

                    # 关键修改：如果是调试模式，在每个点位开始前暂停等待F4继续
                    if self.debug_mode_var.get():
                        self.add_progress_text(f"调试模式：点位 {idx+1} 准备就绪，按 F4 继续...")
                        self.status_bar.config(text=f"调试模式：点位 {idx+1} 就绪，按 F4 继续")
                        self.pause_event.set()  # 自动暂停等待用户确认
                        
                        # 等待用户按F4继续
                        while self.pause_event.is_set() and not self.stop_event.is_set():
                            time.sleep(0.1)
                        
                        if self.stop_event.is_set():
                            break
                    
                    # 执行点击操作
                    click_count = p.get('click_count', 1)
                    for c in range(click_count):
                        if self.stop_event.is_set():
                            break
                        # apply random offset
                        rx = p['x']; ry = p['y']
                        if offset_px > 0:
                            dx = random.randint(-offset_px, offset_px)
                            dy = random.randint(-offset_px, offset_px)
                            rx = p['x'] + dx
                            ry = p['y'] + dy

                        # attempt click
                        try:
                            btn = p.get('button', 'left')
                            pyautogui.click(rx, ry, button=btn)
                            self.stats['successful_clicks'] += 1
                        except Exception as e:
                            self.stats['failed_clicks'] += 1
                            self.add_progress_text(f"点击失败: {e}，任务停止")
                            log(f"点击异常: {e}")
                            self.stop_event.set()
                            break
                        finally:
                            self.stats['total_click_attempts'] += 1

                        # 记录点击时间，用于安全检测判断
                        self.last_click_time = time.time()

                        ops_done += 1
                        progress = (ops_done / total_ops) * 100 if total_ops > 0 else 100
                        self.root.after(0, lambda p=progress: self.progress_var.set(p))
                        self.add_progress_text(f"循环{loop+1}/{loop_count} - 点位{idx+1}/{total_points} - 点击{c+1}/{click_count} - 偏移({rx-p['x']},{ry-p['y']})")
                        log(f"点击: {p['name']} ({rx},{ry}) 按钮:{btn}")

                        # 使用点位独立的点击间隔延时
                        click_interval = p.get('click_interval', 0.1)
                        if rand_delay > 0:
                            click_interval = click_interval + random.uniform(-rand_delay, rand_delay)
                            if click_interval < 0:
                                click_interval = 0
                        if c < click_count - 1 and not self.stop_event.is_set():
                            time.sleep(click_interval)

                    # 清除目标点位记录
                    self.current_target_point = None
                    
                    # interval between points
                    if idx < total_points - 1 and not self.stop_event.is_set():
                        time.sleep(p.get('delay', 1.0))

                # end of loop
                self.stats['loops_completed'] += 1

                # loop interval
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
            self.stats['end_time'] = datetime.datetime.now().isoformat()
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
        self.start_btn.config(state=tk.NORMAL, text="启动任务 (F3)", bg='#2ecc71')
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
        self.root.bind('<F2>', lambda e: self.local_capture())
        self.root.bind('<F3>', lambda e: self._hotkey_toggle_start_stop()())
        self.root.bind('<F4>', lambda e: self.toggle_pause())

        # global keyboard hooks via keyboard lib if available
        if HAVE_KEYBOARD:
            try:
                # 添加全局F2快捷键，但使用防重复机制
                keyboard.add_hotkey('f2', self.global_capture_coord)
                keyboard.add_hotkey('f3', lambda: self._hotkey_toggle_start_stop()())
                keyboard.add_hotkey('f4', lambda: self.toggle_pause())
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
                self.coord_status.config(text=f"已记录坐标: ({x}, {y})", fg='#2ecc71')
                # 在主线程中更新状态
                self.root.after(0, lambda: self.coord_status.config(text=f"已记录坐标: ({x}, {y})", fg='#2ecc71'))
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
                self.coord_status.config(text=f"已记录坐标: ({x}, {y})", fg='#2ecc71')
            except Exception as e:
                log(f"本地捕获失败: {e}")

    # ---------------- Safety: mouse movement detection ----------------
    def check_mouse_movement(self):
        try:
            cur = pyautogui.position()
            prev = self.last_mouse_pos
            dx = cur[0] - prev[0]; dy = cur[1] - prev[1]
            dist = math.hypot(dx, dy)
            
            current_time = time.time()
            
            # 使用UI中设置的动态阈值
            threshold = self.safety_threshold_var.get()
            
            # 只有在任务运行中且未暂停时才检查安全
            if (self.enable_safety_check.get() and 
                self.is_running and 
                not self.pause_event.is_set()):
                
                # 检查是否是任务移动
                is_task_movement = False
                
                # 方法1：检查是否在点击后的短时间内
                if hasattr(self, 'last_click_time'):
                    time_since_last_click = current_time - self.last_click_time
                    if time_since_last_click < 3.0:  # 点击后3秒内认为是任务移动
                        is_task_movement = True
                
                # 方法2：检查鼠标是否正在向目标点位移动
                if (hasattr(self, 'current_target_point') and 
                    self.current_target_point is not None):
                    target_x, target_y = self.current_target_point
                    # 计算到目标点的距离
                    to_target_dist = math.hypot(cur[0]-target_x, cur[1]-target_y)
                    # 如果鼠标正在接近目标点，认为是任务移动
                    if to_target_dist < 100:  # 距离目标点100像素内
                        is_task_movement = True
                
                # 只有当不是任务移动且移动距离超过阈值时才触发安全暂停
                if (dist > threshold and 
                    not is_task_movement and
                    current_time - self.last_safety_trigger > 5):  # 5秒冷却
                    
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
        help_text = (
            "使用说明：\n"
            "- 点击\"开始获取坐标\"进入捕获模式，然后将鼠标移动到目标位置按 F2 记录坐标（支持全局）。\n"
            "- 双击点位可编辑坐标/名称/按钮/延时/点击次数/点击间隔；选中点位可上移/下移/删除/重命名。\n"
            "- 配置好参数后点击\"启动任务\"；开始前会显示任务摘要确认并可倒计时。\n"
            "- 快捷键：F2(记录坐标), F3(开始/停止), F4(暂停/继续)（全局需 keyboard 支持）\n"
            "- 新增安全机制：若程序检测到鼠标快速移动（例如手动操作），会自动暂停任务以防误点。\n"
            "- 支持托盘最小化（最小化后右键托盘可恢复/停止/退出），支持浅/深主题切换。\n"
            "- 点击统计会写入 click_log.txt（可通过文件菜单导出）。\n"
            "- 新增点位独立设置：每个点位可独立设置点击次数和点击间隔延时。"
        )
        messagebox.showinfo("使用说明", help_text)

    def open_log_file(self):
        if os.path.exists(LOG_FILE):
            try:
                if os.name == 'nt':
                    os.startfile(LOG_FILE)
                elif os.name == 'posix':
                    if os.system("which xdg-open >/dev/null 2>&1") == 0:
                        os.system(f"xdg-open {LOG_FILE} >/dev/null 2>&1 &")
                    elif os.system("which open >/dev/null 2>&1") == 0:
                        os.system(f"open {LOG_FILE} >/dev/null 2>&1 &")
                else:
                    messagebox.showinfo("日志文件", f"日志保存在: {os.path.abspath(LOG_FILE)}")
            except Exception:
                messagebox.showinfo("日志文件", f"日志保存在: {os.path.abspath(LOG_FILE)}")
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
        if not s or s['start_time'] is None:
            return "统计：尚未开始任务"
        start = s.get('start_time', 'N/A')
        end = s.get('end_time', 'N/A')
        return (f"统计：尝试 {s['total_click_attempts']} 次，成功 {s['successful_clicks']}，失败 {s['failed_clicks']}，"
                f"循环完成 {s['loops_completed']}，开始 {start}，结束 {end}")

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
            content.append(f"参数: base_delay={self.base_delay_var.get()}, loop={self.loop_var.get()}, random_offset={self.random_offset_var.get()}, random_delay={self.random_delay_var.get()}")
            content.append("点位列表:")
            for i, p in enumerate(self.click_points, 1):
                content.append(f"  {i}. {p['name']} ({p['x']},{p['y']}) 按键:{p['button']} 延时:{p['delay']} 点击:{p['click_count']}次 间隔:{p['click_interval']}")
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
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if not path:
            return
        try:
            with open(CLICK_REPORT, 'r', encoding='utf-8') as src, open(path, 'w', encoding='utf-8') as dst:
                dst.write(src.read())
            messagebox.showinfo("导出成功", f"已导出报告到 {path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    # ---------------- Tray integration ----------------
    def minimize_to_tray(self):
        if not HAVE_PYSTRAY:
            messagebox.showwarning("托盘支持不可用", "未检测到 pystray 或 Pillow，无法最小化到托盘。")
            return
        # hide window and start tray in background thread
        self.hide_window_to_tray()

    def hide_window_to_tray(self):
        # create icon image
        img = Image.new('RGB', (64, 64), color=(30, 144, 255))
        d = ImageDraw.Draw(img)
        d.ellipse((10, 10, 54, 54), fill='white')
        d.text((18, 18), "AC", fill='black')
        menu = Menu(
            MenuItem("打开主窗口", lambda icon, item: self.show_from_tray()),
            MenuItem("开始任务", lambda icon, item: self.start_task()),
            MenuItem("暂停/继续", lambda icon, item: self.toggle_pause()),
            MenuItem("停止任务", lambda icon, item: self.stop_task()),
            MenuItem("退出", lambda icon, item: self.tray_exit())
        )
        self.tray_icon = Icon("AutoClicker", img, f"AutoClicker {VERSION}", menu)
        # withdraw window
        self.root.withdraw()
        # run tray icon in separate thread
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
    def check_update(self):
        if not HAVE_REQUESTS or not UPDATE_CHECK_URL:
            return
        try:
            resp = requests.get(UPDATE_CHECK_URL, timeout=6)
            if resp.status_code == 200:
                data = resp.json()
                remote_version = data.get('version')
                url = data.get('url')
                if remote_version and remote_version != VERSION:
                    # notify user in UI thread
                    def notify():
                        if messagebox.askyesno("检测到新版本", f"发现新版本 {remote_version}，是否打开下载页面？"):
                            try:
                                import webbrowser
                                webbrowser.open(url or UPDATE_CHECK_URL)
                            except Exception:
                                pass
                    self.root.after(0, notify)
        except Exception as e:
            log(f"检查更新失败: {e}")

    def manual_check_update(self):
        if not UPDATE_CHECK_URL:
            messagebox.showinfo("检查更新", "未配置更新检查 URL。")
            return
        threading.Thread(target=self.check_update, daemon=True).start()
        messagebox.showinfo("检查更新", "正在检查更新（后台），如有新版会提示。")

    # ---------------- Exit ----------------
    def on_exit(self):
        if self.is_running:
            if not messagebox.askyesno("退出确认", "任务正在运行，确定退出并停止任务吗？"):
                return
        # set stop
        self.stop_event.set()
        # save current points
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
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
        
        # 居中显示
        about_window.geometry("+%d+%d" % (self.root.winfo_rootx()+200, self.root.winfo_rooty()+150))
        
        main_frame = tk.Frame(about_window, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = tk.Label(main_frame, text=f"AutoClicker {VERSION}", 
                            font=('Arial', 16, 'bold'), fg='#2c3e50')
        title_label.pack(pady=(0, 10))
        
        # 描述
        desc_text = ("高级自动连点器\n"
                    "支持多点位管理、独立点击设置、安全检测等功能")
        desc_label = tk.Label(main_frame, text=desc_text, 
                            font=('Arial', 11), fg='#34495e')
        desc_label.pack(pady=(0, 20))
        
        # 作者信息
        author_frame = tk.LabelFrame(main_frame, text="作者信息", padx=10, pady=10)
        author_frame.pack(fill=tk.X, pady=(0, 15))
        
        author_info = [
            ("作者", "Eli_Morgan"),
            ("版本", VERSION),
            ("开发语言", "Python 3"),
            ("界面库", "Tkinter")
        ]
        
        for i, (label, value) in enumerate(author_info):
            row_frame = tk.Frame(author_frame)
            row_frame.pack(fill=tk.X, pady=2)
            tk.Label(row_frame, text=f"{label}:", width=8, anchor='w', 
                    font=('Arial', 9, 'bold')).pack(side=tk.LEFT)
            tk.Label(row_frame, text=value, anchor='w', 
                    font=('Arial', 9)).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # GitHub链接 - Markdown风格
        github_frame = tk.Frame(author_frame)
        github_frame.pack(fill=tk.X, pady=2)
        tk.Label(github_frame, text="项目地址:", width=8, anchor='w', 
                font=('Arial', 9, 'bold')).pack(side=tk.LEFT)

        # 创建类似Markdown链接的样式
        github_link = tk.Label(github_frame, text="GitHub", 
                            font=('Arial', 9), fg='#0366d6', 
                            cursor='hand2', underline=True)
        github_link.pack(side=tk.LEFT)
        github_link.bind('<Button-1>', lambda e: self.open_github())

        # 在链接后面显示URL（灰色，较小字体）
        # url_label = tk.Label(github_frame, text="(https://github.com/MGHYGitHub/AutoClicker)",
        #                     font=('Arial', 8), fg='#6a737d')
        # url_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # 联系方式
        contact_frame = tk.LabelFrame(main_frame, text="联系方式", padx=10, pady=10)
        contact_frame.pack(fill=tk.X, pady=(0, 15))
        
        contacts = [
            ("邮箱", "Eli_Morgan2025@outlook.com")
        ]
        
        for i, (label, value) in enumerate(contacts):
            row_frame = tk.Frame(contact_frame)
            row_frame.pack(fill=tk.X, pady=2)
            tk.Label(row_frame, text=f"{label}:", width=8, anchor='w', 
                    font=('Arial', 9, 'bold')).pack(side=tk.LEFT)
            tk.Label(row_frame, text=value, anchor='w', 
                    font=('Arial', 9)).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 特别感谢
        thanks_frame = tk.LabelFrame(main_frame, text="特别感谢", padx=10, pady=10)
        thanks_frame.pack(fill=tk.X, pady=(0, 15))
        
        thanks_text = ("• PyAutoGUI - 鼠标键盘自动化\n"
                    "• Keyboard - 全局快捷键支持\n"
                    "• PyStray - 系统托盘集成\n"
                    "• Pillow - 图像处理支持\n"
                    "• 所有贡献者和用户")
        thanks_label = tk.Label(thanks_frame, text=thanks_text, 
                            justify=tk.LEFT, font=('Arial', 9))
        thanks_label.pack(anchor='w')
        
        # 版权信息
        copyright_frame = tk.Frame(main_frame)
        copyright_frame.pack(fill=tk.X, pady=(10, 0))
        
        copyright_text = "© 2025 AutoClicker. 保留所有权利。"
        copyright_label = tk.Label(copyright_frame, text=copyright_text, 
                                font=('Arial', 8), fg='#7f8c8d')
        copyright_label.pack()
        
        # 关闭按钮
        close_btn = tk.Button(main_frame, text="关闭", 
                            command=about_window.destroy,
                            bg='#3498db', fg='white', width=10)
        close_btn.pack(pady=(10, 0))
        
        # 绑定ESC键关闭
        about_window.bind('<Escape>', lambda e: about_window.destroy())
        
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

# ---------------- Dialogs for adding/editing points ----------------
class AddPointDialog:
    def __init__(self, parent, x, y, default_button='left', default_delay=1.0, default_click_count=1, default_click_interval=0.1):
        self.top = tk.Toplevel(parent)
        self.top.title("添加点位")
        self.top.resizable(False, False)
        self.top.transient(parent)  # 设置为主窗口的子窗口
        self.top.grab_set()  # 模态对话框
        self.result = None
        
        # 居中显示
        self.top.geometry("+%d+%d" % (parent.winfo_rootx()+50, parent.winfo_rooty()+50))

        main_frame = tk.Frame(self.top, padx=12, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text=f"当前位置：({x}, {y})", font=('Arial', 9)).pack(anchor='w', pady=(0,8))
        
        # 名称
        name_frame = tk.Frame(main_frame)
        name_frame.pack(fill=tk.X, pady=4)
        tk.Label(name_frame, text="名称:", width=8, anchor='w').pack(side=tk.LEFT)
        self.name_entry = tk.Entry(name_frame)
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.name_entry.insert(0, f"点位({x},{y})")
        
        # 坐标设置
        coord_frame = tk.Frame(main_frame)
        coord_frame.pack(fill=tk.X, pady=4)
        
        tk.Label(coord_frame, text="X坐标:", width=8, anchor='w').pack(side=tk.LEFT)
        self.x_entry = tk.Entry(coord_frame, width=8)
        self.x_entry.pack(side=tk.LEFT, padx=(0,10))
        self.x_entry.insert(0, str(x))
        
        tk.Label(coord_frame, text="Y坐标:", width=8, anchor='w').pack(side=tk.LEFT)
        self.y_entry = tk.Entry(coord_frame, width=8)
        self.y_entry.pack(side=tk.LEFT)
        self.y_entry.insert(0, str(y))

        # 点击键和延时
        settings_frame = tk.Frame(main_frame)
        settings_frame.pack(fill=tk.X, pady=4)
        
        tk.Label(settings_frame, text="点击键:", width=8, anchor='w').pack(side=tk.LEFT)
        self.button_var = tk.StringVar(value=default_button)
        ttk.Combobox(settings_frame, textvariable=self.button_var, values=['left','right'], 
                    state='readonly', width=8).pack(side=tk.LEFT, padx=(0,10))
        
        tk.Label(settings_frame, text="点位延时(秒):", width=10, anchor='w').pack(side=tk.LEFT)
        self.delay_entry = tk.Entry(settings_frame, width=8)
        self.delay_entry.pack(side=tk.LEFT, padx=(0,10))
        self.delay_entry.insert(0, str(default_delay))

        # 点击次数和点击间隔
        click_frame = tk.Frame(main_frame)
        click_frame.pack(fill=tk.X, pady=4)
        
        tk.Label(click_frame, text="点击次数:", width=8, anchor='w').pack(side=tk.LEFT)
        self.click_count_entry = tk.Entry(click_frame, width=8)
        self.click_count_entry.pack(side=tk.LEFT, padx=(0,10))
        self.click_count_entry.insert(0, str(default_click_count))
        
        tk.Label(click_frame, text="点击间隔(秒):", width=10, anchor='w').pack(side=tk.LEFT)
        self.click_interval_entry = tk.Entry(click_frame, width=8)
        self.click_interval_entry.pack(side=tk.LEFT)
        self.click_interval_entry.insert(0, str(default_click_interval))

        # 按钮
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(8,0))
        tk.Button(btn_frame, text="确认", command=lambda: self.on_ok(x, y), 
                 bg='#2ecc71', fg='white', width=8).pack(side=tk.RIGHT, padx=(6,0))
        tk.Button(btn_frame, text="取消", command=self.on_cancel, 
                 width=8).pack(side=tk.RIGHT)

        self.name_entry.focus()
        self.name_entry.select_range(0, tk.END)

    def on_ok(self, x, y):
        name = self.name_entry.get().strip() or f"点位({x},{y})"
        btn = self.button_var.get()
        try:
            delay = float(self.delay_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效的延时")
            return
        try:
            click_count = int(self.click_count_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效的点击次数")
            return
        try:
            click_interval = float(self.click_interval_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效的点击间隔")
            return
        self.result = (name, btn, delay, click_count, click_interval)
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()

class EditPointDialog:
    def __init__(self, parent, x, y, name, button, delay, click_count, click_interval):
        self.top = tk.Toplevel(parent)
        self.top.title("编辑点位")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        self.result = None
        
        # 居中显示
        self.top.geometry("+%d+%d" % (parent.winfo_rootx()+50, parent.winfo_rooty()+50))

        main_frame = tk.Frame(self.top, padx=12, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 名称
        name_frame = tk.Frame(main_frame)
        name_frame.pack(fill=tk.X, pady=4)
        tk.Label(name_frame, text="名称:", width=8, anchor='w').pack(side=tk.LEFT)
        self.name_entry = tk.Entry(name_frame)
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.name_entry.insert(0, name)

        # 坐标设置
        coord_frame = tk.Frame(main_frame)
        coord_frame.pack(fill=tk.X, pady=4)
        
        tk.Label(coord_frame, text="X坐标:", width=8, anchor='w').pack(side=tk.LEFT)
        self.x_entry = tk.Entry(coord_frame, width=8)
        self.x_entry.pack(side=tk.LEFT, padx=(0,10))
        self.x_entry.insert(0, str(x))
        
        tk.Label(coord_frame, text="Y坐标:", width=8, anchor='w').pack(side=tk.LEFT)
        self.y_entry = tk.Entry(coord_frame, width=8)
        self.y_entry.pack(side=tk.LEFT)
        self.y_entry.insert(0, str(y))

        # 点击键和延时
        settings_frame = tk.Frame(main_frame)
        settings_frame.pack(fill=tk.X, pady=4)
        
        tk.Label(settings_frame, text="点击键:", width=8, anchor='w').pack(side=tk.LEFT)
        self.button_var = tk.StringVar(value=button)
        ttk.Combobox(settings_frame, textvariable=self.button_var, values=['left','right'], 
                    state='readonly', width=8).pack(side=tk.LEFT, padx=(0,10))
        
        tk.Label(settings_frame, text="点位延时(秒):", width=10, anchor='w').pack(side=tk.LEFT)
        self.delay_entry = tk.Entry(settings_frame, width=8)
        self.delay_entry.pack(side=tk.LEFT, padx=(0,10))
        self.delay_entry.insert(0, str(delay))

        # 点击次数和点击间隔
        click_frame = tk.Frame(main_frame)
        click_frame.pack(fill=tk.X, pady=4)
        
        tk.Label(click_frame, text="点击次数:", width=8, anchor='w').pack(side=tk.LEFT)
        self.click_count_entry = tk.Entry(click_frame, width=8)
        self.click_count_entry.pack(side=tk.LEFT, padx=(0,10))
        self.click_count_entry.insert(0, str(click_count))
        
        tk.Label(click_frame, text="点击间隔(秒):", width=10, anchor='w').pack(side=tk.LEFT)
        self.click_interval_entry = tk.Entry(click_frame, width=8)
        self.click_interval_entry.pack(side=tk.LEFT)
        self.click_interval_entry.insert(0, str(click_interval))

        # 按钮
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(8,0))
        tk.Button(btn_frame, text="确定", command=self.on_ok, 
                 bg='#2ecc71', fg='white', width=8).pack(side=tk.RIGHT, padx=(6,0))
        tk.Button(btn_frame, text="取消", command=self.on_cancel, 
                 width=8).pack(side=tk.RIGHT)

        self.name_entry.focus()
        self.name_entry.select_range(0, tk.END)

    def on_ok(self):
        try:
            nx = int(self.x_entry.get())
            ny = int(self.y_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效整数坐标")
            return
        nname = self.name_entry.get().strip() or f"点位({nx},{ny})"
        nbutton = self.button_var.get()
        try:
            ndelay = float(self.delay_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效延时")
            return
        try:
            nclick_count = int(self.click_count_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效点击次数")
            return
        try:
            nclick_interval = float(self.click_interval_entry.get())
        except Exception:
            messagebox.showerror("错误", "请输入有效点击间隔")
            return
        self.result = (nx, ny, nname, nbutton, ndelay, nclick_count, nclick_interval)
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()

# ---------------- Main run ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = AutoClickerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_exit)
    root.mainloop()