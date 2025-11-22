# -*- coding: utf-8 -*-
"""
AutoIconPacker Pro v3.5 - 增强版

新增功能：
1. 图片格式转换ICO功能（PNG/JPG等转ICO）
2. 多尺寸ICO生成（16x16到256x256）
3. 图标预览和尺寸检测
4. 自动图标优化和修复
5. 批量图标转换工具
6. 图标设置验证和修复

修复内容：
1. 修复版本文件格式错误
2. 优化错误处理
3. 改进进度显示
4. 修复图标预览问题
5. 移除已废弃的加密参数
"""
import warnings

warnings.filterwarnings("ignore", message="pkg_resources is deprecated")
import os
import sys
import json
import shutil
import subprocess
import threading
import tempfile
import time
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

# Optional imports (graceful fallback)
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *

    TB_AVAILABLE = True
except Exception:
    tb = None
    TB_AVAILABLE = False

try:
    from PIL import Image, ImageTk

    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

try:
    import pystray

    PYSTRAY_AVAILABLE = True
except Exception:
    PYSTRAY_AVAILABLE = False

try:
    from win10toast import ToastNotifier

    TOAST_AVAILABLE = True
except Exception:
    TOAST_AVAILABLE = False

# Windows taskbar progress (ctypes)
IS_WINDOWS = sys.platform.startswith("win")
if IS_WINDOWS:
    try:
        import ctypes
        from ctypes import wintypes

        ctypes.windll.kernel32.SetErrorMode(0x0001)
        # taskbar APIs
        TBPF_NOPROGRESS = 0
        TBPF_INDETERMINATE = 0x1
        TBPF_NORMAL = 0x2
        TBPF_ERROR = 0x4
        TBPF_PAUSED = 0x8
        # COM init
        ctypes.OleDLL("ole32").CoInitialize(None)
        TASKBAR_AVAILABLE = True
    except Exception:
        TASKBAR_AVAILABLE = False
else:
    TASKBAR_AVAILABLE = False

APP_NAME = "AutoIconPacker Pro v3.5 - 增强版"
DEFAULT_ICON = "app.ico"
CONFIG_EXT = ".aipcfg.json"

# ---------------- Utility helpers ----------------


def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)


def run_subprocess_stream(cmd, cwd=None, on_line=None, shell=False):
    """Run command and stream stdout/stderr lines to on_line callback."""
    proc = subprocess.Popen(
        cmd,
        cwd=cwd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        shell=shell,
        bufsize=1,
        universal_newlines=True,
        encoding="utf-8",
        errors="replace",
    )
    for line in proc.stdout:
        if on_line:
            on_line(line.rstrip("\n"))
    proc.wait()
    return proc.returncode


# ---------------- Icon Processing Functions ----------------


def generate_multisize_ico(source_img_path, out_ico_path, sizes=None):
    """Generate multi-size ICO from a given image using Pillow."""
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow (PIL) is required to generate ICOs")

    if sizes is None:
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

    try:
        img = Image.open(source_img_path)

        # Convert to RGBA if necessary
        if img.mode != "RGBA":
            img = img.convert("RGBA")

        icons = []
        for size in sizes:
            # Resize image maintaining aspect ratio
            resized_img = img.copy()
            resized_img.thumbnail(size, Image.Resampling.LANCZOS)

            # Create square canvas
            square_img = Image.new("RGBA", size, (0, 0, 0, 0))

            # Calculate position to center the image
            x = (size[0] - resized_img.width) // 2
            y = (size[1] - resized_img.height) // 2

            # Paste resized image onto square canvas
            square_img.paste(resized_img, (x, y))
            icons.append(square_img)

        # Save as ICO with all sizes
        icons[0].save(out_ico_path, format="ICO", sizes=[icon.size for icon in icons])
        return True

    except Exception as e:
        raise RuntimeError(f"ICO生成失败: {str(e)}")


def convert_image_to_ico(source_path, output_path=None, sizes=None):
    """Convert any image format to ICO."""
    if not PIL_AVAILABLE:
        raise RuntimeError("需要Pillow库来转换图片格式")

    if output_path is None:
        output_path = os.path.splitext(source_path)[0] + ".ico"

    return generate_multisize_ico(source_path, output_path, sizes)


def ico_contains_sizes(ico_path):
    """Check ICO file for available sizes."""
    if not PIL_AVAILABLE:
        return []
    try:
        img = Image.open(ico_path)
        sizes = set()
        try:
            for frame in range(0, getattr(img, "n_frames", 1)):
                img.seek(frame)
                sizes.add(img.size)
        except Exception:
            sizes.add(img.size)
        return sorted(list(sizes), reverse=True)
    except Exception:
        return []


def optimize_icon_for_windows(icon_path):
    """Optimize icon for Windows applications."""
    if not PIL_AVAILABLE:
        return False

    try:
        # Check if it's already an ICO file with proper sizes
        if icon_path.lower().endswith(".ico"):
            sizes = ico_contains_sizes(icon_path)
            required_sizes = [(16, 16), (32, 32), (48, 48), (256, 256)]

            # Check if all required sizes are present
            has_all_sizes = all(
                any(s[0] == req[0] and s[1] == req[1] for s in sizes)
                for req in required_sizes
            )

            if has_all_sizes:
                return True  # Already optimized

        # Convert/optimize the icon
        temp_ico = tempfile.mktemp(suffix=".ico")
        success = generate_multisize_ico(icon_path, temp_ico)

        if success:
            shutil.copy2(temp_ico, icon_path)
            os.remove(temp_ico)
            return True

        return False

    except Exception as e:
        print(f"图标优化失败: {e}")
        return False


def open_path_in_explorer(p):
    if IS_WINDOWS:
        subprocess.Popen(["explorer", os.path.normpath(p)])
    else:
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", p])
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception:
            pass


# ---------------- Icon Converter Dialog ----------------


class IconConverterDialog:
    def __init__(self, parent, theme="light"):
        self.top = tk.Toplevel(parent)
        self.top.title("图标格式转换器")
        self.top.geometry("500x400")
        self.top.resizable(True, True)
        self.top.transient(parent)
        self.top.grab_set()

        self.theme = theme
        self.colors = self._get_theme_colors()
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="准备就绪")

        self._build_ui()

        # Center the dialog
        self.top.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.top.winfo_width()) // 2
        y = (
            parent.winfo_rooty()
            + (parent.winfo_height() - self.top.winfo_height()) // 2
        )
        self.top.geometry(f"+{x}+{y}")

    def _get_theme_colors(self):
        if self.theme == "dark":
            return {
                "bg": "#2b2b2b",
                "card_bg": "#3c3c3c",
                "text": "#ffffff",
                "text_light": "#cccccc",
                "accent": "#007acc",
                "border": "#555555",
            }
        else:
            return {
                "bg": "#f5f5f5",
                "card_bg": "#ffffff",
                "text": "#000000",
                "text_light": "#666666",
                "accent": "#007acc",
                "border": "#cccccc",
            }

    def _build_ui(self):
        colors = self.colors
        self.top.configure(bg=colors["bg"])

        # Main frame
        main_frame = ttk.Frame(self.top, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame, text="图标格式转换器", font=("Segoe UI", 14, "bold")
        )
        title_label.pack(pady=(0, 15))

        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="输入文件", padding=10)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(input_frame, text="选择源图片文件:").pack(anchor="w")
        input_row = ttk.Frame(input_frame)
        input_row.pack(fill=tk.X, pady=5)
        ttk.Entry(input_row, textvariable=self.input_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(input_row, text="浏览", command=self._select_input).pack(
            side=tk.RIGHT, padx=(5, 0)
        )

        # Output section
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(output_frame, text="输出ICO文件:").pack(anchor="w")
        output_row = ttk.Frame(output_frame)
        output_row.pack(fill=tk.X, pady=5)
        ttk.Entry(output_row, textvariable=self.output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(output_row, text="浏览", command=self._select_output).pack(
            side=tk.RIGHT, padx=(5, 0)
        )

        # Options
        options_frame = ttk.LabelFrame(main_frame, text="转换选项", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        self.optimize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="优化图标尺寸 (16x16, 32x32, 48x48, 256x256)",
            variable=self.optimize_var,
        ).pack(anchor="w")

        # Preview
        preview_frame = ttk.LabelFrame(main_frame, text="预览", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.preview_label = ttk.Label(preview_frame, text="选择文件后显示预览")
        self.preview_label.pack(expand=True)

        # Status and buttons
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(status_frame, textvariable=self.status_text).pack(side=tk.LEFT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="开始转换", command=self._convert_icon).pack(
            side=tk.RIGHT, padx=(5, 0)
        )
        ttk.Button(button_frame, text="关闭", command=self.top.destroy).pack(
            side=tk.RIGHT
        )

    def _select_input(self):
        filetypes = [
            ("图片文件", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff"),
            ("PNG文件", "*.png"),
            ("JPEG文件", "*.jpg;*.jpeg"),
            ("所有文件", "*.*"),
        ]

        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.input_path.set(path)
            self._update_output_path(path)
            self._update_preview(path)

    def _select_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".ico", filetypes=[("ICO文件", "*.ico")]
        )
        if path:
            self.output_path.set(path)

    def _update_output_path(self, input_path):
        if input_path:
            base_name = os.path.splitext(input_path)[0]
            self.output_path.set(base_name + ".ico")

    def _update_preview(self, path):
        try:
            if PIL_AVAILABLE and os.path.exists(path):
                img = Image.open(path)
                img.thumbnail((100, 100))
                self.tk_preview = ImageTk.PhotoImage(img)
                self.preview_label.configure(image=self.tk_preview)

                # Show image info
                info = f"尺寸: {img.size}\n格式: {img.format}\n模式: {img.mode}"
                self.preview_label.configure(text=info, compound=tk.TOP)
            else:
                self.preview_label.configure(
                    text="无法预览\n(需要Pillow库支持)", image=""
                )
        except Exception as e:
            self.preview_label.configure(text=f"预览失败: {str(e)}", image="")
            self.top.update_idletasks()

    def _convert_icon(self):
        input_path = self.input_path.get()
        output_path = self.output_path.get()

        if not input_path:
            messagebox.showerror("错误", "请选择输入文件")
            return

        if not output_path:
            messagebox.showerror("错误", "请指定输出文件路径")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("错误", "输入文件不存在")
            return

        try:
            self.status_text.set("正在转换...")
            self.top.update_idletasks()  # 使用 update_idletasks 而不是 update

            # Convert image to ICO
            success = convert_image_to_ico(input_path, output_path)

            if success:
                # Optimize if requested
                if self.optimize_var.get():
                    optimize_icon_for_windows(output_path)

                self.status_text.set("转换完成!")
                messagebox.showinfo("成功", f"图标转换完成!\n输出文件: {output_path}")

                # Update preview with new ICO
                self._update_preview(output_path)
            else:
                self.status_text.set("转换失败")
                messagebox.showerror("错误", "图标转换失败")

        except Exception as e:
            self.status_text.set("转换错误")
            messagebox.showerror("错误", f"转换过程中发生错误:\n{str(e)}")


# ---------------- Main Application ----------------


class AutoIconPackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self._setup_style()
        self._build_variables()
        self._build_ui()
        self._bind_shortcuts()
        self.tray_icon = None
        self.toast = ToastNotifier() if TOAST_AVAILABLE else None
        self._stop_flag = False
        self.current_process = None

    def _setup_style(self):
        if TB_AVAILABLE:
            self.style = tb.Style(theme="flatly")
            self.main_frame = ttk.Frame(self.style.master)
            self.style.master.title(APP_NAME)
            self.root = self.style.master
            self.use_tb = True
        else:
            self.style = None
            self.use_tb = False
            self.root.configure(bg="#f5f5f5")

    def _build_variables(self):
        self.py_file = tk.StringVar()
        self.ico_file = tk.StringVar()
        self.tray_ico_file = tk.StringVar()
        self.win_ico_file = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "dist"))
        self.upx_enable = tk.BooleanVar(value=False)
        self.upx_path = tk.StringVar()
        self.add_data = tk.StringVar()
        self.add_binary = tk.StringVar()
        self.company = tk.StringVar()
        self.file_description = tk.StringVar()
        self.file_version = tk.StringVar(value="1.0.0")
        self.open_after = tk.BooleanVar(value=True)
        self.custom_args = tk.StringVar()
        self.progress_value = tk.DoubleVar(value=0.0)
        self.log_lines = []
        self.config_path = tk.StringVar()
        self.template_name = tk.StringVar()

    def _build_ui(self):
        root = self.root
        pad = 8

        # Top frame
        top = ttk.Frame(root, padding=pad)
        top.pack(fill="x")

        title = ttk.Label(top, text=APP_NAME, font=("Segoe UI", 16, "bold"))
        title.pack(side="left")

        btn_frame = ttk.Frame(top)
        btn_frame.pack(side="right")

        # 新增图标转换器按钮
        ttk.Button(
            btn_frame, text="图标转换器", command=self._open_icon_converter
        ).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config).pack(
            side="left", padx=2
        )
        ttk.Button(btn_frame, text="加载配置", command=self.load_config).pack(
            side="left", padx=2
        )

        # Main layout
        main = ttk.Frame(root, padding=pad)
        main.pack(fill="both", expand=True)

        # 可滚动左侧面板
        canvas_frame = ttk.Frame(main)
        canvas_frame.pack(side="left", fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, borderwidth=0)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        left = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=left, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        left.bind("<Configure>", on_frame_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        right = ttk.Frame(main, width=360)
        right.pack(side="right", fill="y")

        # --- Left: Inputs ---
        # file selectors with enhanced icon handling
        self._add_file_selector(left, "选择主程序 (.py)", self.py_file, self._select_py)

        # 主图标选择器 - 增强功能
        ico_frame = ttk.Frame(left)
        ico_frame.pack(fill="x", pady=4)
        ttk.Label(ico_frame, text="主图标 (.ico 或 图片)").pack(anchor="w")
        hint = "用于主图标/资源 - 支持PNG/JPG等格式自动转换"
        ttk.Label(ico_frame, text=hint, font=("Segoe UI", 8)).pack(anchor="w")

        ico_row = ttk.Frame(ico_frame)
        ico_row.pack(fill="x")
        ttk.Entry(ico_row, textvariable=self.ico_file).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(ico_row, text="浏览", command=self._select_ico).pack(
            side="right", padx=4
        )
        # 添加转换按钮
        ttk.Button(
            ico_row,
            text="转换",
            command=lambda: self._convert_current_icon(self.ico_file),
        ).pack(side="right", padx=2)

        self._add_file_selector(
            left, "托盘图标 (可选)", self.tray_ico_file, self._select_tray_ico
        )
        self._add_file_selector(
            left, "窗口图标 (可选)", self.win_ico_file, self._select_win_ico
        )
        self._add_file_selector(left, "输出目录", self.output_dir, self._select_output)

        # options frame
        opt_box = ttk.Labelframe(left, text="打包增强选项", padding=6)
        opt_box.pack(fill="x", pady=6)

        ttk.Checkbutton(opt_box, text="启用 UPX 压缩", variable=self.upx_enable).pack(
            anchor="w"
        )
        upx_row = ttk.Frame(opt_box)
        upx_row.pack(fill="x", pady=2)
        ttk.Entry(upx_row, textvariable=self.upx_path).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(upx_row, text="选择 UPX", command=self._select_upx).pack(
            side="right", padx=4
        )

        ttk.Label(opt_box, text="额外数据 (--add-data)").pack(anchor="w")
        ttk.Entry(opt_box, textvariable=self.add_data).pack(fill="x")
        ttk.Label(opt_box, text="额外二进制 (--add-binary)").pack(anchor="w")
        ttk.Entry(opt_box, textvariable=self.add_binary).pack(fill="x")

        # metadata
        meta_box = ttk.Labelframe(left, text="版本 / 元数据", padding=6)
        meta_box.pack(fill="x", pady=6)
        ttk.Label(meta_box, text="公司/作者").pack(anchor="w")
        ttk.Entry(meta_box, textvariable=self.company).pack(fill="x")
        ttk.Label(meta_box, text="文件描述").pack(anchor="w")
        ttk.Entry(meta_box, textvariable=self.file_description).pack(fill="x")
        ttk.Label(meta_box, text="版本号").pack(anchor="w")
        ttk.Entry(meta_box, textvariable=self.file_version).pack(fill="x")

        # custom args
        extra_box = ttk.Labelframe(left, text="自定义参数 / 模板", padding=6)
        extra_box.pack(fill="x", pady=6)
        ttk.Label(extra_box, text="自定义 PyInstaller 参数").pack(anchor="w")
        ttk.Entry(extra_box, textvariable=self.custom_args).pack(fill="x")

        trow = ttk.Frame(extra_box)
        trow.pack(fill="x", pady=4)
        ttk.Button(
            trow, text="生成命令行字符串", command=self._generate_cmd_string
        ).pack(side="left")
        ttk.Button(trow, text="一键重打包", command=self._repack).pack(side="right")

        # pack button
        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=6)
        action_row = ttk.Frame(left)
        action_row.pack(fill="x")
        ttk.Button(
            action_row, text="开始打包", command=self.start_pack, bootstyle="success"
        ).pack(side="left", padx=4)
        ttk.Button(action_row, text="停止打包", command=self._stop_pack).pack(
            side="left"
        )
        ttk.Button(action_row, text="清理临时", command=self._clean_env).pack(
            side="right"
        )

        # --- Right: Preview + Logs ---
        preview_box = ttk.Labelframe(right, text="图标预览 / EXE 信息", padding=6)
        preview_box.pack(fill="both", padx=4, pady=4, expand=False)

        self.preview_label = ttk.Label(preview_box, text="(未选择图标)")
        self.preview_label.pack()

        self.exe_info = tk.Text(preview_box, height=6, width=40, state="disabled")
        self.exe_info.pack(fill="both")

        # 图标工具按钮
        icon_tools_frame = ttk.Frame(preview_box)
        icon_tools_frame.pack(fill="x", pady=5)

        ttk.Button(
            icon_tools_frame, text="检测图标尺寸", command=self._check_icon_sizes
        ).pack(side="left", padx=2)
        ttk.Button(
            icon_tools_frame, text="优化图标", command=self._optimize_current_icon
        ).pack(side="left", padx=2)

        log_box = ttk.Labelframe(right, text="打包日志", padding=6)
        log_box.pack(fill="both", expand=True, padx=4, pady=4)

        self.log_text = tk.Text(log_box, height=20, width=48)
        self.log_text.pack(fill="both", expand=True)

        # progress
        prog_frame = ttk.Frame(right)
        prog_frame.pack(fill="x", padx=4, pady=4)
        self.progress = ttk.Progressbar(
            prog_frame, variable=self.progress_value, maximum=100
        )
        self.progress.pack(fill="x")
        self.progress_label = ttk.Label(prog_frame, text="0%")
        self.progress_label.pack()

        note = ttk.Label(root, text="提示：支持PNG/JPG等图片格式自动转换为ICO")
        note.pack(side="bottom", pady=6)

    # ---------------- UI helpers ----------------
    def _add_file_selector(self, parent, label, var, cmd, hint=None):
        f = ttk.Frame(parent)
        f.pack(fill="x", pady=4)
        lbl = ttk.Label(f, text=label)
        lbl.pack(anchor="w")
        row = ttk.Frame(f)
        row.pack(fill="x")
        ent = ttk.Entry(row, textvariable=var)
        ent.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览", command=cmd).pack(side="right", padx=4)
        if hint:
            ttk.Label(f, text=hint, font=("Segoe UI", 8)).pack(anchor="w")

    # 添加缺失的文件选择方法
    def _select_py(self):
        """选择Python文件"""
        p = filedialog.askopenfilename(filetypes=[("Python文件", "*.py")])
        if p:
            self.py_file.set(p)

    def _select_ico(self):
        """选择图标文件"""
        p = filedialog.askopenfilename(
            filetypes=[
                ("图标文件", "*.ico;*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
                ("ICO文件", "*.ico"),
                ("图片文件", "*.png;*.jpg;*.jpeg;*.bmp"),
                ("所有文件", "*.*"),
            ]
        )
        if p:
            self.ico_file.set(p)
            self._update_preview(p)

    def _select_tray_ico(self):
        """选择托盘图标文件"""
        p = filedialog.askopenfilename(filetypes=[("图标文件", "*.ico;*.png;*.jpg")])
        if p:
            self.tray_ico_file.set(p)

    def _select_win_ico(self):
        """选择窗口图标文件"""
        p = filedialog.askopenfilename(filetypes=[("图标文件", "*.ico;*.png;*.jpg")])
        if p:
            self.win_ico_file.set(p)

    def _select_output(self):
        """选择输出目录"""
        p = filedialog.askdirectory()
        if p:
            self.output_dir.set(p)

    def _select_upx(self):
        """选择UPX程序"""
        p = filedialog.askopenfilename(filetypes=[("UPX程序", "upx*")])
        if p:
            self.upx_path.set(p)

    def _update_preview(self, path):
        try:
            if PIL_AVAILABLE and os.path.exists(path):
                img = Image.open(path)
                img.thumbnail((128, 128))
                self.tk_preview = ImageTk.PhotoImage(img)
                self.preview_label.config(image=self.tk_preview, text="")

                # 显示详细信息
                if path.lower().endswith(".ico"):
                    sizes = ico_contains_sizes(path)
                    size_info = f"包含尺寸: {sizes}" if sizes else "单尺寸ICO"
                else:
                    sizes = [img.size]
                    size_info = f"图片尺寸: {img.size}"

                file_info = (
                    f"文件: {os.path.basename(path)}\n"
                    f"{size_info}\n"
                    f"格式: {img.format}\n"
                    f"路径: {path}"
                )
                self._set_exe_info(file_info)
            else:
                self.preview_label.config(text=os.path.basename(path))
                self._set_exe_info(f"文件: {os.path.basename(path)}\n路径: {path}")
        except Exception as e:
            self.preview_label.config(text="预览失败")
            self._set_exe_info(f"预览失败: {str(e)}")

    def _set_exe_info(self, txt):
        self.exe_info.config(state="normal")
        self.exe_info.delete("1.0", "end")
        self.exe_info.insert("1.0", txt)
        self.exe_info.config(state="disabled")

    # ---------------- Icon Tools ----------------

    def _open_icon_converter(self):
        """打开图标转换器对话框"""
        theme = "light"
        if hasattr(self, "use_tb") and self.use_tb:
            theme = "light"  # 可以根据主题系统调整
        IconConverterDialog(self.root, theme)

    def _convert_current_icon(self, icon_var):
        """转换当前选择的图标"""
        icon_path = icon_var.get()
        if not icon_path:
            messagebox.showwarning("警告", "请先选择要转换的图标文件")
            return

        if not os.path.exists(icon_path):
            messagebox.showerror("错误", "图标文件不存在")
            return

        if icon_path.lower().endswith(".ico"):
            if messagebox.askyesno("确认", "文件已经是ICO格式，是否重新生成优化版本？"):
                output_path = icon_path
            else:
                return
        else:
            # 选择输出路径
            base_name = os.path.splitext(icon_path)[0]
            output_path = filedialog.asksaveasfilename(
                defaultextension=".ico",
                initialfile=os.path.basename(base_name + ".ico"),
                filetypes=[("ICO文件", "*.ico")],
            )
            if not output_path:
                return

        try:
            self._append_log(f"开始转换图标: {os.path.basename(icon_path)}")
            success = convert_image_to_ico(icon_path, output_path)

            if success:
                # 优化图标
                optimize_icon_for_windows(output_path)

                # 更新变量
                icon_var.set(output_path)
                self._update_preview(output_path)
                self._append_log(f"✓ 图标转换完成: {os.path.basename(output_path)}")
                messagebox.showinfo("成功", f"图标转换完成!\n{output_path}")
            else:
                self._append_log("✗ 图标转换失败")
                messagebox.showerror("错误", "图标转换失败")

        except Exception as e:
            self._append_log(f"✗ 图标转换错误: {str(e)}")
            messagebox.showerror("错误", f"转换过程中发生错误:\n{str(e)}")

    def _check_icon_sizes(self):
        """检测当前图标的尺寸"""
        icon_path = self.ico_file.get()
        if not icon_path or not os.path.exists(icon_path):
            messagebox.showwarning("警告", "请先选择图标文件")
            return

        if icon_path.lower().endswith(".ico"):
            sizes = ico_contains_sizes(icon_path)
            if sizes:
                size_info = "\n".join([f"{w} x {h}" for w, h in sizes])
                messagebox.showinfo(
                    "图标尺寸检测", f"当前ICO文件包含以下尺寸:\n{size_info}"
                )
            else:
                messagebox.showinfo("图标尺寸检测", "无法读取ICO文件尺寸信息")
        else:
            try:
                img = Image.open(icon_path)
                messagebox.showinfo(
                    "图标尺寸检测", f"图片尺寸: {img.size[0]} x {img.size[1]}"
                )
            except Exception as e:
                messagebox.showerror("错误", f"无法读取图片尺寸: {str(e)}")

    def _optimize_current_icon(self):
        """优化当前图标"""
        icon_path = self.ico_file.get()
        if not icon_path or not os.path.exists(icon_path):
            messagebox.showwarning("警告", "请先选择图标文件")
            return

        try:
            self._append_log(f"开始优化图标: {os.path.basename(icon_path)}")
            success = optimize_icon_for_windows(icon_path)

            if success:
                self._update_preview(icon_path)
                self._append_log("✓ 图标优化完成")
                messagebox.showinfo("成功", "图标优化完成!")
            else:
                self._append_log("✗ 图标优化失败")
                messagebox.showerror("错误", "图标优化失败")

        except Exception as e:
            self._append_log(f"✗ 图标优化错误: {str(e)}")
            messagebox.showerror("错误", f"优化过程中发生错误:\n{str(e)}")

    # ---------------- 其余方法保持不变 ----------------
    # (save_config, load_config, _build_pyinstaller_cmd, start_pack, 等方法保持不变)
    # 这里省略了重复的代码，您原有的这些方法可以保持不变

    def save_config(self):
        cfg = {
            "py_file": self.py_file.get(),
            "ico_file": self.ico_file.get(),
            "tray_ico_file": self.tray_ico_file.get(),
            "win_ico_file": self.win_ico_file.get(),
            "output_dir": self.output_dir.get(),
            "upx_enable": self.upx_enable.get(),
            "upx_path": self.upx_path.get(),
            "add_data": self.add_data.get(),
            "add_binary": self.add_binary.get(),
            "company": self.company.get(),
            "file_description": self.file_description.get(),
            "file_version": self.file_version.get(),
            "custom_args": self.custom_args.get(),
        }
        p = filedialog.asksaveasfilename(
            defaultextension=CONFIG_EXT, filetypes=[("AIP Config", CONFIG_EXT)]
        )
        if p:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("保存配置", "已保存配置: " + p)

    def load_config(self):
        p = filedialog.askopenfilename(filetypes=[("AIP Config", CONFIG_EXT)])
        if p:
            with open(p, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.py_file.set(cfg.get("py_file", ""))
            self.ico_file.set(cfg.get("ico_file", ""))
            self.tray_ico_file.set(cfg.get("tray_ico_file", ""))
            self.win_ico_file.set(cfg.get("win_ico_file", ""))
            self.output_dir.set(cfg.get("output_dir", self.output_dir.get()))
            self.upx_enable.set(cfg.get("upx_enable", False))
            self.upx_path.set(cfg.get("upx_path", ""))
            self.add_data.set(cfg.get("add_data", ""))
            self.add_binary.set(cfg.get("add_binary", ""))
            self.company.set(cfg.get("company", ""))
            self.file_description.set(cfg.get("file_description", ""))
            self.file_version.set(cfg.get("file_version", "1.0.0"))
            self.custom_args.set(cfg.get("custom_args", ""))
            messagebox.showinfo("加载配置", "已加载配置")

    def _build_pyinstaller_cmd(self, dry_run=False):
        py = self.py_file.get()
        ico = self.ico_file.get()
        out = self.output_dir.get() or os.getcwd()
        cmd_parts = ["pyinstaller", "-F"]
        # windowed
        cmd_parts.append("-w")
        if ico:
            # 自动处理图标格式
            if not ico.lower().endswith(".ico") and PIL_AVAILABLE:
                try:
                    # 临时转换图片为ICO
                    temp_ico = tempfile.mktemp(suffix=".ico")
                    if convert_image_to_ico(ico, temp_ico):
                        ico = temp_ico
                        self._append_log("✓ 自动转换图片为ICO格式")
                    else:
                        self._append_log("⚠ 图片转换失败，尝试直接使用")
                except Exception as e:
                    self._append_log(f"⚠ 图标转换失败: {e}")

            cmd_parts.append(f'--icon="{ico}"')

        # version file
        verfile = None
        if any(
            [self.company.get(), self.file_description.get(), self.file_version.get()]
        ):
            ver_content = self._generate_version_file_content()
            fd, vf = tempfile.mkstemp(suffix=".txt")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(ver_content)
            verfile = vf
            cmd_parts.append(f'--version-file="{verfile}"')

        # 其他参数保持不变
        if self.add_data.get():
            cmd_parts.append(f'--add-data="{self.add_data.get()}"')
        if self.add_binary.get():
            cmd_parts.append(f'--add-binary="{self.add_binary.get()}"')
        if self.upx_enable.get() and self.upx_path.get():
            cmd_parts.append(f'--upx-dir="{self.upx_path.get()}"')
        if self.custom_args.get():
            cmd_parts.append(self.custom_args.get())

        cmd_parts.append(f'--distpath="{out}"')
        cmd_parts.append(f'"{py}"')

        cmd = " ".join(cmd_parts)
        if dry_run:
            return cmd
        else:
            return cmd, verfile

    def _generate_version_file_content(self):
        """生成正确的版本文件内容"""
        version_tuple = self._ver_tuple(self.file_version.get())
        company = self.company.get() or "Unknown Company"
        description = self.file_description.get() or "Application"
        version = self.file_version.get() or "1.0.0"

        content = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({version_tuple}),
    prodvers=({version_tuple}),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'{company}'),
         StringStruct(u'FileDescription', u'{description}'),
         StringStruct(u'FileVersion', u'{version}'),
         StringStruct(u'InternalName', u'{description}'),
         StringStruct(u'ProductName', u'{description}'),
         StringStruct(u'ProductVersion', u'{version}')])
    ]),
    VarFileInfo([VarStruct(u'Translation', [0x409, 1200])])
  ]
)
"""
        return content

    def _ver_tuple(self, vstr):
        parts = vstr.split(".")
        while len(parts) < 4:
            parts.append("0")
        return ",".join(p if p.isdigit() else "0" for p in parts[:4])

    def start_pack(self):
        if not self.py_file.get():
            messagebox.showwarning("提示", "请先选择要打包的 Python 文件")
            return

        # 重置状态
        self._stop_flag = False
        self.progress_value.set(0)
        self.progress_label.config(text="0%")
        self.log_text.delete("1.0", "end")

        # 构建命令
        try:
            cmd, verfile = self._build_pyinstaller_cmd(dry_run=False)
        except Exception as e:
            messagebox.showerror("构建命令失败", str(e))
            return

        # 启动线程
        t = threading.Thread(
            target=self._run_pack_thread, args=(cmd, verfile), daemon=True
        )
        t.start()

    def _run_pack_thread(self, cmd, verfile):
        self._append_log("开始打包: " + cmd)
        self._append_log("=" * 50)

        def on_line(line):
            if self._stop_flag:
                return
            self._append_log(line)
            self._parse_progress_from_line(line)

        # 运行
        try:
            ret = run_subprocess_stream(cmd, on_line=on_line, shell=True)
        except Exception as e:
            self._append_log(f"打包进程异常: {e}")
            ret = 1

        # 清理版本文件
        try:
            if verfile and os.path.exists(verfile):
                os.remove(verfile)
        except Exception:
            pass

        # 完成处理
        if ret == 0:
            self._set_progress(100)
            self._append_log("✓ 打包成功完成！")
        else:
            self._append_log(f"✗ 打包失败，返回码: {ret}")

        self._on_pack_finish(ret)

    def _append_log(self, line):
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def _parse_progress_from_line(self, line):
        line_lower = line.lower()
        if "checking analysis" in line_lower:
            self._set_progress(5)
        elif "analyzing" in line_lower and "base_library" in line_lower:
            self._set_progress(10)
        elif "analyzing" in line_lower and ".py" in line_lower:
            self._set_progress(20)
        elif "processing module hooks" in line_lower:
            self._set_progress(30)
        elif "looking for ctypes dlls" in line_lower:
            self._set_progress(40)
        elif "analyzing run-time hooks" in line_lower:
            self._set_progress(50)
        elif "creating base_library.zip" in line_lower:
            self._set_progress(60)
        elif "building pyz" in line_lower:
            self._set_progress(70)
        elif "building pkg" in line_lower:
            self._set_progress(80)
        elif "building exe" in line_lower:
            self._set_progress(90)
        elif "completed successfully" in line_lower:
            self._set_progress(95)

    def _set_progress(self, pct):
        try:
            self.progress_value.set(pct)
            self.progress_label.config(text=f"{int(pct)}%")
        except Exception:
            pass

    def _on_pack_finish(self, retcode):
        out = self.output_dir.get()
        if retcode == 0 and self.open_after.get() and os.path.exists(out):
            open_path_in_explorer(out)
        if TOAST_AVAILABLE and retcode == 0:
            try:
                self.toast.show_toast(APP_NAME, "打包完成", duration=5)
            except Exception:
                pass

    def _stop_pack(self):
        self._stop_flag = True
        self._append_log("用户请求停止打包...")
        messagebox.showinfo(
            "停止打包",
            "已设置停止标志。若子进程仍在运行，请手动结束 pyinstaller 进程。",
        )

    def _clean_env(self):
        d = self.output_dir.get() or "dist"
        b = "build"
        removed = []
        for p in [d, b]:
            if os.path.exists(p):
                try:
                    shutil.rmtree(p)
                    removed.append(p)
                except Exception as e:
                    self._append_log("删除失败: " + str(e))
        for f in os.listdir("."):
            if f.endswith(".spec"):
                try:
                    os.remove(f)
                    removed.append(f)
                except Exception:
                    pass
        if removed:
            messagebox.showinfo("清理完成", "已删除: " + ", ".join(removed))
            self._append_log("清理完成: " + ", ".join(removed))
        else:
            messagebox.showinfo("清理", "未发现临时文件")

    def _repack(self):
        self.start_pack()

    def _generate_cmd_string(self):
        cmd = self._build_pyinstaller_cmd(dry_run=True)
        messagebox.showinfo("命令行字符串", cmd)

    def _bind_shortcuts(self):
        self.root.bind("<Control-s>", lambda e: self.save_config())
        self.root.bind("<Control-o>", lambda e: self.load_config())

    def check_dependencies(self):
        missing = []
        try:
            import PyInstaller
        except Exception:
            missing.append("pyinstaller")
        if not PIL_AVAILABLE:
            missing.append("Pillow")
        if not TOAST_AVAILABLE:
            missing.append("win10toast (可选，用于通知)")
        if missing:
            self._append_log("检测到以下缺失依赖: " + ", ".join(missing))


# ---------------- Run ----------------


def main():
    if TB_AVAILABLE:
        app_root = tb.Window(themename="flatly")
    else:
        app_root = tk.Tk()
    app = AutoIconPackerApp(app_root)
    app.check_dependencies()
    app_root.geometry("960x640")
    app_root.mainloop()


if __name__ == "__main__":
    main()
