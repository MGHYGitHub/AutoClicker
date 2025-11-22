# build_advanced.py - 带进度显示的单文件打包脚本
import os
import sys
import subprocess
import shutil
import time
from pathlib import Path
from threading import Thread, Lock


class ProgressDisplay:
    """进度显示类"""

    def __init__(self):
        self.current_stage = ""
        self.stage_progress = 0
        self.total_stages = 0
        self.completed_stages = 0
        self.lock = Lock()
        self.running = False
        self.animation_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.animation_index = 0

    def start(self):
        """开始显示进度"""
        self.running = True
        self.animation_thread = Thread(target=self._animation_loop, daemon=True)
        self.animation_thread.start()

    def stop(self):
        """停止显示进度"""
        self.running = False
        if hasattr(self, "animation_thread"):
            self.animation_thread.join(timeout=1)

    def _animation_loop(self):
        """动画循环"""
        while self.running:
            with self.lock:
                animation_char = self.animation_chars[self.animation_index]
                self.animation_index = (self.animation_index + 1) % len(
                    self.animation_chars
                )

                # 构建进度条
                progress_bar = self._build_progress_bar()

                # 清空当前行并显示新进度
                sys.stdout.write("\r" + " " * 100 + "\r")
                if self.current_stage:
                    sys.stdout.write(
                        f"{animation_char} {self.current_stage} {progress_bar}"
                    )
                sys.stdout.flush()

            time.sleep(0.1)

    def _build_progress_bar(self):
        """构建进度条"""
        bar_length = 20
        if self.stage_progress <= 0:
            return "[" + " " * bar_length + "]"

        filled = int(bar_length * self.stage_progress)
        bar = "█" * filled + "░" * (bar_length - filled)
        return f"[{bar}] {self.stage_progress*100:.1f}%"

    def update_stage(self, stage_name, progress=0.0):
        """更新当前阶段和进度"""
        with self.lock:
            self.current_stage = stage_name
            self.stage_progress = max(0.0, min(1.0, progress))

    def complete_stage(self):
        """完成一个阶段"""
        with self.lock:
            self.completed_stages += 1
            self.stage_progress = 1.0
            # 显示完成的阶段
            sys.stdout.write("\r" + " " * 100 + "\r")
            sys.stdout.write(f"✓ {self.current_stage} 完成\n")
            sys.stdout.flush()
            self.current_stage = ""
            self.stage_progress = 0


class BuildProcess:
    """构建过程管理"""

    def __init__(self):
        self.progress = ProgressDisplay()
        self.start_time = None

    def print_header(self):
        """打印头部信息"""
        print("🚀 AutoClicker 2.5.3 单文件打包脚本")
        print("=" * 60)
        print("📦 目标: 生成单个包含所有资源的exe文件")
        print("=" * 60)

    def print_footer(self, success=True):
        """打印底部信息"""
        if self.start_time:
            elapsed = time.time() - self.start_time
            print(f"\n⏱️  总耗时: {elapsed:.1f} 秒")

        if success:
            print("🎉 构建过程完成!")
        else:
            print("❌ 构建过程失败!")

    def run_with_progress(self, func, stage_name, *args, **kwargs):
        """带进度显示运行函数"""
        self.progress.update_stage(stage_name, 0.1)
        try:
            result = func(*args, **kwargs)
            if result:
                self.progress.complete_stage()
            else:
                # 即使失败也要完成进度显示
                self.progress.update_stage(f"{stage_name} - 失败", 1.0)
                # 添加一个短暂的延迟，让用户看到失败信息
                time.sleep(0.5)
                # 然后完成这个阶段
                self.progress.complete_stage()
            return result
        except Exception as e:
            self.progress.update_stage(f"{stage_name} - 错误: {e}", 1.0)
            # 添加延迟并完成阶段
            time.sleep(0.5)
            self.progress.complete_stage()
            return False


def check_dependencies(progress):
    """检查必要的依赖"""
    required_packages = [
        "pyinstaller",
        "pyautogui",
        "pystray",
        "Pillow",
        "keyboard",
        "requests",
        "pywin32",
    ]

    progress.update_stage("检查依赖包", 0.1)

    missing_packages = []
    for i, package in enumerate(required_packages):
        progress.update_stage(
            f"检查依赖包: {package}", 0.1 + (i * 0.8 / len(required_packages))
        )
        try:
            __import__(package.replace("-", "_"))
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"\n❌ 缺少依赖包: {', '.join(missing_packages)}")
        return False

    progress.update_stage("所有依赖包已安装", 1.0)
    return True


def install_dependencies(progress):
    """安装依赖包"""
    packages = [
        "pyinstaller",
        "pyautogui",
        "pystray",
        "Pillow",
        "keyboard",
        "requests",
        "pywin32",
    ]

    progress.update_stage("安装依赖包", 0.1)

    for i, package in enumerate(packages):
        progress.update_stage(f"安装 {package}", 0.1 + (i * 0.8 / len(packages)))
        try:
            # 显示pip安装进度
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", package],
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                print(f"\n❌ 安装 {package} 失败: {result.stderr}")
                return False
        except subprocess.CalledProcessError as e:
            print(f"\n❌ 安装 {package} 失败: {e}")
            return False

    progress.update_stage("所有依赖包安装完成", 1.0)
    return True


def create_default_icons(progress):
    """创建默认图标（如果不存在）"""
    icon_dir = Path("ICON")
    if not icon_dir.exists():
        progress.update_stage("创建默认图标", 0.1)
        icon_dir.mkdir(exist_ok=True)

        try:
            from PIL import Image, ImageDraw, ImageFont

            # 创建不同尺寸的图标
            sizes = [16, 32, 48, 64, 128, 256]
            for i, size in enumerate(sizes):
                progress.update_stage(
                    f"创建图标 {size}x{size}", 0.1 + (i * 0.8 / len(sizes))
                )

                # 创建蓝色渐变背景
                img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
                draw = ImageDraw.Draw(img)

                # 绘制渐变圆形背景
                margin = max(2, size // 16)
                for i_pos in range(margin, size - margin):
                    for j_pos in range(margin, size - margin):
                        dist = (
                            (i_pos - size / 2) ** 2 + (j_pos - size / 2) ** 2
                        ) ** 0.5
                        if dist <= (size / 2 - margin):
                            # 蓝色渐变
                            blue = int(30 + (144 - 30) * (dist / (size / 2)))
                            alpha = 255
                            draw.point((i_pos, j_pos), fill=(30, blue, 255, alpha))

                # 添加白色边框
                draw.ellipse(
                    [margin, margin, size - margin, size - margin],
                    outline=(255, 255, 255),
                    width=max(1, size // 32),
                )

                # 添加文字 "AC"
                if size >= 32:
                    try:
                        font_size = max(8, size // 3)
                        font = ImageFont.truetype("arial.ttf", font_size)
                        text = "AC"
                        bbox = draw.textbbox((0, 0), text, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                        x = (size - text_width) // 2
                        y = (size - text_height) // 2
                        draw.text((x, y), text, fill="white", font=font)
                    except:
                        center = size // 2
                        radius = size // 4
                        draw.ellipse(
                            [
                                center - radius,
                                center - radius,
                                center + radius,
                                center + radius,
                            ],
                            fill="white",
                        )

                img.save(icon_dir / f"{size}.png", "PNG")

            progress.update_stage("所有图标创建完成", 1.0)
            return True

        except ImportError:
            progress.update_stage("创建图标失败: 需要 PIL 库", 1.0)
            return False
        except Exception as e:
            progress.update_stage(f"创建图标失败: {e}", 1.0)
            return False
    else:
        progress.update_stage("图标目录已存在", 1.0)
        return True


def check_icon_files(progress):
    """检查图标文件完整性"""
    icon_dir = Path("ICON")
    if not icon_dir.exists():
        progress.update_stage("图标目录不存在", 1.0)
        return False

    progress.update_stage("检查图标文件完整性", 0.1)

    critical_sizes = [16, 32, 64, 256]
    missing = []

    for i, size in enumerate(critical_sizes):
        progress.update_stage(
            f"检查 {size}x{size} 图标", 0.1 + (i * 0.8 / len(critical_sizes))
        )
        if not (icon_dir / f"{size}.png").exists():
            missing.append(f"{size}x{size}")

    if missing:
        progress.update_stage(f"缺少关键图标: {', '.join(missing)}", 1.0)
        return False

    progress.update_stage("所有图标文件完整", 1.0)
    return True


def build_single_exe(progress):
    """构建单文件exe"""
    source_file = "AutoClicker_2.5.py"
    output_name = "AutoClicker_v2.5.3"

    if not Path(source_file).exists():
        progress.update_stage(f"错误: 找不到源文件 {source_file}", 1.0)
        return False

    # 清理旧构建文件
    progress.update_stage("清理旧构建文件", 0.1)
    for folder in ["build", "dist"]:
        if Path(folder).exists():
            shutil.rmtree(folder)
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()

    progress.update_stage("旧构建文件清理完成", 0.3)

    # PyInstaller 命令
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name",
        output_name,
        "--icon=ICON/64.png",
        "--add-data=ICON;ICON",
        "--hidden-import=pystray._win32",
        "--hidden-import=PIL._imaging",
        "--hidden-import=PIL._imagingtk",
        "--hidden-import=PIL._webp",
        "--hidden-import=win32timezone",
        "--hidden-import=win32api",
        "--noconfirm",
        "--clean",
        "--noupx",
    ]

    cmd.append(source_file)

    # 执行构建
    progress.update_stage("启动 PyInstaller 编译", 0.4)

    try:
        # 使用Popen来实时获取输出
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
        )

        # 读取输出并更新进度
        compile_stages = {
            "Analyzing": "分析依赖",
            "Building": "构建程序",
            "Generating": "生成可执行文件",
            "Writing": "写入文件",
            "Completed": "编译完成",
        }

        current_stage = "正在编译"

        for line in process.stdout:
            line = line.strip()
            if line:
                # 检测编译阶段
                for eng_stage, chi_stage in compile_stages.items():
                    if eng_stage in line:
                        current_stage = chi_stage
                        break

                # 显示编译进度
                progress.update_stage(current_stage, 0.5)

                # 显示重要的编译信息
                if "INFO:" in line or "WARNING:" in line or "ERROR:" in line:
                    print(f"\n   {line}")

        # 等待进程完成
        return_code = process.wait()

        if return_code == 0:
            progress.update_stage("编译成功完成", 1.0)
            return True
        else:
            progress.update_stage("编译过程出错", 1.0)
            return False

    except subprocess.CalledProcessError as e:
        progress.update_stage(f"编译失败: {e}", 1.0)
        return False
    except Exception as e:
        progress.update_stage(f"编译过程异常: {e}", 1.0)
        return False


def verify_exe_icon(exe_path, progress):
    """验证exe文件图标"""
    progress.update_stage("验证可执行文件", 0.1)

    try:
        import pefile

        progress.update_stage("检查图标资源", 0.5)
        pe = pefile.PE(exe_path)

        if hasattr(pe, "DIRECTORY_ENTRY_RESOURCE"):
            icon_count = 0
            for resource_type in pe.DIRECTORY_ENTRY_RESOURCE.entries:
                if resource_type.name is not None and (
                    str(resource_type.name) == "RT_ICON" or resource_type.struct.Id == 3
                ):
                    icon_count += len(resource_type.directory.entries)

            if icon_count > 0:
                progress.update_stage(f"找到 {icon_count} 个图标资源", 1.0)
                # 重要：完成这个阶段
                return True
            else:
                progress.update_stage("未找到图标资源", 1.0)
                return False
        else:
            progress.update_stage("未找到图标资源", 1.0)
            return False

    except ImportError:
        progress.update_stage("跳过图标验证 (需要 pefile)", 1.0)
        return True
    except Exception as e:
        progress.update_stage(f"图标验证失败: {e}", 1.0)
        return True


def main():
    build_process = BuildProcess()
    build_process.print_header()

    # 开始计时
    build_process.start_time = time.time()

    # 启动进度显示
    build_process.progress.start()

    try:
        # 检查依赖
        if not build_process.run_with_progress(
            check_dependencies, "检查项目依赖", build_process.progress
        ):
            if not build_process.run_with_progress(
                install_dependencies, "安装缺失依赖", build_process.progress
            ):
                build_process.print_footer(False)
                input("按回车键退出...")
                return

        # 创建图标
        if not build_process.run_with_progress(
            create_default_icons, "准备图标资源", build_process.progress
        ):
            build_process.print_footer(False)
            input("按回车键退出...")
            return

        # 检查图标完整性
        if not build_process.run_with_progress(
            check_icon_files, "验证图标完整性", build_process.progress
        ):
            build_process.print_footer(False)
            input("按回车键退出...")
            return

        # 构建EXE
        if build_process.run_with_progress(
            build_single_exe, "编译可执行文件", build_process.progress
        ):
            output_name = "AutoClicker_v2.5.3"
            exe_path = Path("dist") / f"{output_name}.exe"

            # 在 main() 函数中，找到显示成功信息的地方，在显示信息后立即停止进度：

            if exe_path.exists():
                # 验证文件
                file_size = exe_path.stat().st_size / (1024 * 1024)  # MB

                # 立即停止进度显示，避免卡住
                build_process.progress.stop()

                print(f"\n✅ 构建成功!")
                print(f"📁 文件位置: {exe_path.absolute()}")
                print(f"📏 文件大小: {file_size:.2f} MB")
                print(f"🔧 包含功能: 自动点击器 v2.5.3")

                # 如果需要验证图标，重新启动进度显示
                build_process.progress.start()
                build_process.run_with_progress(
                    lambda progress: verify_exe_icon(exe_path, progress),
                    "最终验证",
                    build_process.progress,
                )
                build_process.progress.stop()

                # 打开输出目录选项
                choice = input("\n是否打开输出目录? (y/n): ").lower()
                if choice in ["y", "yes"]:
                    os.startfile(Path("dist").absolute())

                build_process.print_footer(True)
            else:
                print("\n❌ 输出文件不存在")
                build_process.print_footer(False)
        else:
            build_process.print_footer(False)

    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断构建过程")
        build_process.print_footer(False)
    except Exception as e:
        print(f"\n\n❌ 发生未知错误: {e}")
        import traceback

        traceback.print_exc()  # 打印详细错误信息
        build_process.print_footer(False)
    finally:
        # 停止进度显示
        build_process.progress.stop()

    input("\n按回车键退出...")


if __name__ == "__main__":
    main()
