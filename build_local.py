#!/usr/bin/env python3
"""
本地测试脚本 - 在推送到GitHub前测试编译
"""
import os
import sys
import subprocess
import platform

def run_command(cmd, check=True):
    """运行命令并打印输出"""
    print(f"执行: {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, check=check, 
                              capture_output=True, text=True)
        if result.stdout:
            print(result.stdout)
        return result.returncode == 0
    except subprocess.CalledProcessError as e:
        print(f"错误: {e}")
        if e.stderr:
            print(f"错误输出: {e.stderr}")
        return False

def main():
    print("=== AutoClicker 本地编译测试 ===")
    
    # 检查文件
    if not os.path.exists("autoclicker.py"):
        print("错误: 找不到 autoclicker.py")
        return 1
        
    # 安装依赖
    print("\n1. 安装依赖...")
    deps = ["pyautogui", "pystray", "Pillow", "keyboard", "requests", "pyinstaller"]
    for dep in deps:
        run_command(f"pip install {dep}")
    
    # 清理
    print("\n2. 清理旧文件...")
    for dir_name in ["build", "dist"]:
        if os.path.exists(dir_name):
            run_command(f"rmdir /s /q {dir_name}" if platform.system() == "Windows" 
                       else f"rm -rf {dir_name}")
    
    # 编译
    print("\n3. 开始编译...")
    if platform.system() == "Windows":
        cmd = (
            'pyinstaller --onefile --windowed --noconfirm --clean '
            '--name "AutoClicker_v2.3" '
            '--hidden-import=pystray._win32 '
            '--hidden-import=PIL._imaging '
            '--hidden-import=keyboard._winkeyboard '
            'autoclicker.py'
        )
    else:
        cmd = (
            'pyinstaller --onefile --noconfirm --clean '
            '--name "AutoClicker_v2.3" '
            '--hidden-import=pystray._xorg '
            '--hidden-import=PIL._imaging '
            'autoclicker.py'
        )
    
    success = run_command(cmd)
    
    # 验证输出
    if success and os.path.exists("dist/AutoClicker_v2.3.exe" if platform.system() == "Windows" 
                                 else "dist/AutoClicker_v2.3"):
        print("\n✅ 编译成功!")
        print(f"输出文件: dist/AutoClicker_v2.3{'exe' if platform.system() == 'Windows' else ''}")
        return 0
    else:
        print("\n❌ 编译失败!")
        return 1

if __name__ == "__main__":
    sys.exit(main())