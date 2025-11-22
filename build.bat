# build_advanced.py - 高级打包脚本
import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_dependencies():
    """检查必要的依赖"""
    required_packages = [
        'pyinstaller',
        'pyautogui',
        'pystray', 
        'Pillow',
        'keyboard',
        'requests',
        'pywin32'
    ]
    
    print("检查依赖包...")
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"  ✓ {package}")
        except ImportError:
            print(f"  ✗ {package} 未安装")
            return False
    return True

def install_dependencies():
    """安装依赖包"""
    packages = [
        'pyinstaller',
        'pyautogui', 
        'pystray',
        'Pillow',
        'keyboard',
        'requests',
        'pywin32'
    ]
    
    print("安装依赖包...")
    for package in packages:
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            print(f"  ✓ 成功安装 {package}")
        except subprocess.CalledProcessError:
            print(f"  ✗ 安装 {package} 失败")
            return False
    return True

def create_icon_folder():
    """创建图标文件夹和默认图标（如果不存在）"""
    icon_dir = Path("ICON")
    if not icon_dir.exists():
        print("创建 ICON 文件夹...")
        icon_dir.mkdir(exist_ok=True)
        
        # 创建说明文件
        readme = icon_dir / "README.txt"
        readme.write_text("""图标文件说明
============

请在此文件夹放置以下尺寸的PNG图标文件：
- 16.png   (16x16像素)
- 32.png   (32x32像素) 
- 48.png   (48x48像素)
- 64.png   (64x64像素) - 主要图标
- 128.png  (128x128像素)
- 256.png  (256x256像素)

如果没有这些文件，程序将使用内置默认图标。
""", encoding='utf-8')
        print("  ✓ 已创建 ICON 文件夹")

def build_app():
    """构建应用程序"""
    source_file = "AutoClicker_2.5.py"
    output_name = "AutoClicker_v2.5"
    
    if not Path(source_file).exists():
        print(f"错误: 找不到源文件 {source_file}")
        return False
    
    # 清理旧构建文件
    print("清理旧构建文件...")
    for folder in ['build', 'dist']:
        if Path(folder).exists():
            shutil.rmtree(folder)
    for spec_file in Path('.').glob('*.spec'):
        spec_file.unlink()
    
    # PyInstaller 命令
    cmd = [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--name', output_name,
        '--hidden-import=pystray._win32',
        '--hidden-import=PIL._imaging',
        '--hidden-import=PIL._imagingtk', 
        '--hidden-import=PIL._webp',
        '--hidden-import=win32timezone',
        '--clean'
    ]
    
    # 添加图标（如果存在）
    icon_file = Path("ICON/64.png")
    if icon_file.exists():
        cmd.extend(['--icon', str(icon_file)])
        cmd.extend(['--add-data', 'ICON;ICON'])
        print("使用自定义图标编译...")
    else:
        print("使用默认图标编译...")
    
    cmd.append(source_file)
    
    # 执行构建
    print("开始编译...")
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"编译失败: {e}")
        return False
    
    return True

def post_build():
    """构建后处理"""
    output_name = "AutoClicker_v2.5"
    dist_dir = Path("dist")
    exe_file = dist_dir / f"{output_name}.exe"
    
    if not exe_file.exists():
        return False
    
    print("构建后处理...")
    
    # 复制配置文件示例
    config_files = ['click_points.json', 'autoclicker_log.txt']
    for config_file in config_files:
        if Path(config_file).exists():
            shutil.copy2(config_file, dist_dir)
            print(f"  ✓ 复制 {config_file}")
    
    # 创建使用说明
    readme_content = f"""AutoClicker v2.5 使用说明
==========================

程序文件: {output_name}.exe

主要功能:
• 多点位自动点击管理
• 独立点击参数设置  
• 全局快捷键支持 (F2/F3/F4)
• 智能安全检测机制
• 窗口自动切换功能
• 系统托盘运行

使用步骤:
1. 运行程序
2. 点击"开始获取坐标"，按F2记录点位
3. 配置任务参数
4. 点击"启动任务"或按F3开始

配置文件:
• click_points.json - 点位配置
• autoclicker_log.txt - 运行日志

系统要求:
• Windows 7/10/11
• 需要管理员权限用于全局快捷键

编译时间: {subprocess.getoutput('date /t')} {subprocess.getoutput('time /t')}
"""
    
    readme_file = dist_dir / "使用说明.txt"
    readme_file.write_text(readme_content, encoding='utf-8')
    print("  ✓ 创建使用说明文档")
    
    return True

def main():
    print("AutoClicker 2.5 高级编译脚本")
    print("=" * 40)
    
    # 检查并安装依赖
    if not check_dependencies():
        print("\n缺少依赖包，开始安装...")
        if not install_dependencies():
            print("依赖安装失败，请手动安装所需包")
            input("按回车键退出...")
            return
    
    # 创建图标文件夹
    create_icon_folder()
    
    print("\n开始构建过程...")
    if build_app():
        if post_build():
            print("\n" + "=" * 40)
            print("✓ 编译成功完成!")
            print(f"输出文件: dist\\AutoClicker_v2.5.exe")
            
            # 询问是否打开输出目录
            choice = input("\n是否打开输出目录? (y/n): ").lower()
            if choice in ['y', 'yes']:
                dist_path = Path("dist").absolute()
                os.startfile(dist_path)
        else:
            print("\n构建后处理失败")
    else:
        print("\n编译失败")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()