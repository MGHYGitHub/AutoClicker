@echo off
chcp 65001 >nul
echo AutoClicker 2.5 编译脚本
echo.

set "SOURCE=AutoClicker_2.5.py"
set "OUTPUT=AutoClicker_v2.5"

if not exist "%SOURCE%" (
    echo 错误: 找不到 %SOURCE%
    pause
    exit
)

echo 安装依赖...
pip install pyinstaller pyautogui pystray Pillow keyboard requests

echo 清理旧文件...
rd /s /q build dist 2>nul
del *.spec 2>nul

echo 开始编译...
pyinstaller --onefile --windowed --name "%OUTPUT%" --hidden-import=pystray._win32 --hidden-import=PIL._imaging "%SOURCE%"

if exist "dist\%OUTPUT%.exe" (
    echo 编译成功: dist\%OUTPUT%.exe
    explorer dist
) else (
    echo 编译失败!
)

pause