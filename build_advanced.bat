@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========================================
echo    AutoClicker 2.3 编译脚本
echo ========================================
echo.

set "SOURCE_FILE=autoclicker.py"
set "OUTPUT_NAME=AutoClicker_v2.3"

:: 简单直接的检查
if not exist "%SOURCE_FILE%" (
    echo [错误] 找不到 %SOURCE_FILE%
    pause
    exit /b 1
)

echo 步骤1/5: 检查Python...
python --version >nul 2>&1
if errorlevel 1 goto :error

echo 步骤2/5: 安装依赖...
pip install pyautogui pystray Pillow keyboard requests pyinstaller >nul 2>&1

echo 步骤3/5: 清理旧文件...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "*.spec" del "*.spec" 2>nul

echo 步骤4/5: 选择模式...
echo 1. 单文件模式 (推荐)
echo 2. 调试模式
echo.
choice /c 12 /n /m "请选择"
if errorlevel 2 (
    set "OPTS=--onefile --console"
    set "OUTPUT_NAME=!OUTPUT_NAME!_Debug"
) else (
    set "OPTS=--onefile --windowed"
)

echo 步骤5/5: 开始编译...
echo 请稍候，这需要几分钟...

pyinstaller !OPTS! --noconfirm --clean --name "!OUTPUT_NAME!" --hidden-import=pystray._win32 --hidden-import=PIL._imaging --hidden-import=keyboard._winkeyboard "%SOURCE_FILE%"

if errorlevel 1 goto :error

if exist "dist\!OUTPUT_NAME!.exe" (
    echo.
    echo ========== 编译成功! ==========
    echo 输出文件: dist\!OUTPUT_NAME!.exe
    echo.
    set /p OPEN="是否打开输出目录? [Y/N]: "
    if /i "!OPEN!"=="Y" explorer "dist"
) else (
    goto :error
)

goto :end

:error
echo.
echo [错误] 编译失败!
pause
exit /b 1

:end
pause