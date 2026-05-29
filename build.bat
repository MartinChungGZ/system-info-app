@echo off
chcp 65001 >nul
echo ============================================
echo   系统信息采集工具 - Windows 打包脚本
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到Python，请先安装Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装依赖包...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [2/3] 清理旧的构建文件...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

echo.
echo [3/3] 使用PyInstaller打包为独立exe...
pyinstaller --onefile --windowed --name "系统信息采集工具" ^
    --hidden-import psutil ^
    --hidden-import psutil._common ^
    --hidden-import psutil._pswindows ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import speedtest ^
    --hidden-import speedtest_cli ^
    --hidden-import smtplib ^
    --hidden-import ssl ^
    --hidden-import json ^
    --hidden-import email.mime.multipart ^
    --hidden-import email.mime.base ^
    --hidden-import email.mime.text ^
    --hidden-import email.encoders ^
    --hidden-import urllib.request ^
    --hidden-import xml.etree.ElementTree ^
    --collect-all openpyxl ^
    system_info_app.py

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   打包完成！
    echo   可执行文件位置: dist\系统信息采集工具.exe
    echo   将该exe文件复制到任意Windows电脑双击即可运行
    echo ============================================
) else (
    echo [错误] 打包失败
)

pause
