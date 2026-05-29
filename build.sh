#!/bin/bash
# ============================================
#   系统信息采集工具 - macOS 打包脚本
# ============================================

echo "============================================"
echo "  系统信息采集工具 - macOS 打包脚本"
echo "============================================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未找到Python3，请先安装Python 3.8+"
    echo "下载地址: https://www.python.org/downloads/"
    exit 1
fi

echo "[1/3] 安装依赖包..."
pip3 install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "[错误] 依赖安装失败"
    exit 1
fi

echo ""
echo "[2/3] 清理旧的构建文件..."
rm -rf dist build

echo ""
echo "[3/3] 使用PyInstaller打包为独立应用..."
pyinstaller --onefile --windowed --name "系统信息采集工具_Mac" \
    --hidden-import psutil \
    --hidden-import psutil._common \
    --hidden-import psutil._psosx \
    --hidden-import psutil._psposix \
    --hidden-import openpyxl \
    --hidden-import openpyxl.styles \
    --hidden-import speedtest \
    --hidden-import speedtest_cli \
    --hidden-import smtplib \
    --hidden-import ssl \
    --hidden-import json \
    --hidden-import email.mime.multipart \
    --hidden-import email.mime.base \
    --hidden-import email.mime.text \
    --hidden-import email.encoders \
    --hidden-import urllib.request \
    --hidden-import xml.etree.ElementTree \
    --collect-all openpyxl \
    system_info_app.py

if [ $? -eq 0 ]; then
    echo ""
    echo "============================================"
    echo "  打包完成！"
    echo "  可执行文件位置: dist/系统信息采集工具_Mac"
    echo "  将该文件复制到任意Mac电脑双击即可运行"
    echo "============================================"
else
    echo "[错误] 打包失败"
fi
