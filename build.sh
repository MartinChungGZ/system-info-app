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

echo "[1/4] 安装依赖包..."
pip3 install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "[错误] 依赖安装失败"
    exit 1
fi

echo ""
echo "[2/4] 清理旧的构建文件..."
rm -rf dist build dmg_staging

echo ""
echo "[3/4] 使用PyInstaller构建macOS App Bundle..."
# --onedir on macOS creates a proper .app bundle when --windowed is set
pyinstaller --onedir --windowed --name "系统信息采集工具_Mac" \
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

if [ $? -ne 0 ]; then
    echo "[错误] PyInstaller打包失败"
    exit 1
fi

echo ""
echo "[4/4] 创建专业DMG安装映像..."

# Prepare DMG staging
mkdir -p dmg_staging
cp -R "dist/系统信息采集工具_Mac.app" dmg_staging/

# Create Applications symlink for drag-to-install UX
ln -s /Applications dmg_staging/Applications

# Create DMG
DMG_NAME="dist/系统信息采集工具_Mac.dmg"
hdiutil create -volname "系统信息采集工具" \
    -srcfolder dmg_staging \
    -ov -format UDZO \
    -fs HFS+ \
    "$DMG_NAME"

rm -rf dmg_staging

echo ""
echo "============================================"
echo "  打包完成！"
echo "  DMG: $DMG_NAME"
echo "  Mac用户操作：双击DMG → 拖App到Applications → 完成"
echo "============================================"
