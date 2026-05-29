#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Headless runner - collect system info and generate Excel without GUI."""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from system_info_app import (
    get_cpu_info, get_memory_info, get_network_speed,
    get_browser_versions, create_excel
)

def main():
    import system_info_app as app
    app.set_log_func(lambda msg: print(msg))

    print("=" * 60)
    print("  系统信息采集工具 - 无头运行模式")
    print("=" * 60)

    campus = "初四梅花园校区"
    name = "测试用户"

    print()
    print("[1/5] 采集CPU信息...")
    cpu_info = get_cpu_info()
    for k, v in cpu_info.items():
        print(f"  {k}: {v}")

    print()
    print("[2/5] 采集内存信息...")
    mem_info = get_memory_info()
    for k, v in mem_info.items():
        print(f"  {k}: {v}")

    print()
    print("[3/5] 网络测速（可能需要1-2分钟）...")
    net_info = get_network_speed()
    for k, v in net_info.items():
        print(f"  {k}: {v}")

    print()
    print("[4/5] 采集浏览器版本...")
    browser_info = get_browser_versions()
    for k, v in browser_info.items():
        print(f"  {k}: {v}")

    print()
    print("[5/5] 生成Excel报告...")
    excel_path = create_excel(campus, name, cpu_info, mem_info, net_info, browser_info)

    print()
    print("=" * 60)
    print(f"  Excel文件已生成: {excel_path}")
    print("=" * 60)

if __name__ == "__main__":
    main()
