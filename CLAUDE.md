# 系统信息采集工具 (System Info Collection Tool)

## 项目概述
跨平台桌面应用（Windows/macOS），用于采集电脑系统信息、网络测速、浏览器版本，
生成Excel报表并通过邮件发送。使用Python tkinter + PyInstaller打包为免安装独立exe。

## 项目目录
```
C:\Users\PC001\system_info_app\
  system_info_app.py    # 主程序 (~880行)，包含GUI + 所有采集逻辑
  run_headless.py        # 无头测试脚本（CLI模式运行采集+生成Excel）
  requirements.txt       # Python依赖
  build.bat / build.sh   # PyInstaller打包脚本 (Windows/Mac)
  config.json            # 运行时生成，存储邮箱SMTP配置
  config_template.json   # 邮箱配置模板参考
```

## 核心功能模块 (system_info_app.py)
- `get_cpu_info()` → 注册表/wmic/sysctl获取CPU，`_parse_cpu_brand()` 提取品牌名
- `get_memory_info()` → psutil获取物理内存
- `get_network_speed()` → speedtest-cli(优先) + 多线程HTTP下载验证
- `_get_chrome_version()` → 注册表Uninstall键 + LOCALAPPDATA路径
- `_get_firefox_version()` → 注册表/文件路径/--version命令
- `create_excel()` → 生成简单二维表Excel（项目/值 两列）
- `send_email()` → SMTP/SSL发送，收件人 zhongwenjian@zy.com

## GUI界面
- 所属校区下拉框：初四越秀/梅花园/海珠/天河 + 高四公园前/海珠/梅花园/花都 + 中高复中台
- 姓名文本框：100字符限制，中英文
- "开始搜集"按钮 → 后台线程依次采集(共6步) → Excel（保存到桌面）
- "邮箱设置"按钮已禁用（暂不可用）

## 已知问题/改进方向
- Chrome --version 在Chrome已运行时可能失败，目前用注册表替代
- 邮件功能已禁用（邮箱设置按钮置灰），后续需要时再启用
- CI: Mac用macos-13 Intel原生编译(~15-25min)，产物Intel .app通吃所有Mac（含Apple Silicon通过Rosetta2）

## 已完成的优化 (v1.1-dev)
- ✅ **网络测速国内优化**: speedtest-cli优先CN服务器；HTTP多线程下载加入国内CDN（电信上海/清华TUNA/阿里云OSS），Cloudflare作为备选
- ✅ **上传测速优化**: 多端点尝试（postman-echo.com → httpbin.org），避免单点失败
- ✅ **Mac打包加速**: 移除`--collect-all openpyxl`（极度耗时），改用精确`--hidden-import`；增加timeout防止死循环
- ✅ **Bug修复**: 采集失败日志 `step_num/7` → `step_num/6`

## 运行方式
```
# 开发模式
pip install -r requirements.txt
python system_info_app.py

# 打包为独立exe（免安装）
build.bat    # Windows → dist/系统信息采集工具.exe
./build.sh   # macOS   → dist/系统信息采集工具
```

## 分发说明
- Windows: 双击 `系统信息采集工具.exe` 即可运行（SmartScreen 可能提示，点击"更多信息→仍要运行"）
- macOS: 分发格式为 `.dmg` 映像。用户双击挂载 → 将 `系统信息采集工具_Mac.app` 拖入 Applications 文件夹 → 首次启动时**右键 app → 打开**（Gatekeeper 验证，仅一次）。后续可直接双击运行
- 无需管理员权限、无需安装 Python 或任何依赖

## 用户偏好
- 返回结果使用简洁表格格式
- Excel格式偏好简单二维表（便于后续数据统计）
- 不主动发邮件，需确认后再发送
