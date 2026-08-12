#!/usr/bin/env bash
# 中文文档智能排版工具 —— 国产系统版（统信 UOS / 银河麒麟等 Linux）构建脚本
# 产出：dist/WordFormatterPro（Linux 单文件可执行，无需安装 Python）
#
# 用法（在目标架构机上执行，例如 x86_64 或 ARM64 / LoongArch 的国产机）：
#   chmod +x build_linux.sh
#   ./build_linux.sh
#
# 说明：
#   - 与 Windows 版（build.py 产出 WordFormatterPro.exe）共用同一套源码（src/），
#     功能与界面完全一致，仅在标题 / 关于中标注「国产系统版」。
#   - 需在对应架构的本机上构建：PyInstaller 不支持交叉编译，ARM64 机器打 ARM64 包。
#   - 如需免安装单文件 AppImage，可用 linuxdeploy 对 dist/WordFormatterPro 再封装。
set -e

cd "$(dirname "$0")"
SCRIPT_DIR="$(pwd)"

if ! command -v python3 >/dev/null 2>&1; then
  echo "需要 python3，请先安装： sudo apt install python3 python3-venv python3-pip python3-tk"
  exit 1
fi
if ! python3 -c "import tkinter" >/dev/null 2>&1; then
  echo "缺少 tkinter，请先安装： sudo apt install python3-tk"
  exit 1
fi

# 虚拟环境 + 运行 / 构建依赖（pywin32 因平台标记会自动跳过）
if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install --upgrade pip
fi
.venv/bin/pip install "python-docx>=1.1" "customtkinter>=5.2" "reportlab>=4.0" "pyinstaller>=6.0"

# 构建（单文件、无终端窗口）。Linux 上不加 win32com 相关 hidden-import。
.venv/bin/pyinstaller \
  --name WordFormatterPro \
  --windowed \
  --onefile \
  --paths "$SCRIPT_DIR/src" \
  --hidden-import reportlab \
  --hidden-import reportlab.pdfbase.cidfonts \
  --hidden-import reportlab.pdfbase.ttfonts \
  --hidden-import customtkinter \
  --additional-hooks-dir "$SCRIPT_DIR/pyinstaller_hooks" \
  --noconfirm \
  --clean \
  "$SCRIPT_DIR/launcher.py"

echo "国产系统版构建完成：dist/WordFormatterPro"
