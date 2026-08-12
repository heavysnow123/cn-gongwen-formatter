#!/usr/bin/env bash
# 中文文档智能排版工具 —— Linux / 国产系统（统信 UOS、银河麒麟等）一键启动脚本
# 用法：
#   chmod +x run_linux.sh
#   ./run_linux.sh
set -e

cd "$(dirname "$0")"
SCRIPT_DIR="$(pwd)"

# 1) 检查 python3
if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3，请先安装："
  echo "  Debian/Ubuntu/UOS/麒麟:  sudo apt install python3 python3-venv python3-pip python3-tk"
  exit 1
fi

# 2) 检查 tkinter 运行库（customtkinter 依赖 _tkinter）
if ! python3 -c "import tkinter" >/dev/null 2>&1; then
  echo "缺少 tkinter 运行库，请先安装 python3-tk："
  echo "  sudo apt install python3-tk"
  exit 1
fi

# 3) 创建虚拟环境并安装运行依赖
if [ ! -d .venv ]; then
  echo "首次运行：创建虚拟环境并安装依赖……"
  python3 -m venv .venv
  .venv/bin/pip install --upgrade pip
  # pywin32 在 Linux 上会自动跳过；pyinstaller / pytest 仅开发打包用，此处不装
  .venv/bin/pip install "python-docx>=1.1" "customtkinter>=5.2" "reportlab>=4.0"
fi

# 4) 启动图形界面（src 布局，需将 src 加入 PYTHONPATH）
echo "启动中文文档智能排版工具……"
PYTHONPATH="$SCRIPT_DIR/src" .venv/bin/python -m word_formatter.gui
