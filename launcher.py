"""PyInstaller 入口：启动 GUI。放在项目根目录，独立于包内相对导入。"""

import sys
import os

# 确保打包后能找到 word_formatter 包
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from word_formatter.gui import main

if __name__ == "__main__":
    main()
