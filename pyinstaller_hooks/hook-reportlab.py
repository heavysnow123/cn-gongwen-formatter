"""PyInstaller 钩子：确保 reportlab 的全部数据文件（字体/CMaps）被打进单文件 EXE。

reportlab 的中文字体 CMap 数据位于 _cidfontdata.py（已随 import 自动打包），
但字体目录下的 .pfb/.afm/.ttf 等数据文件 PyInstaller 默认不收集，
补齐以避免“字体数据缺失”类运行时错误。
"""
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = collect_data_files("reportlab", include_py_files=False)
hiddenimports = collect_submodules("reportlab")
