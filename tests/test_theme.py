import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "src"))

import customtkinter as ctk
from word_formatter.gui import WordFormatterGUI

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

gui = WordFormatterGUI()  # __init__ 内部已应用科技风覆盖
gui.root.withdraw()
from customtkinter import ThemeManager
assert ThemeManager.theme["CTkButton"]["fg_color"] == ["#0096E6", "#0B6FBF"], "按钮主色未生效"
assert ThemeManager.theme["CTkSegmentedButton"]["selected_color"] == ["#0096E6", "#0B6FBF"], "标签页选中色未生效"
print("GUI THEME OK - 科技明亮风覆盖生效，窗口创建无崩溃")
gui.root.destroy()
