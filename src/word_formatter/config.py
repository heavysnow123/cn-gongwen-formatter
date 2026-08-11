"""配置定义与默认值。

所有排版参数均可在界面调整并保存为 JSON。默认值面向中文报告/公文排版，
用户可自由覆盖。
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict


@dataclass
class FormatterConfig:
    # ---------- 字体（中文） ----------
    title_font: str = "黑体"
    title_size: float = 22.0
    subtitle_font: str = "黑体"
    subtitle_size: float = 16.0
    h1_font: str = "黑体"
    h1_size: float = 16.0
    h2_font: str = "楷体"
    h2_size: float = 14.0
    # 正文 / 三级 / 四级 共用一套
    body_font: str = "宋体"
    body_size: float = 12.0
    attachment_font: str = "黑体"
    attachment_size: float = 14.0
    figure_caption_font: str = "宋体"
    figure_caption_size: float = 10.5
    table_caption_font: str = "宋体"
    table_caption_size: float = 10.5

    # ---------- 字体（西文/数字） ----------
    use_custom_english_font: bool = True
    english_font: str = "Times New Roman"

    # ---------- 行距（倍数，段落类） ----------
    title_line_spacing: float = 1.3
    subtitle_line_spacing: float = 1.3
    line_spacing: float = 1.5          # 正文/标题

    # ---------- 对齐与缩进（cm） ----------
    left_indent_cm: float = 0.0
    right_indent_cm: float = 0.0
    first_line_indent_chars: int = 2   # 正文首行缩进字符数（0=不缩进）

    # ---------- 页面 ----------
    force_a4: bool = False             # 文件处理默认保持原样；文本输入强制 A4
    margin_top_cm: float = 2.54
    margin_bottom_cm: float = 2.54
    margin_left_cm: float = 2.54
    margin_right_cm: float = 2.54

    # ---------- 页码 ----------
    page_number: bool = True
    page_number_align: str = "center"  # left/center/right
    page_number_font: str = "宋体"
    page_number_size: float = 10.5
    footer_distance_cm: float = 1.75
    page_number_total: bool = False    # 显示“共 Y 页”（需 NUMPAGES 域）

    # ---------- 页眉 ----------
    header_enabled: bool = False        # 是否插入页眉
    header_text: str = ""               # 页眉文字（如文件名/单位）；为空不插
    header_align: str = "center"        # left/center/right
    header_border: bool = True          # 页眉下方加一条分隔线
    header_font: str = ""               # 空 => 沿用页码字体(page_number_font)
    header_size: float = 0.0            # 0 => 沿用页码字号(page_number_size)

    # ---------- 页脚附加文字 ----------
    footer_text: str = ""               # 页脚附加文字（如密级/文件名）；为空仅页码

    # ---------- 大纲级别（生成导航目录） ----------
    set_outline: bool = True

    # ---------- 符号标准化（实验） ----------
    normalize_punctuation: bool = False

    # ---------- 空行处理（TXT/MD） ----------
    # preserve / remove_single / keep_single
    blank_line_mode: str = "remove_single"

    # ---------- 附件格式化 ----------
    enable_attachment_formatting: bool = False

    # ---------- 表格格式化 ----------
    enable_table_formatting: bool = False
    table_font: str = "宋体"
    table_size: float = 10.5
    table_header_font: str = "黑体"
    table_header_bold: bool = True
    table_width_percent: int = 100
    table_row_height_cm: float = 0.8
    table_line_spacing: float = 12.0    # 表格行距（磅，精确值）
    table_border_size_pt: float = 0.5
    table_auto_col_width: bool = True
    table_smart_align: bool = False
    table_unified_borders: bool = True
    table_col_min_pct: int = 5
    table_col_max_pct: int = 40
    table_short_text_len: int = 6

    # ---------- 性能 / 行为 ----------
    large_folder_confirm_threshold: int = 200  # 文件夹文件数超过则提示
    large_file_threshold_mb: int = 50          # 单文件超过此 MB 给出大文件预警
    streaming_mode: bool = False               # 大文件流式模式（lxml iterparse，内存恒定）

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "FormatterConfig":
        known = {k: v for k, v in d.items() if k in cls.__dataclass_fields__}
        return cls(**known)

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "FormatterConfig":
        if not os.path.exists(path):
            return cls()
        with open(path, "r", encoding="utf-8") as f:
            return cls.from_dict(json.load(f))

    @classmethod
    def default_config_path(cls) -> str:
        return os.path.join(os.path.expanduser("~"), ".word_formatter_default.json")


# 空行模式选项（UI 用）
BLANK_LINE_MODE_OPTIONS = {
    "preserve": "不改动任何空行",
    "remove_single": "删除单个空行，多个空行保留至 1 个空行",
    "keep_single": "保留单个空行，多个空行保留至 1 个空行",
}

SUPPORTED_FILE_EXTENSIONS = (".docx", ".doc", ".wps", ".txt", ".md")
