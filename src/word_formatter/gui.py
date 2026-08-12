"""Word Formatter 现代图形界面（CustomTkinter）。

相比原 tkinter 界面的升级：
- 现代化控件与主题；
- 标签页分区参数（字体 / 页面 / 表格 / 其他），带 ? 提示；
- 拖拽添加（tkinterdnd2 不可用时自动回退为按钮）；
- 后台线程处理 + 进度条 + 实时日志；
- 配置保存/加载/恢复默认。
"""

from __future__ import annotations

import os
import platform
import re
import sys
import threading

try:
    import customtkinter as ctk
except ImportError:
    # 友好提示而非堆栈：Linux / 国产系统若未装 customtkinter 时给出明确指引
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _root = _tk.Tk()
    _root.withdraw()
    _mb.showerror(
        "缺少依赖",
        "未找到 customtkinter，请先安装：\n\n"
        "  pip install customtkinter\n\n"
        "国产系统（统信 UOS / 银河麒麟等）也可直接运行项目根目录的\n"
        "run_linux.sh 一键启动（会自动安装依赖）。",
    )
    _root.destroy()
    raise SystemExit(1)


# 产品标识（Windows 版 / 国产系统版共用同一套代码，功能与界面统一）
APP_NAME = "中文文档智能排版工具"
APP_VERSION = "1.2.0"


def platform_label() -> str:
    """返回当前运行版本的平台标注，用于标题 / 关于显示。"""
    if sys.platform.startswith("win"):
        return "Windows 版"
    if sys.platform == "darwin":
        return "macOS 版"
    return "国产系统版"


def platform_arch() -> str:
    """返回 CPU 架构信息，如 x86_64 / aarch64 / loongarch64。"""
    return platform.machine() or "unknown"

from .config import (
    FormatterConfig, BLANK_LINE_MODE_OPTIONS, SUPPORTED_FILE_EXTENSIONS,
    PROCESS_MODE_OPTIONS, NUMBERING_STYLE_OPTIONS,
)

PRESET_FONTS = [
    "宋体", "楷体", "仿宋", "微软雅黑", "等线", "华文仿宋", "华文楷体",
    "SimSun", "KaiTi", "FangSong", "Microsoft YaHei",
    "Times New Roman", "Arial", "Calibri", "Consolas",
]

# 中文字体下拉（公文优先：仿宋_GB2312 / 楷体_GB2312 置顶）
PRESET_CN_FONTS = [
    "仿宋_GB2312", "仿宋", "楷体_GB2312", "楷体", "思源宋体",
    "宋体", "微软雅黑", "等线",
    "华文仿宋", "华文楷体", "方正仿宋", "方正楷体",
    "SimSun", "KaiTi", "FangSong", "Microsoft YaHei",
]

# ---------------- Word 字号（与 Word“开始”选项卡完全一致） ----------------
# 中文命名字号 -> 磅值
WORD_NAMED_SIZES = [
    ("初号", 42.0), ("小初", 36.0), ("一号", 26.0), ("小一", 24.0),
    ("二号", 22.0), ("小二", 18.0), ("三号", 16.0), ("小三", 15.0),
    ("四号", 14.0), ("小四", 12.0), ("五号", 10.5), ("小五", 9.0),
    ("六号", 7.5), ("小六", 6.5), ("七号", 5.5), ("八号", 5.0),
]
# Word 字号下拉里的数字档（磅）
WORD_PT_SIZES = [5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18,
                 20, 22, 24, 26, 28, 36, 48, 72]


def _word_size_label(pt: float) -> str:
    """磅值 -> 下拉显示文本（命中命名字号显示“初号（42 磅）”等）。"""
    for name, p in WORD_NAMED_SIZES:
        if abs(p - pt) < 1e-6:
            return f"{name}（{p:g} 磅）"
    return f"{pt:g} 磅"


def _word_size_options() -> list[str]:
    """字号下拉候选项：命名字号优先，再补数字档（去重）。"""
    named = [f"{name}（{p:g} 磅）" for name, p in WORD_NAMED_SIZES]
    seen = {p for _, p in WORD_NAMED_SIZES}
    pts = [f"{p:g} 磅" for p in WORD_PT_SIZES if p not in seen]
    return named + pts


def _parse_word_size(text: str) -> float:
    """从下拉文本/手输文本解析出磅值；无法解析返回 0.0。"""
    m = re.search(r"(\d+(?:\.\d+)?)", text or "")
    return float(m.group(1)) if m else 0.0


try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
except Exception:  # pragma: no cover
    _DND_AVAILABLE = False


class WordFormatterGUI:
    def __init__(self):
        # 界面主题（亮/暗）来自持久化配置；必须在建窗前设置
        self.cfg = FormatterConfig.load(FormatterConfig.default_config_path())
        mode = self.cfg.appearance_mode
        if mode not in ("Light", "Dark"):
            mode = "Light"
        ctk.set_appearance_mode(mode)
        ctk.set_default_color_theme("blue")
        self._apply_tech_theme()

        # 高分屏（HiDPI）清晰渲染（Windows）
        try:
            ctk.WindowsDPIAware()
        except Exception:
            pass

        if _DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = ctk.CTk()
        self.root.title(f"{APP_NAME} v{APP_VERSION} · {platform_label()}")
        self.root.geometry("1080x760")
        self.root.minsize(900, 640)

        self.file_list: list[dict] = []
        self._running = False
        self.widgets: dict[str, object] = {}
        self._log_queue: list[tuple[str, str]] = []

        self._build_top()
        self._build_main()
        self._build_log()
        self._apply_config(self.cfg)
        self._refresh_file_list()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(200, self._drain_log)

    # ---------------- 科技风主题覆盖 ----------------
    def _apply_tech_theme(self):
        """在内置 blue 主题基础上覆盖科技明亮风配色，保留全部默认键完整。"""
        from customtkinter import ThemeManager
        t = ThemeManager.theme
        # 窗口/顶层背景：浅蓝白
        t["CTk"]["fg_color"] = ["#EEF4FB", "#0F1B2D"]
        t["CTkToplevel"]["fg_color"] = ["#EEF4FB", "#0F1B2D"]
        # 面板（白）/ 顶层面板（更浅）
        t["CTkFrame"]["fg_color"] = ["#FFFFFF", "#16243A"]
        t["CTkFrame"]["top_fg_color"] = ["#F4F8FD", "#1B2A40"]
        t["CTkFrame"]["border_color"] = ["#DCE7F4", "#243349"]
        # 按钮：科技蓝
        t["CTkButton"]["fg_color"] = ["#0096E6", "#0B6FBF"]
        t["CTkButton"]["hover_color"] = ["#0B7FC9", "#0A5FA6"]
        t["CTkButton"]["border_color"] = ["#0096E6", "#0B6FBF"]
        # 复选框：科技蓝
        t["CTkCheckBox"]["fg_color"] = ["#0096E6", "#0B6FBF"]
        t["CTkCheckBox"]["hover_color"] = ["#0B7FC9", "#0A5FA6"]
        t["CTkCheckBox"]["border_color"] = ["#0096E6", "#0B6FBF"]
        # 标签页分段（CTkSegmentedButton）
        t["CTkSegmentedButton"]["selected_color"] = ["#0096E6", "#0B6FBF"]
        t["CTkSegmentedButton"]["selected_hover_color"] = ["#0B7FC9", "#0A5FA6"]
        t["CTkSegmentedButton"]["unselected_color"] = ["#E8F1FA", "#1B2A40"]
        t["CTkSegmentedButton"]["unselected_hover_color"] = ["#D6E6F5", "#243349"]
        t["CTkSegmentedButton"]["fg_color"] = ["#E8F1FA", "#1B2A40"]
        t["CTkSegmentedButton"]["text_color"] = ["#0B2239", "#DCE7F4"]
        # 进度条：青色
        t["CTkProgressBar"]["progress_color"] = ["#00C2C7", "#0B6FBF"]
        # 滚动框标签条
        t["CTkScrollableFrame"]["label_fg_color"] = ["#0096E6", "#0B6FBF"]

    # ---------------- 顶部栏 ----------------
    def _build_top(self):
        f = ctk.CTkFrame(self.root)
        f.pack(fill="x", padx=12, pady=(12, 0))
        ctk.CTkLabel(f, text=APP_NAME, font=ctk.CTkFont(size=20, weight="bold"),
                     text_color="#0B2239").pack(side="left", padx=8)
        ctk.CTkLabel(f, text="一键式中文文档智能排版",
                     text_color="#0096E6").pack(side="left", padx=4)

        for label, cmd in (
            ("+ 添加文件", self.add_files),
            ("+ 文件夹", self.add_folder),
            ("目录树", self.add_tree),
            ("开始排版", self.start_processing),
            ("保存配置", self.save_config),
            ("加载配置", self.load_config),
            ("恢复默认", self.reset_defaults),
            ("使用说明", self.show_help),
        ):
            b = ctk.CTkButton(f, text=label, width=92,
                              fg_color=("#0096E6", "#0B6FBF") if label == "开始排版" else None,
                              command=cmd)
            b.pack(side="right", padx=4)

        ctk.CTkFrame(f, height=3, fg_color="#00C2C7").pack(
            fill="x", side="bottom", pady=(8, 0))

    # ---------------- 主区 ----------------
    def _build_main(self):
        main = ctk.CTkFrame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=6)
        main.grid_columnconfigure(0, weight=1, minsize=300)
        main.grid_columnconfigure(1, weight=2)
        main.grid_rowconfigure(0, weight=1)

        # 左：文件列表
        left = ctk.CTkFrame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ctk.CTkLabel(left, text="待处理文件", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=10, pady=(8, 2))
        self.file_box = ctk.CTkScrollableFrame(left)
        self.file_box.pack(fill="both", expand=True, padx=8, pady=4)
        self.drop_hint = ctk.CTkLabel(
            left,
            text="拖拽文件/文件夹到这里\n或点顶部“添加文件 / 文件夹 / 目录树”",
            text_color=("gray50", "gray60"), justify="center")
        self.drop_hint.pack(fill="x", padx=8, pady=8)
        if _DND_AVAILABLE:
            try:
                self.file_box.drop_target_register(DND_FILES)
                self.file_box.dnd_bind("<<Drop>>", self.handle_drop)
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self.handle_drop)
            except Exception:
                self.drop_hint.configure(text="拖拽不可用，请用按钮添加")
        row = ctk.CTkFrame(left)
        row.pack(fill="x", padx=8, pady=6)
        ctk.CTkButton(row, text="移除选中", width=90, command=self.remove_file).pack(side="left", padx=4)
        ctk.CTkButton(row, text="清空列表", width=90, command=self.clear_list).pack(side="left", padx=4)

        # 左侧底部：快捷工具
        tools = ctk.CTkFrame(left)
        tools.pack(fill="x", padx=8, pady=(2, 8))
        ctk.CTkLabel(tools, text="快捷工具", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=6, pady=(2, 0))
        trow = ctk.CTkFrame(tools)
        trow.pack(fill="x", padx=4, pady=4)
        ctk.CTkButton(trow, text="模板", width=88, command=self.open_template).pack(side="left", padx=4)
        ctk.CTkButton(trow, text="排版质检", width=88, command=self.open_check).pack(side="left", padx=4)
        ctk.CTkButton(trow, text="导出PDF", width=88, command=self.export_pdf_action).pack(side="left", padx=4)
        self.use_builtin_pdf = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(tools, text="优先用内置引擎（不依赖 Office）",
                        variable=self.use_builtin_pdf).pack(anchor="w", padx=6, pady=(0, 4))

    # 右：参数标签页
        right = ctk.CTkFrame(main)
        right.grid(row=0, column=1, sticky="nsew")
        self.tab = ctk.CTkTabview(right)
        self.tab.pack(fill="both", expand=True, padx=8, pady=8)
        for name in ("字体", "页面", "表格", "其他"):
            self.tab.add(name)
            sf = ctk.CTkScrollableFrame(self.tab.tab(name))
            sf.pack(fill="both", expand=True, padx=4, pady=4)
            setattr(self, f"tab_{name}", sf)

        self._build_font_tab()
        self._build_page_tab()
        self._build_table_tab()
        self._build_other_tab()

    # ---------------- 字体页 ----------------
    # 字体 + 字号成对（与 Word“开始”选项卡一致：字体名框 + 字号框 并排）
    FONT_SIZE_PAIRS = [
        ("title_font", "title_size", "主标题"),
        ("subtitle_font", "subtitle_size", "副标题"),
        ("h1_font", "h1_size", "一级标题"),
        ("h2_font", "h2_size", "二级标题"),
        ("body_font", "body_size", "正文/三四级"),
        ("attachment_font", "attachment_size", "附件标识"),
        ("figure_caption_font", "figure_caption_size", "图/表标题"),
        ("table_caption_font", "table_caption_size", "表格标题"),
    ]
    # 仅字体的字段（其字号在“表格”页统一管理）
    FONT_ONLY_FIELDS = [
        ("table_font", "表格正文字体"),
        ("table_header_font", "表头字体"),
    ]

    def _font_size_row(self, parent, font_key, size_key, label):
        """字体 + 字号并排一行（Word 风格）。"""
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=4, pady=2)
        ctk.CTkLabel(row, text=label, width=110, anchor="w").pack(side="left", padx=4)
        fw = ctk.CTkComboBox(row, values=PRESET_CN_FONTS, width=150)
        fw.pack(side="left", padx=(0, 6))
        self.widgets[font_key] = fw
        sw = ctk.CTkComboBox(row, values=_word_size_options(), width=120)
        sw.pack(side="left", padx=0)
        self.widgets[size_key] = (sw, "size", None)

    def _build_font_tab(self):
        sf = self.tab_字体
        ctk.CTkLabel(sf, text="标题与正文（字体/字号均可下拉选择，也可直接输入）",
                     font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(4, 2))
        for font_key, size_key, label in self.FONT_SIZE_PAIRS:
            self._font_size_row(sf, font_key, size_key, label)
        for key, label in self.FONT_ONLY_FIELDS:
            self._field(sf, label, key, font_combo=True, preset=PRESET_CN_FONTS)

        ctk.CTkLabel(sf, text="西文字体", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(8, 2))
        self._check(sf, "自定义数字/字母字体", "use_custom_english_font")
        self._field(sf, "西文字体名称", "english_font", font_combo=True)

        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=8)
        ctk.CTkLabel(sf, text="标题西文字体（可选，留空=沿用上方全局西文字体）",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(2, 2))
        enf = ctk.CTkFrame(sf)
        enf.pack(fill="x", padx=4, pady=2)
        for key, label in (
            ("title_en_font", "主标题"), ("subtitle_en_font", "副标题"),
            ("h1_en_font", "一级标题"), ("h2_en_font", "二级标题"),
            ("h3_en_font", "三/四级标题"), ("h4_en_font", "四级标题"),
        ):
            self._field(enf, label, key, font_combo=True, preset=PRESET_FONTS)

        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=8)
        ctk.CTkLabel(sf, text="公文字体缺失检查", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(2, 2))
        ctk.CTkLabel(
            sf, text="公文规范（GB/T 9704-2012）常用字体：\n"
                     "• 必备·缺失需自行安装：仿宋_GB2312（正文）、楷体_GB2312（二级标题/签发人）\n"
                     "• 大标题（二号·加粗）：思源宋体（开源 SIL OFL，可免费获取并安装）\n"
                     "本工具不安装字体、不改系统；点下方按钮检查本机字体安装情况。",
            text_color=("gray50", "gray60"), justify="left").pack(anchor="w", padx=4, pady=(0, 4))
        self.font_status = ctk.CTkLabel(
            sf, text="点击右侧按钮检查字体",
            text_color=("gray50", "gray60"), justify="left")
        self.font_status.pack(anchor="w", padx=4, pady=(2, 4))
        ctk.CTkButton(sf, text="检查字体状态", width=180,
                      command=self._check_fonts).pack(anchor="e", padx=4, pady=(2, 6))

    # ---------------- 页面页 ----------------
    def _build_page_tab(self):
        sf = self.tab_页面
        self._check(sf, "强制设置为 A4 纸张（文件处理时修改原页面尺寸）", "force_a4")
        for key, label in (
            ("margin_top_cm", "上边距 (cm)"), ("margin_bottom_cm", "下边距 (cm)"),
            ("margin_left_cm", "左边距 (cm)"), ("margin_right_cm", "右边距 (cm)"),
        ):
            self._field(sf, label, key)

        ctk.CTkLabel(sf, text="页码与页脚", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(10, 2))
        self._check(sf, "插入页码", "page_number")
        self._combo(sf, "页码对齐", "page_number_align",
                    ["left", "center", "right"], labels={"left": "左", "center": "居中", "right": "右"})
        self._field(sf, "页脚附加文字", "footer_text")
        self._check(sf, "显示总页数（“第 X 页 / 共 Y 页”）", "page_number_total")
        self._field(sf, "页码/页脚字体", "page_number_font", font_combo=True, preset=PRESET_CN_FONTS)
        self._field(sf, "页码/页脚字号", "page_number_size", size_combo=True)
        self._field(sf, "页脚距 (cm)", "footer_distance_cm")

        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=8)
        ctk.CTkLabel(sf, text="页眉", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(2, 2))
        self._check(sf, "启用页眉", "header_enabled")
        self._field(sf, "页眉文字", "header_text")
        self._combo(sf, "页眉对齐", "header_align",
                    ["left", "center", "right"], labels={"left": "左", "center": "居中", "right": "右"})
        self._check(sf, "页眉下方加分隔线", "header_border")
        self._field(sf, "页眉字体（留空沿用上方）", "header_font", font_combo=True, preset=PRESET_CN_FONTS)
        self._field(sf, "页眉字号（0=沿用上方）", "header_size", size_combo=True)

        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=8)
        ctk.CTkLabel(sf, text="页面背景", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(2, 2))
        self._field(sf, "背景色 (hex，如 F2F2F2，留空不设置)", "page_background_color")

    # ---------------- 表格页 ----------------
    def _build_table_tab(self):
        sf = self.tab_表格
        self._check(sf, "启用表格自动调整（总开关）", "enable_table_formatting")
        self._check(sf, "统一表格边框", "table_unified_borders")
        self._check(sf, "自动调整列宽", "table_auto_col_width")
        self._check(sf, "智能调整单元格对齐", "table_smart_align")
        self._check(sf, "表头加粗", "table_header_bold")
        for key, label in (
            ("table_header_font", "表头字体"), ("table_font", "表格正文字体"),
            ("table_size", "表格正文/表头字号"),
            ("table_caption_font", "表格标题字体"), ("table_caption_size", "表格标题字号"),
            ("table_width_percent", "表格宽度 (%)"),
            ("table_row_height_cm", "行高 (cm)"),
            ("table_line_spacing", "行距 (磅)"),
            ("table_border_size_pt", "边框粗细 (pt)"),
            ("table_col_min_pct", "列最小宽度 (%)"),
            ("table_col_max_pct", "列最大宽度 (%)"),
            ("table_short_text_len", "短文本阈值 (字符)"),
        ):
            if key.endswith("_size"):
                self._field(sf, label, key, size_combo=True)
            else:
                self._field(sf, label, key)

    # ---------------- 其他页 ----------------
    def _build_other_tab(self):
        sf = self.tab_其他
        # 处理模式（全量 / 仅修标点）
        self._combo(sf, "处理模式", "process_mode",
                    list(PROCESS_MODE_OPTIONS.keys()),
                    labels=PROCESS_MODE_OPTIONS)
        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=6)
        ctk.CTkLabel(sf, text="标点与序号", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(2, 2))
        self._check(sf, "标点全半角标准化（中英文标点统一为全角）", "normalize_punctuation")
        self._check(sf, "序号风格统一（1. / 1、 / （一）等归一）", "unify_numbering")
        self._combo(sf, "序号目标风格", "numbering_style",
                    list(NUMBERING_STYLE_OPTIONS.keys()),
                    labels=NUMBERING_STYLE_OPTIONS)
        self._check(sf, "中文换行禁则（避头/避尾 + 启用 Word kinsoku）", "cjk_linebreak_rules")
        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=6)
        # 界面主题（亮/暗），切换即时生效并持久化
        self._combo(sf, "界面主题", "appearance_mode",
                    ["Light", "Dark"],
                    labels={"Light": "浅色", "Dark": "深色"},
                    command=self._on_theme_change)
        ctk.CTkFrame(sf, height=2, fg_color=("#DCE7F4", "#243349")).pack(fill="x", pady=6)
        self._field(sf, "正文行距 (倍数)", "line_spacing")
        self._field(sf, "首行缩进 (字符)", "first_line_indent_chars")
        self._check(sf, "自动设置大纲级别（生成导航目录）", "set_outline")
        self._check(sf, "启用附件格式化", "enable_attachment_formatting")
        self._check(sf, "大文件流式模式（极低内存，超大文档建议开启；自动注入页眉页脚/页码）", "streaming_mode")
        self._combo(sf, "TXT/MD 空行处理", "blank_line_mode",
                    list(BLANK_LINE_MODE_OPTIONS.keys()),
                    labels=BLANK_LINE_MODE_OPTIONS)

        ctk.CTkLabel(sf, text="直接输入文本（强制 A4）", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", pady=(8, 2))
        self.text_input = ctk.CTkTextbox(sf, height=120)
        self.text_input.pack(fill="x", padx=4, pady=4)
        ctk.CTkButton(sf, text="排版上述文本", width=120,
                      command=self.process_text).pack(anchor="e", padx=4, pady=4)

    # ---------------- 字段构造 ----------------
    def _field(self, parent, label, key, font_combo=False, preset=None, size_combo=False):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=4, pady=2)
        ctk.CTkLabel(row, text=label, width=150, anchor="w").pack(side="left", padx=4)
        if font_combo:
            w = ctk.CTkComboBox(row, values=preset if preset is not None else PRESET_FONTS,
                                width=170)
            self.widgets[key] = w
        elif size_combo:
            # 字号下拉：命名字号 + 磅值，可手输任意磅值
            w = ctk.CTkComboBox(row, values=_word_size_options(), width=170)
            self.widgets[key] = (w, "size", None)
        else:
            w = ctk.CTkEntry(row, width=170)
            self.widgets[key] = w
        w.pack(side="right", padx=4)

    def _check(self, parent, label, key):
        w = ctk.CTkCheckBox(parent, text=label)
        w.pack(anchor="w", padx=6, pady=3)
        self.widgets[key] = w

    def _combo(self, parent, label, key, values, labels=None, command=None):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=4, pady=2)
        ctk.CTkLabel(row, text=label, width=150, anchor="w").pack(side="left", padx=4)
        disp = [labels.get(v, v) if labels else v for v in values]
        w = ctk.CTkComboBox(row, values=disp, width=200,
                            command=(lambda _v=None: command()) if command else None)
        w.pack(side="right", padx=4)
        self.widgets[key] = (w, values, labels)

    # ---------------- 弹窗 / 主题 ----------------
    def _popup(self, title, w, h):
        """统一的子窗口创建：标题栏 + 尺寸 + 跟随主窗。"""
        win = ctk.CTkToplevel(self.root)
        win.title(title)
        win.geometry(f"{w}x{h}")
        win.transient(self.root)
        return win

    def _apply_theme(self):
        """按配置文件里的 appearance_mode 应用界面亮/暗主题。"""
        mode = self.cfg.appearance_mode
        if mode not in ("Light", "Dark"):
            mode = "Light"
        ctk.set_appearance_mode(mode)

    def _on_theme_change(self):
        """主题下拉变更：实时切换并持久化。"""
        w, values, labels = self.widgets.get("appearance_mode", (None, None, None))
        if w is None:
            return
        disp = w.get()
        internal = next((v for v in values
                         if (labels.get(v, v) if labels else v) == disp), "Light")
        if internal not in ("Light", "Dark"):
            internal = "Light"
        ctk.set_appearance_mode(internal)
        self.cfg.appearance_mode = internal
        self.cfg.save(FormatterConfig.default_config_path())
        self._log(f"已切换为「{'深色' if internal == 'Dark' else '浅色'}」主题")

    # ---------------- 配置读写 ----------------
    def _apply_config(self, cfg: FormatterConfig):
        d = cfg.to_dict()
        for key, w in self.widgets.items():
            val = d.get(key)
            if val is None:
                continue
            if isinstance(w, tuple):  # combo 或 size
                cw = w[0]
                if w[1] == "size":
                    pt = float(val)
                    cw.set("0（沿用上方）" if pt == 0 else _word_size_label(pt))
                else:
                    values, labels = w[1], w[2]
                    idx = values.index(val) if val in values else 0
                    cw.set(labels.get(values[idx], values[idx]) if labels else values[idx])
            elif isinstance(w, ctk.CTkCheckBox):
                if val:
                    w.select()
                else:
                    w.deselect()
            elif isinstance(w, ctk.CTkComboBox):
                w.set(val)
            else:
                w.delete(0, "end")
                w.insert(0, str(val))

    def _collect_config(self) -> FormatterConfig:
        d = {}
        for key, w in self.widgets.items():
            if isinstance(w, tuple):
                cw = w[0]
                if w[1] == "size":
                    d[key] = _parse_word_size(cw.get())
                else:
                    values, labels = w[1], w[2]
                    disp = cw.get()
                    val = next((v for v in values if (labels.get(v, v) if labels else v) == disp), values[0])
                    d[key] = val
            elif isinstance(w, ctk.CTkCheckBox):
                d[key] = bool(w.get())
            elif isinstance(w, ctk.CTkComboBox):
                d[key] = w.get()
            else:
                txt = w.get().strip()
                # 数值字段
                orig = FormatterConfig.__dataclass_fields__[key].type
                if "float" in str(orig):
                    try:
                        d[key] = float(txt)
                    except ValueError:
                        d[key] = 0.0
                elif "int" in str(orig):
                    try:
                        d[key] = int(float(txt))
                    except ValueError:
                        d[key] = 0
                else:
                    d[key] = txt
        return FormatterConfig.from_dict(d)

    def save_config(self):
        from tkinter import filedialog
        p = filedialog.asksaveasfilename(defaultextension=".json",
                                         filetypes=[("JSON", "*.json")])
        if not p:
            return
        self._collect_config().save(p)
        self.cfg.save(FormatterConfig.default_config_path())
        self._log(f"配置已保存：{p}")

    def load_config(self):
        from tkinter import filedialog
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            cfg = FormatterConfig.load(p)
            self.cfg = cfg
            self._apply_config(cfg)
            self._apply_theme()
            self._log(f"已加载配置：{p}")
        except Exception as e:
            self._log(f"加载配置失败：{e}", "ERROR")

    def reset_defaults(self):
        self.cfg = FormatterConfig()
        self._apply_config(self.cfg)
        self._apply_theme()
        self.cfg.save(FormatterConfig.default_config_path())
        self._log("已恢复内置默认配置")

    # ---------------- 文件列表 ----------------
    def _refresh_file_list(self):
        for w in self.file_box.winfo_children():
            w.destroy()
        if not self.file_list:
            self.drop_hint.pack(fill="x", padx=8, pady=8)
            return
        self.drop_hint.pack_forget()
        for i, entry in enumerate(self.file_list):
            row = ctk.CTkFrame(self.file_box)
            row.pack(fill="x", padx=2, pady=1)
            ctk.CTkLabel(row, text=f"{i+1}. {entry['rel']}",
                         anchor="w").pack(side="left", fill="x", expand=True, padx=4)
            ctk.CTkButton(row, text="✕", width=28, fg_color=("gray70", "gray30"),
                          command=lambda idx=i: self._remove_index(idx)).pack(side="right", padx=2)

    def _remove_index(self, idx):
        if 0 <= idx < len(self.file_list):
            self.file_list.pop(idx)
            self._refresh_file_list()

    def add_files(self):
        from tkinter import filedialog
        files = filedialog.askopenfilenames(
            filetypes=[("文档", "*.docx *.doc *.wps *.txt *.md"), ("全部", "*.*")])
        self._add_paths(list(files))

    def add_folder(self):
        """添加文件夹（仅当前目录，不递归）。"""
        from tkinter import filedialog
        d = filedialog.askdirectory()
        if d:
            files = [os.path.join(d, f) for f in os.listdir(d)
                     if f.lower().endswith(SUPPORTED_FILE_EXTENSIONS)]
            self._add_paths(files)

    def add_tree(self):
        """添加目录树（递归收纳所有子目录中的文档，输出时保持相对目录结构）。"""
        from tkinter import filedialog
        d = filedialog.askdirectory(title="选择目录（递归收纳全部子目录文档）")
        if not d:
            return
        root = os.path.abspath(d)
        collected = []
        for cur, _dirs, files in os.walk(root):
            for f in files:
                if f.lower().endswith(SUPPORTED_FILE_EXTENSIONS):
                    collected.append(os.path.join(cur, f))
        if not collected:
            self._log(f"目录中未找到支持的文档：{d}", "WARN")
            return
        self._add_paths(collected, base=root)

    def handle_drop(self, event):
        files = self.root.tk.splitlist(event.data) if _DND_AVAILABLE else []
        self._add_paths([f for f in files if f.lower().endswith(SUPPORTED_FILE_EXTENSIONS)])

    def _add_paths(self, paths, base=None):
        added = 0
        for p in paths:
            ap = os.path.abspath(p)
            rel = os.path.relpath(ap, base) if base else os.path.basename(ap)
            if any(e["path"] == ap for e in self.file_list):
                continue
            self.file_list.append({"path": ap, "rel": rel})
            added += 1
        if added:
            self._refresh_file_list()
            self._log(f"已添加 {added} 个文件")

    def remove_file(self):
        # 简化：移除最后一个（真实选择需绑定列表选择）；此处移除选中通过行内✕完成
        self._log("请用每行右侧 ✕ 移除文件")

    def clear_list(self):
        self.file_list.clear()
        self._refresh_file_list()
        self._log("已清空文件列表")

    # ---------------- 处理 ----------------
    def start_processing(self):
        if self._running:
            self._log("任务正在进行中", "WARN")
            return
        if not self.file_list:
            self._log("文件列表为空，请先添加文件", "WARN")
            return
        from tkinter import filedialog
        out = filedialog.askdirectory(title="选择输出文件夹")
        if not out:
            return
        cfg = self._collect_config()
        self.cfg = cfg
        cfg.save(FormatterConfig.default_config_path())
        self._preflight_fonts()
        self._running = True
        self.progress.set(0)
        threading.Thread(target=self._worker, args=(out, cfg), daemon=True).start()

    def _render_font_report(self):
        """刷新字体状态区：列出全部公文字体及 ✅/⚠️/ℹ️ 状态。"""
        try:
            from .fonts import gongwen_font_report
            rows = gongwen_font_report()
        except Exception:
            return
        lines = []
        for r in rows:
            if r["installed"]:
                mark = "✅"
            elif r["required"]:
                mark = "⚠️"
            else:
                mark = "ℹ️"
            lines.append(f"{mark} {r['display']}（{r['category']}）")
        self.font_status.configure(text="\n".join(lines))
        required_missing = [r for r in rows if r["required"] and not r["installed"]]
        commercial_missing = [r for r in rows if not r["required"] and not r["installed"]]
        if required_missing:
            color = "#C0392B"
        elif commercial_missing:
            color = "#B7791F"
        else:
            color = "#1A9E5B"
        self.font_status.configure(text_color=color)

    def _preflight_fonts(self):
        """排版前预检公文必备字体：仅检测缺失并提示用户自行安装（不安装、不改系统）。"""
        try:
            from .fonts import check_gongwen_fonts, gongwen_font_report
            self._render_font_report()
            missing = check_gongwen_fonts()
            if not missing:
                return
            self._log("⚠ 检测到缺失的公文字体，排版可能回退到替代字体导致版式不达标。",
                      "WARN")
            for gf in missing:
                note = f"（{gf.get('hint', '')}）" if gf.get("hint") else ""
                self._log(
                    f"⚠ 排版可能需要 {gf['display']} 但系统未安装{note}。请自行安装到系统后再排版。",
                    "WARN")
        except Exception as e:  # 字体预检不应阻塞排版
            self._log(f"字体预检跳过：{e}", "WARN")

    def _check_fonts(self):
        """按钮：检查公文字体安装状态并刷新报告（不安装、不改系统）。"""
        try:
            from .fonts import gongwen_font_report, check_gongwen_fonts
        except Exception as e:
            self._log(f"字体模块加载失败：{e}", "ERROR")
            return
        self._render_font_report()
        missing = check_gongwen_fonts()
        if not missing:
            self._log("✅ 公文必备字体均已安装", "INFO")
        else:
            for gf in missing:
                note = f"（{gf.get('hint', '')}）" if gf.get("hint") else ""
                self._log(
                    f"⚠ {gf['display']} 未安装{note}。请自行安装到系统后再排版。",
                    "WARN")

    def _worker(self, out, cfg):
        from .core import WordFormatter
        fmt = WordFormatter(cfg, log_cb=self._queue_log)
        total = len(self.file_list)
        ok = skip = err = 0
        for i, entry in enumerate(self.file_list):
            src = entry["path"]
            rel_dir = os.path.dirname(entry["rel"])
            target_dir = out if not rel_dir else os.path.join(out, rel_dir)
            r = fmt.format_file(src, target_dir)
            if r["skipped"]:
                skip += 1
            elif r["error"]:
                err += 1
            else:
                ok += 1
            self.root.after(0, lambda v=(i + 1) / total: self.progress.set(v))
        self._queue_log(f"\n🎉 完成：成功 {ok} / 跳过 {skip} / 失败 {err}", "INFO")
        self._running = False

    def process_text(self):
        if self._running:
            self._log("任务正在进行中", "WARN")
            return
        text = self.text_input.get("1.0", "end").strip()
        if not text:
            self._log("文本框内容为空", "WARN")
            return
        from tkinter import filedialog
        out = filedialog.asksaveasfilename(defaultextension=".docx",
                                           filetypes=[("Word", "*.docx")])
        if not out:
            return
        cfg = self._collect_config()
        from .core import WordFormatter
        fmt = WordFormatter(cfg, log_cb=self._queue_log)
        r = fmt.format_text(text, out, is_md=False)
        if r["error"]:
            self._log(f"❌ 失败：{r['error']}", "ERROR")

    # ---------------- 日志 ----------------
    def _build_log(self):
        f = ctk.CTkFrame(self.root)
        f.pack(fill="x", padx=12, pady=(0, 10))
        self.progress = ctk.CTkProgressBar(f)
        self.progress.pack(fill="x", padx=4, pady=(4, 2))
        self.progress.set(0)
        self.log_box = ctk.CTkTextbox(f, height=120, state="disabled")
        self.log_box.pack(fill="x", padx=4, pady=(2, 4))

    def _queue_log(self, msg, level="INFO"):
        self._log_queue.append((msg, level))
        self.root.after(0, self._drain_log)

    def _drain_log(self):
        while self._log_queue:
            msg, level = self._log_queue.pop(0)
            self._log(msg, level)
        self.root.after(200, self._drain_log)

    def _log(self, msg, level="INFO"):
        self.log_box.configure(state="normal")
        tag = {"ERROR": "❌ ", "WARN": "⚠ ", "INFO": ""}.get(level, "")
        self.log_box.insert("end", f"{tag}{msg}\n")
        self.log_box.configure(state="disabled")
        self.log_box.see("end")

    # ---------------- 帮助 / 退出 ----------------
    def show_help(self):
        help_text = (
            f"【{APP_NAME}】 v{APP_VERSION} · {platform_label()}（{platform_arch()}）\n"
            "开源 MIT 许可 · 同一套代码，Windows 与国产系统功能界面完全一致。\n\n"
            "【核心功能】\n"
            "• 批量处理 .docx/.doc/.wps/.txt/.md，或直接在“其他”页粘贴文本排版。\n"
            "• 处理模式：默认“全量排版”；选“仅修标点”可保留原字体段落，只修标点/序号/禁则。\n"
            "• “其他”页可开启：标点全半角标准化、序号风格统一、中文换行禁则。\n"
            "• 标题可在“字体”页分别指定中英文字体（西文字体细分）。\n\n"
            "【智能识别】\n"
            "• 主/副标题：文档开头连续居中、字体字号相同的段落；副标题字号不同。\n"
            "• 一级 一、 二级 （一） 三级 1. 四级 (1)，正文与三/四级共用字体。\n"
            "• 图/表标题：居中且以“图/表”开头的段落。\n"
            "• 附件：识别“附件1/附件：”，段前分页。\n\n"
            "【安全】绝不改原文件，全部在临时副本上操作。\n"
            "【优化】COM 常驻复用、参数全可配、跨平台转换兜底。\n\n"
            "【字号】所有字号均用下拉选择，与 Word“开始”选项卡一致：\n"
            "初号~八号（42→5 磅）及常用数字档，也可直接手输任意磅值。\n\n"
            "【中文字体】中文排版常用仿宋（正文）、楷体（标题）：Windows 用仿宋_GB2312 /\n"
            "楷体_GB2312；Linux / 国产系统可用开源 FandolFang（方政仿宋）、文鼎 AR PL UKai\n"
            "（楷体）、思源宋体 / Noto Serif CJK（大标题）替代。缺失时可在“字体”页查看状态，\n"
            "本工具不安装字体，请自行准备并安装到系统。"
        )
        win = self._popup("使用说明", 560, 420)
        t = ctk.CTkTextbox(win, wrap="word")
        t.pack(fill="both", expand=True, padx=12, pady=12)
        t.insert("end", help_text)
        t.configure(state="disabled")

    def _on_close(self):
        if self._running:
            from tkinter import messagebox
            if not messagebox.askyesno("确认", "任务仍在进行中，确定要退出吗？"):
                return
        try:
            from .legacy import quit_com
            quit_com()
        except Exception:
            pass
        self.root.destroy()

    def run(self):
        self.root.mainloop()

    # ---------------- 模板 / 质检 / 导出 ----------------
    def open_template(self):
        from .templates import TEMPLATES, TEMPLATE_DESCRIPTIONS
        win = self._popup("文档模板", 480, 480)
        ctk.CTkLabel(win, text="选择模板：自动套用排版预设并生成标准骨架文档",
                     font=ctk.CTkFont(weight="bold")).pack(pady=(12, 6))
        sf = ctk.CTkScrollableFrame(win)
        sf.pack(fill="both", expand=True, padx=12, pady=6)

        def card(kind, title, desc):
            f = ctk.CTkFrame(sf)
            f.pack(fill="x", padx=4, pady=6)
            ctk.CTkLabel(f, text=title, font=ctk.CTkFont(weight="bold", size=15)
                         ).pack(anchor="w", padx=10, pady=(6, 0))
            ctk.CTkLabel(f, text=desc, text_color=("gray50", "gray60"),
                         justify="left").pack(anchor="w", padx=10)
            ctk.CTkButton(f, text="生成", width=90,
                          command=lambda k=kind: self._apply_and_build(k)).pack(
                anchor="e", padx=10, pady=6)

        for kind, (label, _fn) in TEMPLATES.items():
            desc = TEMPLATE_DESCRIPTIONS.get(kind, "")
            card(kind, label, desc)

    def _apply_and_build(self, kind):
        from tkinter import filedialog
        from .templates import TEMPLATES, TEMPLATE_DESCRIPTIONS, generate_template
        from docx import Document
        if kind in TEMPLATES:
            label, fn = TEMPLATES[kind]
            self.cfg = fn()
            self._apply_config(self.cfg)
            self.cfg.save(FormatterConfig.default_config_path())
            self._log(f"已套用「{label}」排版预设")
        out = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word", "*.docx")],
            title="保存生成的文档")
        if not out:
            return
        doc = Document()
        generate_template(kind, doc)
        doc.save(out)
        self._log(f"✅ 已生成模板文档：{out}")

    def open_check(self):
        from tkinter import filedialog
        from .checker import check_document
        p = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")], title="选择待质检文档")
        if not p:
            return
        try:
            rep = check_document(p)
        except Exception as e:
            self._log(f"❌ 质检失败：{e}", "ERROR")
            return
        win = self._popup("排版质检报告", 560, 480)
        ctk.CTkLabel(win, text=f"综合评分 {rep['score']} / 100　等级：{rep['level']}",
                     font=ctk.CTkFont(weight="bold", size=16)).pack(pady=10)
        box = ctk.CTkScrollableFrame(win)
        box.pack(fill="both", expand=True, padx=12, pady=6)
        mark = {"ok": "✅", "warn": "⚠️", "fail": "❌"}
        for it in rep["items"]:
            ctk.CTkLabel(box, text=f"{mark.get(it['status'], '•')} {it['name']}",
                         font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=6, pady=(6, 0))
            ctk.CTkLabel(box, text=it["detail"], text_color=("gray50", "gray60"),
                         justify="left", wraplength=500).pack(anchor="w", padx=18)
        if rep["notes"]:
            ctk.CTkLabel(box, text="建议：", font=ctk.CTkFont(weight="bold")
                         ).pack(anchor="w", padx=6, pady=(10, 0))
            for n in rep["notes"]:
                ctk.CTkLabel(box, text="• " + n, justify="left",
                             wraplength=500).pack(anchor="w", padx=18)

    def export_pdf_action(self):
        from tkinter import filedialog
        from . import export_pdf as _pdfmod
        p = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")], title="选择要导出的文档")
        if not p:
            return
        prefer = bool(self.use_builtin_pdf.get())
        self._log(f"正在导出 PDF：{os.path.basename(p)} …"
                  f"{'（内置引擎）' if prefer else ''}")
        def worker():
            try:
                out = _pdfmod.export_pdf(p, prefer_builtin=prefer)
                engine = _pdfmod.LAST_ENGINE or "未知"
                self._queue_log(f"✅ 已导出 PDF：{out}（引擎：{engine}）", "INFO")
            except Exception as e:
                self._queue_log(f"❌ PDF 导出失败：{e}", "ERROR")
        threading.Thread(target=worker, daemon=True).start()


def main():
    WordFormatterGUI().run()


if __name__ == "__main__":
    main()
