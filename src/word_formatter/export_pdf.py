"""将 .docx 导出为 PDF。

导出后端优先级（任一可用即可）：
  1. Microsoft Word COM（Word.Application）—— 版式与 Word 完全一致
  2. WPS COM（KWPS.Application / Kwps.Application / WPS.Application）
  3. LibreOffice headless（soffice --convert-to pdf）
  4. 内置渲染引擎（reportlab，纯 Python，无需安装任何办公软件）

第 4 项「内置引擎」是本工具自带的保底方案：只要装了本软件即可导出 PDF，
不依赖 Word / WPS / LibreOffice。其版面以通用排版规则还原（字体、字号、
对齐、行距、页边距、表格、红头颜色等），与 Word 分页可能略有差异，
但内容、字体、版式要素完整可打印。
"""
from __future__ import annotations

import os
import shutil
import subprocess

# 供 GUI 显示本次实际使用的引擎
LAST_ENGINE = ""


def _ensure_com():
    """子线程调用 COM 必须初始化套间（STA）。主线程已初始化则忽略错误。"""
    try:
        import pythoncom
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except Exception:
        pass


def _detect_soffice() -> str | None:
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/opt/libreoffice/program/soffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    import glob
    for pat in ("/opt/libreoffice*/program/soffice",
                "/opt/LibreOffice*/program/soffice"):
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[-1]
    found = shutil.which("soffice") or shutil.which("libreoffice")
    return found


def _export_word_com(docx_path: str, pdf_path: str) -> bool:
    import win32com.client
    _ensure_com()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(pdf_path),
            ExportFormat=17,  # wdExportFormatPDF
        )
        return True
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass


def _export_wps_com(docx_path: str, pdf_path: str) -> bool:
    import win32com.client
    _ensure_com()
    for pid in ("KWPS.Application", "Kwps.Application", "WPS.Application"):
        try:
            app = win32com.client.Dispatch(pid)
        except Exception:
            continue
        app.Visible = False
        doc = None
        try:
            doc = app.Documents.Open(os.path.abspath(docx_path))
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=os.path.abspath(pdf_path),
                    ExportFormat=17,
                )
            except Exception:
                doc.SaveAs(os.path.abspath(pdf_path), 17)  # wdFormatPDF = 17
            return True
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
    return False


def _export_libreoffice(docx_path: str, pdf_path: str) -> bool:
    soffice = _detect_soffice()
    if not soffice:
        return False
    out_dir = os.path.dirname(os.path.abspath(pdf_path))
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", out_dir, os.path.abspath(docx_path)],
            check=True, capture_output=True, timeout=300,
        )
        base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        produced = os.path.join(out_dir, base)
        if os.path.abspath(produced) != os.path.abspath(pdf_path) and os.path.exists(produced):
            shutil.move(produced, pdf_path)
        return os.path.exists(pdf_path)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 内置渲染引擎（reportlab，纯 Python，无需任何办公软件）
# ---------------------------------------------------------------------------
def _init_builtin_fonts():
    """注册中文字体。优先用系统已装的 TTF（仿宋/楷体/思源宋体等），
    否则回退到 reportlab 内置的 Adobe CID 中文字体（STSong-Light，
    自带且无需字体文件）。返回当前默认中文字体名。"""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfbase.ttfonts import TTFont

    cjk = "Helvetica"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        cjk = "STSong-Light"
    except Exception:
        pass

    # 系统字体候选：name -> 候选文件名列表（按优先级，首个存在者注册）。
    # 同时收录 Windows 专有字体与 Linux / 国产系统常见开源替代
    # （Fandol 仿宋/楷体、文鼎 AR PL UKai、Noto / 思源等）。
    candidates = {
        "FangSong": [
            "仿宋_GB2312.ttf", "FandolFang-Regular.otf", "FandolFang.otf",
            "FZFangSong-Z01S.ttf", "Fandol Fang.otf",
        ],
        "KaiTi": [
            "楷体_GB2312.ttf", "FandolKai-Regular.otf", "UKaiCN.ttf",
            "ARPLUKaiCN.ttf", "FZKai-Z03S.ttf", "Fandol Kai.otf",
        ],
        "SimSun": [
            "simsun.ttc", "NotoSerifCJKsc-Regular.otf",
            "SourceHanSerifSC-Regular.otf",
        ],
        "SimHei": [
            "simhei.ttf", "NotoSansCJKsc-Regular.otf", "SourceHanSansSC-Regular.otf",
        ],
        "SourceHanSerif": ["SourceHanSerifSC-Regular.otf"],
        "SourceHanSerifB": ["SourceHanSerifSC-Bold.otf"],
        "NotoSerifCJK": ["NotoSerifCJKsc-Regular.otf"],
    }
    search_dirs = [
        os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts"),
        os.path.expandvars("%LOCALAPPDATA%/Microsoft/Windows/Fonts"),
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        os.path.expanduser("~/.fonts"),
        os.path.expanduser("~/.local/share/fonts"),
    ]
    registered = {}
    for name, fnames in candidates.items():
        if name in registered:
            continue
        for fname in fnames:
            found_path = None
            for d in search_dirs:
                p = os.path.join(d, fname)
                if os.path.exists(p):
                    found_path = p
                    break
            if found_path:
                try:
                    pdfmetrics.registerFont(TTFont(name, found_path))
                    registered[name] = name
                    break
                except Exception:
                    pass
    # 让 <b> 等标记能落到已注册字体（无独立粗体时复用同字体，避免异常）
    for name in registered:
        try:
            pdfmetrics.registerFontFamily(
                name, normal=name, bold=name, italic=name, boldItalic=name)
        except Exception:
            pass
    _init_builtin_fonts._cjk = cjk
    _init_builtin_fonts._ttf = registered
    return cjk


def _resolve_font(name: str | None, bold: bool):
    cjk = getattr(_init_builtin_fonts, "_cjk", "Helvetica")
    ttf = getattr(_init_builtin_fonts, "_ttf", {})
    if not name:
        return cjk
    low = name.lower()
    if "仿宋" in name or "fangsong" in low:
        key = "FangSong"
    elif "楷" in name or "kai" in low:
        key = "KaiTi"
    elif "思源" in name or "source han" in low or "noto serif cjk" in low:
        key = "SourceHanSerif" if "SourceHanSerif" in ttf else ("SourceHanSerifB" if "SourceHanSerifB" in ttf else "NotoSerifCJK")
    elif "黑" in name or "simhei" in low or "heit" in low:
        key = "SimHei"
    elif "宋" in name or "simsun" in low or "song" in low:
        key = "SimSun"
    else:
        key = None
    if key and key in ttf:
        return key
    return cjk


def _iter_blocks(doc):
    """按文档顺序产出 ('p', paragraph) 或 ('t', table)，含表格内段落顺序。"""
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph as DocxParagraph
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield ("p", DocxParagraph(child, doc))
        elif isinstance(child, CT_Tbl):
            yield ("t", Table(child, doc))


def _align_enum(a):
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    return {
        WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT,
        WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
    }.get(a, TA_LEFT)


def _run_markup(run):
    from xml.sax.saxutils import escape
    esc = escape(run.text or "")
    font = run.font
    fname = _resolve_font(font.name, bool(font.bold))
    size = font.size.pt if font.size else None
    col = None
    try:
        if font.color is not None and font.color.type is not None and font.color.rgb is not None:
            col = "#%s" % str(font.color.rgb)
    except Exception:
        col = None
    attrs = []
    if fname:
        attrs.append(f'name="{fname}"')
    if size:
        attrs.append(f'size="{size:.1f}"')
    if col:
        attrs.append(f'color="{col}"')
    font_tag = f'<font {" ".join(attrs)}>' if attrs else ""
    close_font = "</font>" if font_tag else ""
    bb = "<b>" if font.bold else ""
    ib = "<i>" if font.italic else ""
    close = ("</i>" if font.italic else "") + ("</b>" if font.bold else "") + close_font
    return f"{font_tag}{bb}{ib}{esc}{close}"


def _para_to_flowable(para, default_size=12.0):
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, Spacer
    from docx.enum.text import WD_LINE_SPACING

    pf = para.paragraph_format
    # 取段落中最大字号作为基准
    size = default_size
    for r in para.runs:
        if r.font.size:
            size = max(size, r.font.size.pt)
    # 行距
    leading = size * 1.5
    try:
        if pf.line_spacing_rule == WD_LINE_SPACING.EXACTLY and pf.line_spacing:
            leading = pf.line_spacing.pt
        elif pf.line_spacing_rule == WD_LINE_SPACING.MULTIPLE and pf.line_spacing:
            leading = size * float(pf.line_spacing)
    except Exception:
        pass
    style = ParagraphStyle(
        "p",
        fontName="STSong-Light" if _init_builtin_fonts._cjk == "STSong-Light" else "Helvetica",
        fontSize=size,
        leading=leading,
        alignment=_align_enum(para.alignment),
        spaceBefore=(pf.space_before.pt if pf.space_before else 0),
        spaceAfter=(pf.space_after.pt if pf.space_after else 0),
        firstLineIndent=(pf.first_line_indent.pt if pf.first_line_indent else 0),
        leftIndent=(pf.left_indent.pt if pf.left_indent else 0),
    )
    markup = "".join(_run_markup(r) for r in para.runs) or "&nbsp;"
    return Paragraph(markup, style)


def _table_to_flowable(tbl, default_size=10.5):
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle, Paragraph

    data = []
    for row in tbl.rows:
        cells = []
        for cell in row.cells:
            style = ParagraphStyle(
                "cell", fontName="STSong-Light" if _init_builtin_fonts._cjk == "STSong-Light" else "Helvetica",
                fontSize=default_size, leading=default_size * 1.3)
            txt = (cell.text or "").replace("\n", "<br/>")
            cells.append(Paragraph(txt or "&nbsp;", style))
        data.append(cells)
    t = Table(data, hAlign="LEFT")
    ts = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    t.setStyle(TableStyle(ts))
    return t


def _builtin_footer(canvas, doc):
    from reportlab.lib.units import mm
    canvas.saveState()
    canvas.setFont("Helvetica", 9)
    canvas.setFillColorRGB(0.4, 0.4, 0.4)
    canvas.drawCentredString(doc.pagesize[0] / 2.0, doc.bottomMargin - 12,
                             f"第 {doc.page} 页")
    canvas.restoreState()


def export_pdf_builtin(docx_path: str, pdf_path: str) -> str:
    """用 reportlab 内置引擎导出 PDF，不依赖任何办公软件。返回 pdf 路径。"""
    from docx import Document
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Spacer

    _init_builtin_fonts()

    doc = Document(docx_path)
    sec = doc.sections[0]
    page_w = (sec.page_width.pt if sec.page_width else A4[0])
    page_h = (sec.page_height.pt if sec.page_height else A4[1])
    m_l = sec.left_margin.pt if sec.left_margin else 28 * mm
    m_r = sec.right_margin.pt if sec.right_margin else 26 * mm
    m_t = sec.top_margin.pt if sec.top_margin else 37 * mm
    m_b = sec.bottom_margin.pt if sec.bottom_margin else 35 * mm

    tmpl = SimpleDocTemplate(
        pdf_path, pagesize=(page_w, page_h),
        leftMargin=m_l, rightMargin=m_r, topMargin=m_t, bottomMargin=m_b,
        title=os.path.splitext(os.path.basename(docx_path))[0],
        author="WordFormatterPro",
    )

    from docx.enum.text import WD_ALIGN_PARAGRAPH as _WA
    from reportlab.platypus import KeepTogether

    blocks = list(_iter_blocks(doc))
    flowables = []

    # 分页优化 1：文档开头的连续居中段落（红头/标题/副标题）整体保持同页，
    # 避免“红头或标题被分页割裂到两页”这类与 Word 不一致的现象。
    lead_end = 0
    for k, blk in blocks:
        if k == "p" and getattr(blk, "alignment", None) == _WA.CENTER and lead_end < 6:
            lead_end += 1
        else:
            break
    if lead_end:
        flowables.append(KeepTogether(
            [_para_to_flowable(blocks[j][1]) for j in range(lead_end)]))

    # 分页优化 2：每个表格整体保持（行不被拆散到两页），超大表由 reportlab 自动拆分。
    for k, blk in blocks[lead_end:]:
        if k == "p":
            flowables.append(_para_to_flowable(blk))
        else:
            flowables.append(KeepTogether([_table_to_flowable(blk)]))

    if not flowables:
        from reportlab.platypus import Paragraph as _P
        from reportlab.lib.styles import ParagraphStyle as _S
        flowables.append(_P("&nbsp;", _S("empty")))

    tmpl.build(flowables, onFirstPage=_builtin_footer, onLaterPages=_builtin_footer)
    return pdf_path


def export_pdf(docx_path: str, pdf_path: str | None = None,
               prefer_builtin: bool = False) -> str:
    """把 docx 导出为 pdf，返回 pdf 路径。

    优先顺序：Word COM → WPS COM → LibreOffice → 内置引擎（reportlab）。
    设置 prefer_builtin=True 可跳过外部后端，直接用内置引擎。
    实际引擎名记录在模块变量 LAST_ENGINE，供 GUI 显示。
    """
    global LAST_ENGINE
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"找不到文件：{docx_path}")
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    pdf_path = os.path.abspath(pdf_path)

    if prefer_builtin:
        LAST_ENGINE = "内置引擎(reportlab)"
        return export_pdf_builtin(docx_path, pdf_path)

    errors = []
    try:
        if _export_word_com(docx_path, pdf_path):
            LAST_ENGINE = "Microsoft Word"
            return pdf_path
    except Exception as e:
        errors.append(f"Word: {e}")
    try:
        if _export_wps_com(docx_path, pdf_path):
            LAST_ENGINE = "WPS Office"
            return pdf_path
    except Exception as e:
        errors.append(f"WPS: {e}")
    try:
        if _export_libreoffice(docx_path, pdf_path):
            LAST_ENGINE = "LibreOffice"
            return pdf_path
    except Exception as e:
        errors.append(f"LibreOffice: {e}")

    # 保底：内置引擎
    try:
        LAST_ENGINE = "内置引擎(reportlab)"
        return export_pdf_builtin(docx_path, pdf_path)
    except Exception as e2:
        errors.append(f"内置引擎: {e2}")
        raise RuntimeError(
            "PDF 导出失败，所有后端均不可用：\n"
            + "\n".join(f"• {x}" for x in errors)
        )
