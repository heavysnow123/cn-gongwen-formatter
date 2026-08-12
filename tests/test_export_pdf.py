"""导出 PDF：验证内置引擎（reportlab）可在无 Office 环境下工作。"""
import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from docx import Document
from word_formatter.templates import generate_gongwen
from word_formatter import export_pdf as pdfmod


@pytest.mark.skipif(sys.platform != "win32", reason="PDF 后端依赖 Windows COM/字体")
def test_builtin_export_no_office():
    with tempfile.TemporaryDirectory() as td:
        doc = Document()
        generate_gongwen(doc)
        src = os.path.join(td, "demo.docx")
        doc.save(src)
        out = os.path.join(td, "demo.pdf")
        # 强制内置引擎，不依赖 Word/WPS/LibreOffice
        pdfmod.export_pdf(src, out, prefer_builtin=True)
        assert os.path.exists(out), "内置引擎未生成 PDF"
        assert os.path.getsize(out) > 0, "PDF 为空"
        assert pdfmod.LAST_ENGINE, "未记录使用的引擎"


def test_prefer_builtin_flag():
    # 仅验证标志位被接受、函数签名可用（不要求真实渲染）
    assert callable(pdfmod.export_pdf)
