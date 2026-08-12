import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from docx import Document
from docx.shared import RGBColor
from word_formatter.templates import (
    gongwen_template, report_template,
    generate_gongwen, generate_report, generate_redhead,
)


def test_gongwen_config():
    c = gongwen_template()
    assert c.body_font == "仿宋_GB2312"
    assert c.force_a4 is True
    assert abs(c.margin_top_cm - 3.7) < 1e-6
    assert c.title_size == 22.0


def test_report_config():
    c = report_template()
    assert c.title_font == "黑体"
    assert c.header_enabled is True


def test_gongwen_redhead_color():
    doc = Document()
    doc.add_paragraph("原内容")
    generate_gongwen(doc)
    first = doc.paragraphs[0]
    assert first.runs[0].font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
    assert len(doc.paragraphs) > 5


def test_redhead_only_red():
    doc = Document()
    generate_redhead(doc)
    assert doc.paragraphs[0].runs[0].font.color.rgb == RGBColor(0xFF, 0x00, 0x00)


def test_report_title_bold():
    doc = Document()
    doc.add_paragraph("x")
    generate_report(doc)
    assert doc.paragraphs[0].runs[0].font.bold is True
