"""字体检测模块测试（仅检测，不涉及安装）。

- 归一化（保留中文，跨平台匹配）
- 已装字体集合（Windows 非空，其他平台不报错）
- 公文 / 开源替代别名定义
- Linux / macOS 的 fc-list 解析与目录兜底
- 检测与报告一致性
"""

import os
import subprocess
import sys
from unittest.mock import MagicMock

import pytest

from word_formatter import fonts


def test_normalize_preserves_cjk():
    assert fonts._normalize("仿宋_GB2312") == fonts._normalize("仿宋GB2312")
    assert fonts._normalize("KaiTi") == fonts._normalize("kaiti")
    assert fonts._normalize("Arial Black") == "arialblack"
    # 关键是中文应被保留，便于跨平台匹配开源字体名中的中文（如「文鼎PL仿宋」）
    assert "仿宋" in fonts._normalize("仿宋_GB2312")
    print("✓ 字体名归一化正确（保留中文）")


def test_installed_is_set():
    fonts_list = fonts.list_installed_fonts()
    assert isinstance(fonts_list, set)
    if sys.platform.startswith("win"):
        assert len(fonts_list) > 0, "Windows 上应检测到已安装字体"
    assert not fonts.is_font_installed("WordFormatterNoSuchFontXYZ")
    print("✓ 已装字体集合检测正常")


def test_gongwen_aliases():
    fs = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "fangsong")
    assert "仿宋_GB2312" in fs["aliases"] and "FangSong" in fs["aliases"]
    kt = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "kaiti")
    assert "楷体_GB2312" in kt["aliases"] and "KaiTi" in kt["aliases"]
    print("✓ 公文必备字体别名定义完整")


def test_open_source_aliases_present():
    """国产系统常见开源替代字体名应已收录，确保 Linux 上能正确命中。"""
    fs = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "fangsong")
    kt = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "kaiti")
    sh = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "source_han_serif")
    assert "FandolFang" in fs["aliases"]
    assert any("UKai" in a for a in kt["aliases"])  # 文鼎 AR PL UKai
    assert "Noto Serif CJK SC" in sh["aliases"]
    print("✓ 国产系统开源替代字体别名已收录")


def test_list_fonts_unix_parse_fc_list(monkeypatch):
    # `fc-list : family` 仅输出字体族名（一行可能含多个 family，逗号分隔）
    sample = (
        "FandolFang\n"
        "AR PL UKai CN\n"
        "Noto Serif CJK SC, Noto Serif CJK\n"
    )
    fake = MagicMock(); fake.returncode = 0; fake.stdout = sample
    monkeypatch.setattr(subprocess, "run", lambda cmd, **kw: fake)

    result = fonts._list_fonts_unix()
    assert "FandolFang" in result
    assert "AR PL UKai CN" in result
    assert "Noto Serif CJK SC" in result
    assert "Noto Serif CJK" in result
    print("✓ Linux fc-list 解析正确")


def test_list_fonts_unix_fallback_when_no_fc_list(monkeypatch):
    monkeypatch.setattr(subprocess, "run", lambda cmd, **kw: (_ for _ in ()).throw(FileNotFoundError()))
    monkeypatch.setattr(os.path, "isdir", lambda p: False)
    result = fonts._list_fonts_unix()
    assert result == set()
    print("✓ 无 fc-list / 无字体目录时安全返回空集")


def test_is_font_installed_logic(monkeypatch):
    monkeypatch.setattr(
        fonts, "list_installed_fonts", lambda: {"FandolFang", "AR PL UKai CN"}
    )
    assert fonts.is_font_installed("FandolFang")
    assert fonts.is_font_installed("AR PL UKai CN")
    assert not fonts.is_font_installed("GhostFontXYZ")
    print("✓ 跨平台字体命中逻辑正确")


def test_check_vs_report_consistency():
    missing_keys = {g["key"] for g in fonts.check_gongwen_fonts()}
    report_missing = {r["display"] for r in fonts.gongwen_font_report() if not r["installed"]}
    for g in fonts.GONGWEN_FONTS:
        if g["key"] in missing_keys:
            assert g["display"] in report_missing
    print("✓ 检测与报告结果一致")


if __name__ == "__main__":
    # 仅运行不依赖 monkeypatch 的用例；含 fc-list 的跨平台用例请用 pytest 跑
    test_normalize_preserves_cjk()
    test_installed_is_set()
    test_gongwen_aliases()
    test_open_source_aliases_present()
    test_check_vs_report_consistency()
    print("\n字体模块基础测试通过 ✅（完整跨平台用例请用 pytest 运行）")
