"""字体检测模块测试（仅检测，不涉及安装）。

- 归一化、已装字体集合、别名检测、检测与报告一致性。
"""

import sys

from word_formatter import fonts


def test_normalize():
    assert fonts._normalize("仿宋_GB2312") == fonts._normalize("仿宋GB2312")
    assert fonts._normalize("KaiTi") == fonts._normalize("kaiti")
    assert fonts._normalize("Arial Black") == "arialblack"
    print("✓ 字体名归一化正确")


def test_installed_is_set():
    fonts_list = fonts.list_installed_fonts()
    assert isinstance(fonts_list, set)
    if sys.platform.startswith("win"):
        assert len(fonts_list) > 0, "Windows 上应检测到已安装字体"
    assert not fonts.is_font_installed("WordFormatterNoSuchFontXYZ")
    print("✓ 已装字体集合检测正常（Windows 非空）")


def test_gongwen_aliases():
    fs = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "fangsong")
    assert "仿宋_GB2312" in fs["aliases"] and "FangSong" in fs["aliases"]
    kt = next(g for g in fonts.GONGWEN_FONTS if g["key"] == "kaiti")
    assert "楷体_GB2312" in kt["aliases"] and "KaiTi" in kt["aliases"]
    print("✓ 公文必备字体别名定义完整")


def test_check_vs_report_consistency():
    # check_gongwen_fonts 返回的缺失项，应与 report 中 installed=False 一致
    missing_keys = {g["key"] for g in fonts.check_gongwen_fonts()}
    report_missing = {r["display"] for r in fonts.gongwen_font_report() if not r["installed"]}
    for g in fonts.GONGWEN_FONTS:
        if g["key"] in missing_keys:
            assert g["display"] in report_missing
    print("✓ 检测与报告结果一致")


if __name__ == "__main__":
    test_normalize()
    test_installed_is_set()
    test_gongwen_aliases()
    test_check_vs_report_consistency()
    print("\n字体模块测试通过 ✅")
