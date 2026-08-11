"""公文字体检测与缺失提示（面向中文公文排版）。

公文常用字体：仿宋_GB2312 / 楷体_GB2312（以及不带 GB2312 后缀的现代同名字体）。
这些字体若系统缺失，Word 会回退到替代字体，导致公文版式不达标。

本模块只做**检测与提示**，不安装字体、不修改系统：
- 检测系统已安装字体（读取 HKLM/HKCU 字体注册表）；
- 按别名匹配判断“仿宋/楷体”是否已装；
- 供 GUI 展示字体状态、提示缺失，由用户自行安装所需字体。

字体安装属于系统级修改，且涉及字体授权分发，本工具不代为执行。
"""

from __future__ import annotations

import os
import re
import sys

# 公文规范（GB/T 9704-2012《党政机关公文格式》）常用字体。
# required=True 表示“缺失会导致版式不达标、必须自行安装”；其余为规范推荐/开源字体。
# 现代 Windows（Win7+）已不再内置仿宋_GB2312 / 楷体_GB2312，故二者为核心检测项；
# 思源宋体（大标题，开源 SIL OFL）需用户自行取得并安装到系统。
GONGWEN_FONTS = [
    {
        "key": "fangsong",
        "display": "仿宋_GB2312",
        "category": "正文（三号）",
        "required": True,
        "hint": "微软随 Windows（中文版）分发的专有字体，请自行安装到系统后再使用本工具",
        "aliases": [
            "仿宋_GB2312", "仿宋GB2312", "仿宋", "FangSong_GB2312",
            "FangSongGB2312", "FangSong", "FangSong_GB2312",
        ],
    },
    {
        "key": "kaiti",
        "display": "楷体_GB2312",
        "category": "二级标题 / 签发人",
        "required": True,
        "hint": "微软随 Windows（中文版）分发的专有字体，请自行安装到系统后再使用本工具",
        "aliases": [
            "楷体_GB2312", "楷体GB2312", "楷体", "KaiTi_GB2312",
            "KaiTiGB2312", "KaiTi", "KaiTi_GB2312",
        ],
    },
    {
        "key": "source_han_serif",
        "display": "思源宋体",
        "category": "公文大标题（二号，加粗）",
        "required": False,
        "hint": "开源 SIL OFL 授权，可合法免费获取并安装到系统",
        "aliases": [
            "思源宋体", "思源宋体 Bold", "思源宋体-Bold", "Source Han Serif SC",
            "Source Han Serif CN", "SourceHanSerifSC", "Source Han Serif",
            "Noto Serif CJK SC", "Noto Serif CJK",
        ],
    },
]

_FONT_REGISTRY_HKLM = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
_FONT_REGISTRY_HKCU = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def _normalize(name: str) -> str:
    """归一化字体名：去空格/下划线/大小写，便于别名匹配。"""
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _is_windows() -> bool:
    return sys.platform.startswith("win")


def list_installed_fonts() -> set[str]:
    """返回已安装字体族名集合（合并 HKLM 与 HKCU）。非 Windows 返回空集。"""
    if not _is_windows():
        return set()
    import winreg
    found: set[str] = set()
    for root, sub in ((winreg.HKEY_LOCAL_MACHINE, _FONT_REGISTRY_HKLM),
                      (winreg.HKEY_CURRENT_USER, _FONT_REGISTRY_HKCU)):
        try:
            key = winreg.OpenKey(root, sub)
        except OSError:
            continue
        try:
            i = 0
            while True:
                name, _, _ = winreg.EnumValue(key, i)
                i += 1
                # 注册表值名形如 “宋体 (TrueType)” —— 取括号前的族名
                fam = name.split("(")[0].strip()
                if fam:
                    found.add(fam)
        except OSError:
            pass
        finally:
            winreg.CloseKey(key)
    return found


def is_font_installed(name: str) -> bool:
    """按归一化名称判断字体是否已安装（支持别名直接传入）。"""
    norm = _normalize(name)
    if not norm:
        return False
    return norm in {_normalize(f) for f in list_installed_fonts()}


def _gongwen_installed(gf: dict) -> bool:
    return any(is_font_installed(a) for a in gf["aliases"])


def check_gongwen_fonts() -> list[dict]:
    """仅检测，返回当前缺失的公文字体定义列表（不安装、不改系统）。"""
    return [gf for gf in GONGWEN_FONTS if not _gongwen_installed(gf)]


def gongwen_font_report() -> list[dict]:
    """返回全部公文字体的状态摘要，供 UI 展示。

    每项含：display（展示名）、category（用途）、required（是否必备）、
    installed（当前是否已安装）、hint（缺失时的处理提示）。
    """
    return [
        {
            "display": gf["display"],
            "category": gf.get("category", ""),
            "required": gf.get("required", False),
            "installed": _gongwen_installed(gf),
            "hint": gf.get("hint", ""),
        }
        for gf in GONGWEN_FONTS
    ]
