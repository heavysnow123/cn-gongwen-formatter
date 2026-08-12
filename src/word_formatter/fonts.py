"""中文字体检测与缺失提示（跨平台，面向中文排版 / 公文）。

检测系统已安装的字体，按别名匹配判断「仿宋 / 楷体 / 思源宋体」等常用中文字体是否已装。
- Windows：读取 HKLM / HKCU 字体注册表。
- Linux / macOS（含统信 UOS、银河麒麟等国产系统）：用 fontconfig 的 `fc-list`，
  无 fc-list 时兜底扫描系统字体目录。

本模块只做**检测与提示**，不安装字体、不修改系统；
缺失时给出处理建议（Windows 用专有字体，Linux / 国产系统可用 Fandol、文鼎 AR PL、
Noto 等开源替代）。字体安装属于系统级修改且涉及授权分发，本工具不代为执行。
"""

from __future__ import annotations

import os
import re
import subprocess
import sys

# 常用中文字体（公文规范 GB/T 9704-2012 等场景）。
# required=True 表示「缺失会导致版式不达标」；其余为推荐 / 开源字体。
# aliases 同时收录 Windows 专有字体名与 Linux / 国产系统常见开源替代名，
# 使检测在两平台上都能正确命中。
GONGWEN_FONTS = [
    {
        "key": "fangsong",
        "display": "仿宋（正文三号）",
        "category": "正文（三号）",
        "required": True,
        "hint": "Windows 用「仿宋_GB2312」；Linux / 国产系统可用开源 FandolFang（方政仿宋）"
                "或文鼎 PL 仿宋替代",
        "aliases": [
            "仿宋_GB2312", "仿宋GB2312", "仿宋", "FangSong_GB2312",
            "FangSongGB2312", "FangSong", "FangSong_GB2312",
            "FandolFang", "Fandol Fang", "FandolFang-Regular",
            "文鼎PL仿宋", "AR PL 仿宋", "FZFangSong", "FZFangSong-Z01S",
        ],
    },
    {
        "key": "kaiti",
        "display": "楷体（二级标题 / 签发人）",
        "category": "二级标题 / 签发人",
        "required": True,
        "hint": "Windows 用「楷体_GB2312」；Linux / 国产系统可用文鼎 AR PL UKai（楷体）"
                "或 FandolKai 替代",
        "aliases": [
            "楷体_GB2312", "楷体GB2312", "楷体", "KaiTi_GB2312",
            "KaiTiGB2312", "KaiTi", "KaiTi_GB2312",
            "AR PL UKai", "AR PL UKai CN", "UKai", "FandolKai",
            "Fandol Kai", "文鼎PL楷体", "FZKaiTi", "FZKai-Z03S",
        ],
    },
    {
        "key": "source_han_serif",
        "display": "思源宋体（大标题）",
        "category": "公文大标题（二号，加粗）",
        "required": False,
        "hint": "开源 SIL OFL 授权，可合法免费获取安装；Linux / 国产系统常预装为 Noto Serif CJK SC",
        "aliases": [
            "思源宋体", "思源宋体 Bold", "思源宋体-Bold", "Source Han Serif SC",
            "Source Han Serif CN", "SourceHanSerifSC", "Source Han Serif",
            "Noto Serif CJK SC", "Noto Serif CJK", "Noto Serif CJK JP",
        ],
    },
]

_FONT_REGISTRY_HKLM = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
_FONT_REGISTRY_HKCU = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"

# Linux / macOS 常见字体目录（含国产系统默认路径）
_UNIX_FONT_DIRS = [
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    "/usr/share/fonts/truetype",
    os.path.expanduser("~/.fonts"),
    os.path.expanduser("~/.local/share/fonts"),
]


def _normalize(name: str) -> str:
    """归一化字体名：去空格 / 下划线 / 连字符 / 点 / 括号并转小写，**保留中文**，
    以便跨平台（GB2312 专有名 vs Fandol / AR PL / Noto 开源名）都能匹配。"""
    return re.sub(r"[\s_\-\.(),（）]", "", (name or "").lower())


def _is_windows() -> bool:
    return sys.platform.startswith("win")


def _list_fonts_windows() -> set[str]:
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
            try:
                winreg.CloseKey(key)
            except OSError:
                pass
    return found


def _list_fonts_unix() -> set[str]:
    """Linux / macOS：优先用 fontconfig 的 fc-list；无则扫描字体目录文件名兜底。"""
    found: set[str] = set()
    # 1) fc-list（最准确，覆盖几乎所有 Linux 桌面 / 国产系统）
    try:
        out = subprocess.run(
            ["fc-list", ":", "family"],
            capture_output=True, text=True, timeout=20,
        )
        if out.returncode == 0:
            for line in out.stdout.splitlines():
                for fam in line.split(","):
                    fam = fam.strip()
                    if fam:
                        found.add(fam)
    except Exception:
        pass
    # 2) 兜底：直接扫描字体目录的文件名（无 fc-list 的最小系统 / 容器）
    if not found:
        for d in _UNIX_FONT_DIRS:
            if not os.path.isdir(d):
                continue
            for root, _, files in os.walk(d):
                for f in files:
                    if f.lower().endswith((".ttf", ".otf", ".ttc")):
                        fam = os.path.splitext(f)[0]
                        if fam:
                            found.add(fam)
    return found


def list_installed_fonts() -> set[str]:
    """返回已安装字体族名集合（跨平台：Windows 注册表 / Unix fc-list + 目录扫描）。"""
    if _is_windows():
        return _list_fonts_windows()
    return _list_fonts_unix()


def is_font_installed(name: str) -> bool:
    """按归一化名称判断字体是否已安装（支持别名直接传入）。"""
    norm = _normalize(name)
    if not norm:
        return False
    return norm in {_normalize(f) for f in list_installed_fonts()}


def _gongwen_installed(gf: dict) -> bool:
    return any(is_font_installed(a) for a in gf["aliases"])


def check_gongwen_fonts() -> list[dict]:
    """仅检测，返回当前缺失的字体会用列表（不安装、不改系统）。"""
    return [gf for gf in GONGWEN_FONTS if not _gongwen_installed(gf)]


def gongwen_font_report() -> list[dict]:
    """返回全部字体的状态摘要，供 UI 展示。

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
