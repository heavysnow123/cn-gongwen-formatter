"""旧格式(.doc/.wps)转换 与 COM 预处理（接受修订 / 自动编号转文本）。

优化点（相对原工具）：
- COM 应用常驻单实例，批量处理时复用，避免逐文件启停 WPS/Word；
- 多后端优雅降级：Word COM -> WPS COM -> LibreOffice soffice；
- 缺少任一组件时静默跳过，不崩溃。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import logging
from typing import Optional

IS_WINDOWS = os.name == "nt"
Word_Application = None  # 常驻实例

LOG = logging.getLogger("word_formatter.legacy")


def _log_fallback(log_cb, msg, level="INFO"):
    if log_cb:
        log_cb(msg, level)
    getattr(LOG, level.lower(), LOG.info)(msg)


# ----------------------- LibreOffice -----------------------
def find_soffice() -> Optional[str]:
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/opt/libreoffice/program/soffice",
        "/opt/libreoffice*/program/soffice",
        "/usr/local/bin/soffice", "/usr/local/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/opt/homebrew/bin/soffice", "/snap/bin/libreoffice",
    ]
    for c in candidates:
        if "*" in c:
            import glob
            matches = glob.glob(c)
            if matches:
                return matches[0]
        elif os.path.exists(c):
            return c
    # PATH 探测
    from shutil import which
    for name in ("soffice", "libreoffice"):
        p = which(name)
        if p:
            return p
    return None


def convert_with_soffice(src: str, out_docx: str, log_cb=None) -> bool:
    soffice = find_soffice()
    if not soffice:
        _log_fallback(log_cb, "LibreOffice(soffice) 不可用，无法转换旧格式。", "WARN")
        return False
    try:
        out_dir = os.path.dirname(out_docx)
        cmd = [
            soffice, "--headless", "--convert-to", "docx",
            "--outdir", out_dir, src,
        ]
        env = dict(os.environ)
        env["HOME"] = env.get("HOME", tempfile.gettempdir())
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, timeout=120, env=env)
        # soffice 输出名为 <base>.docx
        base = os.path.splitext(os.path.basename(src))[0]
        generated = os.path.join(out_dir, f"{base}.docx")
        if os.path.exists(generated) and generated != out_docx:
            shutil.move(generated, out_docx)
        return os.path.exists(out_docx)
    except Exception as e:
        _log_fallback(log_cb, f"LibreOffice 转换失败：{e}", "WARN")
        return False


# ----------------------- COM（Windows） -----------------------
def _ensure_com():
    global Word_Application
    if Word_Application is not None:
        return Word_Application
    import pythoncom
    pythoncom.CoInitialize()
    import win32com.client
    # 优先 Word，其次 WPS
    for prog_id in ("Word.Application", "KWPS.Application"):
        try:
            app = win32com.client.DispatchEx(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            Word_Application = (app, prog_id)
            return Word_Application
        except Exception:
            continue
    return None


def quit_com():
    global Word_Application
    if Word_Application is None:
        return
    try:
        Word_Application[0].Quit()
    except Exception:
        pass
    Word_Application = None


def _com_convert(src: str, out_docx: str, log_cb=None) -> bool:
    inst = _ensure_com()
    if inst is None:
        return False
    app, _ = inst
    try:
        doc = app.Documents.Open(os.path.abspath(src))
        # wdFormatXMLDocument = 16
        doc.SaveAs(os.path.abspath(out_docx), 16)
        doc.Close(False)
        return os.path.exists(out_docx)
    except Exception as e:
        _log_fallback(log_cb, f"COM 转换失败：{e}", "WARN")
        return False


def _com_preprocess(docx_path: str, log_cb=None) -> bool:
    """接受所有修订 + 自动编号转文本。仅 Windows 且 Word 可用时生效。"""
    inst = _ensure_com()
    if inst is None:
        return False
    app, _ = inst
    try:
        doc = app.Documents.Open(os.path.abspath(docx_path))
        try:
            doc.Revisions.AcceptAll()
        except Exception:
            pass
        try:
            doc.Range().ListFormat.ConvertNumbersToText()
        except Exception:
            pass
        doc.Save()
        doc.Close(False)
        return True
    except Exception as e:
        _log_fallback(log_cb, f"COM 预处理失败：{e}", "WARN")
        return False


# ----------------------- 统一入口 -----------------------
def convert_legacy_to_docx(src: str, log_cb=None) -> Optional[str]:
    """把 .doc/.wps 转成临时 .docx。返回路径或 None。"""
    out = tempfile.mkstemp(suffix=".docx")[1]
    if IS_WINDOWS:
        if _com_convert(src, out, log_cb):
            # 预处理（接受修订/编号转文本）
            _com_preprocess(out, log_cb)
            return out
    if convert_with_soffice(src, out, log_cb):
        return out
    try:
        os.remove(out)
    except OSError:
        pass
    return None


def preprocess_docx(docx_path: str, log_cb=None) -> bool:
    if IS_WINDOWS:
        return _com_preprocess(docx_path, log_cb)
    return False
