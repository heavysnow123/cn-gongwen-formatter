"""将 .docx 导出为 PDF。

后端探测优先级（任一可用即可）：
  1. Microsoft Word COM（Word.Application）
  2. WPS COM（KWPS.Application / Kwps.Application / WPS.Application）
  3. LibreOffice headless（soffice --convert-to pdf）

纯离线桌面工具，依赖外部渲染引擎（与 Word/WPS 自身导出一致）。
"""
from __future__ import annotations

import os
import shutil
import subprocess


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
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
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


def export_pdf(docx_path: str, pdf_path: str | None = None) -> str:
    """把 docx 导出为 pdf，返回 pdf 路径。失败抛 RuntimeError（含可用后端提示）。"""
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"找不到文件：{docx_path}")
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    pdf_path = os.path.abspath(pdf_path)

    errors = []
    try:
        if _export_word_com(docx_path, pdf_path):
            return pdf_path
    except Exception as e:
        errors.append(f"Word: {e}")
    try:
        if _export_wps_com(docx_path, pdf_path):
            return pdf_path
    except Exception as e:
        errors.append(f"WPS: {e}")
    try:
        if _export_libreoffice(docx_path, pdf_path):
            return pdf_path
    except Exception as e:
        errors.append(f"LibreOffice: {e}")

    raise RuntimeError(
        "未找到可用的 PDF 导出后端。请安装以下任一程序后重试：\n"
        "• Microsoft Word\n"
        "• WPS Office\n"
        "• LibreOffice（https://www.libreoffice.org）\n"
        f"（探测明细：{'; '.join(errors) if errors else '后端均不可用'}）"
    )
