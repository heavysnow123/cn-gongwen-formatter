"""大文件流式排版引擎（突破 python-docx 内存上限）。

原理：用 lxml.iterparse 边读 word/document.xml 边改边写，每处理完一个顶层块
（<w:p>/<w:tbl>）即序列化并清除，内存恒定，与文件总大小无关。
包内其余部件（样式、媒体、关系）原样拷贝，分节符(<w:sectPr>)完整保留。

页眉/页脚/页码：
- 在标准模式下由 python-docx 写入（apply_page_setup → add_header/add_footer）。
- 在流式模式下，因新增页眉页脚需要新建部件并改写包关系，故在 zip 层处理：
  生成 headerN.xml/footerN.xml，更新 word/_rels/document.xml.rels 与
  [Content_Types].xml，并为每个 <w:sectPr> 注入 <w:footerReference>/<w:headerReference>。
  这些操作只新增小部件、不重排原包，内存仍恒定。
- 页边距(A4/页边距)直接改 <w:sectPr> 的 <w:pgMar>/<w:pgSz>，同样安全。
"""

from __future__ import annotations

import re
import zipfile
from typing import Callable, Optional

from lxml import etree
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import parse_xml

from .config import FormatterConfig
from .core import (
    qn, classify_paragraph, apply_para_format, apply_h2_inline_split,
    add_page_break_before, get_para_font_info, normalize_punctuation,
    format_tables, _para_text, footer_segments, header_font_info,
    footer_font_info,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

W_P = qn("w:p")
W_TBL = qn("w:tbl")
W_SECTPR = qn("w:sectPr")
DOC_NAME = "word/document.xml"
RELS_NAME = "word/_rels/document.xml.rels"
CT_NAME = "[Content_Types].xml"

HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

# 分节符内联定位：w:pPr/w:sectPr（段落型分节，用于横向附录页等）
INLINE_SECTPR_XPATH = "%s/%s" % (qn("w:pPr"), qn("w:sectPr"))


def _to_proxy(elem):
    """把 iterparse 产出的裸 lxml 元素转成 python-docx 代理类（CT_P/CT_Tbl）。

    lxml 6.x 的 iterparse 不支持传入带 class-lookup 的 parser，因此裸元素没有
    .alignment/.paragraph_format 等 CT_* 方法。这里序列化后交给 python-docx 的
    oxml_parser 重新解析，得到可正常调用 format 函数的代理元素。代理是独立小树，
    处理完即被 GC 回收，内存仍恒定。
    """
    return parse_xml(etree.tostring(elem))


# ----------------------- 页眉页脚部件生成 -----------------------
def _set_run_font_xml(r_elem, font: str, size: float, en_font: Optional[str]):
    rPr = etree.SubElement(r_elem, qn("w:rPr"))
    rFonts = etree.SubElement(rPr, qn("w:rFonts"))
    rFonts.set(qn("w:eastAsia"), font)
    rFonts.set(qn("w:ascii"), en_font or font)
    rFonts.set(qn("w:hAnsi"), en_font or font)
    if en_font:
        rFonts.set(qn("w:cs"), en_font)
    sz = etree.SubElement(rPr, qn("w:sz"))
    sz.set(qn("w:val"), str(int(round(size * 2))))
    szCs = etree.SubElement(rPr, qn("w:szCs"))
    szCs.set(qn("w:val"), str(int(round(size * 2))))


def _append_field_xml(r_elem, name: str):
    fb = etree.SubElement(r_elem, qn("w:fldChar"))
    fb.set(qn("w:fldCharType"), "begin")
    instr = etree.SubElement(r_elem, qn("w:instrText"))
    instr.set(qn("xml:space"), "preserve")
    instr.text = name
    fe = etree.SubElement(r_elem, qn("w:fldChar"))
    fe.set(qn("w:fldCharType"), "end")


def _make_footer_xml(cfg: FormatterConfig) -> bytes:
    font, size, en = footer_font_info(cfg)
    segs = footer_segments(cfg)

    ftr = etree.Element(qn("w:ftr"), nsmap={"w": W_NS, "r": R_NS})
    p = etree.SubElement(ftr, qn("w:p"))
    pPr = etree.SubElement(p, qn("w:pPr"))
    jc = etree.SubElement(pPr, qn("w:jc"))
    jc.set(qn("w:val"), cfg.page_number_align or "center")
    for kind, val in segs:
        r = etree.SubElement(p, qn("w:r"))
        if kind == "text":
            t = etree.SubElement(r, qn("w:t"))
            t.set(qn("xml:space"), "preserve")
            t.text = val
        else:
            _append_field_xml(r, val)
        _set_run_font_xml(r, font, size, en)
    return etree.tostring(ftr, xml_declaration=True, encoding="UTF-8", standalone=True)


def _make_header_xml(cfg: FormatterConfig) -> bytes:
    font, size, en = header_font_info(cfg)
    hdr = etree.Element(qn("w:hdr"), nsmap={"w": W_NS, "r": R_NS})
    p = etree.SubElement(hdr, qn("w:p"))
    pPr = etree.SubElement(p, qn("w:pPr"))
    jc = etree.SubElement(pPr, qn("w:jc"))
    jc.set(qn("w:val"), cfg.header_align or "center")
    if cfg.header_border:
        pBdr = etree.SubElement(pPr, qn("w:pBdr"))
        bottom = etree.SubElement(pBdr, qn("w:bottom"))
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
    r = etree.SubElement(p, qn("w:r"))
    t = etree.SubElement(r, qn("w:t"))
    t.set(qn("xml:space"), "preserve")
    t.text = cfg.header_text.strip()
    _set_run_font_xml(r, font, size, en)
    return etree.tostring(hdr, xml_declaration=True, encoding="UTF-8", standalone=True)


def _next_part_index(names, prefix: str) -> int:
    n = 0
    for nm in names:
        m = re.match(r"^word/%s(\d+)\.xml$" % re.escape(prefix), nm)
        if m:
            n = max(n, int(m.group(1)))
    return n + 1


def _add_rels(rels_bytes: bytes, target: str, rid: str, rel_type: str) -> bytes:
    root = etree.fromstring(rels_bytes)
    rel = etree.SubElement(root, "{%s}Relationship" % RELS_NS)
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", target)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_content_type(ct_bytes: bytes, partname: str, content_type: str) -> bytes:
    root = etree.fromstring(ct_bytes)
    ov = etree.SubElement(root, "{%s}Override" % CT_NS)
    ov.set("PartName", partname)
    ov.set("ContentType", content_type)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _inject_sectpr_refs(sectpr, hdr_rid, ftr_rid, need_header, need_footer):
    """为分节符注入页眉/页脚引用（已存在则跳过，保留源文件原有页眉页脚）。"""
    if need_footer and sectpr.find(qn("w:footerReference")) is None:
        fr = etree.Element(qn("w:footerReference"))
        fr.set(qn("w:type"), "default")
        fr.set(qn("r:id"), ftr_rid)
        sectpr.insert(0, fr)
    if need_header and sectpr.find(qn("w:headerReference")) is None:
        hr = etree.Element(qn("w:headerReference"))
        hr.set(qn("w:type"), "default")
        hr.set(qn("r:id"), hdr_rid)
        sectpr.insert(0, hr)


def _apply_sectpr_margins(sectpr, cfg: FormatterConfig):
    """设置分节符的页边距（可选强制 A4）。仅改 sectPr 内属性，不涉及包关系。"""
    if cfg.force_a4:
        pgSz = sectpr.find(qn("w:pgSz"))
        if pgSz is None:
            pgSz = etree.SubElement(sectpr, qn("w:pgSz"))
        pgSz.set(qn("w:w"), "11906")
        pgSz.set(qn("w:h"), "16838")
    pgMar = sectpr.find(qn("w:pgMar"))
    if pgMar is None:
        pgMar = etree.SubElement(sectpr, qn("w:pgMar"))
    tw = lambda cm: str(int(round(cm * 567)))
    pgMar.set(qn("w:top"), tw(cfg.margin_top_cm))
    pgMar.set(qn("w:right"), tw(cfg.margin_right_cm))
    pgMar.set(qn("w:bottom"), tw(cfg.margin_bottom_cm))
    pgMar.set(qn("w:left"), tw(cfg.margin_left_cm))
    pgMar.set(qn("w:header"), tw(1.25))
    pgMar.set(qn("w:footer"), tw(cfg.footer_distance_cm))
    pgMar.set(qn("w:gutter"), "0")


class StreamFormatter:
    def __init__(self, cfg: FormatterConfig, log_cb: Optional[Callable[[str, str], None]] = None):
        self.cfg = cfg
        self.log_cb = log_cb
        self.source_type = "docx"
        # 标题检测缓冲状态
        self._title_buffer: list = []   # 开头连续居中段（w:p 元素），待决策
        self._decided = False
        self._title_ids: set = set()
        self._subtitle_ids: set = set()
        self._para_counter = -1
        self._processed = 0
        # 页眉页脚注入参数（run() 中计算后下发）
        self._hdr_rid: Optional[str] = None
        self._ftr_rid: Optional[str] = None
        self._need_header: bool = False
        self._need_footer: bool = False

    def _log(self, msg, level="INFO"):
        if self.log_cb:
            self.log_cb(msg, level)

    # ---------------- 主入口 ----------------
    def run(self, input_path: str, output_path: str):
        cfg = self.cfg
        zin = zipfile.ZipFile(input_path, "r")
        names = zin.namelist()
        rels = zin.read(RELS_NAME) if RELS_NAME in names else None
        ct = zin.read(CT_NAME) if CT_NAME in names else None

        need_header = bool(cfg.header_enabled and cfg.header_text.strip())
        need_footer = bool(cfg.page_number or cfg.footer_text.strip())
        hdr_rid = ftr_rid = None
        hdr_part = ftr_part = None
        hdr_xml = ftr_xml = None

        if need_header or need_footer:
            if rels is None:
                rels = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        b'<Relationships xmlns="%s"></Relationships>' % RELS_NS.encode())
            if ct is None:
                ct = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      b'<Types xmlns="%s"></Types>' % CT_NS.encode())
            if need_header:
                hi = _next_part_index(names, "header")
                hdr_part = "word/header%d.xml" % hi
                hdr_rid = "rIdWfpHeader"
                hdr_xml = _make_header_xml(cfg)
                # rels 的 Target 相对 word/ 目录（rels 文件所在目录）
                rels = _add_rels(rels, hdr_part.split("/", 1)[1], hdr_rid, HEADER_REL_TYPE)
                ct = _add_content_type(ct, "/" + hdr_part, HEADER_CT)
            if need_footer:
                fi = _next_part_index(names, "footer")
                ftr_part = "word/footer%d.xml" % fi
                ftr_rid = "rIdWfpFooter"
                ftr_xml = _make_footer_xml(cfg)
                rels = _add_rels(rels, ftr_part.split("/", 1)[1], ftr_rid, FOOTER_REL_TYPE)
                ct = _add_content_type(ct, "/" + ftr_part, FOOTER_CT)

        self._hdr_rid = hdr_rid
        self._ftr_rid = ftr_rid
        self._need_header = need_header
        self._need_footer = need_footer

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n == DOC_NAME:
                    continue
                if n == RELS_NAME:
                    zout.writestr(n, rels)
                    continue
                if n == CT_NAME:
                    zout.writestr(n, ct)
                    continue
                zout.writestr(n, zin.read(n))
            src = zin.open(DOC_NAME)
            with zout.open(DOC_NAME, "w") as out:
                self._transform(src, out)
            if need_header:
                zout.writestr(hdr_part, hdr_xml)
            if need_footer:
                zout.writestr(ftr_part, ftr_xml)
        zin.close()
        self._log("排版进度 100%", "INFO")

    # ---------------- 流式转换 ----------------
    def _transform(self, src, out):
        context = etree.iterparse(src, events=("start", "end"))
        depth = 0
        in_body = False
        doc_cm = None
        body_cm = None
        with etree.xmlfile(out, encoding="UTF-8") as xf:
            xf.write_declaration()
            for event, elem in context:
                if event == "start":
                    depth += 1
                    if elem.tag == qn("w:document"):
                        doc_cm = xf.element(qn("w:document"), nsmap=elem.nsmap)
                        doc_cm.__enter__()
                    elif elem.tag == qn("w:body"):
                        in_body = True
                        body_cm = xf.element(qn("w:body"), nsmap=elem.nsmap)
                        body_cm.__enter__()
                else:
                    depth -= 1
                    if elem.tag == qn("w:document"):
                        if doc_cm is not None:
                            doc_cm.__exit__(None, None, None)
                    elif elem.tag == qn("w:body"):
                        in_body = False
                        if not self._decided and self._title_buffer:
                            self._flush_title_region(xf)
                        if body_cm is not None:
                            body_cm.__exit__(None, None, None)
                    elif in_body and depth == 2:
                        # 顶层 body 子元素（w:p / w:tbl / w:sectPr / 其他）
                        self._handle_top_level(elem, xf, depth)
                        parent = elem.getparent()
                        if parent is not None:
                            elem.clear()
                            parent.remove(elem)
                        else:
                            elem.clear()

    # ---------------- 顶层块处理 ----------------
    def _handle_top_level(self, elem, xf, depth):
        tag = elem.tag
        if tag == W_SECTPR:
            # 分节符：应用页边距 + 注入页眉页脚引用（标书横向页/页眉分节依赖它）
            _apply_sectpr_margins(elem, self.cfg)
            _inject_sectpr_refs(elem, self._hdr_rid, self._ftr_rid,
                               self._need_header, self._need_footer)
            xf.write(elem)
            return
        if tag == W_P:
            self._handle_paragraph(elem, xf)
        elif tag == W_TBL:
            self._handle_table(elem, xf)
        else:
            # 其他顶层元素（如书签起止）原样透传
            xf.write(elem)

    def _handle_paragraph(self, elem, xf):
        self._para_counter += 1
        proxy = _to_proxy(elem)
        # 段落内联分节符（横向附录页等）：同样应用页边距与页眉页脚引用
        inline_sect = proxy.find(INLINE_SECTPR_XPATH)
        if inline_sect is not None:
            _apply_sectpr_margins(inline_sect, self.cfg)
            _inject_sectpr_refs(inline_sect, self._hdr_rid, self._ftr_rid,
                                self._need_header, self._need_footer)
        p = Paragraph(proxy, None)
        text = _para_text(p).strip()
        if not self._decided:
            if self.source_type in ("txt", "md"):
                if text and not self._title_buffer:
                    # 首个非空段落即题目
                    self._title_buffer.append(proxy)
                    return
                # 题目区结束，先落盘已缓冲的题目
                self._flush_title_region(xf)
            else:
                if text == "":
                    if self._title_buffer:
                        self._flush_title_region(xf)
                    xf.write(proxy)
                    return
                if p.alignment == 1:  # CENTER
                    self._title_buffer.append(proxy)
                    return
                self._flush_title_region(xf)
            # 落盘后继续按普通段落处理当前段
        self._process_and_write_paragraph(proxy, xf)

    def _handle_table(self, elem, xf):
        if not self._decided and self._title_buffer:
            self._flush_title_region(xf)
        if self.cfg.enable_table_formatting:
            proxy = _to_proxy(elem)
            format_tables([Table(proxy, None)], self.cfg)
            xf.write(proxy)
        else:
            xf.write(elem)

    # ---------------- 标题决策与落盘 ----------------
    def _flush_title_region(self, xf):
        self._decided = True
        buf = self._title_buffer
        self._title_buffer = []
        if not buf:
            return
        if self.source_type in ("txt", "md"):
            for e in buf:
                self._apply_role(e, "title")
                xf.write(e)
            return
        # docx：以首个居中段的字体/字号为基准，不同者判为副标题
        infos = [get_para_font_info(Paragraph(e, None)) for e in buf]
        base_name, base_size = infos[0]
        group = "title"
        ref_name, ref_size = base_name, base_size
        pending_normal = []
        for i, (name, size) in enumerate(infos):
            same = (name == ref_name) and (size == ref_size or (name is None and ref_name is None))
            if group == "title":
                if same or (name is None and ref_name is None):
                    self._apply_role(buf[i], "title")
                    xf.write(buf[i])
                else:
                    group = "subtitle"
                    ref_name, ref_size = name, size
                    self._apply_role(buf[i], "subtitle")
                    xf.write(buf[i])
            else:
                if same or (name is None and ref_name is None):
                    self._apply_role(buf[i], "subtitle")
                    xf.write(buf[i])
                else:
                    # 副标题之后再出现不同字体：回归普通段落处理
                    pending_normal.append(buf[i])
        for e in pending_normal:
            self._process_and_write_paragraph(e, xf)

    def _apply_role(self, elem, role):
        p = Paragraph(elem, None)
        apply_para_format(p, role, self.cfg)

    def _process_and_write_paragraph(self, elem, xf):
        p = Paragraph(elem, None)
        ptype = classify_paragraph(p, self._para_counter, self._title_ids, self._subtitle_ids, self.cfg)
        if ptype == "attachment":
            add_page_break_before(p)
            if self.cfg.enable_attachment_formatting:
                apply_para_format(p, "attachment", self.cfg)
        elif ptype == "h2":
            if not apply_h2_inline_split(p, self.cfg):
                apply_para_format(p, "h2", self.cfg)
        else:
            apply_para_format(p, ptype, self.cfg)
        if self.cfg.normalize_punctuation and _para_text(p).strip():
            for r in p.runs:
                r.text = normalize_punctuation(r.text)
        self._processed += 1
        if self._processed % 5000 == 0:
            self._log(f"已流式处理 {self._processed} 段（内存恒定）", "INFO")
        xf.write(elem)


def stream_format_document(input_path: str, output_path: str, cfg: FormatterConfig,
                           source_type: str = "docx",
                           log_cb: Optional[Callable[[str, str], None]] = None):
    """流式排版单个 .docx。返回输出路径或抛出异常。"""
    sf = StreamFormatter(cfg, log_cb)
    sf.source_type = source_type
    sf.run(input_path, output_path)
    return output_path
