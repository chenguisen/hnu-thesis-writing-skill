#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从既有湖南大学学位论文中“原样抽取”前置部分，生成可复用骨架。

与“按规范重建版式”不同，这里直接复制原文档的段落 XML（pPr/rPr、制表位、
段前段后、对齐、分节、页眉页脚、题注/目录域），只替换文字内容。因此骨架的
间距、对齐、缩进与原论文完全一致。

用法：
    python scripts/extract_front_matter.py \
        --source 陈桂森-博士学位论文.docx \
        --out 论文骨架.docx \
        --title "论文题目" --author "张三" --student-id B00000000

默认值都是占位符；不传的参数保留原文档内容（学号、姓名等建议显式传参）。
"""

from __future__ import annotations

import argparse
import re
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
XML = "{http://www.w3.org/XML/1998/namespace}"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

REL_MAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
REL_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
REL_WEB = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

CT_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_SETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
CT_WEB = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
CT_FONT = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

KEEP_TYPES = {
    REL_STYLES: CT_STYLES,
    REL_NUMBERING: CT_NUMBERING,
    REL_SETTINGS: CT_SETTINGS,
    REL_WEB: CT_WEB,
    REL_FONT: CT_FONT,
    REL_THEME: CT_THEME,
    REL_FOOTNOTES: CT_FOOTNOTES,
    REL_ENDNOTES: CT_ENDNOTES,
    REL_HEADER: CT_HEADER,
    REL_FOOTER: CT_FOOTER,
    REL_IMAGE: "image/png",
}

# 前置部分要替换的标签（去空格后匹配），值是 argparse 的属性名
COVER_LABELS = {
    "学号": "student_id",
    "学位申请人姓名": "author",
    "培养单位": "college",
    "导师姓名及职称": "supervisor",
    "学科专业": "major",
    "研究方向": "field",
    "论文提交日期": "date",
}

TABLE_LABELS = {
    "学位申请人姓名": "author",
    "导师姓名及职称": "supervisor",
    "培养单位": "college",
    "学科专业": "major",
    "论文提交日期": "date",
    "论文答辩日期": "defense_date",
    "答辩委员会主席": "chair",
}

RE_CHAPTER = re.compile(r"^第\s*([0-9一二三四五六七八九十]+)\s*章")


# ---------------------------------------------------------------------------
# XML 小工具
# ---------------------------------------------------------------------------


def text_of(el) -> str:
    out = []
    for node in el.iter():
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag == W + "tab":
            out.append("\t")
    return "".join(out)


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text)


def runs(p):
    return p.findall(W + "r")


def run_text(run) -> str:
    return "".join(t.text or "" for t in run.iter(W + "t"))


def set_run_text(run, text: str) -> None:
    for t in list(run.iter(W + "t")):
        parent = t.getparent()
        parent.remove(t)
    t = etree.SubElement(run, W + "t")
    t.set(XML + "space", "preserve")
    t.text = text


def first_rpr(p):
    for r in runs(p):
        rpr = r.find(W + "rPr")
        if rpr is not None:
            return deepcopy(rpr)
    ppr = p.find(W + "pPr")
    if ppr is not None:
        rpr = ppr.find(W + "rPr")
        if rpr is not None:
            return deepcopy(rpr)
    return None


def new_run(text: str, rpr=None):
    r = etree.Element(W + "r")
    if rpr is not None:
        r.append(deepcopy(rpr))
    t = etree.SubElement(r, W + "t")
    t.set(XML + "space", "preserve")
    t.text = text
    return r


def set_paragraph_text(p, text: str) -> None:
    """整段替换：保留第一段的段落属性与第一处 run 的字符格式。"""
    rs = runs(p)
    if not rs:
        p.append(new_run(text, first_rpr(p)))
        return
    set_run_text(rs[0], text)
    for r in rs[1:]:
        p.remove(r)


def replace_after_label(p, label: str, value: str) -> bool:
    """把“标签”之后的值整体替换掉（保留标签、制表符与原有格式）。"""
    key = compact(label)
    positions: list[tuple[str, int]] = []  # (char, run index)
    for ri, r in enumerate(runs(p)):
        for ch in run_text(r):
            if not ch.isspace():
                positions.append((ch, ri))
    flat = "".join(ch for ch, _ in positions)
    pos = flat.find(key)
    if pos < 0:
        return False
    end = pos + len(key)
    if end >= len(positions):
        return False
    start_run = positions[end - 1][1] + 1
    rs = runs(p)
    target = None
    for r in rs[start_run:]:
        if run_text(r).strip():
            target = r
            break
    if target is None:
        return False
    set_run_text(target, value)
    seen = False
    for r in rs[start_run:]:
        if r is target:
            seen = True
            continue
        if seen and run_text(r).strip():
            set_run_text(r, "")
    return True


# ---------------------------------------------------------------------------
# 定位
# ---------------------------------------------------------------------------


def style_id(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return ""
    st = ppr.find(W + "pStyle")
    return st.get(W + "val") if st is not None else ""


def num_id(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None
    np = ppr.find(W + "numPr")
    if np is None:
        return None
    nid = np.find(W + "numId")
    return nid.get(W + "val") if nid is not None else None


def load_style_numbers(z) -> dict[str, str]:
    """styleId -> 样式上挂的 numId（标题自动编号常挂在样式上）。"""
    out: dict[str, str] = {}
    try:
        styles = etree.fromstring(z.read("word/styles.xml"))
    except KeyError:
        return out
    for st in styles.findall(W + "style"):
        sid = st.get(W + "styleId")
        ppr = st.find(W + "pPr")
        if ppr is None:
            continue
        np = ppr.find(W + "numPr")
        if np is None:
            continue
        nid = np.find(W + "numId")
        if nid is not None and sid:
            out[sid] = nid.get(W + "val")
    return out


def effective_num_id(p, style_num: dict[str, str]) -> str | None:
    direct = num_id(p)
    if direct is not None:
        return direct
    return style_num.get(style_id(p))


def find_body_start(blocks, style_num: dict[str, str]) -> int:
    """正文第 1 章标题：Heading 1 且挂了自动编号（numId≠0，可能在样式上）。"""
    for i, el in enumerate(blocks):
        if el.tag != W + "p":
            continue
        if style_id(el) not in ("1", "Heading1"):
            continue
        nid = effective_num_id(el, style_num)
        if (nid and nid != "0") or RE_CHAPTER.match(text_of(el).strip()):
            return i
    raise SystemExit("未找到正文第 1 章标题（Heading 1 + 自动编号），无法确定前置部分边界")


def find_paragraph(blocks, predicate, start=0, end=None):
    for i in range(start, len(blocks) if end is None else end):
        el = blocks[i]
        if el.tag == W + "p" and predicate(el):
            return i
    return None


def collapse_abstract(blocks, heading_compact: str, stop_compacts, placeholder: str):
    h = find_paragraph(blocks, lambda p: compact(text_of(p)) == heading_compact)
    if h is None:
        return
    stop = None
    for i in range(h + 1, len(blocks)):
        el = blocks[i]
        if el.tag != W + "p":
            continue
        c = compact(text_of(el))
        if any(c.startswith(s) for s in stop_compacts):
            stop = i
            break
    if stop is None:
        return
    keep = h + 1
    if keep < stop:
        set_paragraph_text(blocks[keep], placeholder)
        del blocks[keep + 1:stop]


def truncate_field_result(p, placeholder: str) -> bool:
    """保留 TOC 域到 separate 为止，结果替换为一行占位文字。"""
    sep = None
    for fc in p.iter(W + "fldChar"):
        if fc.get(W + "fldCharType") == "separate":
            sep = fc
            break
    if sep is None:
        return False
    sep_run = sep.getparent()
    keep_children = []
    for child in list(p):
        if child is sep_run:
            keep_children.append(child)
            break
        keep_children.append(child)
        if child.tag == W + "pPr":
            continue
        # 只保留 pPr 与 separate 之前的 run
    # 重建：pPr + 直到 sep_run 的 run
    for child in list(p):
        if child.tag == W + "pPr":
            continue
        p.remove(child)
    for child in keep_children:
        if child.tag != W + "pPr":
            p.append(child)
    p.append(new_run(placeholder, first_rpr(p) if runs(p) else None))
    for fc in p.iter(W + "fldChar"):
        if fc.get(W + "fldCharType") == "begin":
            fc.set(W + "dirty", "true")
    return True


def field_end_paragraph(blocks, start) -> int | None:
    depth = 0
    for j in range(start, len(blocks)):
        for fc in blocks[j].iter(W + "fldChar"):
            typ = fc.get(W + "fldCharType")
            if typ == "begin":
                depth += 1
            elif typ == "end":
                depth -= 1
                if depth == 0:
                    return j
    return None


def has_sectpr(p) -> bool:
    ppr = p.find(W + "pPr")
    return ppr is not None and ppr.find(W + "sectPr") is not None


def move_end_field(start_p, end_p) -> None:
    end_run = None
    for r in runs(end_p):
        fc = r.find(W + "fldChar")
        if fc is not None and fc.get(W + "fldCharType") == "end":
            end_run = r
            break
    if end_run is None:
        return
    if has_sectpr(end_p):
        for r in runs(end_p):
            if r is not end_run:
                end_p.remove(r)
    else:
        start_p.append(deepcopy(end_run))


def collapse_toc_fields(blocks) -> None:
    i = 0
    while i < len(blocks):
        el = blocks[i]
        if el.tag == W + "p":
            instr = "".join(t.text or "" for t in el.iter(W + "instrText"))
            if "TOC" in instr:
                end = field_end_paragraph(blocks, i)
                if end is not None and end > i:
                    placeholder = (
                        "（按 Ctrl+A、F9 更新：目录）"
                        if '"图"' not in instr and '"表"' not in instr
                        else "（按 Ctrl+A、F9 更新：索引）"
                    )
                    truncate_field_result(el, placeholder)
                    move_end_field(el, blocks[end])
                    if not has_sectpr(blocks[end]):
                        del blocks[i + 1:end + 1]
                    else:
                        del blocks[i + 1:end]
                    i += 1
                    continue
        i += 1


# ---------------------------------------------------------------------------
# 文本替换
# ---------------------------------------------------------------------------


def rewrite_front_matter(blocks, args) -> None:
    for el in blocks:
        if el.tag == W + "tbl":
            rows = el.findall(W + "tr")
            for tr in rows:
                cells = tr.findall(W + "tc")
                if len(cells) < 2:
                    continue
                label = compact(text_of(cells[0])).rstrip("：:")
                for key, attr in TABLE_LABELS.items():
                    if label == key:
                        value = getattr(args, attr)
                        if value:
                            paras = cells[1].findall(W + "p")
                            if paras:
                                set_paragraph_text(paras[0], value)
                                for extra in paras[1:]:
                                    cells[1].remove(extra)
                        break
            continue
        if el.tag != W + "p":
            continue

        txt = compact(text_of(el))
        # 标题（封面、内封、英文内封）
        if txt and args.title and txt in (compact(args.source_title),) and el.findall(W + "r"):
            set_paragraph_text(el, args.title)
            continue
        if txt and args.title_en and txt == compact(args.source_title_en):
            set_paragraph_text(el, args.title_en)
            continue
        # 学位类别
        if txt in ("博士学位论文", "硕士学位论文") and args.degree_label:
            set_paragraph_text(el, args.degree_label)
            continue
        if txt.startswith("湖南大学") and txt.endswith("学位论文"):
            set_paragraph_text(el, "湖南大学" + args.degree_label)
            continue
        # 封面/内封字段：标签后的值
        for label, attr in COVER_LABELS.items():
            value = getattr(args, attr)
            if not value or label not in compact(text_of(el)):
                continue
            if not replace_after_label(el, label + "：", value):
                replace_after_label(el, label + ":", value) or replace_after_label(el, label, value)
            break
        # 英文内封零散字段
        mapping = {
            "姓名": args.author_en,
            "学位": args.degree_en,
            "博士": args.degree_en,
            "bachelor": args.bachelor,
            "supervisor": args.supervisor_en,
        }
        if args.author_en and compact(text_of(el)).startswith("CHEN"):
            set_paragraph_text(el, args.author_en)
        elif args.bachelor and compact(text_of(el)).startswith("B.E."):
            set_paragraph_text(el, args.bachelor)
        elif args.degree_en and compact(text_of(el)).lower() in (
            "doctorofengineering", "doctorofphilosophy", "masterofengineering"
        ):
            set_paragraph_text(el, args.degree_en)
        elif args.major_en and compact(text_of(el)) in (
            "materialsscienceandengineering", "materialscienceandengineering"
        ):
            set_paragraph_text(el, args.major_en)
        elif args.supervisor_en and compact(text_of(el)).startswith("Professor"):
            set_paragraph_text(el, args.supervisor_en)
        elif args.date_en and re.match(r"^[A-Z][a-z]+,\s*\d{4}$", text_of(el).strip()):
            set_paragraph_text(el, args.date_en)


def rewrite_headers(parts, args) -> None:
    for name, data in parts.items():
        if not re.match(r"word/header\d+\.xml$", name):
            continue
        root = etree.fromstring(data)
        changed = False
        for t in root.iter(W + "t"):
            if not t.text or not t.text.strip():
                continue
            c = compact(t.text)
            if c in ("博士学位论文", "硕士学位论文"):
                t.text = args.degree_label
                changed = True
            elif args.source_title and c == compact(args.source_title):
                t.text = args.title
                changed = True
            elif len(c) >= 8:
                # 模板里残留的旧论文题目（不等于本论文题目）一律换成本论文题目
                t.text = args.title
                changed = True
        if changed:
            parts[name] = etree.tostring(root, xml_declaration=True,
                                         encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# 构建
# ---------------------------------------------------------------------------


def insert_page_break(p) -> None:
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        p.insert(0, ppr)
    if ppr.find(W + "pageBreakBefore") is not None:
        return
    el = etree.Element(W + "pageBreakBefore")
    successors = {
        W + t for t in (
            "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr",
            "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
            "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
            "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
        )
    }
    for child in ppr:
        if child.tag in successors:
            child.addprevious(el)
            return
    ppr.append(el)


def fix_declaration_headers(body) -> None:
    """声明页在原文档里沿用了旧页眉（偶数页是别人的题目），统一成摘要节的页眉。"""
    sects = [s for s in body.iter(W + "sectPr")]
    if len(sects) < 3:
        return
    decl, abstract = sects[1], sects[2]
    refs = [deepcopy(r) for r in abstract.findall(W + "headerReference")]
    if not refs:
        return
    for r in decl.findall(W + "headerReference"):
        decl.remove(r)
    for r in reversed(refs):
        decl.insert(0, r)


def merge_empty_section_breaks(blocks) -> None:
    """把“只装着分节符的空段落”并到上一段，避免分节符前多出一个空行/空白页。"""
    i = 1
    while i < len(blocks):
        el = blocks[i]
        if el.tag == W + "p" and has_sectpr(el) and not text_of(el).strip():
            prev = blocks[i - 1]
            if prev.tag == W + "p" and not has_sectpr(prev):
                prev_ppr = prev.find(W + "pPr")
                if prev_ppr is None:
                    prev_ppr = etree.Element(W + "pPr")
                    prev.insert(0, prev_ppr)
                sect = el.find(W + "pPr").find(W + "sectPr")
                prev_ppr.append(sect)
                del blocks[i]
                continue
        i += 1


def normalize_section_breaks(body) -> None:
    """把“奇数页/偶数页”分节符改成普通下一页分节符。

    原论文为双面印刷设置了 oddPage 分节符；骨架内容变短后会导致分节符前多出空白页。
    连续分节符（continuous）与其他类型保持不动。
    """
    for sect in body.iter(W + "sectPr"):
        for t in sect.findall(W + "type"):
            if t.get(W + "val") in ("oddPage", "evenPage"):
                t.set(W + "val", "nextPage")


def fix_page_numbering(body) -> None:
    """前置部分页码：封面节用负起始值，声明页重新从 Ⅰ 起，其余分节续编。

    WPS 在“声明页从 Ⅰ 重新编号”而封面节也处于 Ⅰ～Ⅲ 时，会在分节前插入一张空白页
    （页码撞号）。把封面节（不显示页码）的起始值设为负数即可避免，且不影响任何可见页码。
    源文档模板又给摘要/Abstract/目录/索引各写了 `w:pgNumType start="1"`，
    会把页码反复重置为 Ⅰ，这里只保留声明页这一处。
    """
    sects = list(body.iter(W + "sectPr"))
    if sects:
        cover = sects[0]
        pg = cover.find(W + "pgNumType")
        if pg is None:
            pg = etree.Element(W + "pgNumType")
            index = len(cover)
            for i, child in enumerate(cover):
                if child.tag in (W + "cols", W + "formProt", W + "vAlign",
                                 W + "noEndnote", W + "titlePg"):
                    index = i
                    break
            cover.insert(index, pg)
        pg.set(W + "fmt", "upperRoman")
        pg.set(W + "start", "-2")

    kept_restart = False
    for sect in sects[1:]:
        pg = sect.find(W + "pgNumType")
        if pg is None or pg.get(W + "fmt") not in ("upperRoman", "lowerRoman"):
            continue
        if pg.get(W + "start") is None:
            continue
        if not kept_restart:
            kept_restart = True
        else:
            del pg.attrib[W + "start"]


def build(source: Path, args) -> tuple[dict[str, bytes], dict]:
    z = zipfile.ZipFile(source)
    style_num = load_style_numbers(z)
    doc = etree.fromstring(z.read("word/document.xml"))
    body = doc.find(W + "body")
    blocks = list(body)

    body_start = find_body_start(blocks, style_num)
    fm = [deepcopy(el) for el in blocks[:body_start]]
    src = blocks  # 原始块，用于取正文示例的格式

    if not getattr(args, "keep_text", False):
        rewrite_front_matter(fm, args)
        collapse_abstract(fm, "摘要", ("关键词",), args.abstract_placeholder)
        collapse_abstract(fm, "Abstract", ("Keywords", "Keywords:"), args.abstract_en_placeholder)
        # 关键词
        for p in fm:
            if p.tag == W + "p":
                c = compact(text_of(p))
                if c.startswith("关键词") and args.keywords:
                    replace_after_label(p, "关键词", args.keywords)
                if c.startswith("Keywords") and args.keywords_en:
                    replace_after_label(p, "Keywords", args.keywords_en)
        collapse_toc_fields(fm)
    merge_empty_section_breaks(fm)

    # 正文示例：章标题 + 节标题 + 一段占位正文
    chapter = deepcopy(src[body_start])
    body_out = [chapter]

    sec_idx = find_paragraph(src, lambda p: style_id(p) == "2", start=body_start)
    if sec_idx is not None:
        sec = deepcopy(src[sec_idx])
        set_paragraph_text(sec, "研究背景与意义")
        body_out.append(sec)

    para = etree.Element(W + "p")
    src_para = None
    for el in src[body_start + 1:body_start + 8]:
        if el.tag == W + "p" and style_id(el) not in ("1", "2", "3", "4"):
            src_para = el
            break
    if src_para is not None:
        ppr = src_para.find(W + "pPr")
        if ppr is not None:
            para.append(deepcopy(ppr))
    para.append(new_run(args.body_placeholder,
                        first_rpr(src_para) if src_para is not None else None))
    body_out.append(para)

    # 后置部分：结论 / 参考文献 / 附录A / 致谢（保持校模板顺序）
    def append_heading(text: str, start: int = body_start):
        idx = find_paragraph(src, lambda p, t=text: compact(text_of(p)) in (t, t + "："),
                             start=start)
        if idx is None:
            return None
        h = deepcopy(src[idx])
        insert_page_break(h)
        body_out.append(h)
        return idx

    append_heading("结论")

    ref_idx = append_heading("参考文献")
    if ref_idx is not None:
        entries = []
        for el in src[ref_idx + 1: ref_idx + 12]:
            if el.tag != W + "p":
                continue
            t = text_of(el).strip()
            if t.startswith("[") and "]" in t:
                entries.append(el)
            if len(entries) >= 2:
                break
        placeholders = [
            "[1] 作者. 文献题名[J]. 期刊名, 2020, 12(3): 45-52.",
            "[2] 作者. 书名[M]. 北京: 出版社, 2019: 10-25.",
        ]
        for el, text in zip(entries, placeholders):
            p = deepcopy(el)
            set_paragraph_text(p, text)
            body_out.append(p)

    idx = find_paragraph(src, lambda p: compact(text_of(p)).startswith("附录A"),
                         start=body_start)
    if idx is not None:
        h = deepcopy(src[idx])
        insert_page_break(h)
        body_out.append(h)

    append_heading("致谢")

    # 正文分节属性（阿拉伯页码从 1 开始）取自第 1 章末尾的 sectPr
    sect = None
    for el in src[body_start:]:
        if el.tag != W + "p":
            continue
        ppr = el.find(W + "pPr")
        if ppr is not None and ppr.find(W + "sectPr") is not None:
            sect = deepcopy(ppr.find(W + "sectPr"))
            break
    for child in list(body):
        body.remove(child)
    for el in fm + body_out:
        body.append(el)
    if sect is None:
        raise SystemExit("未找到正文分节符")
    body.append(sect)
    if not getattr(args, "keep_text", False):
        fix_declaration_headers(body)
        normalize_section_breaks(body)
        fix_page_numbering(body)

    new_document = etree.tostring(doc, xml_declaration=True,
                                  encoding="UTF-8", standalone=True)

    # 包组装：沿用源包（保留样式、编号、页眉页脚、文档属性与自定义 XML），
    # 只去掉正文图片、只保留封面 logo。这样最大程度保证 Word/WPS/LibreOffice 都能打开。
    parts: dict[str, bytes] = {}
    for info in z.infolist():
        name = info.filename
        if name == "word/document.xml":
            parts[name] = new_document
        elif name.startswith("word/media/"):
            if name == "word/media/image1.png":
                parts[name] = z.read(name)
        else:
            parts[name] = z.read(name)
    if not getattr(args, "keep_text", False):
        rewrite_headers(parts, args)

    # document.xml.rels：只删掉已移除图片的关系，其余关系原样保留
    rels_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
    for rel in list(rels_root):
        if rel.get("Type") == REL_IMAGE and rel.get("Target") != "media/image1.png":
            rels_root.remove(rel)
    parts["word/_rels/document.xml.rels"] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return parts, {"front_blocks": len(fm), "body_start": body_start}


def write_package(parts: dict[str, bytes], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="原样抽取学位论文前置部分，生成骨架")
    ap.add_argument("--source", required=True, help="原始 .docx（只读，不修改）")
    ap.add_argument("--out", required=True, help="输出骨架 .docx")
    ap.add_argument("--title", default="（论文题目）")
    ap.add_argument("--title-en", dest="title_en", default="(Title of the Dissertation)")
    ap.add_argument("--author", default="（姓名）")
    ap.add_argument("--author-en", dest="author_en", default="(NAME)")
    ap.add_argument("--student-id", dest="student_id", default="B00000000")
    ap.add_argument("--supervisor", default="（导师姓名及职称）")
    ap.add_argument("--supervisor-en", dest="supervisor_en", default="Professor XX")
    ap.add_argument("--college", default="（培养单位）")
    ap.add_argument("--major", default="（学科专业）")
    ap.add_argument("--major-en", dest="major_en", default="(Major)")
    ap.add_argument("--field", default="（研究方向）")
    ap.add_argument("--date", default="2026年5月1日")
    ap.add_argument("--date-en", dest="date_en", default="May, 2026")
    ap.add_argument("--defense-date", dest="defense_date", default="2026年5月20日")
    ap.add_argument("--chair", default="（答辩委员会主席）")
    ap.add_argument("--degree-label", dest="degree_label", default="博士学位论文")
    ap.add_argument("--degree-en", dest="degree_en", default="Doctor of Engineering")
    ap.add_argument("--bachelor", default="B.E. (Hunan University) 2020")
    ap.add_argument("--keywords", default="关键词1；关键词2；关键词3；关键词4")
    ap.add_argument("--keywords-en", dest="keywords_en",
                    default="keyword1; keyword2; keyword3; keyword4")
    ap.add_argument(
        "--abstract-placeholder", dest="abstract_placeholder",
        default="（在此撰写中文摘要：研究背景、方法、主要结果与结论；博士约 1200 字，硕士约 600 字。）")
    ap.add_argument(
        "--abstract-en-placeholder", dest="abstract_en_placeholder",
        default="(Write the English abstract here; it must match the Chinese abstract.)")
    ap.add_argument(
        "--body-placeholder", dest="body_placeholder",
        default="（在此撰写正文，引用文献在句末右上角标注[1]；宋体小 4 号、首行缩进 2 字符、固定行距 20 磅。）")
    ap.add_argument("--source-title", dest="source_title", default=None,
                    help="原论文中文题目；默认从封面自动识别")
    ap.add_argument("--source-title-en", dest="source_title_en", default=None,
                    help="原论文英文题目；默认从英文内封自动识别")
    ap.add_argument("--keep-text", dest="keep_text", action="store_true",
                    help="只抽取、不替换文字（用于和原文档比对版式）")
    args = ap.parse_args(argv)
    return args


def detect_titles(source: Path) -> tuple[str, str]:
    """从原文档识别中英文题目：封面第二个 22pt 黑体段落 / 英文内封 15pt 段落。"""
    z = zipfile.ZipFile(source)
    style_num = load_style_numbers(z)
    doc = etree.fromstring(z.read("word/document.xml"))
    body = doc.find(W + "body")
    blocks = list(body)
    body_start = find_body_start(blocks, style_num)
    cn_candidates: list[str] = []
    en_candidates: list[str] = []
    for el in blocks[:body_start]:
        if el.tag != W + "p":
            continue
        text = text_of(el).strip()
        if not text:
            continue
        sizes = set()
        for r in el.findall(W + "r"):
            rpr = r.find(W + "rPr")
            sz = rpr.find(W + "sz") if rpr is not None else None
            if sz is not None:
                sizes.add(sz.get(W + "val"))
        if "44" in sizes and text not in ("博士学位论文", "硕士学位论文"):
            cn_candidates.append(text)
        if "30" in sizes and re.search(r"[A-Za-z]", text):
            en_candidates.append(text)
    cn = max(cn_candidates, key=len) if cn_candidates else ""
    en = max(en_candidates, key=len) if en_candidates else ""
    return cn, en


def main(argv=None) -> int:
    args = parse_args(argv)
    source = Path(args.source)
    if not source.exists():
        raise SystemExit(f"找不到文件：{source}")
    cn, en = detect_titles(source)
    if not args.source_title:
        args.source_title = cn
    if not args.source_title_en:
        args.source_title_en = en
    parts, stats = build(source, args)
    out = Path(args.out)
    write_package(parts, out)
    print(f"已生成：{out}")
    print(f"  前置部分块数：{stats['front_blocks']}，正文起点：第 {stats['body_start']} 块")
    print(f"  原题目：{args.source_title!r} / {args.source_title_en!r}")
    print("  版式（间距/对齐/缩进/分节/页眉页脚）沿用原文档，只替换了文字。")
    print("  打开后按 Ctrl+A、F9 更新目录与索引。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
