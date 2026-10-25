#!/usr/bin/env python3
"""生成湖南大学博士/硕士学位论文骨架（.docx）。

这个脚本把“封面—内封—声明—摘要—目录—索引—正文”这一整套前置部分的
**可编辑版式**做出来，避免从空白文档手工搭版式时漏项。结构参照一份
答辩通过的湖南大学博士学位论文（已按校规范核对），但所有内容都是占位符，
不含任何真实论文数据。

用法：
    python scripts/thesis_scaffold.py --out thesis.docx
    python scripts/thesis_scaffold.py --degree master --title "基于××的××研究" --out t.docx

生成后：
1. 打开文件，按提示替换封面/内封占位符（姓名、学号、导师、题目等）。
2. 全选（Ctrl+A）→ F9 更新域，目录、插图索引、附表索引、图表编号会刷新。
3. 用 scripts/thesis_lint.py 自检。

只依赖 python-docx（见 requirements.txt）。
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SONG = "宋体"
HEI = "黑体"
TNR = "Times New Roman"

# 正文版心：A4，上下 2.2 cm，左右 2.5 cm，页眉页脚距边界 1.1 cm（见 references/formatting.md）
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
MARGIN_TB_CM = 2.2
MARGIN_LR_CM = 2.5
HF_DISTANCE_CM = 1.1
TEXT_WIDTH_CM = PAGE_WIDTH_CM - 2 * MARGIN_LR_CM  # 16.0 cm

DEGREE_EN = {
    "doctor": "Doctor of Engineering",
    "master": "Master of Engineering",
    "professional": "Master of Engineering",
}


# ---------------------------------------------------------------------------
# 底层 XML 辅助
# ---------------------------------------------------------------------------


def set_run_font(run, *, ascii_font: str = TNR, eastasia: str = SONG,
                 size: float | None = None, bold: bool | None = None,
                 underline: bool | None = None) -> None:
    """设置一个 run 的西文/中文字体、字号、加粗与下划线。"""
    run.font.name = ascii_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), eastasia)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if underline is not None:
        run.font.underline = underline


def add_text(paragraph, text: str, **kwargs):
    if "eastasia" not in kwargs:
        style_name = (getattr(paragraph.style, "name", "") or "").lower()
        kwargs["eastasia"] = HEI if "heading" in style_name or "标题" in style_name else SONG
    run = paragraph.add_run(text)
    set_run_font(run, **kwargs)
    return run


def add_field(paragraph, instruction: str, cached: str = "", *,
              dirty: bool = False, **kwargs):
    """插入一个 Word 域，并写入“更新前先显示”的缓存文本。

    域在 Word/WPS 里按 F9 更新；缓存文本保证未更新时也能看到合理内容。
    """
    begin = paragraph.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    if dirty:
        fld.set(qn("w:dirty"), "true")
    begin._r.append(fld)

    instr_run = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    instr_run._r.append(instr)

    sep = paragraph.add_run()
    sep_el = OxmlElement("w:fldChar")
    sep_el.set(qn("w:fldCharType"), "separate")
    sep._r.append(sep_el)

    result = paragraph.add_run(cached)

    end = paragraph.add_run()
    end_el = OxmlElement("w:fldChar")
    end_el.set(qn("w:fldCharType"), "end")
    end._r.append(end_el)

    for run in (begin, instr_run, sep, result, end):
        set_run_font(run, **kwargs)
    return result


def add_bookmark_start(paragraph, name: str, bid: int) -> None:
    el = OxmlElement("w:bookmarkStart")
    el.set(qn("w:id"), str(bid))
    el.set(qn("w:name"), name)
    paragraph._p.append(el)


def add_bookmark_end(paragraph, bid: int) -> None:
    el = OxmlElement("w:bookmarkEnd")
    el.set(qn("w:id"), str(bid))
    paragraph._p.append(el)


def set_paragraph_border_bottom(paragraph, *, val: str = "thinThickLargeGap",
                                sz: int = 12, space: int = 1) -> None:
    """给段落加下边框（页眉双线用它）。按 OOXML 顺序插入，避免 Word 修复提示。"""
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.insert_element_before(
            pbdr,
            "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
            "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
            "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
            "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
            "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
            "w:pPrChange",
        )
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), val)
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), "auto")
    pbdr.append(bottom)


def set_page_numbering(section, fmt: str | None = None, start: int | None = None) -> None:
    """设置节的页码格式/起始页码（w:pgNumType）。fmt/start 都为 None 时清除。"""
    sect_pr = section._sectPr
    pg = sect_pr.find(qn("w:pgNumType"))
    if pg is None:
        if fmt is None and start is None:
            return
        pg = OxmlElement("w:pgNumType")
        sect_pr.insert_element_before(
            pg,
            "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
            "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
            "w:printerSettings", "w:sectPrChange",
        )
    if fmt is not None:
        pg.set(qn("w:fmt"), fmt)
    elif start is not None:
        # 只改起始页码时保持原格式；word 会把缺失的 fmt 当作 decimal
        pass
    if start is not None:
        pg.set(qn("w:start"), str(start))
    if start is None and pg.get(qn("w:start")) is not None:
        del pg.attrib[qn("w:start")]


def inherits_continuous_numbering(section) -> None:
    """让本节页码接着上一节连续编排：去掉 add_section 复制来的 pgNumType。"""
    sect_pr = section._sectPr
    pg = sect_pr.find(qn("w:pgNumType"))
    if pg is not None:
        sect_pr.remove(pg)


# ---------------------------------------------------------------------------
# 样式
# ---------------------------------------------------------------------------


def setup_styles(doc) -> None:
    """正文、标题、题注、目录样式。数值见 references/formatting.md 与校模板。"""
    normal = doc.styles["Normal"]
    normal.font.name = TNR
    normal.font.size = Pt(12)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), TNR)
    rfonts.set(qn("w:hAnsi"), TNR)
    rfonts.set(qn("w:eastAsia"), SONG)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = Pt(20)          # 固定 20 磅，约 35 行/页
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(24)     # 首行缩进 2 字符（12 pt × 2）

    # 章 18 pt / 节 15 pt / 条 14 pt / 款 14 pt，黑体；段前段后按校模板
    heading_specs = [
        ("Heading 1", 18, WD_ALIGN_PARAGRAPH.CENTER, 20.45, 20.45),
        ("Heading 2", 15, WD_ALIGN_PARAGRAPH.LEFT, 12, 12),
        ("Heading 3", 14, WD_ALIGN_PARAGRAPH.LEFT, 12, 6),
        ("Heading 4", 14, WD_ALIGN_PARAGRAPH.LEFT, 12, 6),
    ]
    for name, size, align, before, after in heading_specs:
        style = doc.styles[name]
        style.font.name = TNR
        style.font.size = Pt(size)
        style.font.bold = False
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), TNR)
        rfonts.set(qn("w:hAnsi"), TNR)
        rfonts.set(qn("w:eastAsia"), HEI)
        hpf = style.paragraph_format
        hpf.alignment = align
        hpf.space_before = Pt(before)
        hpf.space_after = Pt(after)
        hpf.line_spacing = 1.0
        hpf.first_line_indent = Pt(0)
        hpf.keep_with_next = True
        hpf.keep_together = True
        # 默认模板的标题样式带主题色（蓝色），必须清掉主题色再设为黑色
        color = rpr.find(qn("w:color"))
        if color is not None:
            rpr.remove(color)
        style.font.color.rgb = RGBColor(0, 0, 0)

    # 题注：与答辩版论文一致——12 pt 黑体、居中、段前段后 6 磅
    caption = doc.styles.add_style("题注", WD_STYLE_TYPE.PARAGRAPH)
    caption.base_style = normal
    caption.font.size = Pt(12)
    caption.font.bold = False
    crpr = caption.element.get_or_add_rPr()
    crfonts = crpr.find(qn("w:rFonts"))
    if crfonts is None:
        crfonts = OxmlElement("w:rFonts")
        crpr.insert(0, crfonts)
    crfonts.set(qn("w:ascii"), TNR)
    crfonts.set(qn("w:hAnsi"), TNR)
    crfonts.set(qn("w:eastAsia"), HEI)
    cpf = caption.paragraph_format
    cpf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cpf.space_before = Pt(6)
    cpf.space_after = Pt(6)
    cpf.first_line_indent = Pt(0)
    cpf.line_spacing = 1.0

    # 目录条目：一级黑体、其余宋体；右对齐点线前导符
    toc_specs = [
        ("TOC 1", HEI, 0.0),
        ("TOC 2", SONG, 0.74),
        ("TOC 3", SONG, 1.48),
        ("TOC 4", SONG, 2.22),
    ]
    for name, font, indent_cm in toc_specs:
        try:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        except ValueError:  # 已存在
            style = doc.styles[name]
        style.base_style = normal
        style.font.size = Pt(12)
        style.font.name = TNR
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), TNR)
        rfonts.set(qn("w:hAnsi"), TNR)
        rfonts.set(qn("w:eastAsia"), font)
        tpf = style.paragraph_format
        tpf.left_indent = Cm(indent_cm)
        tpf.first_line_indent = Pt(0)
        tpf.line_spacing = 1.0
        tpf.tab_stops.add_tab_stop(
            Cm(TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )


# ---------------------------------------------------------------------------
# 标题自动编号：第 %1 章 / %1.%2 / %1.%2.%3 / %1.%2.%3.%4
# ---------------------------------------------------------------------------


def build_heading_numbering(doc) -> int:
    """注入多级列表，使章/节/条/款自动编号，题注的 STYLEREF 才能取到章号。"""
    numbering = doc.part.numbering_part.element
    abstract_id, num_id = 90, 90

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)

    levels = [
        (0, "第 %1 章", "Heading1", 36, "center"),
        (1, "%1.%2", "Heading2", 30, "left"),
        (2, "%1.%2.%3", "Heading3", 28, "left"),
        (3, "%1.%2.%3.%4", "Heading4", 28, "left"),
    ]
    for ilvl, lvl_text, style_id, half_points, jc in levels:
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))

        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)

        fmt = OxmlElement("w:numFmt")
        fmt.set(qn("w:val"), "decimal")
        lvl.append(fmt)

        pstyle = OxmlElement("w:pStyle")
        pstyle.set(qn("w:val"), style_id)
        lvl.append(pstyle)

        suff = OxmlElement("w:suff")
        suff.set(qn("w:val"), "space")
        lvl.append(suff)

        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), lvl_text)
        lvl.append(text)

        lvl_jc = OxmlElement("w:lvlJc")
        lvl_jc.set(qn("w:val"), jc)
        lvl.append(lvl_jc)

        ppr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:firstLine"), "0")
        ppr.append(ind)
        lvl.append(ppr)

        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), TNR)
        rfonts.set(qn("w:hAnsi"), TNR)
        rfonts.set(qn("w:eastAsia"), HEI)
        rpr.append(rfonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(half_points))
        rpr.append(sz)
        lvl.append(rpr)

        abstract.append(lvl)

    # OOXML 要求所有 abstractNum 排在所有 num 之前；直接 append 会排到 num 后面，
    # Word 会判定文件损坏。插到第一个 num 之前。
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)
    numbering.append(num)
    return num_id


def apply_numbering(paragraph, num_id: int, level: int) -> None:
    """给标题段落挂上编号；num_id=0 表示“取消编号”（结论/参考文献/致谢/附录）。"""
    ppr = paragraph._p.get_or_add_pPr()
    num_pr = ppr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        ppr.insert_element_before(
            num_pr,
            "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
            "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
            "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
            "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
            "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
            "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
            "w:pPrChange",
        )
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num_ref = OxmlElement("w:numId")
    num_ref.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl)
    num_pr.append(num_ref)


# ---------------------------------------------------------------------------
# 章节内容
# ---------------------------------------------------------------------------


def add_paragraph(doc, text: str = "", *, style: str | None = None,
                  align=None, font: str = TNR, eastasia: str = SONG,
                  size: float | None = 12, bold: bool | None = None,
                  underline: bool | None = None,
                  space_before: float | None = None,
                  space_after: float | None = None,
                  first_line_indent: float | None = None,
                  line_spacing: float | None = None,
                  keep_with_next: bool | None = None):
    p = doc.add_paragraph(style=style)
    if text:
        add_text(p, text, ascii_font=font, eastasia=eastasia, size=size,
                 bold=bold, underline=underline)
    pf = p.paragraph_format
    if align is not None:
        pf.alignment = align
    if space_before is not None:
        pf.space_before = Pt(space_before)
    if space_after is not None:
        pf.space_after = Pt(space_after)
    if first_line_indent is not None:
        pf.first_line_indent = Pt(first_line_indent)
    if line_spacing is not None:
        pf.line_spacing = Pt(line_spacing)
    if keep_with_next is not None:
        pf.keep_with_next = keep_with_next
    return p


def add_values_line(doc, label: str, value: str, *, indent_cm: float = 1.4,
                    tab_cm: float = 5.2) -> None:
    """封面/内封的“标签 + 下划线填写内容”一行。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent_cm)
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 1.5
    pf.tab_stops.add_tab_stop(Cm(tab_cm))
    add_text(p, label, eastasia=HEI, size=12)
    p.add_run("\t")  # 制表符不加下划线，否则会出现一段多余的下划线
    add_text(p, value, eastasia=SONG, size=12, underline=True)


def add_cover(doc, args) -> None:
    """中文封面。"""
    top = add_paragraph(doc, "", first_line_indent=0, line_spacing=28)
    top.paragraph_format.right_indent = Cm(0.45)
    add_text(top, "学校代号", eastasia=HEI, size=12)
    top.add_run("\t")
    add_text(top, args.school_code, eastasia=HEI, size=12, underline=True)
    add_text(top, "\t\t学号", eastasia=HEI, size=12)
    top.add_run("\t")
    add_text(top, args.student_id, eastasia=HEI, size=12, underline=True)

    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=28)
    p.paragraph_format.right_indent = Cm(0.45)
    add_text(p, "分 类 号", eastasia=HEI, size=12)
    p.add_run("\t")
    add_text(p, args.category, eastasia=HEI, size=12, underline=True)
    add_text(p, "\t\t密级", eastasia=HEI, size=12)
    p.add_run("\t")
    add_text(p, args.secret, eastasia=HEI, size=12, underline=True)

    add_paragraph(doc, "", first_line_indent=0)
    add_paragraph(doc, "", first_line_indent=0)
    add_paragraph(doc, args.degree_label, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=HEI, size=22, first_line_indent=0, line_spacing=28)
    add_paragraph(doc, "", first_line_indent=0)
    add_paragraph(doc, args.title, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=HEI, size=22, first_line_indent=0, line_spacing=32)
    add_paragraph(doc, "", first_line_indent=0)

    add_values_line(doc, "学位申请人姓名", args.author)
    add_values_line(doc, "培  养  单  位", args.college)
    add_values_line(doc, "导师姓名及职称", args.supervisor)
    add_values_line(doc, "学  科  专  业", args.major)
    add_values_line(doc, "研  究  方  向", args.field)
    add_values_line(doc, "论文提交日期", args.date)


def add_english_cover(doc, args) -> None:
    """英文内封。行距压到 18 磅，保证整页放得下（否则日期会被挤到下一页）。"""
    line = 18
    for label, value in (
        ("学校代号：", args.school_code),
        ("学号：", args.student_id),
        ("密级：", args.secret),
    ):
        p = add_paragraph(doc, "", first_line_indent=0, line_spacing=line)
        p.paragraph_format.left_indent = Cm(2.74)
        add_text(p, label, eastasia=HEI, size=12)
        add_text(p, value, eastasia=TNR, size=12)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)
    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    p.paragraph_format.left_indent = Cm(2.74)
    add_text(p, "湖南大学" + args.degree_label, eastasia=SONG, size=18)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)
    add_paragraph(doc, args.title, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=HEI, size=22, first_line_indent=0, line_spacing=26)
    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)

    table = doc.add_table(rows=7, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    rows = [
        ("学位申请人姓名：", args.author),
        ("导师姓名及职称：", args.supervisor),
        ("培养单位：", args.college),
        ("学科专业：", args.major),
        ("论文提交日期：", args.date),
        ("论文答辩日期：", args.defense_date),
        ("答辩委员会主席：", args.chair),
    ]
    for row, (label, value) in zip(table.rows, rows):
        row.cells[0].width = Cm(4.1)
        row.cells[1].width = Cm(9.1)
        lp = row.cells[0].paragraphs[0]
        lp.paragraph_format.first_line_indent = Pt(0)
        lp.paragraph_format.line_spacing = Pt(line)
        add_text(lp, label, eastasia=SONG, size=12)
        rp = row.cells[1].paragraphs[0]
        rp.paragraph_format.first_line_indent = Pt(0)
        rp.paragraph_format.line_spacing = Pt(line)
        add_text(rp, value, eastasia=SONG, size=12)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)
    add_paragraph(doc, args.title_en, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=TNR, size=15, first_line_indent=0, line_spacing=20)
    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)

    for text in (
        "by",
        args.author_en,
        args.bachelor,
        "A dissertation submitted in partial satisfaction of the",
        "requirements for the degree of",
        args.degree_en,
        "in",
        args.major_en,
        "in the",
        "Graduate School",
        "of",
        "Hunan University",
    ):
        add_paragraph(doc, text, align=WD_ALIGN_PARAGRAPH.CENTER,
                      eastasia=TNR, size=12, first_line_indent=0, line_spacing=line)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)
    add_paragraph(doc, "Supervisor", align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=TNR, size=12, first_line_indent=0, line_spacing=line)
    add_paragraph(doc, args.supervisor_en, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=TNR, size=12, first_line_indent=0, line_spacing=line)
    add_paragraph(doc, "", first_line_indent=0, line_spacing=line)
    add_paragraph(doc, args.date_en, align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=TNR, size=12, first_line_indent=0, line_spacing=line)


def add_declaration(doc, num_id: int) -> None:
    """原创性声明 + 版权使用授权书（同一页）。"""
    add_paragraph(doc, "湖 南 大 学", align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=HEI, size=18, first_line_indent=0, line_spacing=24)

    heading = doc.add_paragraph(style="Heading 1")
    add_text(heading, "学位论文原创性声明")
    apply_numbering(heading, 0, 0)  # 前置部分不参与章编号

    for text in (
        "本人郑重声明：所呈交的论文是本人在导师的指导下独立进行研究所取得的研究成果。"
        "除了文中特别加以标注引用的内容外，本论文不包含任何其他个人或集体已经发表或撰写过的成果作品。"
        "对本文的研究做出重要贡献的个人和集体，均已在文中以明确方式标明。本人完全意识到本声明的法律后果由本人承担。",
    ):
        add_paragraph(doc, text, first_line_indent=24, line_spacing=24)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(6.0))
    p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0))
    add_text(p, "作者签名：\t日期：\t\t年\t\t月\t\t日", eastasia=SONG, size=12)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    add_paragraph(doc, "学位论文版权使用授权书", align=WD_ALIGN_PARAGRAPH.CENTER,
                  eastasia=HEI, size=18, first_line_indent=0, line_spacing=24)

    add_paragraph(
        doc,
        "本学位论文作者完全了解学校有关保留、使用学位论文的规定，同意学校保留并向国家有关部门"
        "或机构送交论文的复印件和电子版，允许论文被查阅和借阅。本人授权湖南大学可以将本学位论文"
        "的全部或部分内容编入有关数据库进行检索，可以采用影印、缩印或扫描等复制手段保存和汇编本学位论文。",
        first_line_indent=24, line_spacing=24,
    )
    add_paragraph(doc, "本学位论文属于", first_line_indent=24, line_spacing=24)
    for text in (
        "1、保密□，在______年解密后适用本授权书。",
        "2、不保密□。",
        "（请在以上相应方框内打“√”）",
    ):
        p = add_paragraph(doc, text, first_line_indent=24, line_spacing=24)
        p.paragraph_format.left_indent = Cm(2.0)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(6.0))
    p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0))
    add_text(p, "作者签名：\t日期：\t\t年\t\t月\t\t日", eastasia=SONG, size=12)
    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=24)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(6.0))
    p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0))
    add_text(p, "导师签名：\t日期：\t\t年\t\t月\t\t日", eastasia=SONG, size=12)


def add_abstract(doc) -> None:
    heading = doc.add_paragraph(style="Heading 1")
    add_text(heading, "摘   要")
    apply_numbering(heading, 0, 0)

    add_paragraph(
        doc,
        "（在此撰写中文摘要：说明研究背景与意义、采用的方法、主要结果与结论。"
        "博士约 1200 字，硕士约 600 字；摘要不写论文题目，不使用公式和图表，不标注文献序号。）",
        first_line_indent=24, line_spacing=20,
    )
    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=20)
    add_text(p, "关键词：", eastasia=HEI, size=12)
    add_text(p, "关键词1；关键词2；关键词3；关键词4", eastasia=HEI, size=12)


def add_en_abstract(doc) -> None:
    heading = doc.add_paragraph(style="Heading 1")
    add_text(heading, "Abstract", eastasia=TNR, bold=True)
    apply_numbering(heading, 0, 0)

    add_paragraph(
        doc,
        "(Write the English abstract here. It must match the Chinese abstract "
        "word for word in content, with correct grammar and terminology.)",
        eastasia=TNR, first_line_indent=24, line_spacing=20,
    )
    p = add_paragraph(doc, "", first_line_indent=0, line_spacing=20)
    add_text(p, "Key Words: ", eastasia=TNR, size=12, bold=True)
    add_text(p, "keyword1; keyword2; keyword3; keyword4", eastasia=TNR, size=12)


def add_toc(doc) -> None:
    heading = doc.add_paragraph(style="Heading 1")
    add_text(heading, "目   录")
    apply_numbering(heading, 0, 0)

    p = doc.add_paragraph(style="TOC 1")
    # \h 让目录条目成为可点击超链接，\z 取消隐藏页码，\u 使用大纲级别
    add_field(p, ' TOC \\o "1-4" \\h \\z \\u ', "按 Ctrl+A 后 F9 更新目录",
              dirty=True, size=12)

    # 插图索引 / 附表索引（“必要时”）：用 TOC \c 收集带 SEQ 图/表 的题注
    idx = doc.add_paragraph(style="Heading 1")
    add_text(idx, "插图索引")
    apply_numbering(idx, 0, 0)
    p = doc.add_paragraph(style="TOC 1")
    add_field(p, ' TOC \\h \\z \\c "图" ', "按 F9 更新后生成插图索引", dirty=True, size=12)

    idx = doc.add_paragraph(style="Heading 1")
    add_text(idx, "附表索引")
    apply_numbering(idx, 0, 0)
    p = doc.add_paragraph(style="TOC 1")
    add_field(p, ' TOC \\h \\z \\c "表" ', "按 F9 更新后生成附表索引", dirty=True, size=12)


def add_caption(doc, label: str, title: str, bookmark: str, bid: int) -> None:
    """题注 = 标签 + STYLEREF（章号）+ "." + SEQ（章内序号）+ 题名。

    与答辩版论文一致：编号随章自动变化，插/删图后 F9 即重排；
    题注段落带书签，正文用 REF 域引用即可点击跳转。
    """
    p = doc.add_paragraph(style="题注")
    add_bookmark_start(p, bookmark, bid)
    add_text(p, label, eastasia=HEI, size=12)
    add_field(p, " STYLEREF 1 \\s", "1", eastasia=HEI, size=12)
    add_text(p, ".", eastasia=HEI, size=12)
    add_field(p, f" SEQ {label} \\* ARABIC \\s 1", "1", eastasia=HEI, size=12)
    add_bookmark_end(p, bid)  # 书签只包编号，REF 才不会把整条题名带进正文
    add_text(p, " " + title, eastasia=HEI, size=12)


def add_ref(paragraph, bookmark: str, cached: str) -> None:
    """正文中的交叉引用：REF 域 + \\h（可 Ctrl+单击跳转）。"""
    add_field(paragraph, f" REF {bookmark} \\h", cached, size=12)


def make_three_line(table) -> None:
    """三线表：上下粗线、表头下细线、左右无边框。"""
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")
        elif edge == "insideH":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "6")
        else:
            el.set(qn("w:val"), "nil")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    # tblBorders 必须排在 shd/tblLayout/tblCellMar/tblLook 之前，否则 Word 可能判定损坏
    tbl_pr.insert_element_before(
        borders,
        "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange",
    )


def add_body(doc, num_id: int) -> None:
    """正文示例：第 1 章 + 图表公式示例 + 结论 + 参考文献 + 附录 + 致谢。"""
    chapter = doc.add_paragraph(style="Heading 1")
    add_text(chapter, "绪论")
    apply_numbering(chapter, num_id, 0)

    section = doc.add_paragraph(style="Heading 2")
    add_text(section, "研究背景与意义")
    apply_numbering(section, num_id, 1)

    add_paragraph(
        doc,
        "（在此撰写正文。每章先提出本章要解决的问题，再展开论述；章标题不超过 15 字且不用标点。"
        "正文用宋体小 4 号、首行缩进 2 字符、固定行距 20 磅。引用文献在句末右上角标注[1]。）",
        first_line_indent=24, line_spacing=20,
    )

    # 图：占位段落 + 带域题注
    add_paragraph(doc, "（在此插入图片：图宽不超过版心 16 cm，图与图题不得拆页。）",
                  align=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=0, line_spacing=20)
    add_caption(doc, "图", "示例图题（替换为你的图名）", "_RefFigExample", 9001)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=20)
    p = add_paragraph(doc, "如", first_line_indent=24, line_spacing=20)
    add_ref(p, "_RefFigExample", "图1.1")
    add_text(p, "所示，正文用交叉引用域引用图号，按住 Ctrl 单击可跳转到图题。", size=12)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=20)

    # 表：表题在表上，三线表
    add_caption(doc, "表", "示例表题（替换为你的表名）", "_RefTabExample", 9002)
    table = doc.add_table(rows=3, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    make_three_line(table)
    headers = ["项目", "参数 A", "参数 B"]
    for col, text in enumerate(headers):
        cell = table.rows[0].cells[col]
        cell.paragraphs[0].paragraph_format.first_line_indent = Pt(0)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(cell.paragraphs[0], text, eastasia=HEI, size=12)
    for row_idx in (1, 2):
        for col in range(3):
            cell = table.rows[row_idx].cells[col]
            cell.paragraphs[0].paragraph_format.first_line_indent = Pt(0)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_text(cell.paragraphs[0], "—", eastasia=SONG, size=12)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=20)
    p = add_paragraph(doc, "如", first_line_indent=24, line_spacing=20)
    add_ref(p, "_RefTabExample", "表1.1")
    add_text(p, "所示，三线表不加左右边线；表中缺数据填“—”，不写“同上”。", size=12)

    add_paragraph(doc, "", first_line_indent=0, line_spacing=20)

    # 公式：居中 + 右顶格编号
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = Pt(20)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(TEXT_WIDTH_CM / 2), WD_TAB_ALIGNMENT.CENTER)
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)
    add_text(p, "\t（在此插入公式）\t(1.1)", size=12)

    section = doc.add_paragraph(style="Heading 2")
    add_text(section, "本文主要研究内容")
    apply_numbering(section, num_id, 1)
    add_paragraph(
        doc,
        "（在此说明各章安排与相互关系。节标题用黑体小 3 号，条标题用黑体 4 号，"
        "层次以少为宜，全文统一。）",
        first_line_indent=24, line_spacing=20,
    )

    # 结论：单独一章，不加章号
    conclusion = doc.add_paragraph(style="Heading 1")
    add_text(conclusion, "结   论")
    apply_numbering(conclusion, 0, 0)
    conclusion.paragraph_format.page_break_before = True
    add_paragraph(
        doc,
        "（在此写结论：明确指出创造性成果或创新性理论，评价应用前景，并给出后续研究展望；"
        "一般约 2000 字。结论不加“第 X 章”。）",
        first_line_indent=24, line_spacing=20,
    )

    refs = doc.add_paragraph(style="Heading 1")
    add_text(refs, "参考文献")
    apply_numbering(refs, 0, 0)
    refs.paragraph_format.page_break_before = True
    for entry in (
        "[1] 作者. 文献题名[J]. 期刊名, 2020, 12(3): 45-52.",
        "[2] 作者. 书名[M]. 北京: 出版社, 2019: 10-25.",
    ):
        p = add_paragraph(doc, entry, first_line_indent=0, line_spacing=20)
        p.paragraph_format.left_indent = Cm(0.74)
        p.paragraph_format.first_line_indent = Cm(-0.74)

    appendix = doc.add_paragraph(style="Heading 1")
    add_text(appendix, "附录A 攻读学位期间发表的学术论文")
    apply_numbering(appendix, 0, 0)
    appendix.paragraph_format.page_break_before = True
    add_paragraph(
        doc,
        "（按参考文献格式列出攻读学位期间发表/录用的论文；目录中体现为“附录A”。）",
        first_line_indent=24, line_spacing=20,
    )

    thanks = doc.add_paragraph(style="Heading 1")
    add_text(thanks, "致   谢")
    apply_numbering(thanks, 0, 0)
    thanks.paragraph_format.page_break_before = True
    add_paragraph(
        doc,
        "（在此感谢导师、课题组、资助项目与家人；内容简洁、实事求是。）",
        first_line_indent=24, line_spacing=20,
    )


# ---------------------------------------------------------------------------
# 节、页眉页脚
# ---------------------------------------------------------------------------


def configure_page(section) -> None:
    section.page_width = Cm(PAGE_WIDTH_CM)
    section.page_height = Cm(PAGE_HEIGHT_CM)
    section.top_margin = Cm(MARGIN_TB_CM)
    section.bottom_margin = Cm(MARGIN_TB_CM)
    section.left_margin = Cm(MARGIN_LR_CM)
    section.right_margin = Cm(MARGIN_LR_CM)
    section.header_distance = Cm(HF_DISTANCE_CM)
    section.footer_distance = Cm(HF_DISTANCE_CM)


def set_header(section, text: str, *, odd: bool = True) -> None:
    header = section.header if odd else section.even_page_header
    header.is_linked_to_previous = False
    for p in header.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p = header.paragraphs[0]
    for run in list(p.runs):
        run._r.getparent().remove(run._r)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    add_text(p, text, eastasia=SONG, size=9)
    set_paragraph_border_bottom(p)


def set_footer_page_number(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in footer.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p = footer.paragraphs[0]
    for run in list(p.runs):
        run._r.getparent().remove(run._r)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    add_field(p, "PAGE   \\* MERGEFORMAT", "1", eastasia=SONG, size=9)


def build(args) -> Document:
    doc = Document()
    doc.settings.odd_and_even_pages_header_footer = True
    setup_styles(doc)
    num_id = build_heading_numbering(doc)

    first = doc.sections[0]
    configure_page(first)
    first.header.is_linked_to_previous = False
    first.even_page_header.is_linked_to_previous = False

    add_cover(doc, args)
    doc.add_page_break()  # 中文封面后分页，进入英文内封
    add_english_cover(doc, args)

    # 前置部分：声明页起用罗马数字 I
    decl = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_page(decl)
    set_page_numbering(decl, "upperRoman", 1)
    set_header(decl, args.degree_label, odd=True)
    set_header(decl, args.title, odd=False)
    set_footer_page_number(decl)
    add_declaration(doc, num_id)

    abstract = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_page(abstract)
    inherits_continuous_numbering(abstract)
    set_header(abstract, args.degree_label, odd=True)
    set_header(abstract, args.title, odd=False)
    set_footer_page_number(abstract)
    add_abstract(doc)

    en_abstract = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_page(en_abstract)
    inherits_continuous_numbering(en_abstract)
    set_header(en_abstract, args.degree_label, odd=True)
    set_header(en_abstract, args.title, odd=False)
    set_footer_page_number(en_abstract)
    add_en_abstract(doc)

    toc = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_page(toc)
    inherits_continuous_numbering(toc)
    set_header(toc, args.degree_label, odd=True)
    set_header(toc, args.title, odd=False)
    set_footer_page_number(toc)
    add_toc(doc)

    # 正文：阿拉伯数字从 1 开始
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_page(body)
    set_page_numbering(body, "decimal", 1)
    set_header(body, args.degree_label, odd=True)
    set_header(body, args.title, odd=False)
    set_footer_page_number(body)
    add_body(doc, num_id)

    return doc


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="生成湖南大学学位论文骨架（封面/内封/声明/摘要/目录/索引/正文占位）。",
    )
    parser.add_argument("--out", default="thesis_skeleton.docx", help="输出 .docx 路径")
    parser.add_argument("--degree", default="doctor",
                        choices=["doctor", "master", "professional"],
                        help="学位类型，决定“博士学位论文/硕士学位论文”等字样")
    parser.add_argument("--title", default="论文题目", help="中文题目，≤25 字，不设副标题")
    parser.add_argument("--title-en", dest="title_en", default="Title of the Dissertation")
    parser.add_argument("--author", default="张三")
    parser.add_argument("--author-en", dest="author_en", default="ZHANG San")
    parser.add_argument("--supervisor", default="×× 教授")
    parser.add_argument("--supervisor-en", dest="supervisor_en", default="Professor XX")
    parser.add_argument("--college", default="××学院")
    parser.add_argument("--major", default="××专业")
    parser.add_argument("--major-en", dest="major_en", default="XX Engineering")
    parser.add_argument("--field", default="××方向")
    parser.add_argument("--school-code", dest="school_code", default="10532")
    parser.add_argument("--student-id", dest="student_id", default="B00000000")
    parser.add_argument("--category", default="TG1", help="分类号")
    parser.add_argument("--secret", default="公开", help="密级")
    parser.add_argument("--date", default="2026年5月1日")
    parser.add_argument("--date-en", dest="date_en", default="May, 2026")
    parser.add_argument("--defense-date", dest="defense_date", default="2026年5月20日")
    parser.add_argument("--chair", default="×× 教授", help="答辩委员会主席")
    parser.add_argument("--bachelor", default="B.E. (Hunan University) 2020",
                        help="英文内封中的前置学历行")
    parser.add_argument("--degree-en", dest="degree_en", default=None,
                        help="英文内封的学位名称，默认按 --degree 推断")
    parser.add_argument("--degree-label", dest="degree_label", default=None,
                        help="封面/页眉的学位论文类别（默认“博士学位论文”等；"
                             "专业学位按类别改成“工程硕士学位论文”“工商管理硕士学位论文”等）")
    args = parser.parse_args(argv)
    if args.degree_label is None:
        args.degree_label = {"doctor": "博士学位论文",
                             "master": "硕士学位论文",
                             "professional": "专业硕士学位论文"}[args.degree]
    if args.degree_en is None:
        args.degree_en = DEGREE_EN[args.degree]
    return args


def main(argv=None) -> int:
    args = parse_args(argv)
    doc = build(args)
    out = Path(args.out)
    if out.parent and not out.parent.exists():
        out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"已生成：{out}")
    print("下一步：")
    print("  1) 打开文件，替换封面/内封的占位符（姓名、学号、导师、题目、日期）。")
    print("  2) Ctrl+A → F9 更新域：目录、插图索引、附表索引、图表编号。")
    print(f"  3) python scripts/thesis_lint.py {out} --degree {args.degree}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
