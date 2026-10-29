#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把西文/数字的字体统一成 Times New Roman，中文字体保持不变。

规范正文只写“小 4 号宋体”，没有单独说西文用什么；但校模板与 2025 年答辩通过的
湖南大学博士论文里，`Normal` 样式是 `eastAsia=宋体 + ascii=Times New Roman`，
`Heading 1–4` 是 `eastAsia=黑体 + ascii=Times New Roman`。
反过来把 `ascii` 设成宋体/黑体时，英文、数字、单位就会用中文字体排（并影响字宽与
换行），与模板不一致——这是论文里最常见、又最容易被忽略的字体问题。

用法：
    python scripts/fix_latin_fonts.py 论文.docx                    # 只检查，不改文件
    python scripts/fix_latin_fonts.py 论文.docx --out 修好的.docx    # 另存
    python scripts/fix_latin_fonts.py 论文.docx --in-place          # 原地修改
    python scripts/fix_latin_fonts.py 论文.docx --latin "Times New Roman"

只改 `w:rFonts` 的 ascii/hAnsi/cs，不动 `w:eastAsia`，也不改字号、颜色、粗斜体。
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import docx
from docx.oxml.ns import qn

try:  # python-docx >= 1.0
    from docx.parts.hdrftr import FooterPart, HeaderPart
except ImportError:  # pragma: no cover - 旧版 python-docx
    HeaderPart = FooterPart = ()  # type: ignore[assignment]

# 只要 ascii 是这些字体，就说明西文被排成了中文字体
CJK_FONTS = (
    "宋体", "黑体", "仿宋", "楷体", "微软雅黑", "华文宋体", "华文黑体",
    "SimSun", "SimHei", "NSimSun", "FangSong", "KaiTi", "Microsoft YaHei",
)
DEFAULT_LATIN = "Times New Roman"
RE_LATIN = re.compile(r"[A-Za-z0-9]")
XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _needs_fix(rfonts, latin: str) -> bool:
    """这段文字的西文是否被排成了中文字体。"""
    ascii_font = rfonts.get(qn("w:ascii"))
    eastasia = rfonts.get(qn("w:eastAsia"))
    if ascii_font == latin:
        return False
    return ascii_font in CJK_FONTS or (ascii_font is None and eastasia in CJK_FONTS)


def _fix_rfonts(rfonts, latin: str) -> bool:
    """就地把 rFonts 的西文字体改成 latin；返回是否改动。"""
    if not _needs_fix(rfonts, latin):
        return False
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    if rfonts.get(qn("w:cs")) in CJK_FONTS or rfonts.get(qn("w:cs")) is None:
        rfonts.set(qn("w:cs"), latin)
    return True


def _iter_run_elements(document):
    """正文（含表格）与页眉页脚里的所有 w:r。"""
    for r in document.element.body.iter(qn("w:r")):
        yield r
    parts = getattr(document.part.package, "parts", [])
    for part in parts:
        if HeaderPart and isinstance(part, (HeaderPart, FooterPart)):
            element = getattr(part, "element", None)
            if element is not None:
                for r in element.iter(qn("w:r")):
                    yield r


def scan(document, latin: str = DEFAULT_LATIN):
    """统计需要修改的 run（只读，不改文档）。"""
    runs = []
    samples: list[str] = []
    for r in _iter_run_elements(document):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        if not _needs_fix(rfonts, latin):
            continue
        text = "".join(t.text or "" for t in r.iter(qn("w:t")))
        if not RE_LATIN.search(text):
            continue
        runs.append(r)
        if len(samples) < 8:
            samples.append(text.strip()[:24])
    return runs, samples


def used_style_names(document) -> set[str]:
    names = set()
    for p in document.paragraphs:
        style = getattr(p, "style", None)
        if style is not None and style.name:
            names.add(style.name)
    return names


def fix_styles(document, latin: str = DEFAULT_LATIN) -> list[str]:
    changed = []
    used = used_style_names(document)
    for style in document.styles:
        if style.name not in used:
            continue
        element = getattr(style, "element", None)
        if element is None:
            continue
        rpr = element.find(qn("w:rPr"))
        if rpr is None:
            continue
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        if _fix_rfonts(rfonts, latin):
            changed.append(style.name)
    return changed


def fix_runs(runs, latin: str = DEFAULT_LATIN) -> int:
    fixed = 0
    for r in runs:
        rpr = r.find(qn("w:rPr"))
        rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if rfonts is None:
            continue
        if _fix_rfonts(rfonts, latin):
            fixed += 1
    return fixed


def run(path: Path, out: Path | None, in_place: bool, latin: str, quiet: bool) -> int:
    document = docx.Document(str(path))
    runs, samples = scan(document, latin)
    styles = fix_styles(document, latin)
    fixed = fix_runs(runs, latin)

    if not quiet:
        print(f"文件：{path}")
        print(f"西文字体改为 {latin}：run {fixed} 处，样式 {len(styles)} 个"
              + (f"（{'、'.join(styles)}）" if styles else ""))
        if samples:
            print("样例：" + " | ".join(samples))
        if fixed == 0 and not styles:
            print("未发现西文用中文字体的情况。")

    target = out or (path if in_place else None)
    if target is None:
        if not quiet:
            print("\n（仅检查，未写文件。加 --out 另存，或 --in-place 原地修改。）")
        return 0
    document.save(str(target))
    if not quiet:
        print(f"已保存：{target}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="把论文里的西文/数字统一为 Times New Roman（不改中文字体）"
    )
    parser.add_argument("docx", type=Path, help="待处理的 .docx")
    parser.add_argument("--out", type=Path, default=None, help="另存路径")
    parser.add_argument("--in-place", action="store_true", help="原地覆盖（建议先备份）")
    parser.add_argument("--latin", default=DEFAULT_LATIN,
                        help=f"西文字体（默认 {DEFAULT_LATIN}）")
    parser.add_argument("--quiet", action="store_true", help="不打印统计")
    args = parser.parse_args(argv)
    if args.out and args.in_place:
        parser.error("--out 与 --in-place 只能二选一")
    if not args.docx.exists():
        parser.error(f"文件不存在：{args.docx}")
    return run(args.docx, args.out, args.in_place, args.latin, args.quiet)


if __name__ == "__main__":
    sys.exit(main())
