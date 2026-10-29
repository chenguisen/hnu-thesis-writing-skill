#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把希腊字母与单位/数学符号从“中文字体”改回 Times New Roman。

问题现象（Word/WPS 里都能复现）：

    w:rFonts 带 `w:hint="eastAsia"` 时，Word 把“东亚宽度歧义（East_Asian_Width=A）”
    的字符交给中文字体（宋体/黑体）去排。希腊字母（α β γ δ …）、′ ″、℃，°、±、×、
    ÷、·、µ、→、≈、≤、下标 ₂ 都属于歧义字符，于是：

      * 正文里（没有 hint）`δ` 是 Times New Roman，窄体，和西文一致；
      * 题注/目录/表格里（有 hint）同一个 `δ` 变成黑体/宋体的全角字形，又大又宽，
        前面还多出一格空当，γ″ 会撑成大空隙。

    更麻烦的是同一个符号在正文、题注、插图索引里出现三种字形。

本脚本的做法：只在**含符号的 run** 上动手，把 run 按“符号 / 非符号”切段，符号段去掉
`w:hint="eastAsia"`，让它们回到 ascii/hAnsi（Times New Roman）；中文、全角标点、
“”…等仍留在原字体上，字号、粗细、颜色、缩进一概不动。

用法：
    python scripts/fix_symbol_fonts.py 论文.docx                  # 只检查，不改文件
    python scripts/fix_symbol_fonts.py 论文.docx --out 修好的.docx  # 另存
    python scripts/fix_symbol_fonts.py 论文.docx --in-place        # 原地修改
"""

from __future__ import annotations

import argparse
import sys
from copy import deepcopy
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:  # python-docx >= 1.0
    from docx.parts.hdrftr import FooterPart, HeaderPart
except ImportError:  # pragma: no cover - 旧版 python-docx
    HeaderPart = FooterPart = ()  # type: ignore[assignment]

DEFAULT_LATIN = "Times New Roman"
XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"
HINT = qn("w:hint")
EASTASIA = "eastAsia"

# 需要 Times New Roman 的“西文符号”Unicode 区段（见文件头说明）
SYMBOL_RANGES = (
    (0x00B0, 0x00B0),  # ° 度
    (0x00B1, 0x00B1),  # ±
    (0x00B5, 0x00B5),  # µ 微
    (0x00B7, 0x00B7),  # · 间隔号
    (0x00D7, 0x00D7),  # ×
    (0x00F7, 0x00F7),  # ÷
    (0x0370, 0x03FF),  # 希腊字母
    (0x1F00, 0x1FFF),  # 希腊字母扩展
    (0x2032, 0x2037),  # ′ ″ ‴ ‵ ‶ ‷
    (0x2070, 0x209F),  # 上标/下标数字 ⁰¹²³ ₀₁₂₃
    (0x20A0, 0x20BF),  # 货币符号
    (0x2103, 0x2103),  # ℃
    (0x2109, 0x2109),  # ℉
    (0x210E, 0x210E),  # ℎ 普朗克常数
    (0x2113, 0x2113),  # ℓ
    (0x212B, 0x212B),  # Å
    (0x2190, 0x21FF),  # 箭头 → ⇒
    (0x2200, 0x22FF),  # 数学运算符 ≈ ≤ ≥ ∑ ∫ − ∞
    (0x2300, 0x23FF),  # 技术符号 ⌈⌉⌊⌋
    (0x25A0, 0x25FF),  # 几何图形 ■ ○
)


def is_symbol(ch: str) -> bool:
    """该字符是否应使用西文字体（Times New Roman）。"""
    code = ord(ch)
    for lo, hi in SYMBOL_RANGES:
        if lo <= code <= hi:
            return True
    return False


def _run_text(r) -> str:
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def _only_text(r) -> bool:
    """run 里除了 rPr 就只有 w:t（可以安全切段）。"""
    for child in r:
        if child.tag == qn("w:rPr"):
            continue
        if child.tag != qn("w:t"):
            return False
    return True


def _rfonts(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return None
    return rpr.find(qn("w:rFonts"))


def _has_east_hint(r) -> bool:
    rfonts = _rfonts(r)
    return rfonts is not None and (rfonts.get(HINT) or "").lower() == EASTASIA.lower()


def _drop_hint(r) -> None:
    rfonts = _rfonts(r)
    if rfonts is not None and rfonts.get(HINT) is not None:
        del rfonts.attrib[HINT]


def _segments(text: str):
    """按“符号/非符号”切成 [(片段, 是否符号)]，相邻同类合并。"""
    out: list[list] = []
    for ch in text:
        sym = is_symbol(ch)
        if out and out[-1][1] == sym:
            out[-1][0] += ch
        else:
            out.append([ch, sym])
    return [(t, s) for t, s in out]


def _make_run(template, text: str, symbol: bool):
    new = deepcopy(template)
    for child in list(new):
        if child.tag == qn("w:t"):
            new.remove(child)
    t = OxmlElement("w:t")
    t.text = text
    if text != text.strip():
        t.set(XMLSPACE, "preserve")
    new.append(t)
    if symbol:
        _drop_hint(new)
    return new


def _iter_run_elements(document):
    """正文（含表格、目录域结果）与页眉页脚里的所有 w:r。"""
    for r in document.element.body.iter(qn("w:r")):
        yield r
    for part in getattr(document.part.package, "parts", []):
        if HeaderPart and isinstance(part, (HeaderPart, FooterPart)):
            element = getattr(part, "element", None)
            if element is None:
                continue
            for r in element.iter(qn("w:r")):
                yield r


def scan(document):
    """找出“带 eastAsia hint 且含符号”的 run（只读）。"""
    runs = []
    samples: list[str] = []
    for r in _iter_run_elements(document):
        if not _has_east_hint(r):
            continue
        text = _run_text(r)
        if not text or not any(is_symbol(c) for c in text):
            continue
        runs.append(r)
        if len(samples) < 10:
            samples.append(text.strip()[:24])
    return runs, samples


def fix_runs(runs) -> tuple[int, int]:
    """切段并去掉符号段的 hint。返回 (改动的 run 数, 生成的符号段数)。"""
    changed = 0
    sym_segments = 0
    for r in runs:
        if r.getparent() is None:  # pragma: no cover - 防御
            continue
        text = _run_text(r)
        if not _only_text(r):
            # 带制表符/换行/域字符的 run 不切段，整体去掉 hint（只影响歧义字符）
            if not all(is_symbol(c) for c in text):
                continue
            _drop_hint(r)
            changed += 1
            sym_segments += 1
            continue
        segs = _segments(text)
        if all(s for _, s in segs):
            _drop_hint(r)
            changed += 1
            sym_segments += 1
            continue
        parent = r.getparent()
        idx = parent.index(r)
        new_runs = []
        for seg_text, is_sym in segs:
            new_runs.append(_make_run(r, seg_text, is_sym))
            if is_sym:
                sym_segments += 1
        parent.remove(r)
        for offset, nr in enumerate(new_runs):
            parent.insert(idx + offset, nr)
        changed += 1
    return changed, sym_segments


def _style_of(r) -> str:
    """run 所在段落的样式 id（用于统计，取不到就返回 '?'）。"""
    node = r.getparent()
    while node is not None and node.tag != qn("w:p"):
        node = node.getparent()
    if node is None:
        return "?"
    ppr = node.find(qn("w:pPr"))
    if ppr is None:
        return "Normal"
    ps = ppr.find(qn("w:pStyle"))
    return ps.get(qn("w:val")) if ps is not None else "Normal"


def _p(msg: str) -> None:
    """控制台可能不支持某些符号（GBK），打印时降级而不是抛错。"""
    try:
        print(msg)
    except UnicodeEncodeError:  # pragma: no cover - Windows GBK 控制台
        data = (msg + "\n").encode(sys.stdout.encoding or "utf-8", "replace")
        sys.stdout.buffer.write(data)
        sys.stdout.flush()


def run(path: Path, out: Path | None, in_place: bool, quiet: bool, stats: bool = False) -> int:
    document = docx.Document(str(path))
    runs, samples = scan(document)
    breakdown: dict[str, int] = {}
    for r in runs:
        key = _style_of(r)
        breakdown[key] = breakdown.get(key, 0) + 1
    changed, segments = fix_runs(runs)

    if not quiet:
        _p(f"文件：{path}")
        _p(f"含希腊字母/单位符号且被中文字体渲染的 run：{len(runs)} 处；"
           f"已处理 {changed} 处，切出符号段 {segments} 个")
        if stats and breakdown:
            _p("按样式分布：" + "、".join(f"{k}={v}" for k, v in
                                       sorted(breakdown.items(), key=lambda kv: -kv[1])))
        if samples:
            _p("样例：" + " | ".join(samples))
        if not runs:
            _p("未发现符号被中文字体渲染的情况。")

    target = out or (path if in_place else None)
    if target is None:
        if not quiet:
            _p("\n（仅检查，未写文件。加 --out 另存，或 --in-place 原地修改。）")
        return 0
    document.save(str(target))
    if not quiet:
        _p(f"已保存：{target}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="把希腊字母、单位符号的字体改回 Times New Roman（不动中文字体）"
    )
    parser.add_argument("docx", type=Path, help="待处理的 .docx")
    parser.add_argument("--out", type=Path, default=None, help="另存路径")
    parser.add_argument("--in-place", action="store_true", help="原地覆盖（建议先备份）")
    parser.add_argument("--quiet", action="store_true", help="不打印统计")
    parser.add_argument("--stats", action="store_true", help="按段落样式打印分布")
    args = parser.parse_args(argv)
    if args.out and args.in_place:
        parser.error("--out 与 --in-place 只能二选一")
    if not args.docx.exists():
        parser.error(f"文件不存在：{args.docx}")
    return run(args.docx, args.out, args.in_place, args.quiet, args.stats)


if __name__ == "__main__":
    sys.exit(main())
