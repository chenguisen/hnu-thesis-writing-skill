#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把正文引用标注与文后参考文献按“首次引用顺序”重排。

适用条件
--------
* 正文引用写成**上标**的方括号数字：`[3]`、`[1,4]`、`[2-5]`、`[1，4]`（全角逗号也认）；
* 文后参考文献每条以 `[n]` 开头（n 从 1 连续编号），位于标题为“参考文献”的段落之后。

它做三件事
----------
1. 按正文出现顺序计算“旧号→新号”映射（只统计上标里的标记，正文中的晶向
   `[001]`、`[10-10]` 等不会被误判）；
2. 改写全文引用标记：组内升序、连续 ≥2 个压成区间（`[8,9,10]`→`[8-10]`）；
3. 按新号重排文后条目并重写条目前的 `[n]`。

用法
----
    python scripts/renumber_refs.py thesis.docx --report map.txt
    python scripts/renumber_refs.py thesis.docx --out thesis_v2.docx
    python scripts/renumber_refs.py thesis.docx --dry-run      # 只看映射，不改文件

**务必在更新域（目录/题注/REF）之前运行**，否则改完文字还要再更新一次域。
含无法映射编号（文献表里没有的号）的标记会整条跳过并在报告中列出，绝不静默丢弃。
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import OrderedDict

from docx import Document
from docx.oxml.ns import qn

TOKEN = re.compile(r"\[\s*\d+(?:\s*[,\-\u2013~\uff0c\u3001]\s*\d+)*\s*\]")
SPLIT = re.compile(r"[,\uff0c\u3001]")
RANGE = re.compile(r"^(\d+)\s*[-\u2013~]\s*(\d+)$")
HEAD_RE = re.compile(r"^\s*\[(\d+)\]")


def _run_text(t_els):
    return "".join(t.text or "" for t in t_els)


def para_runs(p):
    """返回段落内的 run 列表：{el, tes, text, sup}。"""
    out = []
    for r in p.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        sup = False
        if rpr is not None:
            va = rpr.find(qn("w:vertAlign"))
            if va is not None and va.get(qn("w:val")) == "superscript":
                sup = True
        tes = r.findall(qn("w:t"))
        out.append({"el": r, "tes": tes, "text": _run_text(tes), "sup": sup})
    return out


def set_run_text(run, text):
    tes = run["tes"]
    if not tes:
        return False
    tes[0].text = text
    tes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in tes[1:]:
        t.text = ""
    run["text"] = text
    return True


def expand(inner):
    """'1,4' → [1,4]；'2-5' → [2,3,4,5]；解析失败返回 None。"""
    nums = []
    for part in SPLIT.split(inner):
        part = part.strip()
        if not part:
            continue
        if part.isdigit():
            nums.append(int(part))
        else:
            m = RANGE.match(part)
            if not m:
                return None
            a, b = int(m.group(1)), int(m.group(2))
            if b < a:
                return None
            nums.extend(range(a, b + 1))
    return nums


def fmt(nums):
    """[1,2,3,5] → '[1-3,5]'。"""
    nums = sorted(set(nums))
    parts, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        parts.append("%d-%d" % (nums[i], nums[j]) if j - i >= 1 else str(nums[i]))
        i = j + 1
    return "[" + ",".join(parts) + "]"


def par_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def renumber(path, out=None, report=None, dry_run=False, reorder_list=True):
    doc = Document(path)
    body = doc.element.body
    paras = [c for c in body.iterchildren() if c.tag == qn("w:p")]

    # ---------- 文后参考文献表 ----------
    ref_head = None
    for p in paras:
        if par_text(p).strip() == "参考文献":
            ref_head = p
    if ref_head is None:
        raise SystemExit("未找到标题为“参考文献”的段落；如标题带编号，请先临时改为纯“参考文献”。")

    entries = []
    for p in paras:
        if p is ref_head:
            continue
        m = HEAD_RE.match(par_text(p))
        if m:
            entries.append((int(m.group(1)), p))
    if not entries:
        raise SystemExit("参考文献条目未以“[n]”开头，无法重排（请先统一为手工编号文本）。")
    olds = sorted(n for n, _ in entries)
    if olds != list(range(1, len(olds) + 1)):
        raise SystemExit("参考文献编号不连续：%s…请先修正为 1..N。" % olds[:10])
    nref = len(olds)

    # ---------- 首次引用顺序 ----------
    order = OrderedDict()
    tokens = []
    skipped = []
    for p in paras:
        runs = para_runs(p)
        full = "".join(r["text"] for r in runs)
        if "[" not in full:
            continue
        pos, spans = 0, []
        for r in runs:
            spans.append((pos, pos + len(r["text"])))
            pos += len(r["text"])
        for m in TOKEN.finditer(full):
            covered = [i for i, (s, e) in enumerate(spans) if s < m.end() and e > m.start()]
            if not covered or not all(runs[i]["sup"] for i in covered):
                continue                      # 正文中的晶向/方括号说明，不是引用
            nums = expand(m.group(0)[1:-1])
            if nums is None or any(n < 1 or n > nref for n in nums):
                skipped.append((par_text(p)[:40], m.group(0)))
                continue
            tokens.append((p, covered, m.group(0), nums))
            for n in nums:
                order.setdefault(n, len(order) + 1)

    uncited = [n for n in range(1, nref + 1) if n not in order]
    if uncited:
        raise SystemExit("以下文献在正文中从未被引用，请先补引或删除：%s" % uncited)
    mapping = {old: order[old] for old in order}          # 旧号 -> 新号
    inverse = {v: k for k, v in mapping.items()}
    if sorted(mapping) != list(range(1, nref + 1)):
        raise SystemExit("映射不完整：%s" % sorted(set(range(1, nref + 1)) - set(mapping)))

    lines = ["旧号 -> 新号（按新号排序）"]
    for new in range(1, nref + 1):
        lines.append("%4d -> %4d" % (inverse[new], new))
    text_report = "\n".join(lines)

    if dry_run:
        print(text_report)
        if skipped:
            print("\n跳过（文献表内无此编号）：")
            for ctx, tok in skipped:
                print("  ", tok, "…", ctx)
        print("\n[dry-run] 未修改文件。")
        return mapping

    # ---------- 改写正文标注 ----------
    changed = 0
    for p, covered, tok, nums in tokens:
        newtok = fmt([mapping[n] for n in nums])
        if newtok == tok:
            continue
        runs = para_runs(p)
        starts, pos = [], 0
        for r in runs:
            starts.append(pos)
            pos += len(r["text"])
        full = "".join(r["text"] for r in runs)
        idx = full.find(tok, starts[covered[0]])
        if idx < 0:
            raise SystemExit("内部错误：标记定位失败 %r" % tok)
        first = True
        for i in covered:
            set_run_text(runs[i], newtok if first else "")
            first = False
        changed += 1

    # ---------- 重排文后条目 ----------
    if reorder_list:
        by_old = {n: p for n, p in entries}
        after = ref_head
        for new in range(1, nref + 1):
            el = by_old[inverse[new]]
            el.getparent().remove(el)
            after.addnext(el)
            after = el
        for new in range(1, nref + 1):
            p = by_old[inverse[new]]
            runs = para_runs(p)
            full = "".join(r["text"] for r in runs)
            m = HEAD_RE.match(full)
            target = "[%d]" % new
            pos, placed = 0, False
            for r in runs:
                s, e = pos, pos + len(r["text"])
                pos = e
                if e <= m.start() or s >= m.end():
                    continue
                prefix = r["text"][: m.start() - s] if s < m.start() else ""
                suffix = r["text"][m.end() - s:] if e > m.end() else ""
                if not placed:
                    set_run_text(r, prefix + target + suffix)
                    placed = True
                else:
                    set_run_text(r, prefix + suffix)

    dst = out or path
    doc.save(dst)
    print(text_report)
    print("\n改写引用标记：%d 处；文献条目：%d 条；文件：%s" % (changed, nref, dst))
    if skipped:
        print("跳过（文献表内无此编号，请人工处理）：")
        for ctx, tok in skipped:
            print("  ", tok, "…", ctx)
    if report:
        with open(report, "w", encoding="utf-8") as f:
            f.write(text_report + "\n")
        print("映射表：", report)
    print("提示：如文档含目录/题注/REF 域，请在 Word 或 WPS 中再更新一次域。")
    return mapping


def main(argv=None):
    ap = argparse.ArgumentParser(description="按正文首次引用顺序重排参考文献编号")
    ap.add_argument("docx", help="待处理的 .docx")
    ap.add_argument("--out", help="输出文件（默认原地覆盖）")
    ap.add_argument("--report", help="把“旧号→新号”映射表写到该文件")
    ap.add_argument("--dry-run", action="store_true", help="只打印映射，不修改文件")
    ap.add_argument("--no-reorder-list", action="store_true", help="只改正文标注，不动文后条目顺序")
    a = ap.parse_args(argv)
    renumber(a.docx, a.out, a.report, a.dry_run, not a.no_reorder_list)


if __name__ == "__main__":
    sys.exit(main())
