#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract_front_matter.py 的回归测试。

默认只跑不需要大文件的单元测试；设置环境变量 HNU_SOURCE_DOCX 指向一份
真实的湖南大学学位论文 .docx 时，额外跑一次端到端抽取。
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "extract_front_matter.py"
sys.path.insert(0, str(ROOT / "scripts"))
import extract_front_matter as ex  # noqa: E402

W = ex.W


def p_with_runs(*texts, underline=False):
    p = etree.Element(W + "p")
    for text in texts:
        r = etree.SubElement(p, W + "r")
        if underline:
            rpr = etree.SubElement(r, W + "rPr")
            u = etree.SubElement(rpr, W + "u")
            u.set(W + "val", "single")
        t = etree.SubElement(r, W + "t")
        t.text = text
    return p


class HelperTest(unittest.TestCase):
    def test_compact_strips_whitespace(self):
        self.assertEqual(ex.compact(" 培  养  单  位\n"), "培养单位")

    def test_replace_after_label_keeps_label_and_tab(self):
        p = p_with_runs("学位申请人姓名", "\t", "陈桂", "森", underline=False)
        # 把值 runs 设为下划线
        for r in p.findall(W + "r")[2:]:
            rpr = etree.SubElement(r, W + "rPr")
            u = etree.SubElement(rpr, W + "u")
            u.set(W + "val", "single")
        ok = ex.replace_after_label(p, "学位申请人姓名", "张三")
        self.assertTrue(ok)
        self.assertEqual(ex.text_of(p), "学位申请人姓名\t张三")

    def test_replace_after_label_with_colon(self):
        p = p_with_runs("学", "号", "：", "B1", "913Z", "0546")
        self.assertTrue(ex.replace_after_label(p, "学号：", "B00000000"))
        self.assertEqual(ex.text_of(p), "学号：B00000000")

    def test_set_paragraph_text_keeps_first_run_format(self):
        p = p_with_runs("旧标题", "多", "余")
        rpr = etree.SubElement(p.findall(W + "r")[0], W + "rPr")
        etree.SubElement(rpr, W + "sz").set(W + "val", "44")
        ex.set_paragraph_text(p, "新标题")
        self.assertEqual(ex.text_of(p), "新标题")
        self.assertEqual(len(p.findall(W + "r")), 1)

    def test_find_body_start(self):
        blocks = []
        front = etree.Element(W + "p")
        ppr = etree.SubElement(front, W + "pPr")
        etree.SubElement(ppr, W + "pStyle").set(W + "val", "1")
        np = etree.SubElement(ppr, W + "numPr")
        etree.SubElement(np, W + "numId").set(W + "val", "0")
        blocks.append(front)

        chapter = etree.Element(W + "p")
        ppr = etree.SubElement(chapter, W + "pPr")
        etree.SubElement(ppr, W + "pStyle").set(W + "val", "1")
        np = etree.SubElement(ppr, W + "numPr")
        etree.SubElement(np, W + "numId").set(W + "val", "1")
        etree.SubElement(chapter, W + "r")
        t = etree.SubElement(chapter.findall(W + "r")[0], W + "t")
        t.text = "绪论"
        blocks.append(chapter)

        self.assertEqual(ex.find_body_start(blocks, {}), 1)

    def test_truncate_field_result(self):
        p = etree.Element(W + "p")
        for typ, text in (("begin", None), ("separate", None)):
            r = etree.SubElement(p, W + "r")
            fc = etree.SubElement(r, W + "fldChar")
            fc.set(W + "fldCharType", typ)
            if text:
                it = etree.SubElement(r, W + "instrText")
                it.text = text
        r = etree.SubElement(p, W + "r")
        etree.SubElement(r, W + "t").text = "旧目录"
        r = etree.SubElement(p, W + "r")
        etree.SubElement(r, W + "fldChar").set(W + "fldCharType", "end")

        self.assertTrue(ex.truncate_field_result(p, "（更新域）"))
        self.assertIn("（更新域）", ex.text_of(p))
        self.assertNotIn("旧目录", ex.text_of(p))
        begins = [fc for fc in p.iter(W + "fldChar")
                  if fc.get(W + "fldCharType") == "begin"]
        self.assertEqual(begins[0].get(W + "dirty"), "true")


class EndToEndTest(unittest.TestCase):
    @unittest.skipUnless(os.environ.get("HNU_SOURCE_DOCX"), "未设置 HNU_SOURCE_DOCX")
    def test_extract_real_thesis(self):
        source = Path(os.environ["HNU_SOURCE_DOCX"])
        if not source.exists():
            self.skipTest(f"源文件不存在：{source}")
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "skeleton.docx"
            result = subprocess.run(
                [sys.executable, "-X", "utf8", str(SCRIPT),
                 "--source", str(source), "--out", str(out),
                 "--title", "测试题目", "--author", "测试"],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(out.exists())

            import docx
            document = docx.Document(str(out))
            self.assertGreaterEqual(len(document.sections), 6)
            text = "\n".join(p.text for p in document.paragraphs)
            for needle in ("学校代号", "学位论文原创性声明", "摘   要", "Abstract",
                           "目   录", "插图索引", "附表索引", "绪论",
                           "参考文献", "致谢"):
                self.assertIn(needle, text, needle)


if __name__ == "__main__":
    unittest.main(verbosity=2)
