#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""thesis_lint.py 的回归测试。

运行：
    python -m unittest discover -s tests -v
或：
    python tests/test_lint.py
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import docx

ROOT = Path(__file__).resolve().parent.parent
LINT = ROOT / "scripts" / "thesis_lint.py"
FIXTURE = Path(__file__).resolve().parent / "fixture_master.docx"


class LintTest(unittest.TestCase):
    def run_lint(self, path: Path, *extra: str) -> subprocess.CompletedProcess:
        cmd = [sys.executable, "-X", "utf8", str(LINT), str(path), "--degree", "master", *extra]
        return subprocess.run(
            cmd, capture_output=True, text=True, encoding="utf-8", errors="replace"
        )

    def test_compliant_fixture_has_no_error(self):
        result = self.run_lint(FIXTURE)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("ERROR 0", result.stdout)

    def test_non_a4_page_is_an_error(self):
        with tempfile.TemporaryDirectory() as td:
            bad = Path(td) / "bad.docx"
            document = docx.Document(str(FIXTURE))
            section = document.sections[0]
            section.page_width = docx.shared.Cm(21.59)  # Letter
            section.page_height = docx.shared.Cm(27.94)
            document.save(str(bad))
            result = self.run_lint(bad)
            self.assertEqual(result.returncode, 1, result.stdout)
            self.assertIn("PAGE-02", result.stdout)

    def test_json_report_is_written(self):
        with tempfile.TemporaryDirectory() as td:
            report = Path(td) / "report.json"
            result = self.run_lint(FIXTURE, "--json", str(report))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            data = json.loads(report.read_text(encoding="utf-8"))
            self.assertEqual(data["degree"], "master")
            self.assertIn("references", data["metrics"])
            self.assertIn("issues", data)


if __name__ == "__main__":
    unittest.main(verbosity=2)
