# 湖南大学学位论文写作 Skill

![tests](https://github.com/chenguisen/hnu-thesis-writing-skill/actions/workflows/tests.yml/badge.svg)

从《湖南大学研究生学位论文撰写规范》提取的一套可执行写作与排版规范，包含：

- 一个 AI skill（`SKILL.md`）：把规范变成写作流程、硬指标和常见错误清单；
- 五份规范 reference（论文结构、版式、参考文献、交叉引用、提交前检查清单）和一份来源说明；
- 一个仅依赖 `python-docx` 的自检脚本：检查 `.docx` 的页面、字体字号、标题、摘要、参考文献、图表、交叉引用等；
- 一套回归测试与一个无 ERROR 的样例文档。

> **版本说明**：本仓库依据的规范文档元数据创建于 **2005-08-31**，其中引用的国家标准为
> GB7713-87、GB7714-87、GB3100～3102-86。若湖南大学研究生院/学院已发布新版模板或规范，
> **以新版为准**；本 skill 用于新版未覆盖的细节、格式自检和写作流程。详见
> [`references/source-and-version.md`](references/source-and-version.md)。

## 包含内容

| 文件 | 用途 |
|---|---|
| [`SKILL.md`](SKILL.md) | Skill 入口：何时使用、工作流、硬指标速查、常见错误 |
| [`references/structure.md`](references/structure.md) | 论文构成、各部分的写作要求、字数与参考文献篇数 |
| [`references/formatting.md`](references/formatting.md) | 页面、字体字号、页眉页脚、标题层次、图表公式、印刷装订 |
| [`references/citations.md`](references/citations.md) | 正文引用位置、文后著录格式、GB7714 新旧版差异 |
| [`references/checklist.md`](references/checklist.md) | 提交前逐项检查清单（含机器查不了的项目） |
| [`references/cross-references.md`](references/cross-references.md) | 图/表/公式的题注、书签、REF 域与点击跳转 |
| [`references/source-and-version.md`](references/source-and-version.md) | 来源、提取方法、原文错误与版本说明 |
| [`scripts/thesis_lint.py`](scripts/thesis_lint.py) | 对 `.docx` 做结构与格式自检 |
| [`tests/test_lint.py`](tests/test_lint.py) | 自检脚本的回归测试 |

## 快速开始

### 1. 安装为 AI skill

把仓库放到技能目录（二选一，取决于你的 agent 客户端）：

```bash
# macOS / Linux / Git Bash：DSH / 通用 agent skills
git clone https://github.com/chenguisen/hnu-thesis-writing-skill.git \
  ~/.agents/skills/hnu-thesis-writing

# macOS / Linux / Git Bash：Claude Code
git clone https://github.com/chenguisen/hnu-thesis-writing-skill.git \
  ~/.claude/skills/hnu-thesis-writing
```

```powershell
# Windows PowerShell（~ 不会展开，需用 $env:USERPROFILE）
git clone https://github.com/chenguisen/hnu-thesis-writing-skill.git `
  "$env:USERPROFILE\.agents\skills\hnu-thesis-writing"

git clone https://github.com/chenguisen/hnu-thesis-writing-skill.git `
  "$env:USERPROFILE\.claude\skills\hnu-thesis-writing"
```

之后在对话里说“按湖南大学学位论文规范检查这篇论文”即可触发。Skill 的入口是
`SKILL.md`，引用的细节文件都在 `references/`。

### 2. 直接运行格式自检

```bash
python -m pip install -r requirements.txt
python scripts/thesis_lint.py thesis.docx --degree master
# 只写了一章时：跳过封面/声明/摘要/目录/参考文献/总字数
python scripts/thesis_lint.py chapter3.docx --degree master --chapter
```

`--degree` 可选：

| 值 | 含义 | 正文字数 | 中文摘要 | 参考文献 |
|---|---|---|---|---|
| `doctor` | 博士 | 5～8 万字 | 约 1200 字 | ≥80（外文 ≥40） |
| `master` | 学术硕士 | 3～5 万字 | 约 600 字 | ≥40（外文 ≥10） |
| `professional` | 专业硕士 | 3～5 万字 | 约 600 字 | ≥30（外文 ≥5） |

输出示例（截取）：

```text
========================================================================
湖南大学研究生学位论文格式自检报告
========================================================================
文件：thesis.docx
学位类型：学术硕士
结果：ERROR 0 项，WARN 5 项，INFO 0 项

关键指标：
  - 正文汉字数：419
  - 中文摘要汉字数：70
  - 关键词个数：4
  - 参考文献条数：5
  - REF 交叉引用域：0
  - 正文主字号(pt)：12.0
  - 正文主字体：宋体

[警告（建议核对）]
 1. [BODY-01] 正文汉字约 419 字，低于学术硕士要求（30000～50000 字）
 2. [XREF-01] 检测到 1 个图/表题注，但未检测到 REF 交叉引用域；正文编号可能是手打，插入/删除后不会自动更新，也不能点击跳转
 ...
```

退出码：`0` 无 ERROR（可有 WARN）；`1` 存在 ERROR；`2` 文件无法读取。
加 `--json report.json` 可同时输出机器可读报告。

### 3. 先查哪些项

脚本覆盖：A4 与页边距、页眉/页码域、标题层次与字号、正文字体字号与行距、摘要与关键词个数、
参考文献数量与外文比例、图表题注编号、表格左右边线、REF 交叉引用域与孤图/孤表、原创性声明与目录是否存在。
**不覆盖**：内容质量、逻辑、查重、签名、印刷装订、域是否已按 F9 更新且显示正确——这些用
[`references/checklist.md`](references/checklist.md) 人工核对。

## 规范要点（速查）

| 项目 | 要求 |
|---|---|
| 纸张/版心 | A4；版心 155 mm×253 mm |
| 页边距 | 上/下 22 mm，左/右 25 mm（与 155 mm 版心存在 5 mm 矛盾，见下） |
| 正文 | 小 4 号宋体（12 pt），每页 35 行、每行 35 字 |
| 章标题 | 小 2 号黑体（18 pt），居中，上下各空一行 |
| 节标题 `1.1` | 小 3 号黑体（15 pt），顶格，上下各空 12 磅 |
| 条标题 `1.1.1` | 4 号黑体（14 pt），上下各空 6 磅 |
| 款/项 | 小 4 号黑体（12 pt） |
| 页眉 | 小 5 号宋体（9 pt）；奇数页“学位论文类别”，偶数页“论文题目”；1.5 磅双线 |
| 页码 | 前置罗马数字（原创性声明页起），正文阿拉伯数字；页脚居中 |
| 图表 | 图题在下、表题在上；按章编号；表格不加左右边线 |
| 公式 | 居中；按章编号（1.1）；编号右顶格 |
| 结论 | 单独一章，不加章号，约 2000 字 |

### 图/表/公式能点击跳转吗？

能，但跳转靠的是 Word/WPS 的**交叉引用域**，不是手打编号：

1. 图/表：插入题注 → 正文用“交叉引用”，勾选“插入为超链接”；
2. 公式：给编号加书签，正文用“交叉引用 → 书签”（或 `SEQ` 域自动编号）；
3. Ctrl+A → F9 更新域，导出 PDF 前再更新一次。

完整步骤、PDF 链接保留和常见断链问题见
[`references/cross-references.md`](references/cross-references.md)。`thesis_lint.py` 会检查
`REF` 域是否存在，以及是否有“孤图/孤表”（题注未在正文引用）。

### 两个需要特别注意的矛盾

1. **版心宽度 vs 页边距**：A4 宽 210 mm，左右各 25 mm 后正文宽 160 mm，与“版心 155 mm”
   相差 5 mm。以学校模板为准；从零设置时可用左 30 mm（含 5 mm 装订线）、右 25 mm，
   或左右各 27.5 mm。
2. **旧规范 vs 新国标**：规范引用 GB7714-87；现行标准是 GB/T 7714-2015。学校模板未更新时
   按模板；学校要求新国标时全文统一按新国标，不要混用两套格式。

## 开发与测试

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
# 或
python tests/test_lint.py
```

测试使用 [`tests/fixture_master.docx`](tests/fixture_master.docx) 作为 0 ERROR 的最小回归样例
（示例内容为占位，因此有 4 条 WARN）；`test_non_a4_page_is_an_error` 会把页面改成 Letter，
确认脚本能报出 ERROR。

## 来源与许可

- 规范原文：`湖南大学论文撰写规范.doc`（Word 97-2003，1,051,936 字节，文档元数据创建于
  2005-08-31）。原文件与本仓库的转换副本**不纳入公开仓库**，规范文本的著作权归湖南大学。
- 本仓库只收录整理后的操作指引、检查清单与自检代码；整理时对原文的错字、旧标准号和
  尺寸矛盾做了标注，见 `references/source-and-version.md`。
- 代码采用 [MIT License](LICENSE)；文档与 skill 文本采用 CC BY 4.0，转载或改编请注明
  来源仓库与“依据湖南大学 2005 版规范整理”。详见 [`NOTICE.md`](NOTICE.md)。
- 本 skill 不是学校官方发布物，不构成对规范的解释；最终格式以研究生院/学院当年模板为准。
