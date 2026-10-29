# 文献管理：让编号和著录格式自己更新

这份说明解决一个具体问题：**论文里的参考文献现在是纯文本**（116 条手打条目 + 正文里手打的
上标 `[15]`），加一条、删一条，编号和 180 多处引用都要跟着手改，而且现用格式不是 GB/T 7714。
正确做法是把文献放进文献管理软件，正文引用变成“域”，编号与文后列表由软件生成。

## 1. 选型：Zotero + Word 插件 + GB/T 7714-2015

| 方案 | 是否支持 WPS | 中文 GB/T 7714 样式 | 备注 |
|---|---|---|---|
| **Zotero**（推荐） | ❌ 只支持 Word / LibreOffice / Google Docs | ✅ 官方样式库有（numeric 与 author-date 各一） | 免费、浏览器一键抓取；本机已装 |
| EndNote | ❌ 只支持 Word | ✅ 需装学校/期刊样式 | 2025 答辩版论文用的就是 EndNote |
| NoteExpress / 知网研学 | ✅ 支持 WPS | ✅ 国内文献覆盖好 | 收费/校园版；WPS 用户可选 |
| Word 内置“引用” | ✅（在 Word 里） | ⚠️ 样式受限，中文著录常不合规 | 条目少、要求不严时才用 |

**如果继续用 WPS 排版**，Zotero 帮不上（它没有 WPS 插件）；那种情况选 NoteExpress/知网研学，
或者改用 Word。

## 2. 本机现状（已就绪的部分）

- Zotero 已安装：`C:\Program Files\Zotero`
- Word 插件已装：`%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm`（打开 Word 会有 Zotero 工具条）
- GB/T 7714-2015 样式已放好：
  `%USERPROFILE%\Zotero\styles\china-national-standard-gb-t-7714-2015-numeric.csl`
  （题名：`China National Standard GB/T 7714-2015 (numeric, 中文)`；重启 Zotero 后在
  设置 → 引用 → 样式 里可选，Word 工具条的 Document Preferences 里也能选）

## 3. 一次性迁移（约 1～2 小时）

1. **导入现有条目**：双击仓库根目录的 `references_v6.ris`（或 Zotero：文件 → 导入）。
   116 条里 81 条已按题目匹配到 DOI；35 条没有 DOI，导入后仍是完整条目，只是缺 DOI。
2. **补全缺项**：对没 DOI 的条目，在 Zotero 里右键 → “查找可用的 PDF” / 手动填 DOI，
   或用 Zotero 的“通过标识符添加条目”（DOI/ISBN/arXiv 都行）。
3. **在 Word 里替换纯文本引用**（先在 WPS/Word 外备份一份 `.docx`）：
   - 删掉文后“参考文献”里的 116 条纯文本（保留“参考文献”这个标题）；
   - 删掉正文里手打的上标 `[n]`（用查找 `\[[0-9]` 逐个定位，别误删“图[1]”之外的正常方括号内容）；
   - 光标放到引用位置 → Zotero 工具条 → **Add/Edit Citation** → 选条目；
     多条可一次选，会自动合并成 `[1,2]`、`[1-3]`；
   - 全部引完后 → **Add/Edit Bibliography**，在“参考文献”标题下生成列表。
4. **选样式**：Zotero 工具条 → **Document Preferences** → 选
   `China National Standard GB/T 7714-2015 (numeric, 中文)`。
5. **核对**：文献表顺序应与正文首次出现顺序一致；中文条目显示“等”，期刊有 `[J]`、专著 `[M]`。

## 4. 日常更新（这才是目的）

- **加文献**：Zotero 里加条目（浏览器插件一键抓取）→ Word 里插引用 → 编号、列表自动重排。
- **删文献**：Word 里删掉那条引用标记 → 点 **Refresh** → 列表与编号自动更新。
- **换风格**：Document Preferences 换样式；`numeric` 是顺序编码制（学校要求），
  另一个 `author-date` 是著者-出版年制，别混用。
- 定稿前：Word 里 **Refresh** 一次 → 导出 PDF；**不要**在 WPS 里“更新域”或另存。

## 5. 常见坑

1. **中文文献要显示“等”**：条目的 Language 设成 `zh-CN`（Zotero 右侧信息栏 / 右键）。
2. **`[J]`/`[M]`/`[C]` 由条目类型决定**：期刊 Article、专著 Book、会议 Paper（conference）、
   学位论文 Thesis、标准 Standard；类型错了标记就错。
3. **卷、期、页码、DOI 尽量填全**：GB/T 7714 缺卷期页码会显得不齐（`2023, 15(2): 33-41`）。
4. **顺序编码制要求“按正文首次出现排序”**：`numeric` 样式自动做到，不要用 author-date。
5. **WPS 会把域当文本**：插入 Zotero 域之后，别再用 WPS 打开保存；Word 里编辑最安全。
6. **和图表交叉引用互不影响**：`REF`（图表）与 Zotero 的 citation 域可以共存。
7. **迁移后不要手工改文献表**：列表是域生成的，手改的内容下次 Refresh 会丢。

## 6. 换不动的时候：最低成本的过渡方案

如果暂时不想迁移，至少把“编号”这件事做成可脚本化的：

- 用一个表格（Excel/CSV）维护条目，按正文出现顺序排列；
- 用脚本生成文后列表与正文上标（`scripts/` 里现有工具不覆盖这一步，需要时可再补）；
- 风险：正文里仍有一批手打上标，改一条仍要人工核对——所以只建议作为过渡。

> 结论：**要用 Word 写论文、用 Zotero 管文献、用 GB/T 7714-2015 numeric 样式**，
> 编号与著录格式就永远不用手改了。
