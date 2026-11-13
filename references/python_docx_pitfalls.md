# python-docx 排版避坑（本环境实测）

把这些当成硬约束，跳过任何一个都会在生产文档里出现"表体丢失 / 列宽溢出 / 文档损坏 / 锁文件"等离奇问题。

## 1. 必须显式设 A4 + 页边距
默认模板是 **Letter** 尺寸且版心仅 16.5cm。任何 17cm 及以上的宽表都会溢出。
```python
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.0)
# 纵向版心 = 21.0 - 2.0 - 2.0 = 17.0cm
# 横向版心 = 29.7 - 2.0 - 2.0 = 25.7cm
```

## 2. 固定布局表的 gridCol / tblW 必须手动同步
`doc.add_table(...)` 会把 `w:tblGrid` 的 gridCol **均分**为 8640 twips（15.24cm），**不跟随** `cell.width`。只设 `cell.width` 而不同步 gridCol，Word/WPS 渲染的列宽会与设定不符。

正确做法（见 `scripts/docx_table_toolkit.py` 的 `_set_grid_widths`）：
- 遍历 `tblGrid/gridCol`，逐个写 `w:w = int(round(w_cm * 567))`（1cm = 567 twips）
- 同时显式写 `w:tblW`（`w:type=dxa`, `w:w=合计twips`）。**不写 tblW，Word 的 PDF 导出会对"首页第一张固定表"丢表体。**

## 3. OOXML 子元素顺序（顺序错 → Word 报"文档已损坏"或直接丢内容）
- `w:tblLayout` **必须位于** `w:tblLook` 之前。
- `w:cantSplit` **必须位于** `w:tblHeader` 之前（trPr 内）。
- 所有插入用 `insert_element_before(elem, *successors)`，不要用 `append()`——append 会把元素放到错误位置。

顺序参考：
```
tblPr: ... tblW, tblLayout(fixed), tblCellMar, tblLook ...
trPr : cantSplit, trHeight, tblHeader, tblCellSpacing, jc ...
```

## 4. 重复表头（tblHeader）的隐藏坑
固定布局表"位于页首 + 表头行带 tblHeader"会触发 **Word PDF 导出丢表体**。规避：
- 仅对**会跨页的较长表（>12 数据行）**加 `tblHeader`；短表（如 8 行）不加。
- 表头行**不要**加 `cantSplit`（同样易在页首丢表体）。
- 数据行加 `cantSplit` + `trHeight rule=AT_LEAST` 防止行内跨页断行。

## 5. CJK 字体必须双写
只设 `run.font.name` 不够，Word 会回退宋体。必须同时写 `rPr/rFonts` 的 `w:eastAsia`（以及 ascii/hAnsi）：
```python
rFonts.set(qn("w:eastAsia"), name)
```

## 6. 读取文档时保留段落↔表格顺序
`doc.paragraphs` 和 `doc.tables` 是**各自独立**的列表，丢失相对顺序。遍历正文顺序要用：
```python
for el in doc.element.body.iterchildren():
    if el.tag == qn('w:p'):   ...段落...
    elif el.tag == qn('w:tbl'): ...表格...
```

## 7. PyMuPDF 恒定 segfault —— 用 pypdfium2
本环境 `fitz`/`PyMuPDF` 光栅化任何 PDF 都 exit 139。`pip install pypdfium2` 可正常渲染
PDF→PNG，无系统依赖。**workbuddy-pdf CLI 在此环境也不可用（exit 127/139）。**

## 8. pypdfium2 文本提取的 U+FFFC 伪影（极易误判"表体丢失"）
pypdfium2 对固定布局表单元格提取文本时，会在**字母与数字之间插入 U+FFFC（OBJECT
REPLACEMENT CHARACTER，显示为 ￾）**，例如 `"RP-01"` 变成 `"RP￾01"`。后果：
- 精确字符串搜索 `'RP-01'` 返回 **0 次**；
- 低清渲染图（ViewImage）又漏读 → 一度误判"G.3.1 表体整块丢失"。

**真相：docx / PDF 实际数据完整，该字符只是文本提取伪影，不影响 Word/PDF 显示。**
凡遇"表体数据搜不到"，先打印整页文本层或 ViewImage 直接看图，勿轻信精确子串搜索。

## 9. 预览面板独占锁 → PermissionError
目标 docx 正被预览面板打开时，`doc.save()` 抛 PermissionError（连读也失败，独占锁）。
处理：用环境变量 `OUT_OVERRIDE` 构建到临时路径，再 `copy`；`copy` 同样可能被锁，需重试或另存新名。
判断文件是否被打开：看目录里是否有 `~$xxx.docx` 锁文件。

## 10. Word COM 导出 PDF（按环境分支；本机实测被沙箱拦截）

**环境可用时**（历史 v1~v5 的存量 PDF 就是这样导出的，原实现仍然有效）：
```powershell
$wd = New-Object -ComObject Word.Application   # version 12.0
$doc = $wd.Documents.Open($path)
$doc.Fields.Update() | Out-Null                 # 见 §10b，防页码 0/1
$doc.ExportAsFixedFormat($pdfPath, 17)          # 17 = wdExportFormatPDF
$doc.Close(); $wd.Quit()
```
调用前先 `tasklist` 确认 Word 未运行，避免干扰用户会话；`Quit()` 后偶发 RPC 异常，属无害。

> ★ **本机实测结论（环境已变化）**：Word COM 自动化被沙箱拦截 —— WINWORD.EXE 无法启动，
> phantom v12 注册返回 `E_FAIL`。因此 **DOCX→PDF 不能由脚本完成**：交付 DOCX 并请用户在 Word
> 中手动「另存为 / 导出为 PDF」；可选兜底为 WPS 或 `soffice --convert-to pdf`。
> `scripts/export_pdf.ps1` 保留为「环境修复后可用」的实现，勿当作当前交付步骤。
> 排查顺序：`tasklist | findstr WINWORD` → 尝试 `New-Object -ComObject Word.Application` →
> 报 `E_FAIL` 即判为拦截，直接走降级路径，不要反复重试。
> PDF 可用后再跑 `scripts/verify_revision.py --pdf` 补逐页断言（见 §12 与 revision_playbook §8）。

## 10b. 导出前必须 Fields.Update（否则页码显示 0/1）
（以下为「环境允许 Word COM」时的参考步骤；本机 COM 被拦截时跳过，见 §10。）
Word COM 导出的 PDF 里，PAGE / NUMPAGES 等**域不会自动刷新**——直接导出会得到"第 0 页 / 共 1 页"。
`ExportAsFixedFormat` 之前务必先：
```powershell
$doc.Fields.Update() | Out-Null
$doc.ExportAsFixedFormat($pdfPath, 17)
```
对应脚本见 `scripts/export_pdf.ps1`（已内置该步骤）。

## 11. 页眉 / 页脚插入动态域（PAGE / NUMPAGES）
python-docx 没有原生"插入域" API，必须用 `OxmlElement` 手搓 `fldChar(begin) / instrText / fldChar(end)` 三元组：
```python
from docx.oxml import OxmlElement
def _add_field(paragraph, field_code):
    run = paragraph.add_run()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = f" {field_code} "
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    run._r.append(b); run._r.append(i); run._r.append(e)
```
页脚「第 X 页 / 共 Y 页」：先 `第 ` 文本 → `_add_field(p, "PAGE")` → ` 页 / 共 ` → `_add_field(p, "NUMPAGES")` → ` 页`。
多节文档（如含横向附录）：先在 `sections[0]` 写好页眉页脚，再把后续节 `is_linked_to_previous=True`，顺序反了会覆盖已写内容。
封装见 `scripts/docx_table_toolkit.py` 的 `add_header_footer(doc, header_text)`。

## 12. 原地修订既有成稿的环境约束

- **删除/移动**：本机 safe-delete 钩子拦截 `rm` / `Remove-Item` / `os.remove` / 跨盘 `move`；
  同盘 `os.rename` 不受影响。回滚一律用「覆盖式拷回」，不要用删除。
- **文件锁**：预览面板独占锁会让 `doc.save()` 抛 `PermissionError`（见 §9）。所有修订脚本默认
  **另存 `--output` 新文件、不覆盖输入**；若必须回写原路径，先另存副本，待锁释放再覆盖。
  被 WorkBuddy 进程锁定的文件需重启释放。
- **无 OCR**：本机无 tesseract；纯图片 PDF 用 pypdfium2 渲染成 PNG 后**视觉读取**（见 ocr_pipeline.md §9）。
  禁用 PyMuPDF（segfault）。旧 `.doc` 用 `Word COM Documents.Open → Content.Text` 写 UTF-8 txt
  （勿 SaveAs2），并 `replace("\x07", " ")` 去单元格分隔符。
- **Word COM 分支**（见 §10）：环境允许 → `export_pdf.ps1` + `Fields.Update()`；
  被拦截 → 交付 DOCX + 用户手动导出 PDF，可选 WPS / `soffice --convert-to pdf`。
  **逐页目视是独立一层**：PDF 未产出时视觉级验收整层标 SKIP，并在报告里明确
  「PDF 需手动导出、逐页目视待补」，不得用 XML 级通过冒充版式已验收。
- **三层验收清单**见 `references/manual_structure.md` §7 与 `references/revision_playbook.md` §8。
