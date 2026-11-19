# 第 4 章重写配方：验证项目及指标汇总 + 稳定性结果矩阵（跨品种通用）

将手册某章（典型为第 4 章「方法转移 / 稳定性考察结论摘要」）重写为聚焦的
**「4.1 各法验证项目及指标汇总表 + 4.2 稳定性考察结果（时间点结果矩阵）」**。
脚本：`scripts/build_validation_summary.py`（配置驱动，JSON）。本配方沉淀自《注射用乳糖酸红霉素》
手册 v7 第 4 章的两次重写实战。

> 跨品种通用配方。红霉素的具体数值仅作算例（见 `references/validation_summary_example_lk080.json`），
> **禁止写进脚本默认值**。

## §1 何时用本配方
- 用户要求把某章（如第 4 章）从「简单结论摘要」扩写为「逐法验证项目及指标 + 稳定性结果矩阵」。
- 用户明确要求「只汇总具体结果」「验证项目及指标、稳定性考察结果保留，其它可不要」。
- 配套：源文件含可读 DOCX 验证报告 + 扫描件（需 OCR）+ 稳定性考察报告。

## §2 目标产物结构
章标题保留（仅换正文）。新正文顺序：
1. 章首引导段（数据来源分层说明）。
2. **4.1 各分析方法验证项目及指标汇总** 标题 + 一句说明 + 表（方法｜验证项目｜关键指标｜依据报告）。
   - 法定方法（性状/鉴别/酸碱度/可见异物…）标「按 ChP 通则未单独验证，转移确认见 S0x」；
     已做验证的方法（有关物质★、含量 HPLC★、效价）给完整关键指标。
3. **4.2 稳定性考察结果** 标题 + 引导段 + **时间点结果矩阵**（考察项目 × 0/3/6/9/12/18 月）
   + 加速条件汇总句 + 结论段 + OCR 备注。

## §3 数据来源分层与 OCR 纪律
- **可读 DOCX 验证/稳定性报告**：直接抽文本与表格（`doc.element.body.iterchildren()` 保序，见 `python_docx_pitfalls.md`）。
- **扫描件（有关物质 / 效价验证报告）**：`scripts/ocr_pipeline.py` 或（签批件/数据表）直接渲染 PNG 后视觉读取（`references/ocr_pipeline.md` §9）。
- **不确定值一律标「OCR·需核」**，并在备注中提醒「须与纸质/电子原件逐格核对」；关键限度建议以原件复核。
- **效力顺序**（见 `manual_structure.md`）：签批 PDF ＞ 注册标准 ＞ 电子 DOCX ＞ 历史报告。扫描件数值以原件为准。

## §4 实现要点（对齐本次踩坑）
### 4.1 真实章节标题定位：取「最后出现」者
目录(TOC) 与正文标题同名同文本。定位时取 `max(indices)`（最后出现）= 真实正文章节，
避开 TOC 伪条目。→ 关联 `chapter_refactor_playbook.md §3.3`（目录与正文同名按 style.name 消歧；
本配方用「最后出现」兜底，更省事）。下一章锚点取「真实标题之后第一个匹配 heading」。

### 4.2 孤儿段防护：收集 → 统一插入
`doc.add_paragraph / add_table` 会把元素追加到文档**末尾**。若直接用 `addnext` 把最后建好的子元素
移入章节，会把它甩到文档最末（「本手册完」之后）。正确做法：
1. 把所有新建元素（段落 `._p`、表格 `._tbl`）收集进 `created` 列表；
2. 先 `el.getparent().remove(el)` 全部移出（ detached 元素无父节点，跳过 remove）；
3. 统一 `body.insert(ch4_idx+1, el)` 顺序插入到真实标题之后。
脚本 `insert_created()` 已实现此逻辑。

### 4.3 列宽合计 = 9638 twips = 17cm
A4 + 2cm 边距 → 纵向版心 17cm = 9638 twips。表格 `tblW`、`tblGrid/gridCol`、每个 `tcW` 三者之和必须相等且 = 9638，
**超出会出血/被截断**。横向宽表放 landscape section（版心 25.7cm = 14570 twips）。
脚本对每个表 `assert sum(col_widths) == required_total_twips`（默认 9638）。

### 4.4 OOXML 元素顺序（定宽表）
`tblPr`：`tblBorders → tblW → tblLayout(fixed) → tblCellMar`；再 `tblGrid`；再逐行
`tr`：`trPr` 含 **`cantSplit`**（必须放 trPr，**不可放 tcPr**）→ 每行 `tc`：`tcPr` 含 `tcW` + **`vAlign=center`**。
放错位置（如 cantSplit 误入 tcPr）会导致 Word 打开报错或红框跨页断裂。

### 4.5 字体四属性
每个 run 的 `rPr` 设 `rFonts`：`ascii/hAnsi = Times New Roman`、`eastAsia = 宋体`；`w:sz`/`w:szCs` 成对。
窄列（<5cm）禁加粗，避免宋体下重影（见 `revision_playbook.md §7`）。

### 4.6 标注纪律
- 矩阵中「12 月 / 18 月」凡取自扫描件 OCR 的格，必须标 `(OCR·需核)`。
- **效价不在稳定性考察项目清单内时，不得硬塞进矩阵**（会引出来源错配）。

### 4.7 零内容损失校验（写盘前/后）
- `prefix[0:real_ch4] == backup`、`suffix[real_ch5:] == backup`：章节外的文档零误伤。
- `raw.count('w:type="page"') == 0`：未引入强制换页（连续流式，见 `revision_playbook.md §12`）。
- 表格数净变动合理（删旧 4.2 设计表 + 加 2 表 = 净 +1 或持平，视原结构）。

## §5 零回归验收（复用 `verify_revision.py`）
- **文本级**：验证表行数/矩阵维度全命中、关键限度数值（如分离度、回收率区间）全命中、`**` 残留 0、U+FFFC 伪影 0。
- **OOXML 级**：列宽三同步（tblW/gridCol/tcW）、`cantSplit` 在 `trPr`、`vAlign` 在 `tcPr`、`w:br[type=page]` = 0。
- **尾部契约**：结尾标记「—— 本手册完 ——」独立成段且唯一、增量说明不错挂（见 `chapter_refactor_playbook.md §5`）。

## §6 与 `refactor_chapter.py` 的衔接
本配方是「章节内容**新建/重写**」，不是搬表/压缩，故**不走** `refactor_chapter.py` 的 7 阶段
（merge/delete/renumber/move/compress/tail/release），用独立脚本 `build_validation_summary.py` 完成；
也可在 refactor 流程后作为后处理执行。两者共享定位硬约束（§4.1、§4.4）。

## §7 算例（**非通用常量，禁止写入脚本默认值**）
红霉素全部数值（分离度 2.28、校正因子 1.0、各回收率区间、各时间点结果）仅作示例，
完整 JSON 见 `references/validation_summary_example_lk080.json`，供复现/对照用。脚本 `build_validation_summary.py`
本身**不含**任何产品数值，数据一律来自调用方传入的 config。

## 附录 引用索引
- 定位：`chapter_refactor_playbook.md §3.3`（目录与正文同名）
- 写入：`chapter_refactor_playbook.md §4`（xml:space Clark 记法、禁 `cell.text =`、整段回写取舍）
- 尾部契约：`chapter_refactor_playbook.md §5`
- 版式：`revision_playbook.md §7`（字体）、§11（三类块）、§12（连续流式去换页）
- OCR：`references/ocr_pipeline.md` §9（签批件/数据表视觉读取）
- 脚本：`scripts/build_validation_summary.py`（`--input/--output/--config/--dry-run/--selftest`）
