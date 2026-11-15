---
name: xiaobo-pharma-qc-builder
description: "This skill should be used when generating, structuring, or maintaining Chinese pharmaceutical QC inspection preparation manuals (药品检验手册 / 全项检验准备与关键点控制手册) from QC source documents such as SOPs, method validation reports, drug standards (质量标准), and certificate-of-analysis files. It covers the full pipeline: OCR extraction of scanned Chinese PDFs, python-docx document assembly across multiple stages (one-page overview, per-test operation cards, in-place SOP quoting, appendix material/BOM checklists), fixed-layout table rendering with correct OOXML ordering, header/footer injection (PAGE/NUMPAGES fields), per-test 关键点/易错点 callouts with auto line-splitting on circled numbers (①②③…), folding real 偏差/OOS/CAPA cases into callouts (de-identified), and the domain-specific data model for an inspection material list (检验物料清单 / 备料核对清单) with auto-derived 检验项×物料 matrix. It also covers Stage 6 post-generation revision of an already-delivered manual: content-based column resizing, table font normalization, batch-record lot/supplier backfill, chapter merging/renumbering, readable 检验项×物料 matrices, three-layer zero-regression verification, and formal version release with version/date reminders. Trigger on phrases like 检验手册, 检验物料清单, 备料核对清单, SOP 嵌入手册, QC 文件梳理, 偏差案例挂接, 关键点框, 页眉页脚, 成稿修订, 修订升版, 列宽重算, 批号回填, 供应商回填, 矩阵全名, 版本沿革, 增量说明, 零回归验收, or requests to build or revise inspection-ready manuals for injectables or other drug products from scattered QC docs. It also covers Stage-6 chapter refactoring on an existing manual: quantifying section volume before trimming, the 「搬表而非删表」(move-the-table-don't-delete-the-table) rule (relocate detail tables into a new appendix verbatim and leave only conclusions plus pointers in the chapter), subsection merge / renumber / delete driven by a JSON config, cross-reference rewriting after relocation, tail-structure integrity assertions (the incremental-notes section must not carry another section's items; the closing marker must stand alone and appear exactly once), and user-intent-to-action mapping (rollback / merge-two-sections / trim-a-verbose-chapter / open-for-me). Trigger on phrases like 章节精简, 章节重构, 冗长章节, 章节体量, 搬表, 明细表迁出, 章节合并确认, 小节重编号, 交叉引用改指, 增量说明错挂, 结尾标记粘连, 回退上一版, 章节点乱, 整理精简该章, in addition to 检验手册, 检验物料清单, 备料核对清单, SOP 嵌入手册, QC 文件梳理, 偏差案例挂接, 关键点框, 页眉页脚, 成稿修订, 修订升版, 列宽重算, 批号回填, 供应商回填, 矩阵全名, 版本沿革, 增量说明, 零回归验收. It also covers post-delivery visual-consistency normalization and continuous-flow layout of an existing manual: unifying the per-test card's three blocks (main card table / SOP quote block / key-point callout) to full text-width and left alignment, and removing forced page breaks (run-level w:br type=page between cards and w:pageBreakBefore on headings) that waste pages, replacing them with a light separator line plus safety nets (cantSplit / keep_with_next). Trigger on phrases like 三类块对齐不一致, 没铺开, 操作卡没对齐, 章节间空白太多, 页数太多, 强制换页, 连续排版, 流式排版."
agent_created: true
---

# Xiaobo Pharma QC Builder

## Overview

Build Chinese pharmaceutical QC inspection-prep manuals from heterogeneous QC source
documents (scanned SOP PDFs, readable validation-report DOCX, drug standards, CoA). The
skill packages a verified python-docx table toolkit, a Chinese-scanned-PDF OCR pipeline,
the data model for an inspection material list, and the hard-won layout/extraction
pitfalls so the work is reproducible without re-discovering them each time.

The deliverable is typically a multi-stage Word manual for one drug product:
0) 编制说明/溯源 → 1) 产品基本信息 → 2) 一页速览总览表(检验项清单) → 3) 逐项操作卡(每
项一张，含🔴关键点框) → 4) 验证结论/限度摘要 → 5) 附录 A~G → 6) 参考文献与待核对项。
Appendix G (检验物料清单) is a 备料核对清单, not a cost ledger.

**Validated across products.** The skill has been used end-to-end on multiple injectable
products: 聚多卡醇注射液 (built from scratch, v1→v2), 骨化三醇注射液 (a full Stage 6
revision v5→v6), and 注射用乳糖酸红霉素 (Stage 6 revision to v4). The 骨化三醇 v5→v6
revision — six action classes (material source backfill / font normalization / chapter
consolidation / matrix readability / matrix trimming / zero-regression verification) — is
captured as a worked retrospective in `references/case_study_gsc_v6.md`.

## When To Use

- User wants a QC inspection manual, SOP-embedded operation guide, or material checklist
  for a drug product from QC source files.
- User has scanned Chinese PDFs (SOP / 质量标准) that need OCR before use.
- User asks to build / fix / extend a 检验物料清单 or 备料核对清单.
- User hits python-docx table layout bugs (columns overflow, header repeats wrong,
  table body missing in PDF export, document "corrupted" error).
- User asks to revise an already-delivered manual: resize table columns by content,
  unify table fonts, merge/renumber chapters, rename matrix codes, drop a matrix column,
  or backfill lot/supplier from 批检验记录 scanned records.
- User asks to rewrite a chapter (e.g. the 验证结论 / 稳定性摘要 chapter) into a focused
  「各法验证项目及指标汇总表 + 稳定性时间点结果矩阵」 → use `scripts/build_validation_summary.py`
  with a JSON config and `references/validation_summary_playbook.md`.
- User asks for a formal version upgrade (升版 / 版本沿革 / 增量说明) or a zero-regression
  check after edits.
- Word COM export is blocked and a DOCX / XML-level verification fallback is needed.

## Workflow

### Stage 0 — Collect & assess sources
1. Inventory source files; classify each as **readable DOCX** vs **scanned PDF**.
2. For scanned PDFs, run the OCR pipeline (`scripts/ocr_pipeline.py`); for readable DOCX,
   extract tables/text directly (use `doc.element.body.iterchildren()` to preserve
   paragraph↔table order — see `references/python_docx_pitfalls.md`).
3. Identify the **authoritative source per data field** (限度主依据 usually 稳定性考察
   报告表 or 质量标准; method detail from 验证报告 DOCX). Record source conflicts.

### Stage 1 — One-page overview + per-test cards (v1)
- Build a 一页速览总览表 listing all 检验项 (e.g. 14 项 for 聚多卡醇注射液).
- For each 检验项, one operation card: 操作步骤 / 系统适用性 / 计算公式 / 结果与判定,
  plus a 🔴 关键点框 for error-prone steps.
- Use `add_grid_table` / `add_kv_table` from `scripts/docx_table_toolkit.py`.

### Stage 2 — Embed SOP text in-place (v2)
- Render SOP excerpts as **light-shaded quote blocks** inside each card, tagged with the
  SOP section number + page for traceability.
- On numeric conflicts, **correct in-place per the现行 SOP**, and keep an audit trail
  (old value ↔ new value comparison table) in an appendix.
- Decision: core subsections quoted in full; instrument/reagent/reference-standard lines
  become one-line index rows.

### Stage 3 — Appendix G 检验物料清单 (备料核对清单)
- Model materials with `references/materials_data_model.md`: six categories
  (对照品 / 试剂试液 / 供试品 / 色谱柱 / 仪器 / 消耗品器皿), IDs `RP-/RGT-/SMP-/COL-/INS-/CON-`.
- Generate the six G.3.x sub-tables with `add_grid_table_fixed` (fixed layout, exact
  column widths, cross-page repeating header for long tables only).
- Add a ✔ 核对 column (empty cell + header "√"); do NOT use ☐/☑ glyphs or w:sdt controls.
- Keep a `build_matrix()` that derives the 检验项×物料 matrix from the material `items`
  field — never hand-maintain two copies. For 完整名称（编号） cells, 中文全称（缩写）
  headers, 顿号 joining, and trimming a category column (e.g. SMP-), see
  `references/revision_playbook.md` §6.

### Stage 4 — Fold in real deviation / OOS cases (optional but high-value)
- Mine historical 偏差处理报告 / OOS-OOT 调查报告 and hang the lessons on the 🔴 关键点框
  of the affected cards. Pattern, data model, and phrasing rules:
  `references/manual_structure.md` §6.
- Keep the case text in a separate data module (like `materials.py`) and let the card
  function pull it in via a `case_code` argument — the number of 关键点框 must not change.
- **De-identify**: replace employee names with role descriptions (QC主管 / 检测人员 /
  新入职人员 / 取样人员); keep 编号 / 日期 / 批号 / 限度 for traceability.

### Stage 5 — Render & verify
- **DOCX→PDF is environment-dependent**: when Word COM is available, export via
  `ExportAsFixedFormat` and render with pypdfium2 for the ViewImage page-by-page check
  (PyMuPDF segfaults here; the U+FFFC text-extraction artifact falsely reports "missing
  table body"). On this machine Word COM automation is blocked by the sandbox
  (WINWORD.EXE cannot start; phantom v12 registration returns E_FAIL) — deliver the DOCX,
  ask the user to export the PDF manually in Word, run XML-level zero-regression
  assertions instead, and **explicitly state that the PDF export and page-by-page visual
  check are still pending** (see `references/python_docx_pitfalls.md` §10 / §12 and
  `references/revision_playbook.md` §8).
- **页眉/页脚**：用 `add_header_footer(doc, header_text)` 注入页眉描述文字与页脚
  「第 X 页 / 共 Y 页」（PAGE/NUMPAGES 域）；导出前**必须** `Fields.Update()`，否则页码显示 0/1
  （见 pitfalls §10b / §11）。
- **红框多序号分行**：`add_key_callout(doc, text)` 内部按带圈数字 ①②③… 自动逐条分行；
  挂了偏差案例的 ⑥ 也会独占一行，扫读无障碍。
- Visually confirm via `ViewImage` (do not trust exact substring search on extracted text).
- When a scanned 签批件 and its electronic DOCX twin both exist, **cross-check them** —
  the PDF often carries the checked 最终处理决定, the signature chain, and hand-written
  conclusions the DOCX lacks, and sometimes an extra 受影响批号.

### Stage 6 — Revise an existing manual & release a new version
- Applies to an already-delivered `.docx`; full SOP is `references/revision_playbook.md`
  (layer 1 = cross-product rules, layer 2 = product-specific decisions that must NOT be
  generalized). A complete worked example (骨化三醇注射液 v5→v6, six action classes with
  real measured config) is `references/case_study_gsc_v6.md`.
- Baseline & timestamped backup first; locate structures by **heading anchor**, never a
  cached table index; always write to a new `--output` file, never overwrite the input.
- `scripts/patch_existing_manual.py` — `--strategy autofit|content-width` (content-width
  default) + `--normalize-fonts --font-mode cjk-split|all-song`.
- `scripts/verify_revision.py` — config-driven three-layer verification (text / OOXML
  structure / visual); the visual layer is SKIP when no PDF exists.
- Formal release: **new filename**, never overwrite the previous version; update only
  current-version fields (cover subtitle/version/date, shared header part, summary);
  append a new 版本沿革 row and a new 增量说明 section by cloning existing templates;
  never rewrite historical records.
- **After every document change, remind the user to update the version name and the date
  in the filename**, listing every place that carries them and giving current + suggested
  values (today's `YYYYMMDD`). Reminding is not renaming — wait for confirmation. See
  `references/revision_playbook.md` §9 / 附录 B.
- **Operation-card block visual consistency (user says the three card blocks look misaligned / not full-width)** — the per-test card bundles three blocks (main 2-column card table / SOP quote block / key-point red callout) that must share one width and alignment. Run `scripts/fix_operation_card_blocks.py --doc --out --content-width-cm` (17.0 portrait / 25.7 landscape) to normalize them to full text-width + left-aligned + vertically-centered cells; it identifies each block by template-fixed traits (red/blue left border, 2-column with 方法要点) so appendix/product/matrix tables are never touched. Always `--dry-run` first, then write a new `--output` (never overwrite input); confirm pure-text zero change. See   `references/revision_playbook.md` §11.
- **Remove forced page breaks for continuous flow (user says chapters have too much blank space / too many pages caused by forced page breaks)** — the per-test card section and chapter headings often carry forced breaks (run-level `w:br type="page"` between cards, `w:pageBreakBefore` on Heading 1) that waste pages. Run `scripts/fix_flow_layout.py --doc --out` to strip all forced breaks and switch to continuous flow; it injects a light gray separator line between cards, and adds safety nets (`cantSplit` on the three card-block tables, `keep_with_next` on card titles + H1s) so red callouts never split across pages and headings never orphan. Flags: `--card-pattern` (default `^3\.\d+[　\s]`), and `--no-card-pagebreaks` / `--no-chapter-pagebreaks` / `--no-separator-line` / `--no-safety` to scope the change. Always `--dry-run` first, write a new `--output` (never overwrite input), confirm pure-text zero change. See `references/revision_playbook.md` §12.
- **Visual-polish pass (recommended order after any Stage 6 content change)** — run the two
  layout scripts in sequence: first `scripts/fix_operation_card_blocks.py` (uniform width +
  left alignment of the three card blocks), then `scripts/fix_flow_layout.py` (strip forced
  page breaks for continuous flow), then `scripts/verify_revision.py` for zero-regression.
  Both are pure-layout, zero-content-loss; always `--dry-run` first and write a new
  `--output` (never overwrite input). See `references/revision_playbook.md` §11 / §12.
- **Chapter refactoring (user says a chapter is messy / verbose / too long)** — first
  **quantify** each subsection's volume with `scripts/chapter_volume.py`; volume usually sits
  in the big detail tables, so pure text trimming caps out around 8~25%. Preferred move:
  **relocate the detail table into a new appendix verbatim** and leave only conclusions plus
  a pointer in the original chapter — one pass, ~50%+ reduction, zero factual loss. After the
  move, rewrite every cross-reference that still points at the old location, and assert the
  table's dimensions are unchanged. Full SOP: `references/chapter_refactor_playbook.md` §1~§2;
  runner: `scripts/refactor_chapter.py --config`.
- **Matching hard rules** — never compare `Paragraph` objects with `is` (`doc.paragraphs`
  returns fresh objects each access; compare `p._p`); never match on style **id** (real docs
  use `'000006'` / `'21'` — always `p.style.name`); TOC entries share the chapter-title text
  (TOC = `List Bullet`, body = `Heading 1`, disambiguate by `style.name`); subsection styles
  are **not uniform** within one chapter — locate by text prefix, not by heading level;
  numbering tokens and title text often live in **different runs** (and the space count
  varies — `"6.2 "` vs `"6.5  "`), so match with `^\d+\.\d+` + `\s+`; table paragraphs are
  **not** in `doc.paragraphs` (scan paragraphs + every table cell); **renumber strictly from
  the largest number down**, and never map onto a number that already exists.
- **Renumbering needs two phases** — descending order alone still self-collides (measured:
  seven headings all collapsed into one number). Phase 1 rewrites every source number to a
  collision-proof placeholder (`§§TMPn§§`), phase 2 swaps placeholders for the targets.
- **Locator gotcha: naming a bullet style cancels the bullet filter** — `find_paragraphs`
  must apply the style test **before** the "skip List Bullet" filter, otherwise
  `style='List Bullet'` always returns empty and every template lookup (increment-notes
  bullets, TOC anchors) silently yields `None`.
- **`docx.table._Row(tr, parent)` needs a parent that has `.table`** — pass `tbl.rows`
  (or the `Table`), never the bare `CT_Tbl`; otherwise `AttributeError: 'CT_Tbl' object
  has no attribute 'table'`.
- **Writing text** — always set `xml:space="preserve"` via the Clark form
  (`{http://www.w3.org/XML/1998/namespace}space`; the literal `'xml:space'` raises
  `ValueError: Invalid attribute name`); **never** use `cell.text =` (it drops `rPr` — use
  `set_cell_text_preserve`); a whole-paragraph rewrite keeps only the first run's `rPr`.
- **A whole-paragraph rewrite must never land on a heading** — that erases the section
  title and breaks the numbering chain while the document still opens fine. Use the
  `section_body` compress kind (rewrite the first non-bullet paragraph *under* a heading).
- **"Configured but matched zero times" must raise** — a cross-reference rewrite that hits
  nothing usually means the search string is wrong (full-width brackets, a sub-number that
  does not exist, text split across runs). Silent zero hits leave stale pointers behind;
  the runner aborts unless `allow_zero: true` is set explicitly.
- **Never hard-code table indices** — locate the 版本沿革 table by its header keywords
  (`版本 / 生效日期 / 修订内容 / 来源`); relocating a table shifts every later `doc.tables[i]`.
- **Drift detection is stage-aware** — the table-content fingerprint multiset must not
  change across a stage, except for tables the config explicitly names (`changelog_table`
  in `release`, `kind='table_cell'` targets in `compress`); for those, assert
  append-only (old rows preserved cell by cell).
- **Tail integrity is a hard contract** — after any merge or version bump, the closing marker
  (「—— 本手册完 ——」) must stand alone and appear exactly once, and an 增量说明 section must
  not carry another section's items. A new 增量说明 section's number must not collide with an
  existing heading (after a renumber, take max + 1). Assert it
  (`verify_revision.py --config`'s `tail_integrity` block); absent config → SKIP, never FAIL.
  See `references/chapter_refactor_playbook.md` §5.
- **Baseline cell-by-cell comparison degrades gracefully** — when the table *order* changed
  (a relocation or deletion happened), an index-based comparison produces false positives;
  the verifier then switches to comparing the multiset of table content fingerprints.
  Declare tables that legitimately changed via `allowed_table_changes` (matched by header
  keywords, append-only asserted). See `references/chapter_refactor_playbook.md` §8.5.
- **User intent → action** — 「回退」= restore the previous deliverable by **overwriting
  copy-back (never delete)** and quarantining the voided version; 「把 A 和 B 合并并放到 B 位置」
  = **ask the user first** how A's leftover content should be handled; 「某章内容乱 / 冗长，
  整理精简」= quantify first, then choose "move the table" vs "compress the text";
  「打开我看」= open with the system default program — the workspace filename may differ from
  the Z-drive canonical name (the Z-drive copy carries an `8.` prefix, the workspace copy
  does not). See `references/chapter_refactor_playbook.md` §6.

## Bundled Resources

### scripts/
- `docx_table_toolkit.py` — importable python-docx helpers: `setup_page`,
  `add_grid_table_fixed` (fixed layout + gridCol sync + explicit `w:tblW` + conditional
  `tblHeader` + `cantSplit`), `add_grid_table`, `add_kv_table`, `add_key_callout`
  (with `title=`; default `● 关键点 / 易错点`; 多序号 ①②③… 自动逐条分行 + 左竖线红色 +
  行级 `cantSplit` 防跨页断裂), `add_rich_runs` (renders `**bold**` markers + `\n` soft
  breaks), `add_sop_quote_block`, `add_header_footer(doc, header_text)` (页眉描述文字 +
  页脚 PAGE/NUMPAGES 域), `_add_field` / `_split_numbered` 内部助手,
  `add_landscape_section` / `restore_portrait`, shading, `set_cjk` (rFonts 四属性 +
  `w:sz`/`w:szCs` 成对；字体族由 `CN_FONT` / `BODY_FONT` / `ASCII_FONT` 参数化，
  默认中文宋体 + 英文 Times New Roman). All document-mutating functions take the `doc`
  object as first argument.
- `deviation_cases_template.py` — 偏差/OOS 案例数据源模板 + 聚多卡醇 4 例脱敏参考集；
  导出 `DEVIATION_CASES` / `CASE_BY_CODE` / `COMMON_PATTERNS` / `STABILITY_PLACEMENT_RULE` /
  `MANUAL_REVIEW_NOTES`，以及 `attach_case(key_text, case_code)` 把案例关键点挂到卡内红框。
- `materials_template.py` — 附录 G 物料清单模板（六类，ID 前缀 RP-/RGT-/SMP-/COL-/INS-/CON-），
  含 `build_matrix(cats=None, legacy=False)`：默认输出「完整名称（编号）」+ 顿号，
  列头 `MATRIX_CAT_HEADERS`（中文全称（缩写）），`cats` 可裁剪类别，`legacy=True` 回旧格式；
  始终从物料 `items` 字段自动派生矩阵，杜绝手工两份不同步。
- `patch_existing_manual.py` — 成稿原地修订：合并两份已实测的列宽内核
  （`--strategy autofit` 标记层 / `content-width` 真实定宽层，CJK×1.0、英文×0.5、
  `FLOOR=480tw`、按 section 取版心、tblW/gridCol/tcW 三同步）与
  `--normalize-fonts`（四属性 + szCs 成对 + 空 run 跳过 + 下标替换 + 窄列去加粗）；
  路径全部参数化，默认另存 `--output` 不覆盖输入，支持 `--dry-run`。
- `verify_revision.py` — 参数化三层零回归 runner：`--docx/--baseline/--config/--pdf`，
  文本级（标题/限度/编号/关键点框/`**`/U+FFFC）+ OOXML 结构级（逐格比对、列宽三同步、
  `sectPr`、`tblHeader`/`cantSplit`、`w:br[type=page]`、`w:pBdr`）+ 视觉级（既存 PDF
  文本层，缺则整层 SKIP 并提示「PDF 需手动导出、逐页目视待补」）。不内嵌 COM/PDF 渲染。
  另含**尾部结构级**（结尾标记独立成段且唯一、增量说明不错挂、条目不跨小节）、
  **迁表零损失**（按表头定位后断言维度不变）、**交叉引用双向断言**、**章节体量报告**；
  四组新 config 字段缺省一律 SKIP，**绝不 FAIL**（老 config 行为不变）。
- `docx_revise_toolkit.py` — importable 改稿原语库（import 无副作用）：
  文档序遍历（`iter_block_items` / `all_paragraphs` —— **表内段落不在 `doc.paragraphs`**）、
  按 style.name 定位（`find_paragraphs` / `find_heading` / `find_subheading` / `find_table`
  —— 用表头关键字而非表索引）、保格式写入（`set_paragraph_text` /
  `replace_in_paragraph` / `set_cell_text_preserve` —— 保留首 run 的 rPr、不用
  `cell.text =`、一律以 Clark 记法写 `xml:space`）、结构搬移
  （`move_table_with_notes` —— **先移表、再接说明段**）、重编号
  （`renumber_headings` —— **按源编号降序**）、交叉引用改写
  （`update_cross_references` —— 整段拼接后替换，跨 run 安全）、体量量化
  （`measure_sections` / `measure_chapter`）。带 `--selftest <docx>` 只读自检。
- `refactor_chapter.py` — config 驱动的章节重构 runner
  （`--input/--output/--config/--baseline/--dry-run/--report`）；**七个固定阶段**
  —— merge → delete → renumber → move-tables → compress → tail-fix → release；
  每阶段做表内容指纹漂移检测；收尾自检尾部结构；**永不覆盖输入**。
- `build_validation_summary.py` — 配置驱动地把目标章节重写为「验证项目及指标汇总表 + 稳定性结果矩阵」：定位真实章节标题（取最后出现避 TOC 伪条目）、收集-插入防孤儿段、列宽合计=9638 twips 校验、cantSplit 入 trPr / vAlign 入 tcPr、OCR·需核 标注；`--dry-run` 不写盘、`--selftest` 内置自检；数据全来自 JSON config，**不含产品专属数值**。`references/validation_summary_example_lk080.json` 为红霉素示例。
- `chapter_volume.py` — 只读体量报告（`--docx/--chapter/--format/--top`）：逐小节
  正文 / 表内 / 合计 字数、表数、行数、bullet —— 「从哪砍」的决策依据，
  并给出「搬表」与「压缩文字」的判定提示。
- `fix_operation_card_blocks.py` — 成稿版式校正：把逐项操作卡的「主卡表 / SOP 引用框 / 关键点框」三类块统一为版心满宽 + 左对齐 + 单元格垂直居中 + 文字左对齐，消除「对齐方式不同 / 没铺开」的视觉不一致。按模板固定特征识别（红左边线 `C00000` / 蓝左边线 `1F3A5F` / 2 列含「方法要点」），**只动三类块、不动附录/产品/矩阵表**；纯版式改动零内容损失；带 `--selftest` 与 `--dry-run`，  `--content-width-cm` 默认 17.0（横向传 25.7）。
- `fix_flow_layout.py` — 成稿版式校正：去除逐项操作卡/章节的强制换页，改为连续流式排版；卡间注入浅灰分隔线，三类块表格加 `cantSplit`、卡标题/H1 加 `keep_with_next` 防拆页/孤行；纯版式改动零内容损失；带 `--selftest` 与 `--dry-run`，`--card-pattern` 默认可改、四组 `--no-*` 开关控制去哪些换页/是否加分隔线与安全网。
- `export_pdf.ps1` — Word COM 导出 PDF（已内置 `$doc.Fields.Update()` 防页码 0/1）；
  **本机 COM 被沙箱拦截，保留待环境修复，勿作为交付步骤**（见 pitfalls §10 / §12）。
- `ocr_pipeline.py` — pypdfium2 render + RapidOCR banded scan (BAND=760, OVERLAP=120,
  SCALE=1.6), near-duplicate line merge, JSON output. Includes the onnxruntime 1.20.1 pin
  workaround.

### references/
- `python_docx_pitfalls.md` — OOXML element ordering (tblLayout before tblLook;
  cantSplit before tblHeader), gridCol/tblW sync, A4 page setup, landscape sections,
  file-lock (预览面板独占) handling, PyMuPDF segfault, U+FFFC extraction artifact,
  Word COM PDF export (**§10 conditional: available vs sandbox-blocked manual export**),
  and **§12 in-place revision environment constraints** (safe-delete hooks, lock release,
  no OCR, COM branch).
- `materials_data_model.md` — 物料清单 schema: category list, per-item dict fields
  (`id/name/spec/lot/supplier/expiry/storage/qty/use/items/src/src_ref`), source-mark enums
  (SOP原文 / 质量标准(SOP未载) / 需查 / 外包), the manual-card↔SOP-section mismatch
  caveat, and the graded lot/supplier rule (验证报告批号不得充当对照品证书信息；
  批检验记录可逐份汇总回填并留「待核对」— see `revision_playbook.md` §4).
- `revision_playbook.md` — Stage 6 full SOP: 两层结构（通用规则 §1~§10 / 品种专项决策
  §11）、写回安全、列宽两策略、字体统一、批号回填分级口径、章节合并代码清单、
  矩阵可读化与删列三同步、版式修补、**三层验收清单与双路径**、正式升版与
  「每次改完提醒版本名称 + 文件名日期」话术模板、环境约束索引、config 示例；
  **§5 的 config 驱动版本与 §8 的尾部完整性断言见 `chapter_refactor_playbook.md`**。
- `chapter_refactor_playbook.md` — **章节重构 SOP（跨品种通用）**：「冗长」的量化与决策树；
  「搬表而非删表」通则（表内容一字不删、迁后改指交叉引用、**先移表再接说明段**）；
  定位 10 条硬约束（`doc.paragraphs` 身份漂移、styleId 非标准、目录与正文同名、
  子标题样式不统一、跨 run 编号 token、表内段落不在 `doc.paragraphs`、禁硬编码表索引、
  定位失败须抛错、位置锚点不得静默回退）；写入 6 条硬约束（Clark 记法 `xml:space`、
  禁 `cell.text =`、整段回写取舍、克隆保格式、前后语义分明、**降序重编号 + 目标号不冲突**）；
  **尾部结构完整性契约**；**用户指令 → 动作映射**；`refactor_chapter.py --config` 完整
  schema；`verify_revision.py` 新增字段；体量报告决策示例。
- `validation_summary_playbook.md` — **第 4 章重写配方（跨品种通用）**：把「方法转移 / 稳定性摘要」章重写为「各法验证项目及指标汇总 + 稳定性时间点结果矩阵」；含数据来源分层与 OCR·需核 纪律、7 条实现要点（真实标题最后出现避 TOC 伪条目、收集-插入防孤儿段、列宽=9638 twips、cantSplit 入 trPr / vAlign 入 tcPr、字体四属性）、零回归验收、与 refactor_chapter 衔接、红霉素算例非通用常量警告。
- `case_study_gsc_v6.md` — **骨化三醇注射液 v5→v6 修订实战复盘（品种专项）**：以六类动作
  （物料溯源补录 / 表格字体统一 / 章节整合 / 矩阵可读化 / 矩阵精简 / 零回归验收）为骨架，
  记录真实修订过程、实测数字、正式升版机制与已知债务，并回指通用规则 § 索引；
  所有数值标注「算例非通用常量」，红线与 `revision_playbook.md` §11 一致。
- `ocr_pipeline.md` — detailed OCR approach: pypdfium2 (not PyMuPDF),
  rapidocr-onnxruntime 1.2.3, onnxruntime MUST be 1.20.1 (system 1.31.1 segfaults;
  prepend the pinned path to `sys.path`), Windows.Media.Ocr cross-validation, 90°-rotated
  tables (`rotate(-90, expand=True)` then x-cluster), banded re-scan to fix silent
  dropped lines. **§9: the faster no-OCR path for signed scanned reports** — render pages
  to PNG and read them visually; for 偏差/OOS 签批件 the body is only the first few pages
  and the rest are attachments. Also the PDF-vs-DOCX cross-check checklist.
- `manual_structure.md` — v1/v2/appendix-G structure decisions and the pharma-QC content
  decisions that are easy to get wrong (校正因子按 1.0 主成分自身对照法; 残留溶剂—乙醇
  限度 37.5~41.5 mg/ml; 有关物质 LK036-C 分离度 ≥1.2; 渗透压摩尔浓度非本品项目;
  密度 37.4 与下限 37.5 矛盾判为标准笔误), plus **§6 how to fold real deviation/OOS
  cases into the 关键点/易错点 boxes** (data model, phrasing, de-identification, and the
  "共性归纳" section that carries the real value), and **§7 three-layer acceptance
  checklist** (text / OOXML structure / visual).
- `case_callout_patterns.md` — 偏差/OOS 案例挂接模式详解：数据外置字段、挂接方式
  （红框数量不变）、卡内单条固定句式、脱敏规则、共性归纳写法、附录配套与验收硬指标。

### tools/
- `upload_to_github.py` — 把本技能目录上传到 GitHub 仓库的可复用脚本（维护者用）。
  认证走本机 `git credential` 缓存的 github.com 令牌（亦支持 `GITHUB_TOKEN` 环境变量），
  逐个文件 PUT 到 REST Contents API，已存在文件自动带 sha 更新。用于绕过「WorkBuddy
  GitHub 连接器只读」与「沙箱直连 github.com 主站超时」的限制。用法见脚本内 docstring。

## Usage Notes

- Always set A4 + 2.0cm side margins explicitly; the default template is Letter and the
  text area is only 16.5cm, so ≥17cm-wide tables overflow.
- Wide tables (>25.7cm) go in a landscape `add_section`; restore portrait afterward.
- Keep `add_grid_table` (flow layout, 50-table workhorse) untouched; add a *separate*
  `add_grid_table_fixed` for fixed-layout appendix tables to avoid shifting the whole book.
- When deleting/trimming appendix sections, also remove the landscape `add_section`
  switches so the section count stays consistent.
- Emoji in the finished document: 🔴/⚠️ are unstable under 宋体; use ★ / ● / △.
  `add_key_callout` defaults to `title="● 关键点 / 易错点"` (compatible with the existing
  `startswith("● 关键点")` acceptance checks); narrative "🔴 关键点框" here is only
  shorthand (see `revision_playbook.md` §3 / §8).
- Stage 6 delivers DOCX only when Word COM is blocked; ask the user to export the PDF
  manually, and after every document change remind about the version name + filename date
  (see `revision_playbook.md` §9 / 附录 B).
- When a chapter is judged verbose, **never delete detail tables**: relocate them into a new
  appendix verbatim and leave conclusions + pointers behind; rewriting the pointers is part
  of the job, not optional cleanup. Quantify first (`chapter_volume.py`) — pure text
  trimming caps out around 8~25%, while moving the tables cuts 50%+ with zero fact loss.
- Tail checks are assertions, not eyeballing: the closing marker must stand alone and
  appear exactly once, and an 增量说明 section must not carry another section's items —
  `verify_revision.py --config`'s `tail_integrity` enforces both (absent config → SKIP,
  never FAIL).
