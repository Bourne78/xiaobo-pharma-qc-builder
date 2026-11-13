# 成稿修订、正式升版与交付收尾（跨品种通用）

Stage 6 的操作手册。`SKILL.md` Stage 0~5 讲「从零散 QC 源文件从零生成手册」；本文件只讲  
**对一份已经成稿的 .docx 做修订、验收、升版**——表格重排、物料溯源补录、章节整合、  
矩阵可读化、版本沿革、交付收尾。

本文件分两层，**读之前先分清**：

- **第一层（§1~§10）＝通用规则**：换任何一个品种都适用，可直接照做。
- **第二层（§11）＝品种专项决策**：某一次手册的当次决定与单次实测数字，**禁止**升格为通用规则。

> 复用/指向：OOXML 元素顺序与 gridCol/tblW 同步见 `python_docx_pitfalls.md` §2~§4；  
> 六类物料与 items 单一数据源见 `materials_data_model.md` §1、§5；  
> 文本级与视觉级验收见 `manual_structure.md` §7；无 OCR 视觉读图见 `ocr_pipeline.md` §9。  
> 完整引用索引见本文附录 A。

---

# 第一层 · 通用规则

## 1. 总流程、红线与写回安全

**顺序**（不要跳步）：

```
只读基线 → 带时间戳完整备份 → 标题锚定定位 → 结构/内容修订 → 版式修补
→ 字体归一 → 三层自动验收 → 逐页目视（环境允许时）→ 正式升版交付 → 收尾提醒
```

**三条红线**：

1. `items` 单一数据源不变：矩阵永远由物料 `items` 字段派生，绝不手写第二份清单。
2. 历史记述不就地改写：旧版增量说明、版本沿革表旧行只在升版时**新增**一节/一行。
3. 正式升版用**新文件名**，不覆盖上一版成品。

**写回安全**：

- 动手前记录「输入 docx、上一版成品、构建脚本」的哈希基线；整目录 `cp -a` 备份。
- 修订脚本一律**另存 `--output` 新文件、不覆盖输入**；每批 ≤10 个文件、逐批校验、异常即停。
- 区分两个不同的输出参数：生成期用环境变量 `OUT_OVERRIDE`（`python_docx_pitfalls.md` §9），  
  修订期用 `--output`；**别混为一谈**。
- 目标 docx 被预览面板独占锁时 `doc.save()` 抛 `PermissionError`：先另存副本，待锁释放  
  （必要时重启 WorkBuddy 释放进程锁）再覆盖；判断依据是目录里的 `~$xxx.docx` 锁文件。
- 本机 safe-delete 钩子拦截 `rm` / `Remove-Item` / `os.remove` / 跨盘 `move`；同盘  
  `os.rename` 不受影响。回滚用**覆盖式拷回**，不要用删除。
- 覆盖后做 MD5 / 字节一致性校验。

> 复用/指向：`python_docx_pitfalls.md` §9（预览锁）、§12（原地修订环境约束）。

## 2. 表格列宽：两层策略，互补不可互替

| 策略                                       | 做什么                                                                         | 何时用                   |
| ---------------------------------------- | --------------------------------------------------------------------------- | --------------------- |
| `--strategy autofit`（标记层）                | `tblLayout=autofit`、`tblW type=auto w=0`、每格 `tcW type=auto w=0`，把决定权交还 Word | 用户明确要 Word「根据内容调整」标记时 |
| `--strategy content-width`（**默认**，真实定宽层） | 实测每格文本宽度，写死 `tblLayout=fixed` + `tblW` + `gridCol` + 每格 `tcW`               | 需要确定性列宽、宽表不溢出、零回归可断言时 |

已知差异：`autofit` 只是标记，轻量预览器可能仍按旧 `gridCol` 显示；要「内容少的列窄、  
内容多的列宽」必须用 `content-width` 层。

**content-width 算法链**（`scripts/patch_existing_manual.py`）：

```
char_width_twips(ch, sz)      CJK(W/F)=1.0、模糊(A)=0.6、其余=0.5；×字号pt×20 (1pt=20tw)
cell_ideal_width_twips(tc)    逐 run 读 w:sz；遇 w:br / \n 断行取最长行；w:tab 按一个空格
cell_pad_twips(tc)            默认左右各 115tw（读 tcMar 覆盖）
get_content_width(doc, tbl)   按表所在 section 取 pgSz@w 与 pgMar；
                              横向 pgSz@w 已是 16838tw，直接算得 14570tw=25.70cm，
                              **勿按 orient 再互换 w/h**
compute_col_widths(tbl)        处理 gridSpan（跨列约束按 s 均摊）与 vMerge 续格
resize_table(...)             MAX_COL=版心×0.6、FLOOR=480tw；
                              超版心时**只压缩超出下限的部分**，最多 10 轮；极端才整体缩放
```

OOXML 写入顺序：`ensure_child` 用 `addprevious` 插到正确后继元素之前，勿用 `append()`；  
`tblLayout=fixed` 必须在 `tblLook` 之前，`cantSplit` 在 `tblHeader` 之前。同步写  
`tblW`(`type=dxa`, `w=合计`)、`tblGrid/gridCol`、每行每格 `tcW`——**三者缺一不可**。

CLI：`python patch_existing_manual.py --input X.docx --output Y.docx --strategy content-width [--max-col-ratio 0.6] [--floor-twips 480] [--max-rounds 10] [--dry-run]`

> 复用/指向：`python_docx_pitfalls.md` §2（gridCol/tblW 同步）、§3（元素顺序）、§1（A4 版心）；  
> `materials_data_model.md` §7（纵向列宽合计 17.00cm）。

## 3. 表格字体统一（save 前全 run 后处理）

**方法为规则、字体族为参数**：

| 模式              | ascii / hAnsi / cs | eastAsia | 适用              |
| --------------- | ------------------ | -------- | --------------- |
| `cjk-split`（默认） | Times New Roman    | 宋体       | 新生成 / 新修订的标准配方  |
| `all-song`      | 宋体                 | 宋体       | 兼容既有「全宋体最小改动」成品 |

规则（缺一条就会出现重影、豆腐块或行距被撑高）：

- `rPr/rFonts` **四属性全写**（`ascii` / `hAnsi` / `cs` / `eastAsia`）；`docx_table_toolkit.set_cjk`  
  已按此实现，`CN_FONT` / `BODY_FONT` / `ASCII_FONT` 可参数化。
- `w:sz` 与 `w:szCs` **成对**：用 `sz.addnext(OxmlElement("w:szCs"))`，**勿用 `rPr.append`**  
  （顺序错会破坏 OOXML）。
- 空 run（`not text.strip()`）**跳过**，否则会撑高行距。
- `Normal` 样式同步写四属性，否则新增 run 仍回退。
- 窄列（<5cm、7pt）**禁加粗**（易出重影）。
- Unicode 下标（MgSO₄ / V₁）先替换成 `MgSO4` / `V1`，宋体下会成豆腐块。
- `**加粗**` 只在 `add_rich_runs` / `set_cell_text` / `add_key_callout` 中解析；成品 `**` 残留必须为 0。
- **emoji**：宋体下 🔴 / ⚠️ 不稳，成品文档内统一 **★ / ● / △**；`add_key_callout` 默认  
  `title="● 关键点 / 易错点"`。技能叙述里的「🔴 关键点框」仅为指代，不是成品写法。  
  **兼容性**：既有验收脚本用 `startswith("● 关键点")` 计红框，新默认串仍以其开头，不会漏计；  
  目录假阳性用精确串 `● 关键点 / 易错点` 排除。旧稿可在调用处显式传 `title=` 保持原样。

CLI：`python patch_existing_manual.py --input X.docx --output Y.docx --normalize-fonts --font-mode cjk-split|all-song`

> 复用/指向：`python_docx_pitfalls.md` §5（CJK 双写）、`manual_structure.md` §5（富文本 `**`）。

## 4. 批检验记录的批号 / 供应商回填（分级口径）

**分级口径（不要一刀切填「需查」，也不要拿验证报告批号冒充证书信息）**：

| 来源            | 能否回填 `lot` / `supplier` | 要求                               |
| ------------- | ----------------------- | -------------------------------- |
| 方法学验证报告中的批号   | **禁止**当对照品证书信息          | 属验证用样品/试剂信息，来源不同不可互替             |
| 现行 SOP / 质量标准 | 优先，直接采信                 | 标 `SRC_SOP` / `SRC_STD` + 章节页码   |
| **批检验记录扫描件**  | **可逐份汇总回填**             | 属实际执行记录；每条标来源文件+页码，保留「待核对」       |
| 未签批草稿 / 合同    | 仅作背景，数值不采信              | 合同须判 `covers_product`，不符不得写「已闭环」 |

**视觉读取路径**（本机无 tesseract/OCR）：`pypdfium2` **250dpi 单页渲染** → `ViewImage` 视觉读取。  
禁用 2×2 低清拼图（会误读手写批号）、禁用 PyMuPDF（segfault）。用 `rec00~rec09` 稳定命名 +  
`rec_mapping.json` 保存「原文件名 → recNN」映射，防止中文括号 `split` 造成覆盖。  
定向读 5 类页：**封面 / 目录 / 标准规定页 / 变更历史 / 结论**；同一目录下不同时间点常是不同版本。

回填到 F.2.1 / F.2.2（或对应物料表）时：逐份汇总**所有**批号并入「名称/规格/关键属性」列，  
加「供应商」列；供应商多种写法（同一试剂不同来源）**以主表为准并标「待核对」**；  
留 `batch_lot_extract.json` 台账。

**效力顺序**：签批 PDF（签署链 / 受控章 / 生效日期）＞ 注册标准 ＞ 电子 DOCX 配套稿 ＞  
历史验证报告。判定看两份：签批件「修订历史」依据 + 记录首页「检验依据」编号。  
冲突以签批件**就地更正**，旧值留审计表。「名实不符」（文件名 ≠ 首页文件编号/版本号）按内页定编号，  
B-xx 就地更正 + ★高待 QA。

> 复用/指向：`ocr_pipeline.md` §9（无 OCR 视觉读图 + PDF/DOCX 交叉核对）、  
> `manual_structure.md` §3（数值冲突就地更正 + 审计）、`case_callout_patterns.md` §四（签批件为准）。

## 5. 章节合并与重编号（可复制代码清单）

章节结构改动**登记为品种级决定**（见 §11），主结构模板不因此改写。技术上是三件事：

**① 标题锚定定位**：表格搬移会让 `doc.tables[i]` 索引整体前移，一律按标题文本定位，  
不要缓存表索引。注意目录与章头可能出现同名标题（如「1 产品基本信息」），用 `occurrence` / `after` 消歧。

**② 保留格式的文本替换**：

```python
def set_para_text(p_el, new_text):
    """整段替换，保留首 run 的 rPr（字体/字号/加粗不丢）。"""
    ts = [t for t in p_el.iter(qn('w:t'))]
    ts[0].text = new_text
    ts[0].set(qn('xml:space'), 'preserve')
    for t in ts[1:]:
        t.text = ""

def replace_in(p_el, old, new):
    """跨 run 片段替换：按字符区间切分，只在首个命中 run 插入新串。"""
    ts = [t for t in p_el.iter(qn('w:t'))]
    full = "".join(t.text or "" for t in ts)
    idx = full.find(old); end = idx + len(old); pos = 0; placed = False
    for t in ts:
        seg = t.text or ""; s, e = pos, pos + len(seg); pos = e
        if e <= idx or s >= end: continue
        b, a = seg[:max(s, idx)-s], seg[min(e, end)-s:]
        if not placed: t.text = b + new + a; placed = True
        else: t.text = b + a
    return True
```

**③ 克隆保格式**：新增「增量说明」标题段 / 概述段 / List Bullet 条目时，克隆上一版同类型段落；  
版本沿革表新增行时克隆上一版行。勿 `doc.add_paragraph()` 裸建（格式会漂）。

**④ 交叉引用清点**：重编号后清点全文引用（含目录、页眉、附录指向），并防数值串误替换  
（如 `0.5μg` / `0.81` 不得被当成章号）。

> 复用/指向：`python_docx_pitfalls.md` §6（body.iterchildren 保序）、§3（元素顺序）。

## 6. 检验项 × 物料矩阵：可读化、裁剪列与三同步

- 单元格：`完整名称（编号）`，例如 `XX 对照品（RP-01）`；同类多物料用顿号 `、`；无物料填 `—`。
- 列头：`中文全称（缩写）`，例如 `试剂、试液与标准溶液（RGT-）`；用  
  `MATRIX_CAT_HEADERS`（`MATRIX_CAT_SHORT` 已是它的兼容别名，**值同为全称**，勿再当简称列头用）。
- `build_matrix(cats=None, legacy=False)`：默认输出上述格式；`cats` 可裁剪类别（去掉某类即少一列）；  
  `legacy=True` 回旧「名称 编号 + 换行」格式仅供旧稿兼容。
- **删列必须三同步**：`tblW` 合计、`tblGrid/gridCol` 数量与宽度、每行每格 `tcW`；  
  表头 `tblHeader` 与数据行 `cantSplit` 保留；横向表列宽合计恰等于版心 14570tw。
- `items` 字段仍是唯一数据源，`MATERIALS` 六类与 ID 前缀不动；删矩阵某类**不等于**删物料表。

**验收误报规避**：「裸代号 0 命中」只扫矩阵**数据格 / 数据行**——全文扫会命中 F.2 物料表的  
编号列、总览表关键物料列与操作卡表；顿号计数排除列头（列头本身含中文顿号）。

> 复用/指向：`materials_data_model.md` §5（矩阵自动生成）、§7（列宽）、`SKILL.md` Stage 3。

## 7. 版式修补经验

- **章节起新页**：用 `add_h1_page()`（段前分页 `page_break_before`）替代裸  
  `doc.add_page_break()`；裸分页符落在页底会造出整页空白。
- **表格续页残片**：在该表所属 `add_h2` 段落设 `page_break_before=True`，整块移到新页。
- **孤立红框页**：`cantSplit` 会把整块红框推到下一页 → 精简该卡 SOP 引文（只留色谱条件 +  
  核心配制 + 公式 + 标准规定）、合并 `key` 的 `\n` 分段、案例外置附录 B。
- **取消卡间分页后**：卡间用段落下边框轻量分隔（  
  `pPr.insert_element_before(pBdr, 'w:shd', …, 'w:jc', …)`，本地 CT_PPr 不含 `w:pBdr` 时必须这样插），  
  整卡 `keepNext` 软约束；**链尾红框段落不赋值** keepNext，否则会把下一章一起吸住。
- 章节段前分页 / 横向附录小节分页保持不动。

> 复用/指向：`python_docx_pitfalls.md` §4（cantSplit/tblHeader 顺序）、`manual_structure.md` §5。

## 8. 三层验收与零回归（双路径）

对应 `scripts/verify_revision.py`（`--config` JSON 驱动，不引 PyYAML）。

**文本级**

- 章节标题全命中、关键数值/限度全命中、物料编号 0 缺失且六类前缀齐全。
- 关键点框数 = 卡片数（用**精确串** `● 关键点 / 易错点` 计数，防目录假阳性）。
- U+FFFC 伪影 0；`**` 残留 0（**段落 + 所有表格单元格一起扫**）。
- 零回归四联断言：旧汇总数字/旧标题 0 命中 + 新增编号全命中 + 主依据限度全命中且  
  红框数=卡片数 + 附录行数联动。

**OOXML 结构级**（PDF 导不出时的主力，技能此前缺这一层）

- 与 `--baseline` **逐格内容比对**，除 `config.allowed_cells` 白名单外零改动。
- `tblW` / `gridCol` / 每行每格 `tcW` 三者一致且 `tblLayout=fixed`；列宽合计 = 所在 section 版心。  
  **容差**：`tblW` 与 `gridCol` 合计常因逐列 `round()` 累计差几个 twips（实测 7 列差 9~11tw  
  ≈0.016cm），断言按列数自适应容差，勿用绝对相等；每格 `tcW` 仍逐格精确比对。  
  新脚本写表时先取整再用「取整后列宽之和」写 `tblW`，从源头消除该误差。
- `sectPr` 数量与内容一字未动；总表数、总段落数、关键表维度按 `config.expected_counts` 联动。
- `tblHeader` / `cantSplit` 保留；`w:br[type=page]` = 0；`w:pBdr` 分隔段计数 = 卡数 − 1。
- `rFonts` 四属性与 `w:sz` / `w:szCs` 成对齐全。

**视觉级（双路径）**

- (a) COM/PDF 可用 → pypdfium2 渲染 + `ViewImage` 逐页目视 + `orphan_kc` 双判据：  
  每页红框数 ≤ 该页卡标题数，**且**逐卡配对「第 k 个红框页 == 第 k 个卡标题页」；  
  含卡标题页数 ≤ 卡片数。
- (b) **本机 COM 被沙箱拦截（当前实况）→ 视觉级整层标 `SKIP`**。报告末尾固定打印  
  「PDF 需用户在 Word 手动导出、逐页目视待补」。**不得把结构级通过表述为版式已验收**。  
  脚本不内嵌 COM 导出、不内嵌 PDF 渲染；`--pdf` 仅对既存 PDF 文本层做可选断言，缺则 SKIP。

**需人工目视项**（写进交付说明，不假装自动化）：逐页无溢出/截断/乱码/豆腐块、  
横向页右边界在页内、红框不跨页撕裂、无孤立红框页、封面/页眉/沿革表/增量说明的版本号与日期一致、  
供应商多写法「待核对」标注是否恰当。

> 复用/指向：`manual_structure.md` §7（生成态三层清单）、`python_docx_pitfalls.md` §8（U+FFFC）、§10/§12。

## 9. 正式升版与交付收尾

1. **新文件名**，不覆盖上一版成品（历史 v1~v5 成品原样保留）。
2. **就地更新**（这些是「当前版本」信息，允许改）：文件名、封面副标题、版本栏、编制日期、  
   封面摘要段、**共享页眉 part**（多节常共用同一 header part，按元素 id 去重，改 1 处全生效）。
3. **新增**（不改写历史）：版本沿革表克隆上一版行 + 新增一行；新增「增量说明」一节，  
   克隆上一版模板段的标题/概述/List Bullet 格式。
4. **历史记述不就地改写**：旧增量说明各节、沿革表旧行一律保留。
5. **每次改完文档后，最终回复必须主动提醒更新「版本名称」+「文件名中的日期」**。  
   提醒要具体列出实际写着版本号/日期的位置：文件名、封面副标题/版本栏/日期栏、各节页眉、  
   版本沿革表、增量说明；并给出**当前值**与**建议值**（当天 `YYYYMMDD`）。  
   **提醒 ≠ 擅自改名**：升哪个版本号、是否新增增量说明由用户确认后再动；改名牵动活动预览锁、  
   备份引用、生成脚本路径与对外引用。话术模板见附录 B。
6. 本机只交付 DOCX；**PDF 由用户在 Word 手动导出**（或环境修复后再用 `export_pdf.ps1`）。  
   PDF 产出后再补跑视觉级验收。

> 复用/指向：`python_docx_pitfalls.md` §9（写回锁）、§10（COM 分支）、§12（环境约束）。

## 10. 环境约束（索引式，不重复展开）

| 事项       | 结论                                                                                     | 详见                            |
| -------- | -------------------------------------------------------------------------------------- | ----------------------------- |
| 删除/移动    | safe-delete 钩子拦截 `rm` / `Remove-Item` / `os.remove` / 跨盘 move；同盘 `os.rename` 可用        | `python_docx_pitfalls.md` §12 |
| 文件锁      | 预览面板独占 → `PermissionError`；另存副本待锁释放                                                    | §9 / §12                      |
| OCR      | 本机无 tesseract；纯图片 PDF 用 pypdfium2 渲染 + 视觉读取；禁 PyMuPDF                                  | `ocr_pipeline.md` §9 / §12    |
| 旧 `.doc` | Word COM `Documents.Open → Content.Text` 写 UTF-8 txt（勿 SaveAs2）+ `replace("\x07"," ")` | §12                           |
| COM 导出   | 按环境分支：可用则 `export_pdf.ps1` + `Fields.Update()`；被拦截则手动导出                                | §10 / §10b / §12              |

---

# 第二层 · 品种专项决策（勿通用化）

以下均为**某一次手册的当次决定**，记录供理解上下文，不得写成跨品种默认结构。


- 骨化三醇注射液 v5→v6 的六类动作：物料溯源补录 / 表格字体统一 / 章节整合 /
  矩阵可读化 / 矩阵精简 / 零回归验收。
- **F.3 是「附录 G 顺延为 F」后的品种编号**；其他品种可能是 G.4 或别的编号。
- **删「供试品与样品制备（SMP-）」列**是当次「矩阵精简」决策；六类物料模型与供试品表
  本身完好保留，技能模板不删类。
- **第 0 章并入第 6 章、`0.1~0.8 → 6.5.1~6.5.8`** 是品种级结构决定，
  `SKILL.md` Overview 的「0) 编制说明/溯源」独立章节结构不因此改动。
- 12 份批检验记录、某几份不含内毒素/无菌/热原项 → 对应物料无来源可提，留「需查」，
  这不是通用规则。
- 单次实测数字（**算例，非通用常量**）：横向矩阵列宽 `[1900,2020,2590,1930,3130,3000]`tw
  合计 `14570`tw；矩阵 `18×6`；段落 `240→248`、表 `76` 不变、3 个 `sectPr` 不变、
  其余 74 表逐格零改动。通用脚本**不得**把这些写成默认值。

**已知 TODO / 债务（仅记录，不改技能脚本内容）**

- v6 成品的表格字体是「全宋体最小改动」妥协；中文宋体 + 英文 TNR 的分离版尚未在成品验证。
  下次生成可用 `--font-mode cjk-split` 并目视确认。
- 项目工作区的 `build_gsc_v5.py` / `data_gsc_v5.py` 仍是 v5 代号版，**重跑会回退**；
  这属于工作区债务、与技能目录无关，执行期禁止重跑。

---

## 附录 A 引用索引表（避免重复）

| 本文件 | 现有资料 |
|---|---|
| §1 写回安全 | `python_docx_pitfalls.md` §9 / §12 |
| §2 列宽 | `python_docx_pitfalls.md` §1~§4；`materials_data_model.md` §7 |
| §3 字体 | `python_docx_pitfalls.md` §5；`manual_structure.md` §5 |
| §4 批号 / 效力 | `ocr_pipeline.md` §9；`manual_structure.md` §3；`case_callout_patterns.md` §四 |
| §5 章节改写 | `python_docx_pitfalls.md` §3 / §6 |
| §6 矩阵 | `materials_data_model.md` §5；`SKILL.md` Stage 3 |
| §7 版式 | `python_docx_pitfalls.md` §4；`manual_structure.md` §5 |
| §8 验收 | `manual_structure.md` §7；`python_docx_pitfalls.md` §8 |
| §9 升版 | `python_docx_pitfalls.md` §9 / §10 |
| §10 环境 | `python_docx_pitfalls.md` §12；`ocr_pipeline.md` §9 |

## 附录 B 交付收尾提醒话术模板

> 本次改动已完成，按约定提醒更新「版本名称」与「文件名中的日期」：
> - 文件名：当前 `{当前文件名}`，建议 `{品种}_…_v{N+1}_{YYYYMMDD}.docx`
> - 封面：副标题 / 版本栏 / 编制日期，当前值 `…`，建议 `…`
> - 各节页眉（共享 header part，改 1 处全生效）：当前 `…`
> - 版本沿革表：需**新增**一行（旧行不改写）
> - 增量说明：需**新增**一节（历史各节不改写）
>
> 升哪个版本号（v6 / v5.x）、是否新增增量说明由你确认后再动；我不会擅自改名。
> 本机 Word COM 被拦截，PDF 需你在 Word 手动导出后我再补跑逐页目视验收。

## 附录 C `verify_revision.py --config` 最小示例

```json
{
  "chapter_titles": ["1 产品基本信息", "6 参考文献与溯源说明"],
  "limits": ["90.0%～115.0%", "≥1.2"],
  "old_strings": ["手册 v5", "共识别 8 处冲突/存疑"],
  "new_strings": ["v6 物料溯源补录与附录整合升版"],
  "id_prefixes": ["RP", "RGT", "SMP", "COL", "INS", "CON"],
  "callout_exact": "● 关键点 / 易错点",
  "expected_counts": {"tables": 76, "paragraphs": 248, "callouts": 17},
  "matrix_table_title": "F.3  检验项 × 物料矩阵",
  "matrix_expected_dims": [18, 6],
  "check_fonts": true,
  "allowed_cells": []
}
```
所有阈值与字符串都来自 config；换品种只换 config，不改脚本。
