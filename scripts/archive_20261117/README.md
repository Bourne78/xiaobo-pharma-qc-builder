# scripts/archive_20261117/ — 历史一次性脚本归档（红霉素 v8）

本目录保存《注射用乳糖酸红霉素 全项检验准备与关键点控制手册 v8》在 2026-11-16~17
一轮改版的**当次执行脚本**，按用户确认「原样归档、不泛化」收存，文件内容逐字节保持原样。

> ## ⚠️ 历史执行记录，不可直接重跑
>
> 这些脚本是**算例**，不是通用工具：
> - 含产品硬编码路径 `E:\Workbuddy\文献整理\注射用乳糖酸红霉素_全项检验准备与关键点控制手册_v8_20261115.docx`
>   及同名 `TMP` 临时输出；在其他机器/品种上必然找不到输入。
> - 含绝对 `sys.path` 注入（指向本机 skill 的 `scripts/` 目录）。
> - 含 `body` 索引常量（如 `208/209/224/225/229`、`EXPECT_N=231`）、文本字面量锚点、
>   写死的编号（`2.2 / 2.2.1 / 2.2.2`）与表格维度（`16×6 / 16×7`）。
> - 部分脚本直接产出新版本 v8 成品，重跑会覆盖或回退已交付版本。
>
> 若确需复用：先复制到工作区，改 `SRC/TMP` 路径与锚点/索引，先跑 `--dry-run`
> 核对命中数，另存 `--output` 新文件，绝不覆盖输入。产品名、批号、编号均为当次算例，
> 不是通用常量。

## 归档清单

| 脚本 | 当次动作 |
|---|---|
| `move_g_into_ch2.py` | 把「附录 G 检验物料清单」整章搬进第 2 章一页速览之后：整块元素搬迁（载 `sectPr` 的空段随迁、横向节方向 `[纵,横,纵]` 不变）、目录条目单独重定位、标题 `附录 G→2.2 / G.1→2.2.1 / G.2→2.2.2` 重编号与样式调整、悬空引用段落+表格全扫描改写（历史修订表豁免）、改前逐块 c14n 指纹快照。 |
| `merge_g2_into_g3.py` | 合并①：把 G.2.1~G.2.6 六张物料明细表并入 G.3，顺延小节重编号。 |
| `merge_d4_into_g2.py` | 合并②：把 D.4 对照品信息并入 G.2「检验项 × 物料矩阵」，删 D.4 原文，清理 E-01 与说明段的悬空引用。 |
| `merge_g3_into_g1.py` | 合并③：把「G.3 首次开检前置条件（逐项打勾）」的 12 条清单并入 G.1 用途与口径，删 G.3 标题。 |
| `integrate_g_matrix.py` | 矩阵整合（属**版本产出**，非「三次合并」之一）：把物料明细总表的「名称 / 规格」嵌进矩阵单元格、删 65×6 总表、取消单元格代号、改写悬空引用，另存 v8。 |

## 与通用能力的关系

当次验证过的**通用做法**已沉淀在 skill 的正式脚本里，新任务应优先用这些，而非重跑本目录脚本：

- 逐块定位 / 保格式改写 / 跨 run 替换 / 表格单元格全扫描：
  `scripts/docx_revise_toolkit.py`（`replace_in_paragraph` / `set_cell_multiline` /
  `set_cell_text_preserve` / `move_block_after` / `renumber_headings` / `update_cross_references`）。
- 章节合并 / 重编号 / 迁表：`scripts/refactor_chapter.py`、`scripts/delete_chapter.py`
  及 `references/chapter_refactor_playbook.md`。
- 零回归验收：`scripts/verify_revision.py`、`references/revision_playbook.md` §8。
- 矩阵 `cell_mode="name_spec"`：`scripts/materials_template.py` + `set_cell_multiline`。
