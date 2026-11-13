# xiaobo-pharma-qc-builder

WorkBuddy 技能：从零散 QC 文档（SOP、方法验证报告、药品注册标准、检验报告）生成中文
《药品全项检验准备与关键点控制手册》。

> 这是技能本身的说明文档。如果你是要**把本技能推送到 GitHub**，请看
> [`README_PUSH.md`](./README_PUSH.md)；如果你在找**可复用的上传脚本**，见
> [`tools/upload_to_github.py`](./tools/upload_to_github.py)。

---

## 这个技能能干什么

- **扫描版中文 PDF 提取**：用 pypdfium2 渲染 + RapidOCR 做带中文的 OCR 流水线。
- **多阶段 python-docx 装配**：一页速览总览表 → 逐项操作卡（含关键点框；叙述用「🔴」，
  成品文档内实际写 `● 关键点 / 易错点`）→ 验证结论 → 附录 A~G。
- **固定布局表格**：正确的 OOXML 元素顺序、`gridCol`/`tblW` 同步、跨页表头重复、
  行级 `cantSplit` 防止红框跨页断裂。
- **页眉 / 页脚**：注入页眉描述文字 + 页脚「第 X 页 / 共 Y 页」（PAGE/NUMPAGES 域，
  Word COM 导出前 `Fields.Update()` 防页码 0/1；本机 COM 被沙箱拦截，PDF 需手动导出）。
- **成稿修订与正式升版**：对既有手册按内容重算列宽、统一表格字体、回填补录物料批号/供应商、
  章节合并重编号、矩阵可读化与删列、零回归三层验收、版本号与日期交付收尾
  （见 `references/revision_playbook.md`）。
- **红框多序号分行**：`add_key_callout()` 内部按带圈数字 ① ② ③… 自动逐条分行，
  挂了偏差案例的 ⑥ 也会独占一行，扫读无障碍。
- **偏差 / OOS / CAPA 案例挂接**：把历史偏差调查报告的教训脱敏后挂到对应卡片的红框，
  数据源外置于独立模块，红框数量不变。
- **附录 G 检验物料清单**：六类物料（对照品 / 试剂试液 / 供试品 / 色谱柱 / 仪器 /
  消耗品器皿）+ 自动派生「检验项 × 物料」矩阵。

---

## 目录结构

| 路径 | 作用 |
|---|---|
| `SKILL.md` | 技能入口与全流程说明（WorkBuddy 据此加载） |
| `README.md` | 本文件：技能简介与用法 |
| `README_PUSH.md` | 把技能推送到 GitHub 的人工指引 |
| `scripts/docx_table_toolkit.py` | python-docx 助手集（表格 / 红框 / 页眉页脚 / 富文本；默认中文宋体+英文 TNR） |
| `scripts/deviation_cases_template.py` | 偏差/OOS 案例数据源模板 + 脱敏参考集 |
| `scripts/materials_template.py` | 物料清单模板 + `build_matrix()`（完整名称（编号）/ 中文化列头 / 类别可裁剪） |
| `scripts/patch_existing_manual.py` | 成稿原地修订：按内容重算列宽、字体归一（`--strategy` / `--font-mode`） |
| `scripts/verify_revision.py` | 参数化三层零回归验收 runner（`--config` 驱动） |
| `scripts/export_pdf.ps1` | Word COM 导出 PDF（含 `Fields.Update()`）；本机 COM 被拦截，保留待环境修复 |
| `scripts/ocr_pipeline.py` | 扫描 PDF OCR 流水线 |
| `references/revision_playbook.md` | Stage 6 成稿修订、零回归与正式升版 SOP（通用规则 + 品种专项两层） |
| `references/*.md` | 避坑集与数据模型（OOXML / 物料模型 / 案例挂接模式 / OCR / 手册结构） |
| `tools/upload_to_github.py` | 可复用上传脚本（维护者用） |

---

## 在 WorkBuddy 中调用

- **显式点名**（最稳）：
  ```
  @skill:xiaobo-pharma-qc-builder 把这份 STP/SOP 整理成检验手册
  ```
- **自然语言**：「帮我把这批 SOP 和验证报告梳理成一份检验准备手册」
- **触发词**：检验手册 / 检验物料清单 / 备料核对清单 / SOP 嵌入手册 / 偏差案例挂接 /
  关键点框 / 页眉页脚 / 成稿修订 / 修订升版 / 列宽重算 / 批号回填 / 供应商回填 /
  矩阵全名 / 版本沿革 / 增量说明 / 零回归验收

> 本技能为**用户级**，装在 `~/.workbuddy/skills/` 下，跨所有项目可用。

---

## 把技能安装到 WorkBuddy（使用者）

将本仓库 clone 后，把整个目录（要求 `SKILL.md` 在根）复制为：

```
~/.workbuddy/skills/xiaobo-pharma-qc-builder/
```

重启 / 刷新 WorkBuddy 即可在对话中通过上面的方式调用。

---

## 脱敏与许可说明

技能内的示例数据已脱敏（真实姓名 → 角色，如「QC 主管 / 检测人员 / 新入职人员」），
但保留真实产品名（聚多卡醇注射液）与文件编号（STP-SP-036-5001）用于演示。
仓库默认公开，若需保密请在 GitHub 网页将仓库改为 Private。
