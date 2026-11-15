# xiaobo-pharma-qc-builder

WorkBuddy 用户级技能：从零散 QC 文档（SOP、方法验证报告、药品注册标准、检验报告）生成中文《药品全项检验准备与关键点控制手册》。

## 能做什么
- **扫描版中文 PDF 提取**：pypdfium2 渲染 + RapidOCR 带中文 OCR 流水线
- **多阶段 python-docx 装配**：一页速览 → 逐项操作卡（含关键点框）→ 验证结论 → 附录 A~G
- **固定布局表格**：正确的 OOXML 元素顺序、跨页表头重复、行级 `cantSplit` 防红框跨页断裂
- **页眉 / 页脚**：PAGE/NUMPAGES 域（本机 Word COM 被沙箱拦截，PDF 需手动导出）
- **偏差 / OOS / CAPA 案例挂接**：历史教训脱敏后挂到对应卡片红框，红框数量不变
- **附录 G 检验物料清单**：六类物料（对照品 / 试剂试液 / 供试品 / 色谱柱 / 仪器 / 消耗品器皿）+ 自动派生「检验项 × 物料」矩阵
- **成稿修订与正式升版**：列宽重算、字体归一、批号/供应商回填、章节合并重编号、零回归三层验收、版本号与日期交付收尾
- **成稿版式校正（纯版式零内容损失）**：`fix_operation_card_blocks.py` 把操作卡三类块统一为版心满宽 + 左对齐；`fix_flow_layout.py` 去除强制换页改连续流式，卡间浅灰分隔线 + `cantSplit`/`keep_with_next` 安全网
- **第 4 章重写（验证摘要 + 稳定性矩阵）**：配置驱动把「方法转移 / 稳定性摘要」章重写为「各法验证项目及指标汇总表 + 稳定性时间点结果矩阵」（`build_validation_summary.py` + `validation_summary_playbook.md`），含 OCR·需核 标注纪律与孤儿段防护
- **删除章节/附录 + 悬空引用清理**：配置驱动删除指定章节/附录（`delete_chapter.py` + `delete_chapter_playbook.md`），段落+表格单元格全扫描定位引用、三分类改写、残留零命中/保留锚点/结尾唯一验收，吸取「只扫段落漏表格致回退」教训

## 目录
| 路径 | 作用 |
|---|---|
| `SKILL.md` | 技能入口与全流程说明（WorkBuddy 据此加载） |
| `scripts/` | python-docx 助手集 + 修订 / 验收 / 上传脚本（含 `fix_operation_card_blocks.py` 操作卡三类块版式统一、`fix_flow_layout.py` 连续流式排版去除强制换页、`build_validation_summary.py` 第4章重写、`delete_chapter.py` 删章节清引用） |
| `references/` | 避坑集、数据模型、修订 SOP、实战复盘、第 4 章重写配方（`validation_summary_playbook.md`）、删章节清引用 SOP（`delete_chapter_playbook.md` + 红霉素示例 JSON） |
| `tools/upload_to_github.py` | 把本技能推送到 GitHub 的脚本 |

## 在 WorkBuddy 中调用
- 显式点名：`@skill:xiaobo-pharma-qc-builder 把这份 STP/SOP 整理成检验手册`
- 自然语言：「帮我把这批 SOP 和验证报告梳理成一份检验准备手册」
- 触发词：检验手册 / 检验物料清单 / 备料核对清单 / 关键点框 / 成稿修订 / 修订升版 / 零回归验收 / 三类块没铺开 / 章节间空白太多 / 连续排版 …

## 安装（使用者）
将本仓库 clone 后，把整个目录（含根 `SKILL.md`）复制到：
```
~/.workbuddy/skills/xiaobo-pharma-qc-builder/
```
重启 / 刷新 WorkBuddy 即可在对话中调用。

> 仓库默认公开；技能内已脱敏（无真实姓名 / 内网信息），但含真实产品名与文件编号等内部 QC 信息，需保密请在 GitHub 网页将仓库改为 Private。
