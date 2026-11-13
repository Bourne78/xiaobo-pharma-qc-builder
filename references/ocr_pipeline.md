# OCR 管线细节（中文扫描件 QC PDF）

扫描版 SOP / 质量标准的文本提取走这条管线。关键约束都是本环境实测踩出来的。

## 1. 工具选型
| 环节 | 工具 | 备注 |
|---|---|---|
| 光栅化 | `pypdfium2` | `pip install pypdfium2`，无系统依赖，可用 |
| 识别 | `rapidocr-onnxruntime` 1.2.3 | 自带 PP-OCRv3 中文模型 |
| 交叉校验 | `Windows.Media.Ocr`（zh-Hans-CN） | 第二识别引擎，互相印证 |

**禁用**：PyMuPDF（`fitz`）——本环境恒定 segfault（exit 139）。workbuddy-pdf CLI 也不可用（exit 127/139）。

## 2. ⚠️ onnxruntime 必须用 1.20.1
系统版 `onnxruntime 1.31.1` 在本机 import 即 exit 139。隔离装 1.20.1 并把其 site-packages
用 `sys.path.insert(0, ...)` 插到最前，再 `import rapidocr_onnxruntime`。
```bash
python -m venv .venv_ort120
.venv_ort120\Scripts\activate
pip install "onnxruntime==1.20.1" "rapidocr-onnxruntime==1.2.3"
```
运行时设 `ORT120_PATH=.venv_ort120\Lib\site-packages` 环境变量（`scripts/ocr_pipeline.py` 会读取）。

## 3. 分块重叠补扫（否则静默漏行）
整页直接 OCR 会**漏行**（实测把"3倍(1.5%)"漏成"0.6倍(0.3%)"）。必须分块扫描：
```
BAND = 760      # 每块高度(px)
OVERLAP = 120   # 块间重叠(px)
SCALE = 1.6     # 渲染放大，提升小字识别率
```
按 `step = BAND - OVERLAP` 滑动切带，逐带 OCR，再按 y 坐标去重（同 y±45px 内保留最长文本）。
聚多卡醇 SOP 590 行、质量标准经此补扫到 1029 行。

## 4. 90° 旋转横向表
质量标准 P3~P6 是 90° 旋转横向表，直接 OCR 会全乱。先 `rotate(-90, expand=True)` 转正，
再按 **x 坐标聚类**合并单元格文字。

## 5. 章节切分（sop_extract）
- 标题判定用**章节号白名单（约 88 个）+ BAD_IN_TITLE**（标题不得含数字/单位符号），否则
  "1）取本品……" 会被误判为标题。
- `TITLE_FIX` 修 OCR 退化标题（如 "3.2.3溶液颜色" → "3.2.3.2"）。
- 最终聚多卡醇 SOP 切出 87 节 / 508 行，43 个必需小节齐备。

## 6. conf 字段是字符串
`rapidocr` 返回的 `conf` 是**字符串**，须 `float(conf)` 后再格式化，否则 TypeError。

## 7. Windows.Media.Ocr 注意事项
- 走 `SoftwareBitmap` 纯内存路径（动态编译/反射被本机安全策略拦截）。
- `StorageFile.GetFileFromPathAsync` 对本机路径报 WinError -2147024735 → 改用内存位图。
- 仅作交叉校验，不作为主识别。

## 8. 输出
`ocr.json`：`{ "pages": [ { "page":1, "lines":[{"text","conf","y"}] } ] }`，
下游脚本按章节号/表格线聚类还原结构。

---

## 9. ★ 纸质签批扫描件的轻量读取路径（优先于 OCR）

**适用**：偏差处理报告 / OOS-OOT 调查报告 / 变更记录等**签批件**——这类文件往往只有
「正文几页 + 大量附图（色谱图/照片）」，页数多但有效信息密度低。

**实测对比**：重建 onnxruntime 1.20.1 隔离环境（下载 + 安装 + 配 PATH）动辄十几分钟且易失败；
而「渲染成 PNG → 视觉模型逐页目视」**零依赖、零安装**，对印刷体中文扫描件的读取准确率足够高。

```python
import pypdfium2 as pdfium
pdf = pdfium.PdfDocument(path)
# 只渲染需要的页，别整本渲染
for i in [0, 1, 2]:
    pdf[i].render(scale=2.0).to_pil().save(f"probe_p{i+1}.png")
```
然后用 **ViewImage** 逐张读图，query 里**明确要求**：
- 「逐字读出……」
- 「手写部分若不清晰请标注'手写不清晰'，不要猜测补全」

**页数判断技巧**：先只渲染首页 + 末页 + 中间页看结构。
签批件的正文往往在**前 3~5 页**结束，其余是附图附件（图表谱、照片、培训签名表）。
例：133 页的 OOS 报告正文只有 3 页；17 页的偏差报告正文只有 4 页 —— 读前几页即可拿全。

**PDF 与电子版 DOCX 必须交叉核对**（同一份报告常有扫描签批件 + 电子版两套）：
签批 PDF 常带有电子版没有的**勾选项、签署链、手写补充结论**，
而电子版可能有 PDF 漏掉的批号。核对清单：
- 最终处理决定栏勾了哪一项
- 手写补充口径（如「采用重新配制后重新检测的结果为最终报告数据」）
- QA 经理 / 质量负责人签署与日期
- Section E 追踪评估、Section F 关闭勾选
- 实施人 / 完成期限 / 验收人 / 验收日期四个字段的填写值
- **受影响批号数量**（PDF 常比 DOCX 多）
- 培训签名表（CAPA 闭环证据，一般在最后一页）

**差异处理原则**：涉及「最终决定 / 关闭 / 生效 / 责任人 / 批号」一律以**签批 PDF 为准**，
并在手册中标注来源页码；电子版仅作技术内容主体。

