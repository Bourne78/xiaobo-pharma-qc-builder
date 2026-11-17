# -*- coding: utf-8 -*-
"""把 D.4 对照品信息并入 G.2 矩阵 RP- 列，并删除 D.4 整节。

用户决策（知情）：
  - 把 D.4 的批号 / 含量 / 贮藏 / 来源全部填入 G.2 矩阵 RP- 列对应物料条目
  - D.4 标题 + 表 + 警示段整节删除
风控（脚本自动执行，不可关）：
  - 填入内容保留「S03 P4 表 5 · 备料线索 · 开检须查当批证书原件（E-01）」标注
  - D.4 警示段里「禁止补入 SOP」的硬约束并入 E-01 依据栏与 G.1 口径段，不随表蒸发
  - 原 D.4 三条来源依据（S03 P4 表 5 / 规格效期未载 / S02 §3.3 未载批号）逐条保留

改动块（body 索引，改前 235 块）：
  [188] P  D.4 标题      -> 删
  [189] T  D.4 表        -> 删
  [190] P  D.4 警示段    -> 删（核心约束并入 [192] 与 [213][214]）
  [192] T  E-01 依据栏   -> 改写 R1C2（补入 SOP 禁令 + 线索标注）
  [213] P  G.1 口径段    -> 追加对照品线索说明
  [214] P  单一数据源段  -> 追加例外口径
  [218] T  矩阵表        -> RP 列 R2/R6/R14 三格写入批号等
  [219] P  矩阵说明段    -> 追加线索与证书原件说明

用法：python merge_d4_into_g2.py [--dry-run]
"""
import sys
import copy
import hashlib
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, r"C:\Users\bourne-office\.workbuddy\skills\xiaobo-pharma-qc-builder\scripts")
from docx_revise_toolkit import set_cell_multiline, set_paragraph_text  # noqa: E402

SRC = Path(r"E:\Workbuddy\文献整理\注射用乳糖酸红霉素_全项检验准备与关键点控制手册_v8_20261115.docx")
TMP = SRC.with_name("_tmp_merge_d4.docx")
DRY = "--dry-run" in sys.argv

# D.4 两条对照品数据（逐字取自 [189] 表）
SST_LINES = [
    "乳糖酸红霉素系统适用性对照品",
    "批号 LK070-RS-SST-20240701；含量 N/A；贮藏 2~8℃ 保存；来源 苏州朗科生物技术股份有限公司（S03 P4 表 5，备料线索）",
    "每 1ml 约含红霉素 A 4mg、LK070-H 10μg、LK070-A 10μg、LK070-B 20μg、红霉素 C 8μg、LK070-L 3μg、LK070-C 16μg、LK070-D 32μg、红霉素 B 5μg、LK070-F 0.5μg、LK070-E 2μg",
    "规格 / 效期未载，须查当批证书原件（E-01）；禁止补入 SOP",
]
ERY_LINES = [
    "红霉素 A 对照品",
    "批号 LK07002-231214-81RS；含量 93.8%；贮藏 密封，在干燥处保存；来源 宁夏启元药业有限公司（S03 P4 表 5，备料线索）",
    "规格 / 效期未载，须查当批证书原件（E-01）；禁止补入 SOP",
]
# R14 = 红霉素 A + 系统适用性对照品，两段之间空行分隔
R14_LINES = ERY_LINES + [""] + SST_LINES

# 改写文案
E01_NEW = ("S01 / S02 全文 0 命中；S03 P4 表 5 批号 / 含量 / 贮藏 / 来源已并入 G.2 矩阵 RP- 列，"
           "仅作备料线索、开检前须查当批证书原件，禁止据此补入 SOP 或质量标准（E-01）；"
           "红霉素 C 对照品 7.0 / 8.0 批次含量差异见附录 B.4")
G1_213_OLD_TAIL = "现场如需逐项核对，可自行在矩阵相应行加注。"
G1_213_NEW_TAIL = ("现场如需逐项核对，可自行在矩阵相应行加注。对照品批号 / 含量 / 贮藏 / 来源取自 S03 P4 表 5，"
                   "已列入 G.2 矩阵 RP- 列作备料线索，开检前须查当批证书原件、禁止补入 SOP（E-01）。")
G1_214_OLD = "凡现行 SOP 与质量标准均未载明的信息（对照品证书、柱寿命、仪器资产编号、部分器皿规格）一律填「需查」。"
G1_214_NEW = ("凡现行 SOP 与质量标准均未载明的信息（对照品证书、柱寿命、仪器资产编号、部分器皿规格）一律填「需查」；"
              "S03 P4 表 5 对照品线索为例外，已并入 G.2 矩阵并标注来源，开检前仍须查当批证书原件、禁止补入 SOP（E-01）。")
NOTE_219_OLD = "本矩阵由 items 字段自动生成；现场若实际使用清单外物料，须回写 items 后重新生成，不得手工改表。"
NOTE_219_NEW = ("本矩阵由 items 字段自动生成；现场若实际使用清单外物料，须回写 items 后重新生成，不得手工改表。"
                "对照品格内 S03 P4 表 5 批号 / 含量 / 贮藏 / 来源仅作备料线索，开检前须查当批证书原件、禁止补入 SOP（E-01）。")


def body_blocks(doc):
    return [c for c in doc.element.body.iterchildren()
            if c.tag in (qn("w:p"), qn("w:tbl"))]


def fp(el):
    from lxml import etree
    return hashlib.sha256(etree.tostring(el, method="c14n")).hexdigest()


def para_text(el):
    return Paragraph(el, None).text


def main():
    doc = Document(str(SRC))
    body = body_blocks(doc)
    print(f"改前块数={len(body)}")
    assert len(body) == 235, len(body)
    # 关键：指纹快照必须在任何改写之前取，否则改后 vs 改后恒相等
    orig_fp = [fp(x) for x in body]

    # ---------- 断言改前状态 ----------
    assert para_text(body[188]).startswith("D.4  对照品信息"), para_text(body[188])[:60]
    assert body[189].tag == qn("w:tbl")
    assert "禁止用该批号" in para_text(body[190]), para_text(body[190])[:80]
    assert para_text(body[213]).endswith(G1_213_OLD_TAIL), para_text(body[213])[-60:]
    assert G1_214_OLD in para_text(body[214])
    assert NOTE_219_OLD in para_text(body[219])

    # E-01 表：[192]
    e01 = Table(body[192], doc)
    assert e01.rows[1].cells[0].text.strip() == "E-01"
    assert "禁止直接补入 SOP" in e01.rows[1].cells[2].text

    # 矩阵：[218]
    mat = Table(body[218], doc)
    assert "检验项（卡号）" in mat.rows[0].cells[0].text
    assert "LK07002-231214-81RS" not in mat.rows[14].cells[1].text

    if DRY:
        print("[dry-run] 全部前置断言通过，拟改 5 处正文 + 删 3 块 + 写 3 格。未落盘。")
        return

    # ---------- 1) 矩阵 RP 列写入 ----------
    for ri, lines in ((2, SST_LINES), (6, SST_LINES), (14, R14_LINES)):
        cell = mat.rows[ri].cells[1]
        nbr = set_cell_multiline(cell, lines)
        print(f"  矩阵 R{ri}C1 写入 {len(lines)} 行 / {nbr} 个 br")

    # ---------- 2) E-01 依据栏改写（整格首 run 保格式） ----------
    c = e01.rows[1].cells[2]
    p = c.paragraphs[0]
    set_paragraph_text(p, E01_NEW)
    print("  E-01 R1C2 依据栏改写完成")

    # ---------- 3) G.1 两段 + 矩阵说明段改写 ----------
    for idx, old, new, label in (
        (213, para_text(body[213]), para_text(body[213]).replace(G1_213_OLD_TAIL, G1_213_NEW_TAIL), "G.1 口径段"),
        (214, para_text(body[214]), para_text(body[214]).replace(G1_214_OLD, G1_214_NEW), "单一数据源段"),
        (219, para_text(body[219]), para_text(body[219]).replace(NOTE_219_OLD, NOTE_219_NEW), "矩阵说明段"),
    ):
        assert old != new, f"{label} 未发生改写"
        set_paragraph_text(Paragraph(body[idx], doc), new)
        print(f"  {label} 改写完成")

    # ---------- 4) 删除 D.4 三块 ----------
    parent = doc.element.body
    for el in (body[188], body[189], body[190]):
        parent.remove(el)
    print("  D.4 标题 + 表 + 警示段已摘除")

    # ---------- 5) 零回归验收 ----------
    after = body_blocks(doc)
    assert len(after) == 232, f"改后块数={len(after)}，预期 232"
    # 索引映射：改后 j = i (<188) / i-3 (>190)
    def jmap(i):
        return i if i < 188 else i - 3
    CHANGED = {192, 213, 214, 218, 219}
    DELETED = {188, 189, 190}
    new_fp = [fp(x) for x in after]
    mism = []
    for i in range(len(body)):
        if i in DELETED:
            continue
        j = jmap(i)
        if orig_fp[i] != new_fp[j] and i not in CHANGED:
            mism.append((i, j))
    print(f"\n白名单外指纹不一致: {mism}")
    assert not mism, "白名单外有改动！"
    for i in sorted(CHANGED):
        assert orig_fp[i] != new_fp[jmap(i)], f"应改块 {i} 未变"
    print("白名单内 5 块全部变化，其余 227 块逐字未动 ✓")

    # 关键数值命中
    mat2 = [t for t in doc.tables if t.rows and "检验项（卡号）" in t.rows[0].cells[0].text][0]
    all_cell = "\n".join(c.text for r in mat2.rows for c in r.cells)
    for v in ("LK07002-231214-81RS", "93.8%", "宁夏启元药业有限公司",
              "LK070-RS-SST-20240701", "苏州朗科生物技术股份有限公司", "2~8℃ 保存"):
        assert v in all_cell, f"矩阵缺少关键值 {v}"
        print(f"  关键值命中: {v}")
    assert all_cell.count("禁止补入 SOP") == 4, all_cell.count("禁止补入 SOP")

    # D.4 残留零命中（正文 + 表格全扫）
    resid = []
    for el in after:
        if el.tag == qn("w:p"):
            t = Paragraph(el, doc).text
        else:
            t = " ".join(c.text for r in Table(el, doc).rows for c in r.cells)
        if "D.4  对照品信息" in t or "禁止用该批号" in t:
            resid.append(t[:60])
    print(f"  D.4 标题/旧警示残留: {resid}")
    assert not resid

    # 结构未塌
    assert len(doc.sections) == 3
    assert [s.orientation for s in doc.sections] == [0, 1, 0]
    assert len(mat2.rows) == 16 and len(mat2.columns) == 7
    print(f"  分节={[s.orientation for s in doc.sections]} 矩阵={len(mat2.rows)}x{len(mat2.columns)} ✓")

    doc.save(str(TMP))
    print(f"\n临时文件: {TMP} ({TMP.stat().st_size}B)，用同盘 os.replace 覆盖。")


if __name__ == "__main__":
    main()
