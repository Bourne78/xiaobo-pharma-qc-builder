# -*- coding: utf-8 -*-
"""把附录 G 整章物理搬移进第 2 章（横向矩阵连同分节一起跟过去）。

用户确认口径：
  - 整章移动（非复制），原位置不保留
  - 横向 16×7 矩阵连同横向分节一起搬，保持横向排版
  - 编号：2.2 / 2.2.1 / 2.2.2（主标题降 Heading2，G.1→2.2.1 保 Normal，G.2→2.2.2 升 Heading3）
  - 目录条目移到「2 一页速览」后并改号
  - TBL64「各组分浓度见附录 G RP-02」→ 以对照品证书 / SOP 原件为准

分节原理：段落内 sectPr 定义"它之前"内容的属性。把 [208..229]（D+E+F+G，
含两个载节空段）整体移到 body[25] 后，节自动分为
  SEC1=A+B+D(纵) / SEC2=F(横) / SEC3=C+H(纵)，方向 [纵,横,纵] 不变。

用法： python move_g_into_ch2.py [--dry-run]
脚本只保存 TMP，不自动覆盖源文件；全部验收在独立复核脚本完成。
"""
import sys, os, hashlib
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table
from lxml import etree

SKILL_SCRIPTS = r"C:\Users\bourne-office\.workbuddy\skills\xiaobo-pharma-qc-builder\scripts"
sys.path.insert(0, SKILL_SCRIPTS)
from docx_revise_toolkit import replace_in_paragraph, replace_in_cell

SRC = Path(r"E:\Workbuddy\文献整理\注射用乳糖酸红霉素_全项检验准备与关键点控制手册_v8_20261115.docx")
TMP = Path(r"E:\Workbuddy\文献整理\_tmp_move_g_into_ch2.docx")

EXPECT_N = 231


def body_blocks(doc):
    return [c for c in doc.element.body.iterchildren() if c.tag in (qn('w:p'), qn('w:tbl'))]


def fp(el):
    return hashlib.sha256(etree.tostring(el, method='c14n')).hexdigest()


def txt(block, doc):
    if block.tag == qn('w:p'):
        return Paragraph(block, doc).text
    return ''


def main():
    dry = '--dry-run' in sys.argv
    doc = Document(str(SRC))
    blocks = body_blocks(doc)
    print(f"块数={len(blocks)}")
    assert len(blocks) == EXPECT_N, len(blocks)

    # ---------- 前置锚点 ----------
    assert '本手册完' in txt(blocks[230], doc)
    assert '备料前先按附录 G' in txt(blocks[25], doc)
    assert txt(blocks[208], doc).strip().startswith('5  附录 G'), txt(blocks[208], doc)[:40]
    assert txt(blocks[224], doc).strip().startswith('G.2'), txt(blocks[224], doc)[:40]
    assert txt(blocks[209], doc).strip().startswith('G.1'), txt(blocks[209], doc)[:40]
    for i in (225, 229):
        assert txt(blocks[i], doc).strip() == '', (i, txt(blocks[i], doc))
        assert blocks[i].find('.//'+qn('w:sectPr')) is not None, i
    # 矩阵表
    assert blocks[227].tag == qn('w:tbl')
    mat = Table(blocks[227], doc)
    assert len(mat.rows) == 16 and len(mat.columns) == 7
    # 速览表
    assert blocks[24].tag == qn('w:tbl')
    quick = Table(blocks[24], doc)
    assert len(quick.rows) == 16 and len(quick.columns) == 6
    # 目录条目
    assert txt(blocks[18], doc).strip() == '5  附录 G  检验物料清单（备料核对清单）★'
    assert txt(blocks[9], doc).strip().startswith('2  一页速览')

    # 改前指纹快照
    orig_fp = [fp(b) for b in blocks]

    if dry:
        print("[dry-run] 全部前置断言通过。计划：")
        print("  搬移 [208..229] 22 块到 body[25] 之后（含 2 个载节空段 + 横向矩阵）")
        print("  编号: 5 附录G→2.2(Heading2) / G.1→2.2.1(Normal) / G.2→2.2.2(Heading3)")
        print("  引用改写 6 处 + 目录条目移动到 2 一页速览后改号 2.2")
        print("  节方向预期保持 [纵,横,纵]")
        return

    # ---------- 3) 整体搬移 [208..229] → body[25] 后 ----------
    moving = blocks[208:230]
    assert len(moving) == 22
    anchor = blocks[25]
    for el in moving:
        anchor.addnext(el)
        anchor = el
    print("[3] 已搬移 22 块到 body[25] 之后")

    # 索引映射（搬移后、目录处理前）：用于定位改写元素（引用原块对象，不受位置影响）
    # 原 208..229 -> 26..47；原 26..207 -> 48..229；其余不变
    # 我们用块对象引用直接改写，无需重算索引

    # ---------- 4) 标题编号 + 样式 ----------
    p208 = Paragraph(blocks[208], doc)
    assert replace_in_paragraph(p208, '5  附录 G  ', '2.2  ', count=1) == 1
    p208.style = doc.styles['Heading 2']

    p209 = Paragraph(blocks[209], doc)
    assert replace_in_paragraph(p209, 'G.1', '2.2.1', count=1) == 1
    # 样式保持 Normal

    p224 = Paragraph(blocks[224], doc)
    assert replace_in_paragraph(p224, 'G.2', '2.2.2', count=1) == 1
    p224.style = doc.styles['Heading 3']
    print("[4] 标题编号改为 2.2 / 2.2.1 / 2.2.2，样式 H2/Normal/H3")

    # ---------- 5) 悬空引用改写（expected_hits 锁定） ----------
    # P25
    assert replace_in_paragraph(Paragraph(blocks[25], doc), '按附录 G 核对', '按 2.2 检验物料清单核对', count=1) == 1
    # P210: G.2 矩阵 -> 2.2.2 矩阵（1）
    assert replace_in_paragraph(Paragraph(blocks[210], doc), 'G.2 矩阵', '2.2.2 矩阵', count=1) == 1
    # P211: G.2 -> 2.2.2（同段 2 处，全部替换）
    n211 = replace_in_paragraph(Paragraph(blocks[211], doc), 'G.2', '2.2.2', count=0)
    assert n211 == 2, n211
    print(f"[5] P25/P210/P211 段落引用改写 OK (P211命中{n211})")

    # 表格单元格引用：用全表定位（避免依赖表格索引），expected_hits 锁定
    def cell_replace(needle, new, expect_cells):
        hits = 0
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    if needle in c.text:
                        got = replace_in_cell(c, needle, new, count=1)
                        assert got == 1, (needle, got, c.text[:60])
                        hits += 1
        assert hits == expect_cells, (needle, hits, expect_cells)
        return hits

    # TBL64: 各组分浓度见附录 G RP-02 -> 以证书/SOP 原件为准
    h1 = cell_replace('各组分浓度见附录 G RP-02',
                      '各组分浓度以系统适用性对照品证书 / 现行 SOP 原件为准', 1)
    # TBL189: G.2 矩阵 -> 2.2.2 矩阵
    h2 = cell_replace('G.2 矩阵', '2.2.2 矩阵', 1)
    print(f"[5] 表格单元格引用改写 OK (各组分浓度{h1}, G.2矩阵{h2})")
    # TBL182 历史记述豁免，不动

    # ---------- 6) 目录条目移动 + 改号 ----------
    toc_entry = blocks[18]
    toc_anchor = blocks[9]
    toc_anchor.addnext(toc_entry)
    assert replace_in_paragraph(Paragraph(toc_entry, doc), '5  附录 G  ', '2.2  ', count=1) == 1
    print("[6] 目录条目已移到「2 一页速览」后并改为 2.2")

    # ---------- 保存 TMP ----------
    doc.save(str(TMP))
    print(f"\n已保存: {TMP} size={os.path.getsize(TMP)}")
    print("请用独立复核脚本验收后再原子覆盖源文件。")


if __name__ == '__main__':
    main()
