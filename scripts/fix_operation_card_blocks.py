# -*- coding: utf-8 -*-
"""
fix_operation_card_blocks.py
===========================
成稿修订（Stage 6）通用版式校正：把「逐项操作卡」里的三类块
  - 主操作卡表（2 列：方法要点 / 操作步骤 … / 关键物料速查）
  - SOP 引用框（单格浅蓝底 F2F6FB + 左竖线 1F3A5F）
  - 关键点框（单格浅红底 FBE4E4 + 红色左竖线 C00000）
统一为「铺满版心 + 左对齐 + 单元格垂直居中 + 文字左对齐」，
消除「三类块对齐方式不同 / 没铺开」的视觉不一致。

识别依据（跨品种模板固定生成特征，非品种专用 —— 见 docx_table_toolkit.py
的 add_key_callout / add_sop_quote_block / 主卡表手写模板）：
  - 主卡表  ：cols==2 且 首列文本含「方法要点」/「操作步骤」/「关键物料速查」
  - SOP 框   ：cols==1 且 单元格左竖线色 == 1F3A5F
  - 关键点框 ：cols==1 且 单元格左竖线色 == C00000
  - 其余表（附录 / 产品信息 / 矩阵 / 普通表）一律不动。

参数：
  --doc PATH            输入 docx（不覆盖，写 --out）
  --out PATH            输出 docx
  --content-width-cm F  版心宽度 cm，纵向 17.0 / 横向 25.7（默认 17.0）
  --dry-run             只报告将改哪些块、改动前后宽度，不写文件
  --selftest            构造最小文档端到端自测，断言识别 + 设置正确

依赖：python-docx>=1.0。纯版式改动，不触碰任何单元格文本内容（零内容损失）。

经验来源：注射用乳糖酸红霉素 v7 巡检发现三类块宽度 8.02/10.2/12.62cm 且
卡表居中、另两者靠左，视觉不一致；统一为版心满宽 + 左对齐后复核纯文本零变化。
"""
import argparse
import sys
import os

from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CM2TW = 567.0  # 1cm = 567 twips
CARD_KEY_TOKENS = ("方法要点", "操作步骤", "关键物料速查")
KEY_EDGE_COLOR = "C00000"   # docx_table_toolkit.add_key_callout 左边线
SOP_EDGE_COLOR = "1F3A5F"   # docx_table_toolkit.add_sop_quote_block 左边线


# ---------------------------------------------------------------------
# 识别层
# ---------------------------------------------------------------------
def _cell_left_border_color(cell):
    """返回单元格左竖线色（大写 hex），无则 None。"""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        return None
    left = borders.find(qn('w:left'))
    if left is None:
        return None
    return (left.get(qn('w:color')) or '').upper()


def classify_table(table):
    """返回 'card' / 'sop' / 'key' / None（None = 非三类块，不处理）。"""
    cols = len(table.columns)
    if cols == 1:
        c0 = table.cell(0, 0)
        edge = _cell_left_border_color(c0)
        if edge == KEY_EDGE_COLOR:          # 关键点框（红左边线）
            return 'key'
        if edge == SOP_EDGE_COLOR:          # SOP 引用框（蓝左边线）
            return 'sop'
        return None
    if cols == 2:
        c0 = table.cell(0, 0)
        txt = c0.text or ''
        if any(tok in txt for tok in CARD_KEY_TOKENS):
            return 'card'
        return None
    return None


# ---------------------------------------------------------------------
# 写入层（OOXML 级，固定布局 + 列宽 + 对齐）
# ---------------------------------------------------------------------
def _table_width(table):
    tblW = table._tbl.tblPr.find(qn('w:tblW'))
    if tblW is None:
        return None
    v = tblW.get(qn('w:w'))
    return float(v) if v is not None else None


def _set_table_width(table, total_tw):
    """设置 tblW（dxa）。"""
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert(0, tblW)
    tblW.set(qn('w:w'), str(int(round(total_tw))))
    tblW.set(qn('w:type'), 'dxa')


def _set_fixed_layout(table):
    """w:tblLayout type=fixed。OOXML 顺序：tblLayout 必须在 tblLook 之前。"""
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(old)
    el = OxmlElement('w:tblLayout')
    el.set(qn('w:type'), 'fixed')
    tblPr.insert_element_before(el, 'w:tblCellMar', 'w:tblLook',
                                'w:tblCaption', 'w:tblDescription', 'w:tblPrChange')


def _set_grid_col_widths(table, widths_tw):
    """设置各列 gridCol 宽度 + 每个单元格 tcW（fixed 布局下 Word 以 tcW 为准）。"""
    grid = table._tbl.tblGrid
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths_tw):
        gc.set(qn('w:w'), str(int(round(w))))
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.insert(0, tcW)
            tcW.set(qn('w:w'), str(int(round(widths_tw[ci]))))
            tcW.set(qn('w:type'), 'dxa')


def _set_cells_left_align(table):
    """单元格垂直居中 + 段落左对齐。"""
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _card_old_col_widths(table):
    """读主卡表首行各列的旧 tcW；读不到则回退默认比例 2.42:10.2。"""
    try:
        old = []
        for c in table.rows[0].cells:
            tcW = c._tc.find(qn('w:tcW'))
            old.append(float(tcW.get(qn('w:w'))) if tcW is not None else None)
        if all(x is not None for x in old) and sum(old) > 0:
            return old
    except Exception:
        pass
    return [2.42 / 12.62, 10.2 / 12.62]


def fix_document(doc, content_width_cm=17.0):
    """原地修正三类块，返回改动记录 [(kind, old_tblw, new_tblw), ...]。"""
    total_tw = content_width_cm * CM2TW
    reports = []
    for t in doc.tables:
        kind = classify_table(t)
        if kind is None:
            continue
        cols = len(t.columns)
        if cols == 1:
            widths_tw = [total_tw]
        else:  # 主卡表 2 列，按比例缩放旧列宽到版心
            ratios = _card_old_col_widths(t)
            s = sum(ratios)
            widths_tw = [r / s * total_tw for r in ratios]
        old_tblw = _table_width(t)
        _set_table_width(t, total_tw)
        _set_grid_col_widths(t, widths_tw)
        _set_fixed_layout(t)
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        _set_cells_left_align(t)
        reports.append((kind, old_tblw, total_tw))
    return reports


# ---------------------------------------------------------------------
# CLI / selftest
# ---------------------------------------------------------------------
def _selftest():
    try:
        from docx_table_toolkit import (add_key_callout, add_sop_quote_block,
                                        add_grid_table_fixed, setup_page)
    except Exception as e:
        print('[selftest] 需要 docx_table_toolkit 在同目录：', e)
        return 1
    doc = Document()
    setup_page(doc.sections[0], landscape=False)
    # 1) 普通多列附录表（应不被识别）
    add_grid_table_fixed(doc, ['编号', '物料', '用途'], [['T01', 'xx', 'yy']], [3, 6, 8])
    # 2) SOP 引用框（应识别为 sop）
    add_sop_quote_block(doc, 'S02 §1', '原文内容…')
    # 3) 关键点框（应识别为 key）
    add_key_callout(doc, '易错点说明')
    # 4) 主卡表（2 列，含方法要点；应识别为 card）
    card = doc.add_table(rows=2, cols=2)
    card.style = 'Table Grid'
    card.cell(0, 0).text = '方法要点'
    card.cell(0, 1).text = '操作步骤'
    card.cell(1, 0).text = 'a'
    card.cell(1, 1).text = 'b'
    # 5) 普通 2 列表（首格不含关键词；应不被识别）
    plain = doc.add_table(rows=1, cols=2)
    plain.style = 'Table Grid'
    plain.cell(0, 0).text = '参数'
    plain.cell(0, 1).text = '数值'

    # 运行修正（纵向版心 17cm）
    rep = fix_document(doc, content_width_cm=17.0)
    kinds = [r[0] for r in rep]
    assert kinds.count('card') == 1 and kinds.count('sop') == 1 and kinds.count('key') == 1, kinds

    # 三类块宽度应 = 17cm（含取整误差 <=2 twips）
    for t in doc.tables:
        k = classify_table(t)
        if k:
            w = _table_width(t)
            assert abs(w - 17.0 * CM2TW) <= 2, (k, w)
    # 附录表 / 普通 2 列表 不应被识别
    assert classify_table(doc.tables[0]) is None, '附录表被误识别'
    assert classify_table(doc.tables[4]) is None, '普通2列表被误识别'
    print('[selftest] OK  blocks=', dict((kk, kinds.count(kk)) for kk in ('card', 'sop', 'key')))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='统一操作卡三类块为版心满宽 + 左对齐')
    ap.add_argument('--doc', help='输入 docx')
    ap.add_argument('--out', help='输出 docx')
    ap.add_argument('--content-width-cm', type=float, default=17.0,
                    help='版心宽度 cm，纵向 17.0 / 横向 25.7（默认 17.0）')
    ap.add_argument('--dry-run', action='store_true', help='只报告，不写文件')
    ap.add_argument('--selftest', action='store_true', help='端到端自测')
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()

    if not args.doc or not args.out:
        ap.error('--doc 与 --out 必填（除非 --selftest）')
    if args.dry_run and os.path.exists(args.out):
        pass  # dry-run 仍允许 out 存在（不会写）
    doc = Document(args.doc)
    rep = fix_document(doc, content_width_cm=args.content_width_cm)
    by_kind = {}
    for kind, old, new in rep:
        by_kind.setdefault(kind, []).append((old, new))
    print('[fix] 识别并修正三类块：', {k: len(v) for k, v in by_kind.items()})
    for kind, old, new in rep:
        print('   - %-4s tblW %.0f -> %.0f twips (%.2f cm)' % (kind, old or 0, new, new / CM2TW))
    if args.dry_run:
        print('[dry-run] 未写文件。')
        return 0
    doc.save(args.out)
    print('[fix] 已写出：', args.out, os.path.getsize(args.out), 'B')
    return 0


if __name__ == '__main__':
    sys.exit(main())
