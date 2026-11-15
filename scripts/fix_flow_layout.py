# -*- coding: utf-8 -*-
"""
fix_flow_layout.py
==================
成稿修订（Stage 6）版式校正：去除「逐项操作卡」里多余的强制换页，改为连续流式排版，
消除「章节间强制换页留白过多、总页数偏多」的问题。

两类强制换页源（跨品种通用，均按文档实际结构识别，不写死品种专用文本）：
  - 卡间分页：操作卡标题之间的空段落内嵌 run 级 `w:br type="page"`
    （成稿模板常 `doc.add_page_break()` 把每张卡逼到新页）
  - 章节分页：Heading 1 章节标题的 `w:pageBreakBefore`
    （成稿模板常 `add_h1_page()` 让每章从新页起）

可选动作：
  - 卡间空段落注入浅灰下边框分隔线（轻量视觉分界，不占满页、不影响连续排版）
  - 安全网：三类块表格（主卡表 / SOP 框 / 关键点框）设 `cantSplit`（防红框拆页）；
           操作卡标题 + H1 设 `keep_with_next`（防标题落页底孤行）

参数：
  --doc PATH              输入 docx（不覆盖，写 --out）
  --out PATH             输出 docx
  --card-pattern REGEX   操作卡标题正则（默认 `^3\\.\\d+[　\\s]`，匹配 3.1 / 3.2 …）
  --no-card-pagebreaks   不去除卡间分页（默认去除）
  --no-chapter-pagebreaks 不去除章节分页（默认去除）
  --no-separator-line    不在卡间注入分隔线（默认注入）
  --no-safety            不加 cantSplit / keep_with_next 安全网（默认加）
  --dry-run              只报告将改什么，不写文件
  --selftest             构造最小文档端到端自测，断言识别 + 移除 + 安全网正确

依赖：python-docx>=1.0；复用 `fix_operation_card_blocks.classify_table` 识别三类块。
纯版式改动，不触碰任何单元格/段落文本内容（零内容损失）。

经验来源：注射用乳糖酸红霉素 v7 巡检发现 27 处强制换页（14 卡间 page br + 13 章节
pageBreakBefore）撑出大量空白；去除后卡间改为浅灰分隔线、整本连续顺排，纯文本零变化。
"""
import argparse
import re
import sys
import os

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from fix_operation_card_blocks import classify_table

DEFAULT_CARD_PATTERN = r'^3\.\d+[　\s]'


# ---------------------------------------------------------------------
# 识别 / 移除层
# ---------------------------------------------------------------------
def _set_page_break_before(paragraph, value=True):
    """在段落 pPr 上设置/清除 w:pageBreakBefore。"""
    pPr = paragraph._p.get_or_add_pPr()
    pb = pPr.find(qn('w:pageBreakBefore'))
    if value:
        if pb is None:
            pb = OxmlElement('w:pageBreakBefore')
            pPr.append(pb)
        pb.set(qn('w:val'), 'true')
    else:
        if pb is not None:
            pPr.remove(pb)


def _add_bottom_border(p_el, color='D9D9D9', sz=6, space=4, before=40, after=40):
    """给段落加浅灰细下边框（轻量分隔线）。"""
    pPr = p_el.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'), str(after))
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    bottom = pBdr.find(qn('w:bottom'))
    if bottom is None:
        bottom = OxmlElement('w:bottom')
        pBdr.append(bottom)
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), str(space))
    bottom.set(qn('w:color'), color)


def remove_forced_pagebreaks(doc, args):
    """原地去除强制换页并加安全网。返回统计 dict。"""
    paras = list(doc.paragraphs)
    stats = {
        'page_br_removed': 0,
        'chapter_pbb_removed': 0,
        'separators': 0,
        'card_titles': 0,
        'h1_with_keepnext': 0,
        'blocks_cant_split': 0,
    }

    card_pat = re.compile(args.card_pattern)

    # ---- 1) 移除 run 级 page br，记录其所在段落元素 ----
    pagebr_els = set()
    if args.remove_card_pagebreaks:
        for p in paras:
            for r in p._p.findall(qn('w:r')):
                for br in r.findall(qn('w:br')):
                    if br.get(qn('w:type')) == 'page':
                        r.remove(br)
                        stats['page_br_removed'] += 1
                        if len(r.getchildren()) == 0:
                            r.getparent().remove(r)
                        pagebr_els.add(p._p)

    # ---- 2) 卡间分隔线：相邻卡标题之间、且含过 page br 的空段落 ----
    card_idx = [i for i, p in enumerate(paras)
                if card_pat.match((p.text or '').strip())]
    stats['card_titles'] = len(card_idx)
    sep_done = set()
    if args.separator_line and args.remove_card_pagebreaks and len(card_idx) >= 2:
        for k in range(1, len(card_idx)):
            lo, hi = card_idx[k - 1], card_idx[k]
            for i in range(lo + 1, hi):
                el = paras[i]._p
                if el in pagebr_els and el not in sep_done:
                    _add_bottom_border(el)
                    sep_done.add(el)
                    stats['separators'] += 1

    # ---- 3) 移除章节(H1) pageBreakBefore ----
    if args.remove_chapter_pagebreaks:
        for p in paras:
            if p.style.name == 'Heading 1':
                pPr = p._p.find(qn('w:pPr'))
                if pPr is not None:
                    pb = pPr.find(qn('w:pageBreakBefore'))
                    if pb is not None:
                        pPr.remove(pb)
                        stats['chapter_pbb_removed'] += 1
                    if args.safety:
                        p.paragraph_format.keep_with_next = True
                        stats['h1_with_keepnext'] += 1

    # ---- 4) 卡标题 keep_with_next ----
    if args.safety:
        for i in card_idx:
            paras[i].paragraph_format.keep_with_next = True

    # ---- 5) 三类块表格 cantSplit ----
    if args.safety:
        for t in doc.tables:
            if classify_table(t) is not None:
                tblPr = t._tbl.tblPr
                if tblPr.find(qn('w:cantSplit')) is None:
                    tblPr.append(OxmlElement('w:cantSplit'))
                stats['blocks_cant_split'] += 1

    return stats


# ---------------------------------------------------------------------
# CLI / selftest
# ---------------------------------------------------------------------
def _selftest():
    try:
        from docx_table_toolkit import add_key_callout, setup_page
    except Exception as e:  # pragma: no cover
        print('[selftest] 需要 docx_table_toolkit 在同目录：', e)
        return 1

    doc = Document()
    setup_page(doc.sections[0], landscape=False)

    # 章节（带 pageBreakBefore）
    h1 = doc.add_heading('1 产品基本信息', level=1)
    _set_page_break_before(h1, True)
    doc.add_paragraph('intro text')
    # 操作卡章节标记
    doc.add_paragraph('3.1 性状')
    # 卡间空段落，内嵌 run 级 page br
    empty = doc.add_paragraph()
    run = empty.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)
    doc.add_paragraph('3.2 鉴别')
    # 关键点框（三类块之一，应被加 cantSplit）
    add_key_callout(doc, '易错点说明')

    class A:
        card_pattern = DEFAULT_CARD_PATTERN
        remove_card_pagebreaks = True
        remove_chapter_pagebreaks = True
        separator_line = True
        safety = True

    stats = remove_forced_pagebreaks(doc, A())

    # 断言：两类换页均被移除
    assert stats['page_br_removed'] >= 1, stats
    assert stats['chapter_pbb_removed'] >= 1, stats
    assert stats['separators'] >= 1, stats
    assert stats['card_titles'] == 2, stats
    assert stats['blocks_cant_split'] == 1, stats

    # 断言：无残留 page br
    for p in doc.paragraphs:
        for r in p._p.findall(qn('w:r')):
            for b in r.findall(qn('w:br')):
                assert b.get(qn('w:type')) != 'page', '残留 page br'
    # 断言：H1 无 pageBreakBefore
    for p in doc.paragraphs:
        if p.style.name == 'Heading 1':
            pPr = p._p.find(qn('w:pPr'))
            assert pPr is None or pPr.find(qn('w:pageBreakBefore')) is None
    # 断言：关键点框有 cantSplit
    key_tbl = [t for t in doc.tables if classify_table(t) == 'key']
    assert key_tbl, '关键点框未识别'
    assert key_tbl[0]._tbl.tblPr.find(qn('w:cantSplit')) is not None, 'cantSplit 缺失'
    # 断言：分隔线段落含 pBdr bottom
    assert any((p._p.find(qn('w:pPr')) is not None
                and p._p.find(qn('w:pPr')).find(qn('w:pBdr')) is not None
                and p._p.find(qn('w:pPr')).find(qn('w:pBdr')).find(qn('w:bottom')) is not None)
               for p in doc.paragraphs), '分隔线未注入'

    print('[selftest] OK ', stats)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(
        description='去除强制换页，改为连续流式排版（操作卡三类块配套）')
    ap.add_argument('--doc', help='输入 docx')
    ap.add_argument('--out', help='输出 docx')
    ap.add_argument('--card-pattern', default=DEFAULT_CARD_PATTERN,
                    help='操作卡标题正则（默认 ^3\\.\\d+[　\\s]）')
    ap.add_argument('--no-card-pagebreaks', action='store_true',
                    help='保留卡间分页（默认去除）')
    ap.add_argument('--no-chapter-pagebreaks', action='store_true',
                    help='保留章节分页（默认去除）')
    ap.add_argument('--no-separator-line', action='store_true',
                    help='不在卡间注入分隔线（默认注入）')
    ap.add_argument('--no-safety', action='store_true',
                    help='不加 cantSplit / keep_with_next 安全网（默认加）')
    ap.add_argument('--dry-run', action='store_true', help='只报告，不写文件')
    ap.add_argument('--selftest', action='store_true', help='端到端自测')
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()

    if not args.doc or not args.out:
        ap.error('--doc 与 --out 必填（除非 --selftest）')

    class Opt:
        card_pattern = args.card_pattern
        remove_card_pagebreaks = not args.no_card_pagebreaks
        remove_chapter_pagebreaks = not args.no_chapter_pagebreaks
        separator_line = not args.no_separator_line
        safety = not args.no_safety

    doc = Document(args.doc)
    stats = remove_forced_pagebreaks(doc, Opt())
    print('[fix] 识别并修正：', stats)
    if args.dry_run:
        print('[dry-run] 未写文件。')
        return 0
    doc.save(args.out)
    print('[fix] 已写出：', args.out, os.path.getsize(args.out), 'B')
    return 0


if __name__ == '__main__':
    sys.exit(main())
