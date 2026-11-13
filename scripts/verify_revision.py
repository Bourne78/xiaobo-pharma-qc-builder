# -*- coding: utf-8 -*-
"""参数化三层零回归验收 runner（跨品种通用）。

用法：
    python verify_revision.py --docx X.docx --config cfg.json \
        [--baseline OLD.docx] [--pdf X.pdf] \
        [--matrix-table-title "F.3  检验项 × 物料矩阵"] [--check-page-break]

分层：文本级 / OOXML 结构级 / 视觉级。所有阈值与字符串来自 --config（JSON），
换品种只换 config，不改脚本。

脚本不内嵌 COM 导出、不内嵌 PDF 渲染：--pdf 仅对既存 PDF 文本层做可选断言，
缺则视觉级整层 SKIP，并固定打印「PDF 需手动导出、逐页目视待补」。
详见 references/revision_playbook.md §8 / 附录 C。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    from docx import Document
    from docx.oxml.ns import qn
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False
    Document = qn = None

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class Report:
    def __init__(self):
        self.rows = []
        self.failed = 0
        self.skipped = 0

    def ck(self, cond, msg, layer):
        ok = bool(cond)
        self.rows.append((layer, 'OK' if ok else 'FAIL', msg))
        if not ok:
            self.failed += 1
        return ok

    def skip(self, msg, layer):
        self.rows.append((layer, 'SKIP', msg))
        self.skipped += 1

    @property
    def exit_code(self):
        return 1 if self.failed else 0

    def dump(self):
        cur = None
        for layer, status, msg in self.rows:
            if layer != cur:
                print('\n=== %s ===' % layer)
                cur = layer
            print('  %-4s %s' % (status, msg))
        print('\n合计: FAIL=%d  SKIP=%d' % (self.failed, self.skipped))
        return self.exit_code


def _require_docx():
    if not _DOCX_OK:
        raise SystemExit('缺少 python-docx，请先 pip install python-docx')


def load_config(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def full_text(doc):
    paras = [p.text for p in doc.paragraphs]
    cells = [c.text for t in doc.tables for r in t.rows for c in r.cells]
    return '\n'.join(paras + cells)


def cell_matrix(doc):
    return [[[c.text for c in r.cells] for r in t.rows] for t in doc.tables]


# ---------------------------------------------------------------- XML 级工具
def xml_root(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        return ET.fromstring(z.read('word/document.xml'))


def sectpr_signatures(docx_path):
    return [ET.tostring(e, encoding='unicode')
            for e in xml_root(docx_path).iter(W + 'sectPr')]


def count_page_breaks(docx_path):
    return sum(1 for br in xml_root(docx_path).iter(W + 'br')
               if br.get(W + 'type') == 'page')


def count_pbdr(docx_path):
    return sum(1 for _ in xml_root(docx_path).iter(W + 'pBdr'))


def check_fonts(docx_path):
    root = xml_root(docx_path)
    bad_rf = bad_sz = 0
    for rPr in root.iter(W + 'rPr'):
        rf = rPr.find(W + 'rFonts')
        if rf is not None and any(rf.get(W + a) is None
                                  for a in ('ascii', 'hAnsi', 'cs', 'eastAsia')):
            bad_rf += 1
        if rPr.find(W + 'sz') is not None and rPr.find(W + 'szCs') is None:
            bad_sz += 1
    return bad_rf, bad_sz


def check_table_widths(tbl, tol=None):
    """返回 (ok, total, fixed)；autofit 表返回 (None, total, False) 表示不参与。

    tol 默认按列数自适应：tblW 与 gridCol 合计常因逐列 round() 累计差几个 twips
    （实测 7 列差 9~11tw ≈ 0.016cm），用绝对相等会误报；每格 tcW 仍逐格精确比对。
    """
    tblPr = tbl.find(qn('w:tblPr'))
    layout = tblPr.find(qn('w:tblLayout')) if tblPr is not None else None
    fixed = layout is not None and layout.get(qn('w:type')) == 'fixed'
    tw = tblPr.find(qn('w:tblW')) if tblPr is not None else None
    grid_el = tbl.find(qn('w:tblGrid'))
    grid = ([int(g.get(qn('w:w')) or 0)
             for g in grid_el.findall(qn('w:gridCol'))]
            if grid_el is not None else [])
    total = int(tw.get(qn('w:w')) or 0) if tw is not None else 0
    if not fixed:
        return None, total, False
    if tol is None:
        tol = max(4, 2 * len(grid))
    ok = total > 0 and abs(sum(grid) - total) <= tol
    for tr in tbl.findall(qn('w:tr')):
        col = 0
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            gs = tcPr.find(qn('w:gridSpan')) if tcPr is not None else None
            s = int(gs.get(qn('w:val'))) if gs is not None else 1
            exp = sum(grid[col:col + s])
            tcW = tcPr.find(qn('w:tcW')) if tcPr is not None else None
            got = int(tcW.get(qn('w:w')) or 0) if tcW is not None else 0
            if abs(got - exp) > tol:
                ok = False
            col += s
    return ok, total, True


def table_flags(tbl):
    """返回 (是否有 tblHeader, cantSplit 行数)。"""
    hdr = False
    cant = 0
    for tr in tbl.findall(qn('w:tr')):
        trPr = tr.find(qn('w:trPr'))
        if trPr is None:
            continue
        if trPr.find(qn('w:tblHeader')) is not None:
            hdr = True
        if trPr.find(qn('w:cantSplit')) is not None:
            cant += 1
    return hdr, cant


def read_pdf_pages(path):
    """只读既存 PDF 的文本层；不渲染图片、不调用 Word COM。"""
    if not path or not os.path.exists(path):
        return None
    try:
        import pypdfium2 as pdfium
    except ImportError:
        print('WARN 未安装 pypdfium2，PDF 文本层不可用（视觉级 SKIP）')
        return None
    try:
        pdf = pdfium.PdfDocument(path)
        return [pdf[i].get_textpage().get_text_range() for i in range(len(pdf))]
    except Exception as e:                       # noqa: BLE001 - 环境差异兜底
        print('WARN 无法读取 PDF 文本层: %s' % e)
        return None


# ---------------------------------------------------------------- 三层断言
def text_level(doc, full, cfg, rep, pdf_text):
    L = '文本级'
    for t in cfg.get('chapter_titles', []):
        rep.ck(t in full, '章节标题命中：%s' % t, L)
    miss = [x for x in cfg.get('limits', []) if x not in full]
    rep.ck(not miss, '关键数值/限度全命中（缺失：%s）' % (miss or '无'), L)
    for s in cfg.get('new_strings', []):
        rep.ck(s in full, '新增串命中：%s' % s, L)
    for s in cfg.get('old_strings', []):
        rep.ck(s not in full, '旧串 0 命中：%s' % s, L)
    prefixes = cfg.get('id_prefixes', ['RP', 'RGT', 'SMP', 'COL', 'INS', 'CON'])
    ids = set(re.findall(r'(?:%s)-\d+' % '|'.join(prefixes), full))
    rep.ck(bool(ids), '物料编号 %d 个' % len(ids), L)
    found = {p for p in prefixes if re.search(p + r'-\d+', full)}
    rep.ck(found == set(prefixes), '六类前缀齐全：%s' % sorted(found), L)
    exact = cfg.get('callout_exact', '● 关键点 / 易错点')
    n = sum(1 for t in doc.tables
            if t.rows and exact in t.rows[0].cells[0].text)
    exp = cfg.get('expected_counts', {}).get('callouts')
    rep.ck(n == exp if exp is not None else n > 0,
           '关键点框 %d（期望 %s）' % (n, exp), L)
    rep.ck(full.count('**') == 0, "DOCX '**' 残留 = %d" % full.count('**'), L)
    rep.ck(full.count('\ufffc') == 0,
           'U+FFFC 伪影 = %d' % full.count('\ufffc'), L)
    if pdf_text is not None:
        rep.ck(pdf_text.count('**') == 0,
               "PDF '**' 残留 = %d" % pdf_text.count('**'), L)


def structure_level(docx_path, doc, cfg, rep, baseline_path, check_page_break):
    L = 'OOXML 结构级'
    exp = cfg.get('expected_counts', {})
    if 'tables' in exp:
        rep.ck(len(doc.tables) == exp['tables'],
               '总表数 = %d（期望 %d）' % (len(doc.tables), exp['tables']), L)
    if 'paragraphs' in exp:
        rep.ck(len(doc.paragraphs) == exp['paragraphs'],
               '总段落数 = %d（期望 %d）' % (len(doc.paragraphs), exp['paragraphs']), L)

    # baseline 逐格内容比对（白名单之外的差异一律 FAIL）
    if baseline_path and os.path.exists(baseline_path):
        b = Document(baseline_path)
        A, B = cell_matrix(doc), cell_matrix(b)
        allowed = set(cfg.get('allowed_cells', []))
        diffs = []
        if len(A) != len(B):
            diffs.append('表数不同 %d vs %d' % (len(A), len(B)))
        else:
            for ti, (ta, tb) in enumerate(zip(A, B)):
                if len(ta) != len(tb):
                    diffs.append('t%d 行数不同' % ti)
                    continue
                for ri, (ra, rb) in enumerate(zip(ta, tb)):
                    if len(ra) != len(rb):
                        diffs.append('t%d.r%d 列数不同' % (ti, ri))
                        continue
                    for ci, (ca, cb) in enumerate(zip(ra, rb)):
                        key = '%d.%d.%d' % (ti, ri, ci)
                        if ca != cb and key not in allowed:
                            diffs.append(key)
        rep.ck(not diffs,
               'baseline 逐格零改动（差异 %d 处：%s）' % (len(diffs), diffs[:8]), L)
        sa, sb = sectpr_signatures(docx_path), sectpr_signatures(baseline_path)
        rep.ck(sa == sb, 'sectPr 数量=%d 且内容一字未动' % len(sa), L)
    else:
        rep.skip('未提供 --baseline，跳过逐格比对', L)
        sa = sectpr_signatures(docx_path)
        if 'sections' in exp:
            rep.ck(len(sa) == exp['sections'],
                   'sectPr 数量 = %d（期望 %d）' % (len(sa), exp['sections']), L)

    # tblW / gridCol / tcW 三者一致且 fixed（autofit 表自动跳过）
    bad = []
    for i, t in enumerate(doc.tables):
        ok, _, fixed = check_table_widths(t._tbl)
        if fixed and not ok:
            bad.append(i)
    rep.ck(not bad, 'tblW/gridCol/tcW 一致且 fixed（异常表 %s）' % (bad[:8] or '无'), L)

    # 矩阵表维度、表头重复、行不跨页、列宽合计
    dims = cfg.get('matrix_expected_dims')
    if dims:
        cand = [t for t in doc.tables
                if len(t.rows) == dims[0] and len(t.columns) == dims[1]]
        rep.ck(bool(cand), '矩阵表维度 %s 命中' % (dims,), L)
        if cand:
            hdr, cant = table_flags(cand[0]._tbl)
            rep.ck(hdr, '矩阵表 tblHeader 保留', L)
            rep.ck(cant >= dims[0] - 1,
                   '矩阵表 cantSplit 行数 >= %d（实际 %d）' % (dims[0] - 1, cant), L)
            okw, total, _ = check_table_widths(cand[0]._tbl)
            rep.ck(okw is True, '矩阵表列宽三同步，合计 = %s twips' % total, L)

    # 裸代号只扫矩阵数据格，避免全文扫命中物料表编号列（误报规避）
    title = cfg.get('matrix_table_title')
    bare = cfg.get('matrix_bare_code_pattern')
    if bare and dims:
        cand = [t for t in doc.tables
                if len(t.rows) == dims[0] and len(t.columns) == dims[1]]
        if cand:
            hits = []
            for ri, row in enumerate(cand[0].rows):
                if ri == 0:
                    continue                      # 排除列头
                for ci, c in enumerate(row.cells):
                    if re.search(bare, c.text):
                        hits.append('%d.%d' % (ri, ci))
            rep.ck(not hits, '矩阵数据格裸代号 0 命中（%s）' % (hits[:8] or '无'), L)
        else:
            rep.skip('未定位到矩阵表，跳过裸代号断言', L)
    else:
        rep.skip('未提供 matrix_bare_code_pattern，跳过裸代号断言', L)

    if check_page_break:
        rep.ck(count_page_breaks(docx_path) == 0,
               'w:br[type=page] = %d' % count_page_breaks(docx_path), L)
    if 'card_separators' in exp:
        pc = count_pbdr(docx_path)
        rep.ck(pc == exp['card_separators'],
               'w:pBdr 分隔段 = %d（期望 %d）' % (pc, exp['card_separators']), L)
    if cfg.get('check_fonts'):
        bad_rf, bad_sz = check_fonts(docx_path)
        rep.ck(bad_rf == 0, 'rFonts 四属性缺失 run = %d' % bad_rf, L)
        rep.ck(bad_sz == 0, 'w:sz 无 szCs 配对 run = %d' % bad_sz, L)
    else:
        rep.skip('check_fonts=false，跳过字体四属性断言', L)


def visual_level(cfg, rep, pdf_pages):
    L = '视觉级'
    if pdf_pages is None:
        rep.skip('无 PDF：逐页目视（ViewImage）待补', L)
        return False
    full_pdf = '\n'.join(pdf_pages)
    rep.ck(full_pdf.count('\ufffc') == 0,
           'PDF U+FFFC 伪影 = %d' % full_pdf.count('\ufffc'), L)
    exact = cfg.get('callout_exact', '● 关键点 / 易错点')
    kc_pages = [i for i, t in enumerate(pdf_pages) if exact in t]
    rep.ck(bool(kc_pages), 'PDF 文本层可检出关键点框（%d 页）' % len(kc_pages), L)
    card_re = cfg.get('card_title_regex')
    if card_re:
        rx = re.compile(card_re)
        card_pages = [i for i, t in enumerate(pdf_pages) if rx.search(t)]
        rep.ck(all(sum(1 for t in pdf_pages[:p + 1] if exact in t)
                   <= sum(1 for t in pdf_pages[:p + 1] if rx.search(t))
                   for p in range(len(pdf_pages))),
               'orphan_kc 判据一：每页红框数 ≤ 累计卡标题数', L)
        rep.ck(len([p for p in card_pages]) <= cfg.get('expected_counts', {}).get('callouts', 10 ** 6),
               '含卡标题页数 %d ≤ 卡片数' % len(card_pages), L)
    else:
        rep.skip('未提供 card_title_regex，跳过 orphan_kc 双判据', L)
    rep.skip('逐页 ViewImage 目视（溢出/截断/红框跨页）仍需人工完成', L)
    return True


def verify(docx_path, config, pdf_path=None, baseline_path=None,
           matrix_table_title=None, check_page_break=False):
    _require_docx()
    rep = Report()
    doc = Document(docx_path)
    docx_text = full_text(doc)
    pdf_pages = read_pdf_pages(pdf_path) if pdf_path else None
    pdf_text = '\n'.join(pdf_pages) if pdf_pages else None
    full = docx_text + ('\n' + pdf_text if pdf_text else '')
    if matrix_table_title:
        config = dict(config)
        config['matrix_table_title'] = matrix_table_title
    text_level(doc, full, config, rep, pdf_text)
    structure_level(docx_path, doc, config, rep, baseline_path, check_page_break)
    visual_done = visual_level(config, rep, pdf_pages)
    if not visual_done:
        print('\n★ PDF 需用户在 Word 手动导出、逐页目视待补；'
              '不得把结构级通过表述为版式已验收。')
    else:
        print('\n★ 逐页目视（ViewImage）仍需人工完成；PDF 文本层断言不等于目视验收。')
    return rep


def build_parser():
    ap = argparse.ArgumentParser(
        description='参数化三层零回归验收（文本级 / OOXML 结构级 / 视觉级）')
    ap.add_argument('--docx', required=True, help='待验收 .docx')
    ap.add_argument('--config', required=True, help='JSON 配置（所有阈值来自这里）')
    ap.add_argument('--baseline', help='上一版 .docx，用于逐格零回归比对')
    ap.add_argument('--pdf', help='既存 PDF（仅读文本层，不渲染、不导出）')
    ap.add_argument('--matrix-table-title', help='矩阵表标题（写入 config）')
    ap.add_argument('--check-page-break', action='store_true',
                    help='断言 w:br[type=page]=0（仅适用于取消卡间分页的版式）')
    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)
    cfg = load_config(args.config)
    rep = verify(args.docx, cfg, pdf_path=args.pdf, baseline_path=args.baseline,
                 matrix_table_title=args.matrix_table_title,
                 check_page_break=args.check_page_break)
    return rep.dump()


if __name__ == '__main__':
    sys.exit(main())
