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


# ------------------------------------------------- 表指纹 / 表头（结构比对用）
def _table_header(tbl):
    """表头行文本元组（用于判断两版表的**顺序**是否一致）。"""
    if not tbl.rows:
        return ()
    return tuple(c.text.strip() for c in tbl.rows[0].cells)


def _table_fingerprint(tbl):
    """表的内容指纹：维度 + 每格文本。

    与 `refactor_chapter.py` 的同名函数保持同一口径——**按内容**而非按索引，
    迁表只改位置不改内容，指纹不变。
    """
    cells = []
    for r in tbl.rows:
        cells.append('\x1f'.join(c.text for c in r.cells))
    return '%dx%d|%s' % (len(tbl.rows), len(tbl.columns), '\x1e'.join(cells))


def _table_fingerprints(doc):
    return [_table_fingerprint(t) for t in doc.tables]


def _tables_by_header(doc, keywords, mode='all'):
    """按表头关键字定位表（禁索引——迁表会让索引整体漂移）。"""
    out = []
    kws = keywords or []
    if not kws:
        return out
    for t in doc.tables:
        if not t.rows:
            continue
        head = ' '.join(c.text.strip() for c in t.rows[0].cells)
        hit = (all(k in head for k in kws) if mode == 'all'
               else any(k in head for k in kws))
        if hit:
            out.append(t)
    return out


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
        # 描述式白名单（新通道）：供「无法用索引表达」的场景使用
        allowed_pats = list(cfg.get('allowed_cell_patterns', []))
        diffs = []
        # 表**顺序**是否一致（表头序列相同）。中间若发生迁表 / 删节，
        # `doc.tables` 索引会整体漂移，此时「按索引逐格比对」必然假阳性
        # —— 必须降级为**按表内容指纹的多重集合**比对。
        hdr_a = [_table_header(t) for t in doc.tables]
        hdr_b = [_table_header(t) for t in b.tables]
        same_order = (len(A) == len(B)) and (hdr_a == hdr_b)
        if not same_order:
            fa = _table_fingerprints(doc)
            fb = _table_fingerprints(b)
            allowed_fp_a, allowed_fp_b = set(), set()
            for spec in (cfg.get('allowed_table_changes') or []):
                kws = spec.get('header_keywords', [])
                mode = spec.get('match_mode', 'all')
                for t in _tables_by_header(doc, kws, mode):
                    allowed_fp_a.add(_table_fingerprint(t))
                for t in _tables_by_header(b, kws, mode):
                    allowed_fp_b.add(_table_fingerprint(t))
            lost = sorted(set(fb) - set(fa) - allowed_fp_b)
            added = sorted(set(fa) - set(fb) - allowed_fp_a)
            # 「表序已变」本身不是差异（迁表/删节是合法操作），
            # 只有**内容**变了才算差异。
            rep.ck(True, '表序已变 → baseline 降级为指纹集合比对'
                          '（丢失 %d / 新增 %d）' % (len(lost), len(added)), L)
            if lost:
                diffs.append('丢失/被改 %d 张' % len(lost))
            if added:
                diffs.append('新增/被改 %d 张' % len(added))
            for x in (lost + added)[:6]:
                diffs.append(x[:48])
            # 白名单表正向断言：只允许**末尾追加**，旧行必须逐格保留
            for spec in (cfg.get('allowed_table_changes') or []):
                kws = spec.get('header_keywords', [])
                mode = spec.get('match_mode', 'all')
                olds = _tables_by_header(b, kws, mode)
                news = _tables_by_header(doc, kws, mode)
                for to, tn in zip(olds, news):
                    if len(tn.rows) < len(to.rows):
                        diffs.append('白名单表行数减少')
                        continue
                    for ri in range(len(to.rows)):
                        if [c.text for c in to.rows[ri].cells] != \
                                [c.text for c in tn.rows[ri].cells]:
                            diffs.append('白名单表旧行 r%d 未原样保留' % ri)
                            break
        elif len(A) != len(B):
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
                        if ca == cb:
                            continue
                        if key in allowed:
                            continue
                        # 描述式白名单：模式串命中差异上下文即放行
                        ctx = '%s' % ca
                        if any(re.search(p, ctx) for p in allowed_pats):
                            continue
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


# ------------------------------------------------- 尾部结构级（升版出口契约）
def _para_texts(doc):
    """正文段 [(style_name, text)]，保持文档序。"""
    out = []
    for p in doc.paragraphs:
        st = p.style.name if p.style is not None else ''
        out.append((st, (p.text or '').strip()))
    return out


def tail_level(doc, cfg, rep):
    """尾部结构完整性 —— 合并/升版脚本的**出口契约**。

    要防的三种真实事故（都实际发生过）：
      1. 「增量说明」标题下**正文为空**——本该属于上一小节的条目被挂了上去；
      2. 结尾标记「—— 本手册完 ——」**粘在 bullet 末尾**，不是独立成段；
      3. 小节条目**跨小节错挂**。

    **向后兼容**：本层的所有 config 字段缺失时，整层 SKIP（绝不 FAIL）。
    旧 config 是别的品种的既有资产，不能因为脚本升级就变红。
    """
    L = '尾部结构级'
    ti = cfg.get('tail_integrity')
    if not ti:
        rep.skip('未提供 tail_integrity，跳过尾部结构断言', L)
        return

    paras = _para_texts(doc)
    bullets = {'List Bullet', 'List Bullet 2', 'List Paragraph'}

    # --- 结尾标记 ---
    marker = ti.get('end_marker')
    if marker:
        hits = [(st, t) for st, t in paras if marker in t]
        exp = ti.get('end_marker_occurrences', 1)
        rep.ck(len(hits) == exp,
               '结尾标记出现 %d 次（期望 %d）' % (len(hits), exp), L)
        if ti.get('end_marker_standalone', True) and hits:
            standalone = [h for h in hits if h[1].strip() == marker]
            rep.ck(bool(standalone),
                   '结尾标记独立成段（命中 %d 处，其中独立 %d 处）'
                   % (len(hits), len(standalone)), L)
    else:
        rep.skip('未提供 end_marker，跳过结尾标记断言', L)

    # --- 增量说明正文非空 ---
    #
    # 判定口径（重要）：**不是所有增量说明节都要求概述段**。
    # 真实文档里有两类增量说明：
    #   (a) bullet 型——标题下直接跟若干条要点，本身就没有概述段（合法）；
    #   (b) 概述型——标题下先有一段 Normal 概述，再跟要点（升版节常用）。
    # 真正的事故是 (b) 型被写坏了：概述段丢失，本该属于**上一小节**的条目
    # 被挂到了增量说明标题之下。
    #
    # 因此本断言只针对「本应概述型却写成 bullet 型」的节：
    # 判定 = 该增量说明节的**第一条 bullet**，若其内容在文档别处已作为
    # 独立小节的内容出现（即它本属别节），则报错。
    # 简化且稳健的实现：只检查**文档最后一个**增量说明节（升版节），
    # 且当该节之后紧跟收尾标记时，要求其下至少有一条 bullet 且这些 bullet
    # 不得包含收尾标记（收尾标记必须独立成段，已由上面断言覆盖）。
    inc_re = ti.get('increment_heading_regex', '增量说明')
    rx = re.compile(inc_re)
    if ti.get('require_body_paragraph', True):
        inc_idx = []
        for i, (st, t) in enumerate(paras):
            if st in bullets or not t or not rx.search(t):
                continue
            if not re.match(r'^\s*\d', t):        # 必须是编号开头的标题段
                continue
            inc_idx.append(i)

        # 收集「本应属于上一小节」的候选：出现在增量说明节之前的最后一个 bullet
        bad = []
        for i in inc_idx:
            prev_bullet = None
            for j in range(i - 1, -1, -1):
                if paras[j][1] and paras[j][0] in bullets:
                    prev_bullet = paras[j][1]
                    break
            if prev_bullet is None:
                continue
            # 下一段是 bullet，且与上一小节的末条**内容相同** → 错挂
            nxt = None
            for j in range(i + 1, len(paras)):
                if paras[j][1]:
                    nxt = paras[j]
                    break
            if nxt and nxt[0] in bullets and nxt[1][:24] == prev_bullet[:24]:
                bad.append('%s → 首条与上一小节末条重复（错挂）' % paras[i][1][:26])
        rep.ck(not bad,
               '增量说明节无「错挂上一小节条目」（检查 %d 个增量说明节，异常：%s）'
               % (len(inc_idx), bad[:4] or '无'), L)

        # 概述型增量说明节（config 显式指定其编号）必须有概述段
        strict = ti.get('strict_body_heading_numbers') or []
        if strict:
            sbad = []
            for num in strict:
                idx = None
                for i, (st, t) in enumerate(paras):
                    if st in bullets:
                        continue
                    if re.match(r'^\s*' + re.escape(str(num)) + r'(?![0-9.])', t):
                        idx = i
                        break
                if idx is None:
                    sbad.append('%s 未找到' % num)
                    continue
                nxt = None
                for j in range(idx + 1, len(paras)):
                    if paras[j][1]:
                        nxt = paras[j]
                        break
                if nxt is None:
                    sbad.append('%s 下无正文段' % num)
                elif nxt[0] in bullets:
                    sbad.append('%s 下首个段是 bullet，概述段缺失' % num)
            rep.ck(not sbad,
                   '指定增量说明节有概述段（异常：%s）' % (sbad[:4] or '无'), L)
        else:
            rep.skip('未提供 strict_body_heading_numbers，'
                     '跳过「概述型增量说明必须有概述段」断言', L)
    else:
        rep.skip('require_body_paragraph=false，跳过增量说明正文断言', L)

    # --- 跨小节错挂：bullet 段若自带编号，该编号不得大于其上方最近的小节标题 ---
    if ti.get('forbid_cross_section_attach', True):
        def num_of(t):
            m = re.match(r'^\s*(\d+(?:\.\d+)*)(?![0-9.])', t)
            if not m:
                return None
            return tuple(int(x) for x in m.group(1).split('.'))

        cur_num = None       # 上方最近的小节标题编号（至少两级）
        cross = []
        for st, t in paras:
            if not t:
                continue
            if st in bullets:
                bn = num_of(t)
                # bullet 自带编号且大于其上方的小节号 → 错挂
                if bn is not None and cur_num is not None and bn > cur_num:
                    cross.append(t[:30])
                continue
            n = num_of(t)
            if n is not None and len(n) >= 2:
                cur_num = n
        rep.ck(not cross,
               '无跨小节错挂条目（异常 %d 条：%s）' % (len(cross), cross[:4] or '无'), L)
    else:
        rep.skip('forbid_cross_section_attach=false，跳过错挂断言', L)


# ------------------------------------------------- 章节体量（决策辅助）
def volume_level(doc, cfg, rep):
    """章节体量报告 + 可选硬断言。

    这是「从哪砍」的决策依据：**表内文字与正文文字必须分开统计**，
    否则会得出「全是文字，砍不动」的错误结论。
    """
    L = '章节体量'
    sv = cfg.get('section_volume')
    if not sv:
        rep.skip('未提供 section_volume，跳过章节体量报告', L)
        return
    try:
        from docx_revise_toolkit import measure_sections
    except ImportError:
        rep.skip('未安装 docx_revise_toolkit，跳过章节体量报告', L)
        return

    chapters = sv.get('chapters') or []
    secs = []
    for ch in chapters:
        secs.extend(measure_sections(doc, chapter=str(ch)))
    if not secs:
        secs = measure_sections(doc)
    total = sum(s['total_chars'] for s in secs)
    tbl_total = sum(s['table_chars'] for s in secs)
    ratio = (tbl_total / float(total) * 100) if total else 0
    rep.ck(total > 0, '章节体量合计 %d 字（表内 %d 字，占 %.1f%%）'
           % (total, tbl_total, ratio), L)
    top = sorted(secs, key=lambda s: -s['total_chars'])[:sv.get('report_top_n', 5)]
    for s in top:
        rep.skip('体量 %s %s：正文 %d + 表内 %d = %d 字 / %d 表'
                 % (s['number'] or '-', (s['title'] or '')[:24], s['chars'],
                    s['table_chars'], s['total_chars'], s['tables']), L)
    if sv.get('assert_no_empty_sections'):
        empty = [s['number'] for s in secs
                 if s['total_chars'] == 0 and s['style'] != 'Heading 1']
        rep.ck(not empty, '无空小节（异常：%s）' % (empty[:6] or '无'), L)
    mc = sv.get('min_body_chars')
    if mc is not None:
        thin = ['%s(%d)' % (s['number'], s['chars']) for s in secs
                if s['chars'] < mc and s['number']]
        rep.ck(not thin, '正文段字符数 ≥ %d（偏薄：%s）' % (mc, thin[:6] or '无'), L)


# ------------------------------------------------- 交叉引用专项
def crossref_level(doc, cfg, rep):
    """迁表后的交叉引用双向断言：必须出现的新指向 + 必须消失的旧指向。"""
    L = '交叉引用'
    cr = cfg.get('cross_reference_checks')
    if not cr:
        rep.skip('未提供 cross_reference_checks，跳过交叉引用断言', L)
        return
    if cr.get('scan_cells', True):
        full = full_text(doc)
    else:
        full = '\n'.join(p.text for p in doc.paragraphs)
    for s in cr.get('must_contain', []):
        rep.ck(s in full, '新指向命中：%s' % s, L)
    for s in cr.get('must_not_contain', []):
        rep.ck(s not in full, '旧指向 0 命中：%s' % s, L)


# ------------------------------------------------- 迁表零损失
def moved_tables_level(doc, cfg, rep):
    """迁出表的维度必须不变（表内容一字不删，只挪位）。

    **必须按表头关键字定位**，绝不能用索引——迁表会让 `doc.tables` 索引整体漂移。
    """
    L = '迁表零损失'
    specs = cfg.get('moved_tables')
    if not specs:
        rep.skip('未提供 moved_tables，跳过迁表零损失断言', L)
        return
    for spec in specs:
        kws = spec.get('header_keywords', [])
        mode = spec.get('match_mode', 'all')
        hits = [t for t in doc.tables if t.rows and (
            all(k in ' '.join(c.text.strip() for c in t.rows[0].cells) for k in kws)
            if mode == 'all' else
            any(k in ' '.join(c.text.strip() for c in t.rows[0].cells) for k in kws))]
        rep.ck(len(hits) >= 1,
               '迁出表命中（%s）：%d 张' % (kws, len(hits)), L)
        dims = spec.get('expect_dims')
        if dims and hits:
            ok = any(len(t.rows) == dims[0] and len(t.columns) == dims[1]
                     for t in hits)
            rep.ck(ok, '迁出表维度保持 %sx%s（实际 %s）'
                   % (dims[0], dims[1],
                      ['%dx%d' % (len(t.rows), len(t.columns)) for t in hits]), L)


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
    # 尾部结构级：合并/升版的出口契约。
    # 放在 OOXML 结构级之后、视觉级之前；config 缺字段时整层 SKIP（不 FAIL）。
    tail_level(doc, config, rep)
    moved_tables_level(doc, config, rep)
    crossref_level(doc, config, rep)
    volume_level(doc, config, rep)
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
