# -*- coding: utf-8 -*-
"""成稿原地修订：表格列宽两策略 + 字体归一（跨品种通用）。

用法：
    python patch_existing_manual.py --input X.docx --output Y.docx \
        --strategy content-width [--normalize-fonts] [--font-mode cjk-split] \
        [--backup B.docx] [--dry-run]

两层列宽策略（互补，不可互替）：
    autofit        Word「根据内容调整」标记层：tblLayout=autofit + tblW/tcW auto
    content-width  实测每格文本宽度后写死 fixed 布局（默认）

设计约束：
    - 默认另存 --output，绝不覆盖输入；
    - 顶层不执行任何文件读写，import 无副作用；
    - 无品种专名、无绝对路径、无单次实测常量默认值。

详见 references/revision_playbook.md §2 / §3。
"""
from __future__ import annotations

import argparse
import shutil
import sys
import unicodedata

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:                       # --help 仍可用
    _DOCX_OK = False
    Document = qn = OxmlElement = None


def _require_docx():
    if not _DOCX_OK:
        raise SystemExit("缺少 python-docx，请先 pip install python-docx")


# ---------------------------------------------------------------- 通用工具
def ensure_child(parent, tag, after_of):
    """取/建 parent 下 tag 子元素，按 OOXML 顺序插到 after_of 中首个已有子元素之前。

    after_of 可写 'w:jc' 或完整 '{ns}jc'；内部统一转成 qn 形式再比较，
    避免 child.tag 与短名比较永远不命中、静默退化成 append 的坑。
    """
    el = parent.find(qn(tag))
    if el is not None:
        return el
    el = OxmlElement(tag)
    successors = {qn(t) if not t.startswith("{") else t for t in after_of}
    for child in parent:
        if child.tag in successors:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


_TBLW_AFTER = {'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders',
               'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblPrChange'}
_TCW_AFTER = {'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd',
              'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFit', 'w:tcPrChange'}


# ---------------------------------------------------------------- 列宽：标记层
def mark_all_autofit(doc):
    """Word「根据内容调整」标记层。返回处理的表数。"""
    _require_docx()
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        layout = tblPr.get_or_add_tblLayout()
        layout.set(qn('w:type'), 'autofit')
        tblW = ensure_child(tblPr, 'w:tblW', _TBLW_AFTER)
        tblW.set(qn('w:type'), 'auto')
        tblW.set(qn('w:w'), '0')
        for row in t.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = ensure_child(tcPr, 'w:tcW', _TCW_AFTER)
                tcW.set(qn('w:type'), 'auto')
                tcW.set(qn('w:w'), '0')
    return len(doc.tables)


# ---------------------------------------------------------------- 列宽：实测层
def char_width_twips(ch, sz_pt):
    """CJK(W/F)=1.0、模糊(A)=0.6、其余=0.5；1pt = 20 twips。"""
    ea = unicodedata.east_asian_width(ch)
    f = 1.0 if ea in ('W', 'F') else (0.6 if ea == 'A' else 0.5)
    return sz_pt * 20.0 * f


def _run_size_pt(r):
    rPr = r.find(qn('w:rPr'))
    if rPr is not None:
        szEl = rPr.find(qn('w:sz'))
        if szEl is not None:
            try:
                return int(szEl.get(qn('w:val'))) / 2.0
            except (TypeError, ValueError):
                pass
    return 10.5                          # 表格默认五号


def cell_ideal_width_twips(tc):
    """单元格最长文本行宽度；逐 run 读 w:sz，处理 w:br / 换行 / w:tab。"""
    max_line = 0.0
    for p in tc.findall(qn('w:p')):
        line_w = 0.0
        for r in p.findall(qn('w:r')):
            sz = _run_size_pt(r)
            for child in r:
                if child.tag == qn('w:t'):
                    for ch in (child.text or ''):
                        if ch in ('\n', '\r'):
                            max_line = max(max_line, line_w)
                            line_w = 0.0
                        else:
                            line_w += char_width_twips(ch, sz)
                elif child.tag == qn('w:br'):
                    max_line = max(max_line, line_w)
                    line_w = 0.0
                elif child.tag == qn('w:tab'):
                    line_w += char_width_twips(' ', sz)
        max_line = max(max_line, line_w)
    return max_line


def cell_pad_twips(tc):
    """左右单元格边距合计；默认各 115tw（0.08"）。"""
    left = right = 115
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is not None:
        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is not None:
            e = tcMar.find(qn('w:left'))
            if e is not None and e.get(qn('w:w')):
                left = int(e.get(qn('w:w')))
            e = tcMar.find(qn('w:right'))
            if e is not None and e.get(qn('w:w')):
                right = int(e.get(qn('w:w')))
    return left + right


def get_content_width(doc, tbl_el):
    """按表所在 section 返回可用版心宽度（twips）。

    横向时 pgSz@w 已是 16838tw，直接减左右边距即可；勿按 orient 互换 w/h。
    """
    body = doc.element.body
    children = list(body.iterchildren())
    breaks = []
    for child in children:
        if child.tag == qn('w:p'):
            pPr = child.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                breaks.append(child)
    sp = body.find(qn('w:sectPr'))
    try:
        ib = children.index(tbl_el)
    except ValueError:
        ib = -1
    for p in breaks:
        try:
            if children.index(p) > ib:
                sp = p.find(qn('w:pPr')).find(qn('w:sectPr'))
                break
        except ValueError:
            continue
    if sp is None:
        return 11906 - 2268
    pgSz = sp.find(qn('w:pgSz'))
    pgMar = sp.find(qn('w:pgMar'))
    pageW = int(pgSz.get(qn('w:w')) or 11906) if pgSz is not None else 11906
    if pgMar is not None:
        ml = int(pgMar.get(qn('w:left')) or 1134)
        mr = int(pgMar.get(qn('w:right')) or 1134)
    else:
        ml = mr = 1134
    return pageW - ml - mr


def compute_col_widths(tbl_el):
    """逐格取理想宽度约束；gridSpan 跨列约束按 s 均摊，vMerge 续格跳过。"""
    constraints = []
    max_col = 0
    for tr in tbl_el.findall(qn('w:tr')):
        col = 0
        for tc in tr.findall(qn('w:tc')):
            gs = tc.find(qn('w:gridSpan'))
            s = int(gs.get(qn('w:val'))) if gs is not None else 1
            tcPr = tc.find(qn('w:tcPr'))
            vMerge = tcPr.find(qn('w:vMerge')) if tcPr is not None else None
            is_continue = (vMerge is not None and
                           (vMerge.get(qn('w:val')) is None or
                            vMerge.get(qn('w:val')) != 'restart'))
            if not is_continue:
                constraints.append((col, s, cell_ideal_width_twips(tc) + cell_pad_twips(tc)))
            col += s
        max_col = max(max_col, col)
    widths = [0.0] * max_col
    for (c0, s, W) in constraints:
        if s == 1:
            widths[c0] = max(widths[c0], W)
        else:
            cur = sum(widths[c0:c0 + s])
            if cur < W:
                add = (W - cur) / s
                for k in range(c0, c0 + s):
                    widths[k] += add
    return widths


def resize_table(table, doc, *, max_col_ratio=0.6, floor_twips=480.0, max_rounds=10):
    """按内容重算并写死该表列宽；返回 (widths, content_width)。"""
    _require_docx()
    tbl_el = table._tbl
    content_w = get_content_width(doc, tbl_el)
    max_col = content_w * max_col_ratio
    floor = float(floor_twips)
    widths = compute_col_widths(tbl_el)
    widths = [max(min(w, max_col), floor) for w in widths]
    # 超出版心时只压缩「超出下限的部分」，保住窄列下限；极端情况才整体缩放
    for _ in range(int(max_rounds)):
        total = sum(widths)
        if total <= content_w:
            break
        excess = total - content_w
        shrinkable = sum(max(0.0, w - floor) for w in widths)
        if shrinkable <= 1e-6:
            scale = content_w / total
            widths = [w * scale for w in widths]
            break
        for i, w in enumerate(widths):
            if w > floor:
                widths[i] = w - excess * (w - floor) / shrinkable
    # 先统一取整，tblW 用「已取整列宽之和」，避免逐列 round 的累计舍入造成
    # tblW 与 gridCol 合计差几个 twips（历史文档里常见，验收时会误报）
    rounded = [int(round(w)) for w in widths]
    tblPr = tbl_el.tblPr
    tblW = ensure_child(tblPr, 'w:tblW', _TBLW_AFTER)
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(sum(rounded)))
    layout = tblPr.get_or_add_tblLayout()
    layout.set(qn('w:type'), 'fixed')
    # gridCol 同步
    grid = tbl_el.find(qn('w:tblGrid'))
    if grid is None:
        grid = OxmlElement('w:tblGrid')
        tbl_el.insert(list(tbl_el).index(tblPr) + 1, grid)
    gcs = grid.findall(qn('w:gridCol'))
    for i, w in enumerate(rounded):
        v = str(w)
        if i < len(gcs):
            gcs[i].set(qn('w:w'), v)
        else:
            ng = OxmlElement('w:gridCol')
            ng.set(qn('w:w'), v)
            grid.append(ng)
    # 每行每格 tcW 同步（含 gridSpan 累加）
    for tr in tbl_el.findall(qn('w:tr')):
        col = 0
        for tc in tr.findall(qn('w:tc')):
            gs = tc.find(qn('w:gridSpan'))
            s = int(gs.get(qn('w:val'))) if gs is not None else 1
            wsum = sum(rounded[col:col + s])
            tcPr = tc.get_or_add_tcPr()
            tcW = ensure_child(tcPr, 'w:tcW', _TCW_AFTER)
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(wsum))
            col += s
    return widths, content_w


def resize_all(doc, **kw):
    """对全部表按内容定宽；返回 [(列数, 最窄, 最宽, 版心), ...]（单位 twips）。"""
    return [(len(w), min(w), max(w), cw)
            for w, cw in (resize_table(t, doc, **kw) for t in doc.tables)]


# ---------------------------------------------------------------- 字体归一
SUBSCRIPT_MAP = {c: str(i) for i, c in enumerate("₀₁₂₃₄₅₆₇₈₉")}


def _all_paragraphs(doc):
    """产出 (paragraph, cell)；正文段落的 cell 为 None。"""
    for p in doc.paragraphs:
        yield p, None
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p, cell


def normalize_fonts(doc, *, font_mode='cjk-split', east_asia='宋体',
                    ascii_font=None, skip_empty=True, subscript_map=None):
    """save 前全 run 后处理：rFonts 四属性 + sz/szCs 成对 + 下标替换 + 窄列去加粗。

    font_mode='cjk-split'（默认）：ascii/hAnsi/cs=Times New Roman、eastAsia=宋体；
    font_mode='all-song'：四属性全宋体，兼容既有「全宋体最小改动」成品。
    """
    _require_docx()
    if font_mode == 'all-song':
        af = east_asia
    elif ascii_font:
        af = ascii_font
    else:
        af = 'Times New Roman'
    smap = SUBSCRIPT_MAP if subscript_map is None else subscript_map
    stats = {'runs': 0, 'skipped_empty': 0, 'subscript': 0,
             'szcs_added': 0, 'narrow_unbold': 0, 'font_mode': font_mode}
    for p, cell in _all_paragraphs(doc):
        narrow = False
        if cell is not None and cell.width is not None:
            try:
                narrow = cell.width.cm < 5.0
            except (AttributeError, TypeError):
                narrow = False
        for r in p.runs:
            if skip_empty and not (r.text or '').strip():
                stats['skipped_empty'] += 1
                continue
            for t in r._element.iter(qn('w:t')):
                if t.text and any(ch in smap for ch in t.text):
                    t.text = ''.join(smap.get(ch, ch) for ch in t.text)
                    stats['subscript'] += 1
            rPr = r._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
                rFonts.set(qn(attr), af)
            rFonts.set(qn('w:eastAsia'), east_asia)
            sz = rPr.find(qn('w:sz'))
            if sz is not None:
                szCs = rPr.find(qn('w:szCs'))
                if szCs is None:
                    szCs = OxmlElement('w:szCs')
                    sz.addnext(szCs)          # 必须紧跟 w:sz，勿 rPr.append
                    stats['szcs_added'] += 1
                szCs.set(qn('w:val'), sz.get(qn('w:val')) or '21')
            if narrow and r.font.size is not None and r.font.size.pt <= 7:
                r.font.bold = False
                stats['narrow_unbold'] += 1
            stats['runs'] += 1
    # Normal 样式同步，避免新增 run 回退
    try:
        style = doc.styles['Normal']
        style.font.name = af
        srFonts = style.element.rPr.rFonts
        for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
            srFonts.set(qn(attr), af)
        srFonts.set(qn('w:eastAsia'), east_asia)
    except (KeyError, AttributeError):
        pass
    return stats


# ---------------------------------------------------------------- CLI
def build_parser():
    ap = argparse.ArgumentParser(
        description='成稿原地修订：表格列宽两策略 + 字体归一（默认另存、不覆盖输入）')
    ap.add_argument('--input', required=True, help='输入 .docx')
    ap.add_argument('--output', required=True, help='输出 .docx（另存，不覆盖输入）')
    ap.add_argument('--backup', help='可选：动手前把输入复制到该路径')
    ap.add_argument('--strategy', choices=['autofit', 'content-width'],
                    default='content-width', help='列宽策略，默认 content-width')
    ap.add_argument('--max-col-ratio', type=float, default=0.6)
    ap.add_argument('--floor-twips', type=float, default=480.0)
    ap.add_argument('--max-rounds', type=int, default=10)
    ap.add_argument('--normalize-fonts', action='store_true')
    ap.add_argument('--font-mode', choices=['cjk-split', 'all-song'], default='cjk-split')
    ap.add_argument('--east-asia', default='宋体')
    ap.add_argument('--ascii-font', default=None)
    ap.add_argument('--dry-run', action='store_true', help='只统计不保存')
    return ap


def main(argv=None):
    _require_docx()
    args = build_parser().parse_args(argv)
    if args.dry_run:
        print('[dry-run] 不写任何文件')
    elif args.backup:
        shutil.copy2(args.input, args.backup)
        print('backup:', args.backup)
    doc = Document(args.input)
    # 先归一（下标替换会改变文本宽度），再按内容定宽
    if args.normalize_fonts:
        st = normalize_fonts(doc, font_mode=args.font_mode,
                             east_asia=args.east_asia, ascii_font=args.ascii_font)
        print('normalize_fonts:', st)
    if args.strategy == 'autofit':
        print('mark_all_autofit: %d tables' % mark_all_autofit(doc))
    else:
        info = resize_all(doc, max_col_ratio=args.max_col_ratio,
                          floor_twips=args.floor_twips, max_rounds=args.max_rounds)
        print('resize_all: %d tables' % len(info))
        for i, (nc, mn, mx, cw) in enumerate(info[:6]):
            print('  tbl#%d cols=%d min=%.2fcm max=%.2fcm content=%.2fcm'
                  % (i, nc, mn / 567, mx / 567, cw / 567))
    if args.dry_run:
        print('[dry-run] 结果未保存')
    else:
        doc.save(args.output)
        print('saved:', args.output)
    return 0


if __name__ == '__main__':
    sys.exit(main())
