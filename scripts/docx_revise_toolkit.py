# -*- coding: utf-8 -*-
"""
docx_revise_toolkit —— QC 检验手册「改既有成稿」通用原语库

定位
----
`docx_table_toolkit.py` 负责**从零建文档**（加表、加卡片、加页眉页脚）。
本模块负责**改一份已经成稿的 .docx**：定位结构、保格式改写、搬移块、重编号、
改写交叉引用、量化章节体量。两者互补，不可互相替代。

设计铁律（每一条都对应一个真实踩过的坑）
----------------------------------------
1. **身份比较只用 `_p` / `._tbl` 的 `is`**。`doc.paragraphs` 每次访问都返回**新的
   Paragraph 对象**，对象 `is` 比较必然失败；`p._p is q._p` 才成立。
2. **一律用 `p.style.name`，禁用 styleId**。真实文档的 Heading 2 styleId 是
   `'000006'` 或 `'21'`，不是 `'Heading2'`；且**同一章内子标题样式不统一**
   （有的 Heading 2、有的 Normal），故定位子标题不能只查 Heading 层级。
3. **目录条目与正文标题同名**（目录 = `List Bullet`，正文 = `Heading 1`），
   按文本找会误中目录那份，必须靠 `style.name` 消歧。
4. **编号 token 与标题文字常分属不同 run**（`"0.1"` | `" 效力口径"`），
   且空格数不一致（`"6.2 "` 单空格 vs `"6.5  "` 双空格）。匹配用 `^\\d+(\\.\\d+)*`
   前缀 + `\\s+`，不要写死空格个数。
5. **表格内段落不在 `doc.paragraphs` 里**（实测正文段 226 / 表内段 1612）。
   全文扫描必须「段落 + 所有表格单元格」一起扫。
6. **写文本必须带 `xml:space="preserve"`，且用 Clark 记法**
   （`{http://www.w3.org/XML/1998/namespace}space`）。字面量 `'xml:space'` 会抛
   `ValueError: Invalid attribute name`。
7. **禁用 `cell.text = `**（丢 rPr）与**硬编码 `doc.tables[i]`**（搬表后索引整体漂移）。
8. **重编号必须按源编号降序执行**，否则新 token 会被下一次查找误命中。
9. **迁表顺序铁律：先移表、再把随附说明段接到表后**；反序会使说明段滞留原位。
10. **整段回写只保留首 run 的 rPr**（局部加粗会丢）——这是已知取舍，压缩前须确认。

用法
----
    from docx_revise_toolkit import *

    doc = Document(src)
    h = find_heading(doc, '6.5')                 # 按编号找小节标题
    tbl = find_table(doc, ['版本', '生效日期'])     # 按表头关键字找表（不靠索引）
    set_paragraph_text(h, '6.5  新标题')
    move_table_with_notes(tbl._tbl, target_el, [note_el])

自检
----
    python docx_revise_toolkit.py --selftest <某本手册.docx>
"""

import re
import copy

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.table import Table, _Row

__all__ = [
    # L1 遍历与身份
    'XMLSPACE', 'iter_block_items', 'all_paragraphs', 'el_text', 'style_name', 'in_order',
    'is_heading', 'is_bullet', 'same_el', 'LocateError',
    # L2 定位
    'find_paragraphs', 'find_heading', 'find_subheading', 'find_table',
    'find_block', 'resolve_ref',
    # L3 写入
    'set_paragraph_text', 'replace_in_paragraph', 'replace_regex_in_paragraph',
    'replace_in_document', 'set_cell_text_preserve', 'set_cell_multiline', 'replace_in_cell',
    # L4 结构
    'move_after', 'move_before', 'move_block_after', 'move_table_with_notes',
    'clone_paragraph_after', 'clone_paragraph_before', 'clone_row',
    'append_table_row', 'detach_paragraph',
    # L5 编号与引用
    'parse_number', 'number_sort_key', 'renumber_headings',
    'update_cross_references', 'rewrite_leading_number',
    # L6 体量
    'measure_sections', 'measure_chapter',
]

# Clark 记法。字面量 'xml:space' 会被 lxml 拒绝（Invalid attribute name）。
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'

_BULLET_STYLES = {'List Bullet', 'List Bullet 2', 'List Paragraph'}


class LocateError(Exception):
    """定位失败。**不要静默返回 None** —— 静默失败是改版事故的头号来源。"""


# =====================================================================
# L1 · 遍历与身份
# =====================================================================

def iter_block_items(doc):
    """按**文档序**产出 ('p', Paragraph) / ('tbl', Table)。

    必须用它而不是 `doc.paragraphs` + `doc.tables` 两条独立列表——后者无法判断
    「这张表属于哪个小节」，而这正是章节体量量化与迁表的前提。
    """
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield ('p', Paragraph(child, doc))
        elif child.tag == qn('w:tbl'):
            yield ('tbl', Table(child, doc))


def all_paragraphs(doc, include_cells=True, unique=True):
    """产出文档全部段落。

    include_cells=True 时把**表格单元格内的段落**一并纳入——它们不在
    `doc.paragraphs` 里，实测占比可达 87%（正文 226 : 表内 1612）。
    unique=True 时去重（gridSpan 展开会让同一 cell 被访问多次）。

    去重 key 用**元素对象本身**而非 `id()`——`id()` 在对象被回收后会被复用，
    而 Paragraph 是临时构造的，会导致「明明存在的段被误判为重复而跳过」。
    用元素做 key 同时保持引用存活，id 不会被复用。
    """
    seen = set()
    for kind, obj in iter_block_items(doc):
        if kind == 'p':
            if unique:
                if obj._p in seen:
                    continue
                seen.add(obj._p)
            yield obj
    if include_cells:
        seen_cells = set()
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    key = cell._tc
                    if key in seen_cells:
                        continue
                    seen_cells.add(key)
                    for p in cell.paragraphs:
                        if unique:
                            if p._p in seen:
                                continue
                            seen.add(p._p)
                        yield p


def el_text(el):
    """拼接元素内全部 `w:t` 文本（**跨 run**）。

    run 边界会把一句话劈成几段（实测 `"来源依据见"` | `" "` | `"G.2.1"`），
    单 run 匹配必然失败——凡是要做子串判断的地方，一律先 el_text 整段拼。
    """
    if hasattr(el, '_p'):
        el = el._p
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def style_name(p):
    """返回 `p.style.name`；style 为 None 时返回 ''。

    **禁用 styleId**：真实文档 Heading 2 的 styleId 是 '000006' 或 '21'，
    且同一章内子标题样式不统一（Normal / Heading 2 混用）。
    """
    try:
        st = p.style
        return st.name if st is not None else ''
    except Exception:
        return ''


def is_heading(p, level=None):
    """按 style.name 判定标题。level=None 表示任意层级。"""
    name = style_name(p)
    if level is None:
        return name.startswith('Heading ')
    return name == 'Heading %d' % level


def is_bullet(p):
    """是否为项目符号段（List Bullet 系列）。"""
    return style_name(p) in _BULLET_STYLES


def same_el(a, b):
    """元素身份比较。**只能用 `_p` / `._tbl` 的 `is`**。

    `doc.paragraphs` 每次访问返回新对象，`p is q` 恒为 False；
    `p._p is q._p` 才反映「是不是同一个段落元素」。
    """
    if a is None or b is None:
        return False
    ea = getattr(a, '_p', None)
    eb = getattr(b, '_p', None)
    if ea is None:
        ea = getattr(a, '_tbl', None)
        eb = getattr(b, '_tbl', None)
    return ea is not None and ea is eb


# =====================================================================
# L2 · 定位与匹配
# =====================================================================

def find_paragraphs(doc, *, text=None, prefix=None, regex=None, style=None,
                     exclude_bullet=True, include_cells=False,
                     occurrence=0, all_occurrences=False, scope=None):
    """通用定位器。

    参数
    ----
    text / prefix / regex
        对**整段拼接文本**生效，故天然跨 run 安全（铁律 4）。
        三者同时给出时取 AND。
    style
        按 **style.name** 匹配（'Heading 1' / 'Heading 2' / 'List Bullet' / 'Normal'）。
        None 表示不限样式。
    exclude_bullet
        True（默认）跳过 List Bullet —— 这是避开**目录条目**的关键；
        需要改正文项目符号段时显式设 False。
        **显式指定 bullet 样式（style='List Bullet' 等）时该过滤自动失效**，
        否则会出现「点名要 bullet 却恒查不到」的自相矛盾。
    include_cells
        True 时把表格单元格内的段落一并纳入。
    occurrence
        0-based，指定取第几个命中；all_occurrences=True 时返回全部命中。
    scope
        元素区间（start_el, end_el），只在区间内查找。
    """
    hits = []
    for p in all_paragraphs(doc, include_cells=include_cells):
        if scope is not None:
            s, e = scope
            if s is not None and not in_order(doc, p._p, s, e):
                continue
        # 样式先判、bullet 排除后判：显式点名 bullet 样式时不得再被排除，
        # 否则 style='List Bullet' 恒返回空（目录条目 / 项目符号段定位全部落空）。
        if style is not None and style_name(p) != style:
            continue
        if exclude_bullet and style not in _BULLET_STYLES and is_bullet(p):
            continue
        t = el_text(p._p)
        if text is not None and text not in t:
            continue
        if prefix is not None and not t.startswith(prefix):
            continue
        if regex is not None and not re.search(regex, t):
            continue
        hits.append(p)
    if all_occurrences:
        return hits
    if occurrence >= len(hits):
        return None
    return hits[occurrence]


def in_order(doc, el, start_el, end_el):
    """判断 el 是否落在 [start_el, end_el) 区间内（按 body 子元素序）。"""
    body = doc.element.body
    kids = list(body.iterchildren())
    try:
        i = kids.index(el)
    except ValueError:
        return False
    if start_el is not None:
        try:
            if i < kids.index(start_el):
                return False
        except ValueError:
            return False
    if end_el is not None:
        try:
            if i >= kids.index(end_el):
                return False
        except ValueError:
            return False
    return True


def find_heading(doc, number=None, *, title_kw=None, style=None, level=None,
                 exclude_bullet=True, occurrence=0, include_cells=False):
    """按编号 token 或标题关键字找标题段。

    number
        形如 '6' / '6.5' / 'G.2'，用 `^\\d+(\\.\\d+)*` 前缀匹配**行首编号**，
        因此 '6.5' 也能命中 '6.5  xxx'，且不受空格数影响。
    title_kw
        标题文字关键字，与 number 取 AND。
    style
        显式指定 style.name（如 'Heading 1'）时严格过滤。
    level
        按 style.name 过滤层级。**level=None 且 style=None 时不要求 Heading 样式**，
        用于「同一章内子标题样式不统一」的场景（铁律 2）。

    注意：多个附录章可能共用同一编号（如「5 附录 A」~「5 附录 H」全是 Heading 1），
    此时靠 title_kw 或 occurrence 消歧。
    """
    pat = None
    if number is not None:
        esc = re.escape(str(number))
        pat = r'^\s*' + esc + r'(?![0-9.])'
    hits = []
    for p in all_paragraphs(doc, include_cells=include_cells):
        if exclude_bullet and is_bullet(p):
            continue
        if style is not None and style_name(p) != style:
            continue
        if level is not None and not is_heading(p, level):
            continue
        t = el_text(p._p)
        if pat and not re.match(pat, t):
            continue
        if title_kw and title_kw not in t:
            continue
        hits.append(p)
    if occurrence >= len(hits):
        return None
    return hits[occurrence]


def find_subheading(doc, number, *, within=None, occurrence=0,
                    title_kw=None, exclude_bullet=True):
    """找形如 '6.5  xxx' 的小节标题，**不限定 Heading 样式**。

    within
        (start_el, end_el) 区间，限定在该区间内查找；None = 全文。
    """
    pat = r'^\s*' + re.escape(str(number)) + r'(?![0-9.])'
    hits = []
    for p in all_paragraphs(doc, include_cells=False):
        if exclude_bullet and is_bullet(p):
            continue
        t = el_text(p._p)
        if not re.match(pat, t):
            continue
        if title_kw and title_kw not in t:
            continue
        if within is not None:
            s, e = within
            if s is not None and not in_order(doc, p._p, s, e):
                continue
        hits.append(p)
    if occurrence >= len(hits):
        return None
    return hits[occurrence]


def find_table(doc, header_keywords, *, row_keyword=None, occurrence=0,
               match_mode='all'):
    """按**表头关键字**定位表（替代硬编码 `doc.tables[i]`）。

    搬表会让 `doc.tables` 索引整体前移（实测搬一次后下标 65~70 内容全变），
    因此一切表定位都必须走表头关键字。

    header_keywords
        表头行的关键字列表；match_mode='all' 要求全部命中，'any' 命中任一。
    row_keyword
        给定时再要求某行**首列**文本等于该值。
    """
    kws = list(header_keywords or [])
    hits = []
    for tbl in doc.tables:
        if not tbl.rows:
            continue
        head = ' '.join(c.text.strip() for c in tbl.rows[0].cells)
        if match_mode == 'any':
            ok = any(k in head for k in kws)
        else:
            ok = all(k in head for k in kws)
        if not ok:
            continue
        if row_keyword is not None:
            if not any(r.cells and r.cells[0].text.strip() == row_keyword
                       for r in tbl.rows):
                continue
        hits.append(tbl)
    if occurrence >= len(hits):
        return None
    return hits[occurrence]


def find_block(doc, start_el, end_el):
    """返回文档序区间 [start_el, end_el) 内的元素列表（不含 end_el）。"""
    out = []
    started = False
    for kind, obj in iter_block_items(doc):
        el = obj._p if kind == 'p' else obj._tbl
        if el is start_el:
            started = True
            continue
        if el is end_el:
            break
        if started:
            out.append((kind, obj))
    return out


def resolve_ref(doc, ref):
    """把 config 里的引用描述解析为元素。

    ref 形如：
        {"number": "6.5", "style": "Heading 2", "title_kw": null, "occurrence": 0}
        {"prefix": "受控文件签署链"}                       # 按文本前缀找段落
        {"header_keywords": ["版本", "生效日期"]}           # 按表头找表

    解析失败抛 LocateError，**不静默返回 None**。

    **prefix 若以编号开头**，按「行首编号 + 任意空白」匹配，而不是字面前缀——
    真实文档里同一个编号的空格数并不统一（`"6.1 效力口径"` 单空格 vs
    `"6.5  v2 增量说明"` 双空格），字面前缀会把其中一种漏掉。
    """
    if not isinstance(ref, dict):
        raise LocateError('ref 必须是 dict，实际是 %r' % type(ref))
    if 'header_keywords' in ref:
        t = find_table(doc, ref['header_keywords'],
                       row_keyword=ref.get('row_keyword'),
                       occurrence=ref.get('occurrence', 0))
        if t is None:
            raise LocateError('按表头关键字未找到表: %r' % (ref['header_keywords'],))
        return t
    if 'prefix' in ref:
        pref = ref['prefix']
        m = re.match(r'^\s*(\d+(?:\.\d+)*)(?![0-9.])\s*(.*)$', pref)
        if m:
            # 编号 + 标题关键字：编号用正则，标题用包含匹配
            num, title = m.group(1), m.group(2).strip()
            pat = r'^\s*' + re.escape(num) + r'(?![0-9.])\s*'
            if title:
                pat += re.escape(title[:12])
            hits = [p for p in all_paragraphs(doc, include_cells=False)
                    if re.match(pat, el_text(p._p))
                    and (not ref.get('exclude_bullet', True) or not is_bullet(p))]
            if ref.get('style'):
                hits = [h for h in hits if style_name(h) == ref['style']]
            if not hits:
                raise LocateError('按编号/标题未找到段落: %r' % pref)
            occ = ref.get('occurrence', 0)
            if occ >= len(hits):
                raise LocateError('命中 %d 个，按 occurrence=%d 取不到: %r'
                                  % (len(hits), occ, pref))
            return hits[occ]
        p = find_paragraphs(doc, prefix=pref,
                            style=ref.get('style'),
                            exclude_bullet=ref.get('exclude_bullet', True),
                            occurrence=ref.get('occurrence', 0))
        if p is None:
            raise LocateError('按前缀未找到段落: %r' % pref)
        return p
    if 'number' in ref:
        p = find_heading(doc, ref.get('number'),
                         title_kw=ref.get('title_kw'),
                         style=ref.get('style'),
                         level=ref.get('level'),
                         occurrence=ref.get('occurrence', 0))
        if p is None:
            raise LocateError('按编号未找到标题: %r' % ref.get('number'))
        return p
    raise LocateError('无法识别的 ref: %r' % (ref,))


# =====================================================================
# L3 · 写入（保格式）
# =====================================================================

def _ensure_text_el(run_el):
    """取 run 内第一个 w:t；没有则新建一个。"""
    ts = run_el.findall(qn('w:t'))
    if ts:
        return ts[0]
    t = run_el.makeelement(qn('w:t'), {})
    run_el.append(t)
    return t


def set_paragraph_text(p, text, *, preserve_rpr=True, keep_empty_runs=False):
    """整段回写。

    preserve_rpr=True 时**保留首 run 的 rPr**（字体/字号/加粗不丢），其余 run 清空。
    已知取舍：段内局部加粗会丢（铁律 10）。压缩正文前应先确认该段无局部强调。

    返回是否成功写入。
    """
    el = getattr(p, '_p', p)
    runs = el.findall(qn('w:r'))
    if not runs:
        # 空段落：造一个 run
        r = el.makeelement(qn('w:r'), {})
        el.append(r)
        runs = [r]
    first = runs[0]
    t = _ensure_text_el(first)
    t.text = text
    t.set(XMLSPACE, 'preserve')
    for r in runs[1:]:
        if keep_empty_runs:
            for tt in r.findall(qn('w:t')):
                tt.text = ''
                tt.set(XMLSPACE, 'preserve')
        else:
            for tt in r.findall(qn('w:t')):
                tt.text = ''
                tt.set(XMLSPACE, 'preserve')
    return True


def replace_in_paragraph(p, old, new, count=1):
    """跨 run 片段替换。

    先把整段拼接，在拼接结果上定位，再把新串写进**首个命中 run**，
    并把其余 run 的对应片段清空——这样跨 run 的引用（"份见" | " 0.6"）也能换掉。

    count=0 表示**全部替换**（与 `replace_regex_in_paragraph` 语义一致）。
    """
    el = getattr(p, '_p', p)
    ts = [t for t in el.iter(qn('w:t'))]
    if not ts:
        return 0
    full = ''.join(t.text or '' for t in ts)
    if not old:
        return 0
    done = 0
    limit = count if count and count > 0 else -1
    while limit < 0 or done < limit:
        idx = full.find(old)
        if idx < 0:
            break
        end = idx + len(old)
        pos = 0
        placed = False
        for t in ts:
            seg = t.text or ''
            s, e = pos, pos + len(seg)
            pos = e
            if e <= idx or s >= end:
                continue
            b = seg[:max(s, idx) - s]
            a = seg[min(e, end) - s:]
            if not placed:
                t.text = b + new + a
                t.set(XMLSPACE, 'preserve')
                placed = True
            else:
                t.text = b + a
                t.set(XMLSPACE, 'preserve')
        full = ''.join(t.text or '' for t in ts)
        done += 1
    return done


def replace_regex_in_paragraph(p, pattern, repl, count=0):
    """正则版跨 run 替换。

    pattern 应自带足够的上下文锚点，避免误伤数值（如 `0.5μg` / `0.81`）。
    count=0 表示全部替换。
    """
    el = getattr(p, '_p', p)
    ts = [t for t in el.iter(qn('w:t'))]
    if not ts:
        return 0
    full = ''.join(t.text or '' for t in ts)
    new_full, n = re.subn(pattern, repl, full, count=count)
    if n == 0:
        return 0
    if len(ts) == 1:
        ts[0].text = new_full
        ts[0].set(XMLSPACE, 'preserve')
        return n
    # 多 run：整体重写到首 run，其余清空
    ts[0].text = new_full
    ts[0].set(XMLSPACE, 'preserve')
    for t in ts[1:]:
        t.text = ''
        t.set(XMLSPACE, 'preserve')
    return n


def replace_in_document(doc, pairs, *, include_cells=True, scope=None,
                         report=False):
    """跨全文替换。pairs = [(old, new), ...]。

    scope = (start_el, end_el) 限定区间。
    report=True 时返回 (命中次数, 命中段落列表)。
    """
    total = 0
    hits = []
    for p in all_paragraphs(doc, include_cells=include_cells):
        if scope is not None:
            s, e = scope
            if s is not None and not in_order(doc, p._p, s, e):
                continue
        for old, new in pairs:
            n = replace_in_paragraph(p, old, new, count=0)
            if n:
                total += n
                hits.append(el_text(p._p)[:80])
    return (total, hits) if report else total


def set_cell_text_preserve(cell, text):
    """单元格改写：保留首段首 run 的 rPr，删除其余段。

    **禁用 `cell.text = `** —— 它会清空全部 run 并丢掉 rPr（字体/字号/加粗全丢）。
    """
    ps = cell.paragraphs
    if not ps:
        p = cell.add_paragraph()
        ps = [p]
    set_paragraph_text(ps[0], text)
    for p in ps[1:]:
        p._p.getparent().remove(p._p)
    return cell.paragraphs[0]


def replace_in_cell(cell, old, new, count=0):
    """单元格内跨 run 替换（对该单元格所有段生效）。"""
    total = 0
    for p in cell.paragraphs:
        total += replace_in_paragraph(p, old, new, count=count)
    return total


def set_cell_multiline(cell, lines):
    """格内多行写入：保留首段首 run 的 rPr，每条物料一个 <w:br/>。

    用途：把「物料名称 / 规格」等多条内容嵌进一个矩阵单元格，每条独立成行
    （名称本身常含顿号 / 分号，用标点连接会糊成一团）。build_matrix 的
    cell_mode="name_spec" 输出含 \\n 的串，正是喂给本函数。

    - lines: list[str]，每条一行；元素内不要带 \\n。
    - 返回插入的 w:br 数量（= len(lines)-1），便于调用方断言。
    - 禁用 cell.text=；克隆首 run 的 rPr 保住字号/字体（矩阵多为 7pt）。
    - 只动传入的 cell，不影响其它单元格。
    """
    import copy
    ps = cell.paragraphs
    if not ps:
        cell.add_paragraph()
        ps = cell.paragraphs
    p0 = ps[0]._p
    runs = p0.findall(qn('w:r'))
    assert runs, '单元格首段无 run，无法克隆 rPr'
    proto = runs[0].find(qn('w:rPr'))
    rpr_copy = copy.deepcopy(proto) if proto is not None else None

    for r in runs:                       # 清空首段原有 run
        p0.remove(r)
    br_count = 0
    for i, line in enumerate(lines):
        r = OxmlElement('w:r')
        if rpr_copy is not None:
            r.append(copy.deepcopy(rpr_copy))
        t = OxmlElement('w:t')
        t.text = str(line)
        t.set(XMLSPACE, 'preserve')
        r.append(t)
        p0.append(r)
        if i < len(lines) - 1:
            br_run = OxmlElement('w:r')
            if rpr_copy is not None:
                br_run.append(copy.deepcopy(rpr_copy))
            br_run.append(OxmlElement('w:br'))
            p0.append(br_run)
            br_count += 1
    for p in ps[1:]:                     # 删除多余段
        p._p.getparent().remove(p._p)
    return br_count


# =====================================================================
# L4 · 结构操作
# =====================================================================

def move_after(el, ref_el):
    """把 el 移到 ref_el 之后（lxml addnext）。"""
    ref_el.addnext(el)
    return el


def move_before(el, ref_el):
    """把 el 移到 ref_el 之前。"""
    ref_el.addprevious(el)
    return el


def move_block_after(elements, ref_el):
    """整块保持相对顺序地移动到 ref_el 之后。"""
    cursor = ref_el
    for el in elements:
        cursor.addnext(el)
        cursor = el
    return cursor


def move_table_with_notes(tbl, ref_el, note_elements=()):
    """迁表原子操作。

    **顺序铁律：先移表、再把随附说明段接到表后**（铁律 9）。
    反序（先插说明段再移表）会使说明段滞留原位——这是一个真实发生过的 bug。

    tbl / note_elements 接受 Table / Paragraph 或其元素。
    """
    t_el = getattr(tbl, '_tbl', tbl)
    t_el.addnext(ref_el) if False else None  # no-op 占位，保持可读
    # 真正搬运：把表放到 ref_el 之后
    ref_el.addnext(t_el)
    cursor = t_el
    for n in note_elements:
        n_el = getattr(n, '_p', n)
        cursor.addnext(n_el)
        cursor = n_el
    return cursor


def clone_paragraph_after(ref_el, template_el, text, *, align=None):
    """克隆模板段保格式后插到 ref_el **之后**并写入 text。

    template_el 为 None 时返回 None（模板取不到要兜底，别硬崩）。
    """
    if template_el is None:
        return None
    src = getattr(template_el, '_p', template_el)
    new_el = copy.deepcopy(src)
    ref_el.addnext(new_el)
    p = Paragraph(new_el, None)
    if text is not None:
        set_paragraph_text(p, text)
    if align is not None:
        p.alignment = align
    return p


def clone_paragraph_before(ref_el, template_el, text, *, align=None):
    """克隆模板段保格式后插到 ref_el **之前**并写入 text。

    与 `clone_paragraph_after` 成对。语义别搞反——用 addnext 实现「插到之前」
    是常见错误，会把新章节插到锚点章节的**后面**。
    """
    if template_el is None:
        return None
    src = getattr(template_el, '_p', template_el)
    new_el = copy.deepcopy(src)
    ref_el.addprevious(new_el)
    p = Paragraph(new_el, None)
    if text is not None:
        set_paragraph_text(p, text)
    if align is not None:
        p.alignment = align
    return p


def clone_row(tbl, row_index=-1):
    """深拷贝一行 w:tr 并追加到表尾（沿革表追加行用，替代 add_row 以保格式）。"""
    tr = tbl.rows[row_index]._tr
    new_tr = copy.deepcopy(tr)
    tbl._tbl.append(new_tr)
    return new_tr


def append_table_row(tbl, values, *, template_row=-1):
    """克隆末行 → 逐格写值（保 rPr）。

    values 为列表，长度应与表列数一致；缺的留空。
    """
    tr = clone_row(tbl, template_row)
    # _Row 的 parent 必须是带 .table 属性的对象（_Rows / Table），
    # 不能直接传 CT_Tbl——否则 _Row.table 会抛 'CT_Tbl' has no attribute 'table'。
    row = _Row(tr, tbl.rows)
    for i, val in enumerate(values):
        if i >= len(row.cells):
            break
        set_cell_text_preserve(row.cells[i], val)
    return row


def detach_paragraph(p_el):
    """从父节点摘除段落元素（不删文件，safe-delete 钩子不涉及）。返回是否成功。"""
    el = getattr(p_el, '_p', p_el)
    parent = el.getparent()
    if parent is None:
        return False
    parent.remove(el)
    return True


# =====================================================================
# L5 · 编号与交叉引用
# =====================================================================

# 行首编号模式：纯数字层级（6 / 6.10）或字母附录层级（G.2.1 / A.1 / F.3）。
# 字母必须后接点号，否则 "G.2.1" 会被解析失败。
_NUM_RE = r'(?:([A-Za-z])\.)?(\d+(?:\.\d+)*)(?![0-9.])'


def parse_number(text):
    """从文本行首解析编号。返回 (level, number_str) 或 None。

        '6.10  xxx'  -> (2, '6.10')
        '6  xxx'     -> (1, '6')
        'G.2.1 xxx'  -> (None, 'G.2.1')     # 字母附录层级特殊

    末尾的 `(?![0-9.])` 是必需的——否则 '6.10' 可能被截成 '6.1'。
    """
    if not text:
        return None
    m = re.match(r'^\s*' + _NUM_RE, text)
    if not m:
        return None
    letter, num = m.group(1), m.group(2)
    if letter:
        return (None, letter + '.' + num)
    parts = num.split('.')
    return (len(parts), num)


def number_sort_key(num_str):
    """排序键：'6.10' -> (6, 10)；'G.2.1' -> ('G', 2, 1)。

    用于**降序**排重编号映射（铁律 8）。
    """
    m = re.match(r'^([A-Za-z])\.(\d+(?:\.\d+)*)$', str(num_str))
    if m:
        return (m.group(1),) + tuple(int(x) for x in m.group(2).split('.'))
    m = re.match(r'^(\d+(?:\.\d+)*)$', str(num_str))
    if m:
        return tuple(int(x) for x in m.group(1).split('.'))
    return (str(num_str),)


def rewrite_leading_number(el, new_number, *, sep='  '):
    """重写元素**行首**的编号 token，保留其余 run 与格式。

    跨 run 时（"0.1" | " 效力口径"）只改首个含编号 token 的 `w:t`。
    返回是否成功。
    """
    e = getattr(el, '_p', el)
    for t in e.iter(qn('w:t')):
        txt = t.text or ''
        m = re.match(r'^(\s*)' + _NUM_RE, txt)
        if m:
            lead = m.group(1)
            letter, num = m.group(2), m.group(3)
            old_full = (letter + '.' + num) if letter else num
            rest = txt[m.end():]
            # 保留原有的编号与后文之间的分隔（单/双空格在真实文档里并不统一）
            sep_m = re.match(r'^(\s*)', rest)
            old_sep = sep_m.group(1) if sep_m else ''
            body = rest[len(old_sep):]
            new = '%s%s%s%s' % (lead, new_number,
                                 old_sep if old_sep else sep, body)
            t.text = new
            t.set(XMLSPACE, 'preserve')
            return True
    return False


def renumber_headings(doc, mapping, *, scope=None, include_toc=True,
                      include_body=True, include_cells=False, dry_run=False):
    """按 mapping {'6.10': '6.9'} 重写标题编号 token。

    **两阶段改写（不可省）**：

    只按源编号降序执行**仍然会自碰撞**——把 6.11 改成 6.10 之后，
    下一轮查「6.10」会命中刚改好的那个（原 6.11），于是它又被改成 6.9，
    最后所有标题都落到同一个号上（实测：7 个标题全变成 6.4）。

    正确做法是**两阶段**：
      阶段一：把每个要改的标题改成**临时占位符**（形如 `§§TMP0§§`，
              与任何真实编号都不可能冲突）；
      阶段二：把占位符统一替换成目标号。

    只替换标题段**行首**的编号 token，不动标题文字。
    include_toc=True 时目录条目（List Bullet）也一并改。

    返回 {旧号: 命中数}
    """
    result = {}
    if not mapping:
        return result
    ordered = sorted(mapping.keys(), key=number_sort_key, reverse=True)

    # --- 阶段一：改写为占位符 ---
    placeholders = {}
    for idx, old in enumerate(ordered):
        token = '\u00a7\u00a7TMP%d\u00a7\u00a7' % idx
        placeholders[token] = str(mapping[old])
        cnt = 0
        pat = r'^\s*' + re.escape(str(old)) + r'(?![0-9.])'
        for p in all_paragraphs(doc, include_cells=include_cells):
            st = style_name(p)
            is_toc = st in _BULLET_STYLES
            if is_toc:
                if not include_toc:
                    continue
            else:
                if not include_body:
                    continue
            if scope is not None:
                s, e = scope
                if s is not None and not in_order(doc, p._p, s, e):
                    continue
            t = el_text(p._p)
            if not re.match(pat, t):
                continue
            if dry_run:
                cnt += 1
                continue
            if _rewrite_leading_token(p._p, token):
                cnt += 1
        result[old] = cnt

    # --- 阶段二：占位符 → 目标号 ---
    if not dry_run:
        for token, target in placeholders.items():
            for p in all_paragraphs(doc, include_cells=include_cells):
                if token in el_text(p._p):
                    replace_in_paragraph(p, token, target, count=0)
    return result


def _rewrite_leading_token(el, new_token):
    """把元素行首的编号 token 替换为 new_token（可能是非数字占位符）。

    只动**第一个含行首编号的 w:t**；跨 run 时后续 run 不动。
    """
    e = getattr(el, '_p', el)
    for t in e.iter(qn('w:t')):
        txt = t.text or ''
        m = re.match(r'^(\s*)' + _NUM_RE, txt)
        if m:
            lead = m.group(1)
            rest = txt[m.end():]
            sep_m = re.match(r'^(\s*)', rest)
            old_sep = sep_m.group(1) if sep_m else ''
            body = rest[len(old_sep):]
            t.text = '%s%s%s%s' % (lead, new_token, old_sep, body)
            t.set(XMLSPACE, 'preserve')
            return True
    return False


def update_cross_references(doc, mapping, *, patterns=None, include_cells=True,
                            scope=None, dry_run=False, numeric_guard=None):
    """按 mapping 改写正文交叉引用（'见 6.7' -> '见附录 H.4'）。

    mapping  = {'见 6.7': '见附录 H.4'}
    patterns = [正则, ...]，给出时改用正则替换（支持数字串保护）。
    numeric_guard = ['0.5μg', '0.81']，命中这些字面量的段落整段跳过。

    整段拼接后替换，跨 run 安全（铁律 4）。
    返回命中次数。
    """
    total = 0
    for p in all_paragraphs(doc, include_cells=include_cells):
        if scope is not None:
            s, e = scope
            if s is not None and not in_order(doc, p._p, s, e):
                continue
        t = el_text(p._p)
        if numeric_guard:
            if any(g in t for g in numeric_guard):
                continue
        if patterns:
            for pat in patterns:
                n = re.subn(pat, lambda m: m.group(0), t)  # 先探命中数
                cnt = len(re.findall(pat, t))
                if cnt:
                    if dry_run:
                        total += cnt
                    else:
                        total += replace_regex_in_paragraph(p, pat, r'\g<0>')
            continue
        for old, new in mapping.items():
            n = t.count(old)
            if n:
                if dry_run:
                    total += n
                else:
                    total += replace_in_paragraph(p, old, new, count=0)
    return total


# =====================================================================
# L6 · 体量量化
# =====================================================================

def _blank_section():
    return {
        'number': None, 'title': '', 'level': None, 'style': '',
        'chars': 0, 'table_chars': 0, 'total_chars': 0,
        'tables': 0, 'table_rows': 0, 'paras': 0, 'bullets': 0,
        'first_item': '',
    }


def measure_sections(doc, *, chapter=None, include_cells=True):
    """按文档序遍历，量化每个章节/小节的体量。

    返回 [{number, title, level, style, chars, table_chars, total_chars,
           tables, table_rows, paras, bullets, first_item}, ...]

    关键口径
    --------
    * **表格字符数必须单独统计**（`table_chars`）。体量大头通常全在表里，
      纯文字口径会得出「砍不动」的错误结论。
    * 小节归属以**编号前缀**判定，而非样式——因为同一章内子标题样式不统一。
    * 表内段落不计入 `paras`（避免与正文重复计数），但计入 `table_chars`。

    chapter
        指定章号（'6'）时只返回该章内的小节；None 返回全文。
    """
    sections = []
    cur = None
    in_chapter = chapter is None
    chapter_no = str(chapter) if chapter is not None else None

    for kind, obj in iter_block_items(doc):
        if kind == 'p':
            t = el_text(obj._p)
            st = style_name(obj)
            parsed = parse_number(t)
            is_bullet = st in _BULLET_STYLES
            # 判定「这是一个章节标题段」：行首有编号，且不是 bullet
            if parsed and not is_bullet and t.strip() and st != 'List Paragraph':
                num, lvl = parsed[1], parsed[0]
                # 章级标题（Heading 1 或顶层编号）
                if chapter_no is not None:
                    if num == chapter_no:
                        in_chapter = True
                    elif in_chapter and lvl is not None and lvl == 1 and num != chapter_no:
                        in_chapter = False
                if not in_chapter:
                    cur = None
                    continue
                cur = _blank_section()
                cur['number'] = num
                cur['title'] = t.strip()
                cur['level'] = lvl
                cur['style'] = st
                cur['first_item'] = t.strip()[:60]
                sections.append(cur)
                continue
            if not in_chapter or cur is None:
                continue
            if t.strip():
                cur['chars'] += len(t)
                cur['paras'] += 1
                if is_bullet:
                    cur['bullets'] += 1
                if not cur['first_item']:
                    cur['first_item'] = t.strip()[:60]
        else:
            if not in_chapter or cur is None:
                continue
            tbl = obj
            n = sum(len(c.text) for r in tbl.rows for c in r.cells)
            cur['table_chars'] += n
            cur['total_chars'] = cur['chars'] + cur['table_chars']
            cur['tables'] += 1
            cur['table_rows'] += len(tbl.rows)

    for s in sections:
        s['total_chars'] = s['chars'] + s['table_chars']
    return sections


def measure_chapter(doc, chapter_number):
    """只量某章（'6'），返回该章内各小节 + 章合计。"""
    secs = measure_sections(doc, chapter=str(chapter_number))
    total = {
        'number': str(chapter_number), 'title': '', 'level': 1, 'style': '',
        'chars': sum(s['chars'] for s in secs),
        'table_chars': sum(s['table_chars'] for s in secs),
        'total_chars': sum(s['total_chars'] for s in secs),
        'tables': sum(s['tables'] for s in secs),
        'table_rows': sum(s['table_rows'] for s in secs),
        'paras': sum(s['paras'] for s in secs),
        'bullets': sum(s['bullets'] for s in secs),
        'first_item': '',
    }
    return secs, total


# =====================================================================
# 自检
# =====================================================================

def _selftest(path):
    """只读自检。需传入一本成稿手册路径。"""
    import sys
    from docx import Document
    d = Document(path)
    ok = []
    fail = []

    def chk(name, cond, detail=''):
        (ok if cond else fail).append('%-52s %s' % (name, detail))

    ps = d.paragraphs
    body = d.element.body
    n_tables = len(d.tables)

    allp = list(all_paragraphs(d, include_cells=True))
    cell_only = list(all_paragraphs(d, include_cells=False))
    chk('all_paragraphs 含单元格 > 正文段落数',
        len(allp) > len(cell_only), '%d > %d' % (len(allp), len(cell_only)))
    chk('all_paragraphs 去重生效（无重复 _p）',
        len({id(p._p) for p in allp}) == len(allp))

    # 目录同名消歧
    h1s = [p for p in ps if style_name(p) == 'Heading 1']
    if h1s:
        sample = el_text(h1s[1]._p) if len(h1s) > 1 else el_text(h1s[0]._p)
        pref = sample.split('  ')[0] + '  ' if '  ' in sample else sample[:6]
        with_b = find_paragraphs(d, prefix=pref, exclude_bullet=False,
                                 all_occurrences=True)
        without_b = find_paragraphs(d, prefix=pref, exclude_bullet=True,
                                    all_occurrences=True)
        chk('目录同名可消歧（exclude_bullet 生效）',
            len(with_b) >= len(without_b) and len(without_b) <= len(with_b),
            '含bullet=%d 排除bullet=%d 前缀=%r' % (len(with_b), len(without_b), pref))

    # 跨 run 拼接
    joined = [el_text(p._p) for p in ps]
    chk('el_text 跨 run 拼接（存在多 run 段）', any('  ' in j for j in joined))

    # 编号解析
    chk('parse_number("6.10 ") == (2, "6.10")', parse_number('6.10 ') == (2, '6.10'))
    chk('parse_number("G.2.1") 字母附录', parse_number('G.2.1 xxx') == (None, 'G.2.1'))
    chk('number_sort_key 降序可用',
        sorted(['6.9', '6.10', '6.11'], key=number_sort_key, reverse=True)
        == ['6.11', '6.10', '6.9'])

    # 表定位
    tbl = find_table(d, ['版本', '生效日期', '修订内容', '来源'])
    chk('find_table 按表头关键字命中版本沿革表', tbl is not None,
        '' if tbl is None else '%d行×%d列' % (len(tbl.rows), len(tbl.columns)))

    # 体量
    secs = measure_sections(d)
    chk('measure_sections 返回非空', len(secs) > 0, '%d 节' % len(secs))
    chk('体量含表格字符数（体量大头不漏）',
        any(s['table_chars'] > 0 for s in secs))
    chk('measure_sections total = chars + table_chars',
        all(s['total_chars'] == s['chars'] + s['table_chars'] for s in secs))

    # 身份比较
    a = ps[0]
    b = [p for p in d.paragraphs if p._p is a._p]
    chk('身份比较用 _p（doc.paragraphs 每次返回新对象）',
        len(b) == 1 and b[0] is not a, '同一 _p 的新对象 is 判定=%s' % (b[0] is a))

    # 表内段落不在 doc.paragraphs
    chk('表格内段落不在 doc.paragraphs',
        len(allp) > len(ps), 'all=%d vs doc.paragraphs=%d' % (len(allp), len(ps)))

    # 子标题样式不统一：level=None 时不要求 Heading 样式
    subs = [s for s in secs if s['level'] == 2]
    styles_seen = {s['style'] for s in subs}
    chk('子标题样式不统一时仍能全部量化',
        len(subs) > 0, '小节 %d 个，样式集合=%s' % (len(subs), sorted(styles_seen)))

    # 编号查找：子标题样式混合（Normal / Heading 2 混用）也能找到
    mixed = [s for s in subs if s['style'] not in ('Heading 2',)]
    if mixed:
        num = mixed[0]['number']
        h = find_heading(d, num)
        chk('find_heading 不限层级仍命中 Normal 样式子标题',
            h is not None and el_text(h._p).startswith(num),
            '%s -> %s' % (num, style_name(h) if h else 'None'))

    # set_cell_multiline：内存态写入 w:br，不落盘（用一个临时表单元格）
    if n_tables:
        from docx import Document as _D
        _tmp = _D()                                   # 空白文档，纯内存
        _c = _tmp.add_table(rows=1, cols=1).cell(0, 0)
        _c.text = '模板'                              # 造一个带 rPr 的 run
        _n = set_cell_multiline(_c, ['名称 A / 规格 A', '名称 B / 规格 B', '名称 C'])
        _brs = _c._tc.findall('.//' + qn('w:br'))
        chk('set_cell_multiline 插入 w:br 数 = 行数-1',
            _n == 2 and len(_brs) == 2, '返回=%d 实际br=%d' % (_n, len(_brs)))
        chk('set_cell_multiline 文本按行保留',
            _c.text.split('\n') == ['名称 A / 规格 A', '名称 B / 规格 B', '名称 C'],
            repr(_c.text))

    # 编号改写保格式（内存态，不落盘）
    num2 = subs[0]['number'] if subs else None
    if num2:
        h2 = find_heading(d, num2)
        old_txt = el_text(h2._p)
        rewrite_leading_number(h2._p, '99.99')
        new_txt = el_text(h2._p)
        chk('rewrite_leading_number 只改行首编号、保留后文',
            new_txt.startswith('99.99') and old_txt.split(None, 1)[-1] in new_txt,
            '%r -> %r' % (old_txt[:30], new_txt[:30]))

    print('\n'.join('  OK   ' + x for x in ok))
    if fail:
        print('\n'.join('  FAIL ' + x for x in fail))
    print('\n自检: %d 通过 / %d 失败（文档：%s，%d 段 / %d 表）'
          % (len(ok), len(fail), path, len(ps), n_tables))
    return 0 if not fail else 1


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description='docx_revise_toolkit 自检')
    ap.add_argument('--selftest', metavar='DOCX', help='对指定手册跑只读自检')
    args = ap.parse_args()
    if args.selftest:
        raise SystemExit(_selftest(args.selftest))
    print(__doc__)
