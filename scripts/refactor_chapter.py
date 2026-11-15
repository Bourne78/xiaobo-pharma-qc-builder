# -*- coding: utf-8 -*-
"""
refactor_chapter —— config 驱动的章节重构 runner

用途
----
把「合并两章 / 重编号 / 删小节 / 明细表迁入新附录 / 压缩正文 / 修尾部结构 /
正式升版」这些改版动作，从一次性脚本固化成 **JSON 配置 + 固定阶段流水线**。
下次改版只写配置，不改代码。

为什么是固定阶段，而不是自由 steps 数组
--------------------------------------
章节重构有**内在因果序**：

    merge → delete → renumber → move_tables → compress → tail_fix → release

  * 先合并/删除，才有稳定的编号空间可重排；
  * 先重编号，迁表时的「交叉引用改指」才知道旧号是什么；
  * 迁表会改变文档序，压缩阶段必须在其后做；
  * 尾部修复依赖最终的章节结构；
  * 升版（封面/沿革/增量说明）必须最后做。

用自由 steps 数组极易写出「重编号先于合并」这类顺序错——而顺序错恰恰是
重编号碰撞（同源坑）的温床。固定阶段把顺序固化在代码里，config 只描述
**每个阶段做什么**。

安全约定
--------
* **只写 `--output`，永不覆盖 `--input`**（除非显式 `allow_overwrite_input`）。
* 每个阶段前后记录结构签名（段数 / 表数 / 每表维度），出现预期外的结构变化
  即抛 `StructureDrift` 并中止——专门防「迁表后说明段滞留原位」被静默接受。
* `--dry-run` 只打印「将命中什么 → 将执行什么」，不写文件。
* 阶段 7（release）收尾自检：增量说明标题下须有非 bullet 正文段；结尾标记
  须独立成段且全文唯一。不满足直接抛错，不写文件。

用法
----
    python refactor_chapter.py --input X.docx --output Y.docx --config cfg.json
    python refactor_chapter.py --input X.docx --output Y.docx --config cfg.json --dry-run
"""

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                                    # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH               # noqa: E402

from docx_revise_toolkit import (                            # noqa: E402
    LocateError, all_paragraphs, append_table_row, clone_paragraph_after,
    clone_paragraph_before, el_text, find_heading,
    find_paragraphs, find_table, in_order, is_bullet, measure_sections, move_after,
    move_block_after, move_table_with_notes, parse_number, renumber_headings,
    replace_in_cell, replace_in_paragraph, resolve_ref, set_cell_text_preserve,
    set_paragraph_text, style_name, update_cross_references,
)

SCHEMA_VERSION = '1.0'

STAGES = ('merge_sections', 'delete_sections', 'renumber', 'move_tables',
          'compress', 'tail_fix', 'release')


class StructureDrift(Exception):
    """阶段执行后文档结构出现预期外变化。"""


class ConfigError(Exception):
    """配置不合法。"""


# =====================================================================
# 结构签名
# =====================================================================

def structure_signature(doc):
    """文档结构签名。

    关键：**表的指纹按内容计算且用多重集合比对**，不能用「按序索引的维度」——
    迁表会改变 `doc.tables` 的顺序索引（同一张表跑到别处去了），
    按索引比对会把「表被搬动」误报成「表被改坏了」。
    """
    paras = len(doc.paragraphs)
    dims = tuple((len(t.rows), len(t.columns)) for t in doc.tables)
    fingerprints = tuple(sorted(_table_fingerprint(t) for t in doc.tables))
    return {'paras': paras, 'tables': len(dims), 'dims': dims,
            'fingerprints': fingerprints}


def _table_fingerprint(tbl):
    """表的内容指纹：维度 + 每格文本。用于「表有没有被改坏」的判定。"""
    cells = []
    for r in tbl.rows:
        cells.append('\x1f'.join(c.text for c in r.cells))
    return '%dx%d|%s' % (len(tbl.rows), len(tbl.columns), '\x1e'.join(cells))


def _dims_changed(before, after, idx):
    """比对第 idx 张表的维度是否变化。"""
    b = before['dims']
    a = after['dims']
    if idx >= len(b) or idx >= len(a):
        return True
    return b[idx] != a[idx]


# =====================================================================
# Runner
# =====================================================================

class RefactorRunner(object):

    def __init__(self, doc, cfg, dry_run=False, verbose=True):
        self.doc = doc
        self.cfg = cfg
        self.dry_run = dry_run
        self.verbose = verbose
        self.log = []
        self.actions = []
        # 阶段内已创建的附录标题（文本 → 元素），避免重复创建同名附录
        self._created_appendices = {}
        # 本轮已删除的小节号（其号已空出，可被重编号目标号占用）
        self._freed_numbers = set()

    # ---------- 日志 ----------
    def _say(self, msg):
        self.log.append(msg)
        if self.verbose:
            print(msg)

    def _record(self, stage, action, detail=''):
        self.actions.append({'stage': stage, 'action': action, 'detail': detail})

    # ---------- 阶段 ----------
    def run(self):
        self._check_schema()
        for stage in STAGES:
            spec = self.cfg.get(stage)
            if not spec:
                continue
            before = structure_signature(self.doc)
            # 本阶段**声明**要改动的表：它们的指纹变化是合法预期，
            # 不能被漂移检测误报成「表被改坏」（如 release 给沿革表加一行）。
            declared = self._declared_tables(stage, spec)
            declared_before = [_table_fingerprint(t) for t in declared]
            self._say('\n--- 阶段 %s ---' % stage)
            getattr(self, '_stage_' + stage)(spec)
            after = structure_signature(self.doc)
            allowed = [(bf, _table_fingerprint(tbl))
                       for tbl, bf in zip(declared, declared_before)]
            self._check_drift(stage, before, after, allowed)
        return self

    def _declared_tables(self, stage, spec):
        """定位本阶段**声明**允许改动内容的表（漂移检测白名单）。

        只有「配置里明确点名要改的表」才进白名单：
        - release → changelog_table（版本沿革表新增行）
        - compress → kind='table_cell' 的目标表（单元格内容压缩）
        其余任何表的内容变化都算漂移。
        """
        out = []
        try:
            if stage == 'release':
                ct = (spec or {}).get('changelog_table')
                if ct:
                    t = find_table(self.doc, ct.get('header_keywords', []),
                                   match_mode=ct.get('match_mode', 'all'))
                    if t is not None:
                        out.append(t)
            elif stage == 'compress':
                for item in spec or []:
                    if not isinstance(item, dict) or item.get('kind') != 'table_cell':
                        continue
                    tgt = item.get('target') or {}
                    hdr = tgt.get('header_keywords')
                    if not hdr:
                        continue
                    t = find_table(self.doc, hdr,
                                   match_mode=tgt.get('match_mode', 'all'))
                    if t is not None:
                        out.append(t)
        except Exception:
            return []
        return out

    def _check_schema(self):
        v = str(self.cfg.get('$schema_version', ''))
        if v != SCHEMA_VERSION:
            raise ConfigError(
                'config 的 $schema_version 为 %r，本脚本只支持 %r。'
                '（不认识就报错——防止旧配置静默跑偏。）' % (v, SCHEMA_VERSION))
        io = self.cfg.get('io') or {}
        if not io.get('input'):
            raise ConfigError('config 缺少 io.input')
        if not io.get('output'):
            raise ConfigError('config 缺少 io.output')

    def _check_drift(self, stage, before, after, allowed=None):
        """结构漂移检测。

        判定口径：**表内容指纹的多重集合**必须保持一致。
        - 迁表 / 移动：指纹集合不变 → 放行（顺序变了不算漂移）。
        - 误删列 / 误改单元格内容：指纹集合变化 → 报漂移并中止。

        表**数量**变化是合法的（删小节会带走表），不在此处拦截；
        数量变化后的内容一致性由 `moved_tables` 验收断言负责。

        `allowed` 是 (改前指纹, 改后指纹) 对——**本阶段配置里明确声明要改的表**。
        这类表的内容变化是预期行为，不算漂移（release 给沿革表加一行是常态）。
        """
        if before['tables'] != after['tables']:
            return
        if before['fingerprints'] == after['fingerprints']:
            return
        allowed = list(allowed or [])
        allowed_before = {b for b, _ in allowed}
        allowed_after = {a for _, a in allowed}
        bset = set(before['fingerprints'])
        aset = set(after['fingerprints'])
        lost = (bset - aset) - allowed_before
        added = (aset - bset) - allowed_after
        if not lost and not added:
            return
        raise StructureDrift(
            '阶段 %s 后表格内容指纹发生变化（表数 %d 未变）：\n'
            '      丢失/被改：%s\n      新增：%s\n'
            '      迁表只应改变表的位置，不应改变表内容。已中止。'
            % (stage, before['tables'],
               ' / '.join(sorted(x[:60] for x in list(lost)[:2])) or '（无）',
               ' / '.join(sorted(x[:60] for x in list(added)[:2])) or '（无）'))

    # ---------- 阶段 1：合并 ----------
    def _stage_merge_sections(self, spec):
        for m in spec:
            src = resolve_ref(self.doc, m['source_heading'])
            dst = resolve_ref(self.doc, m['target_heading'])
            self._say('  合并 %s → %s'
                      % (el_text(src._p)[:28], el_text(dst._p)[:28]))

            # 收集源章的区块（到下一个同级标题为止）
            blocks = self._collect_section_blocks(src)
            self._say('    源章区块 %d 项（含 %d 张表）'
                      % (len(blocks), sum(1 for k, _ in blocks if k == 'tbl')))

            keep = m.get('keep_position', 'target')
            if keep == 'target':
                # 把源区块整体搬到目标章标题之前
                anchor = dst._p
                els = [o._p if k == 'p' else o._tbl for k, o in blocks]
                move_block_after(els, anchor.getprevious())
                self._record('merge', 'move_source_blocks_to_target',
                             '%d 项' % len(els))
                # 源标题本身搬走
                move_after(src._p, els[-1] if els else anchor)
                self._record('merge', 'move_source_heading', el_text(src._p)[:40])
            else:
                self._record('merge', 'keep_source_position', el_text(src._p)[:40])

            # 改标题
            if m.get('rename_source_to'):
                set_paragraph_text(src, m['rename_source_to'])
                self._record('merge', 'rename_heading', m['rename_source_to'])
            if m.get('retitle_target'):
                set_paragraph_text(dst, m['retitle_target'])
                self._record('merge', 'retitle_target', m['retitle_target'])

            # 子编号映射
            numbering = m.get('subsection_numbering') or {}
            mapping = numbering.get('map') or {}
            if mapping:
                r = renumber_headings(self.doc, mapping,
                                      include_toc=numbering.get('apply_to_toc', True),
                                      dry_run=self.dry_run)
                self._say('    子节重编号 %d 项：%s' % (len(mapping), r))
                self._record('merge', 'renumber_subsections', str(r))

            # 目录条目处理
            if m.get('delete_source_toc_entries', True):
                self._delete_toc_entry(m.get('source_heading'))

    def _collect_section_blocks(self, heading_p):
        """收集标题段之后、到下一个同级标题之前的全部区块。"""
        from docx_revise_toolkit import iter_block_items
        body = self.doc.element.body
        kids = list(body.iterchildren())
        try:
            start = kids.index(heading_p._p)
        except ValueError:
            return []
        lvl = parse_number(el_text(heading_p._p))
        level = lvl[0] if lvl else 1
        out = []
        for kind, obj in iter_block_items(self.doc):
            el = obj._p if kind == 'p' else obj._tbl
            try:
                i = kids.index(el)
            except ValueError:
                continue
            if i <= start:
                continue
            if kind == 'p':
                t = el_text(el)
                parsed = parse_number(t)
                if parsed and not is_bullet(obj) and parsed[0] is not None \
                        and parsed[0] <= level:
                    break  # 遇到同级或更高级标题，结束
            out.append((kind, obj))
        return out

    def _delete_toc_entry(self, heading_ref):
        """删除目录中与该标题同名的条目（List Bullet）。

        匹配口径：目录条目与正文标题**文本相同但样式不同**（目录=List Bullet，
        正文=Heading N）。这里按**编号前缀 + 标题关键字**匹配，
        不依赖样式——目录条目的样式本来就与正文标题不同。
        """
        if not heading_ref:
            return
        number = heading_ref.get('number')
        title = heading_ref.get('title_kw')
        pat = None
        if number:
            pat = r'^\s*' + re.escape(str(number)) + r'(?![0-9.])'
        for p in find_paragraphs(self.doc, style='List Bullet',
                                 all_occurrences=True):
            t = el_text(p._p)
            if pat and not re.match(pat, t):
                continue
            if title and title not in t:
                continue
            if not pat and not title:
                continue
            self._say('    删除目录条目: %s' % t.strip()[:50])
            self._record('delete', 'delete_toc_entry', t.strip()[:50])
            if not self.dry_run:
                p._p.getparent().remove(p._p)

    # ---------- 阶段 2：删除小节 ----------
    def _stage_delete_sections(self, spec):
        for d in spec:
            h = resolve_ref(self.doc, d['heading'])
            blocks = self._collect_section_blocks(h)
            n_bullet = sum(1 for k, o in blocks
                           if k == 'p' and is_bullet(o))
            self._say('  删除小节 %s（区块 %d 项 / bullet %d 条）'
                      % (el_text(h._p)[:34], len(blocks), n_bullet))
            # 记录被删小节的号：其号已空出，可被后续重编号的目标号占用
            from docx_revise_toolkit import parse_number as _pn
            _pnum = _pn(el_text(h._p))
            if _pnum and _pnum[1]:
                self._freed_numbers.add(_pnum[1])
            self._record('delete', 'delete_section',
                         '%s | %s' % (el_text(h._p)[:40], d.get('reason', '')))
            if self.dry_run:
                continue
            for kind, obj in blocks:
                el = obj._p if kind == 'p' else obj._tbl
                el.getparent().remove(el)
            h._p.getparent().remove(h._p)
            # **同步删除目录条目**——否则目录里仍留着小节号，
            # 后续重编号会把该号分配出去，产生重复编号。
            self._delete_toc_entry(d.get('heading'))

    # ---------- 阶段 3：重编号 ----------
    def _stage_renumber(self, spec):
        mapping = spec.get('map') or {}
        if not mapping:
            return
        scope = None
        if spec.get('scope', {}).get('chapter'):
            ch = find_heading(self.doc, str(spec['scope']['chapter']))
            if ch:
                scope = (ch._p, None)

        # **冲突预检**：目标号若已存在于文档中，重编号必然产生重复编号。
        # 例：把 6.10→6.9、6.11→6.10，而文档里本来就有一个 6.9，
        # 结果会出现三个「6.9」。这是配置错误，必须在动手前报出来。
        #
        # 例外：**本轮已删除的小节，其号已空出**，可以被目标号占用
        # （删掉 6.4 后再把 6.5→6.4 是合法的连续区间调整）。
        existing = set(self._collect_existing_numbers(scope))
        freed = set(self._freed_numbers)
        new_nums = {str(v) for v in mapping.values()}
        src_nums = {str(k) for k in mapping.keys()}
        conflict = sorted((new_nums & existing) - src_nums - freed)
        if conflict:
            raise ConfigError(
                '重编号目标号与文档中已存在的编号冲突：%s。\n'
                '      文档现有编号示例：%s\n'
                '      请把映射设计成**连续区间**（如 6.9→6.10、6.10→6.11、'
                '6.11→6.12），而不是交换。已中止，未写文件。'
                % (conflict, sorted(existing)[:14]))

        r = renumber_headings(self.doc, mapping, scope=scope,
                              include_toc=spec.get('include_toc', True),
                              include_body=spec.get('include_body', True),
                              dry_run=self.dry_run)
        self._say('  重编号 %d 项：%s' % (len(mapping), r))
        self._record('renumber', 'renumber', str(r))

        # 事后校验：不得出现重复编号
        after_list = self._collect_existing_numbers(scope)
        after = set(after_list)
        dups = sorted(n for n in after if after_list.count(n) > 1)
        if dups:
            raise StructureDrift(
                '重编号后出现重复编号：%s。降序执行虽能避免自碰撞，'
                '但目标号与现存号相撞仍会重复。已中止，未写文件。' % dups)

        if spec.get('update_cross_references', True):
            pairs = {}
            for old, new in mapping.items():
                pairs['见 %s' % old] = '见 %s' % new
                pairs['见 %s）' % old] = '见 %s）' % new
            if pairs:
                n = update_cross_references(
                    self.doc, pairs,
                    patterns=spec.get('cross_reference_patterns'),
                    numeric_guard=spec.get('numeric_guard'),
                    dry_run=self.dry_run)
                self._say('    交叉引用改写 %d 处' % n)
                self._record('renumber', 'update_cross_refs', '%d 处' % n)

    def _collect_existing_numbers(self, scope=None):
        """收集范围内已存在的行首编号（含目录条目，去重前保留重复以便检测）。"""
        nums = []
        for p in all_paragraphs(self.doc, include_cells=False):
            if scope is not None:
                s, e = scope
                if s is not None and not in_order(self.doc, p._p, s, e):
                    continue
            t = el_text(p._p)
            m = re.match(r'^\s*(\d+(?:\.\d+)*)(?![0-9.])', t)
            if m:
                nums.append(m.group(1))
        return nums

    # ---------- 阶段 4：迁表 ----------
    def _stage_move_tables(self, spec):
        for mv in spec:
            tbl = resolve_ref(self.doc, mv['table'])
            self._say('  迁表：表头 %s' % mv['table'].get('header_keywords'))

            # 目标锚点
            tgt = mv.get('target_appendix') or {}
            anchor = self._ensure_appendix(tgt)

            # 随附说明段（必须随表一起迁）
            notes = []
            for cn in mv.get('carry_notes') or []:
                np_ = find_paragraphs(self.doc, prefix=cn['prefix'],
                                      all_occurrences=True)
                if np_:
                    notes.append(np_[0])
                    self._say('    随附说明段: %s' % cn['prefix'][:40])

            if self.dry_run:
                self._record('move_tables', 'would_move', str(mv['table']))
                continue

            # 顺序铁律：先移表，再接说明段
            move_table_with_notes(tbl._tbl, anchor, [n._p for n in notes])
            self._record('move_tables', 'moved',
                         '%s (%d行×%d列) + %d 说明段'
                         % (mv['table'].get('header_keywords'),
                            len(tbl.rows), len(tbl.columns), len(notes)))

            # 迁后引用改指
            for ur in mv.get('update_references') or []:
                scope_mode = ur.get('scope', 'paragraph_and_cells')
                old, new = ur['from'], ur['to']
                occ = ur.get('occurrence', 0)
                if occ == -1:
                    n = self._replace_occurrence(old, new, scope_mode, -1)
                else:
                    n = self._replace_occurrence(old, new, scope_mode, occ)
                self._say('    引用改指 %r → %r（%d 处）'
                          % (old, new, n))
                self._record('move_tables', 'update_ref',
                             '%s→%s (%d)' % (old, new, n))
                # 硬契约：迁表后指向旧位置的交叉引用**必须**被改指。
                # 0 处命中通常意味着引用串写错（含全角括号与否、子节号是否存在），
                # 静默通过会留下指向旧位置的引用——必须报错。
                if n == 0 and not ur.get('allow_zero'):
                    raise ConfigError(
                        '引用改指 0 处命中：%r → %r。'
                        '请核对文档中该引用的**真实写法**'
                        '（是否含全角括号 / 子节号是否存在 / 是否跨 run 拆分）；'
                        '确实无需改指时显式加 "allow_zero": true。'
                        % (old, new))

    def _replace_occurrence(self, old, new, scope_mode, occurrence):
        """替换第 occurrence 处 old。scope_mode 决定是否含单元格。

        occurrence = -1 表示**全部替换**；否则只改第 N 处（0-based）。
        """
        n = 0
        targets = []
        if scope_mode in ('paragraph', 'paragraph_and_cells'):
            for p in self.doc.paragraphs:
                if old in el_text(p._p):
                    targets.append(p)
        if scope_mode in ('cells', 'paragraph_and_cells'):
            for t in self.doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        if old in c.text:
                            targets.append(c)
        if occurrence >= 0:
            targets = targets[occurrence:occurrence + 1]
        for obj in targets:
            if hasattr(obj, '_tc'):
                n += replace_in_cell(obj, old, new)
            else:
                n += replace_in_paragraph(obj, old, new)
        return n

    def _resolve_heading_loose(self, ref, exclude_bullet=True):
        """宽松解析标题引用。

        真实文档里**附录标题的编号不在行首**——形态是 `5  附录 H  源文件清单…`，
        字母 H 嵌在标题中间。因此按 `number='H'` 匹配必然落空。
        这里按三级回退：number+style → title_kw → 标题含关键词。

        `insert_before` 这类**位置锚点解析失败必须抛错**，不能静默回退到文末
        （静默回退会把附录插到错误位置，且看起来"成功"了）。

        exclude_bullet=True 时跳过 List Bullet —— **这是必须的**：
        目录里也有「6 编制说明…」条目，且与正文第 6 章同名。
        不排除 bullet 会把锚点解析到目录那一条，附录就被插到目录后面去了
        （实测症状：附录被插到 §6 之后而不是之前）。
        """
        number = ref.get('number')
        title_kw = ref.get('title_kw')
        style = ref.get('style')
        if number:
            p = find_heading(self.doc, str(number), title_kw=title_kw, style=style,
                             exclude_bullet=exclude_bullet)
            if p is not None:
                return p
        if title_kw:
            hits = [x for x in find_paragraphs(self.doc, all_occurrences=True,
                                               exclude_bullet=exclude_bullet)
                    if title_kw in el_text(x._p)]
            if style:
                hits = [x for x in hits if style_name(x) == style]
            if hits:
                return hits[0]
        raise LocateError('无法解析标题引用: %r' % (ref,))

    def _ensure_appendix(self, tgt):
        """确保目标附录章存在，返回其标题元素（作为插入锚点）。

        同一阶段内多次迁表到**同一个新附录**时，必须复用同一个标题段，
        否则会创建出两个同名附录（实测会生成两个「5  附录 I」）。
        复用键 = 标题文本。
        """
        head = tgt.get('heading') or {}
        text = head.get('text')
        num = head.get('number')
        # 阶段内已创建过同名附录 → 复用其标题段（但仍要走子标题/插入点解析）
        anchor = None
        if text and text in self._created_appendices:
            anchor = self._created_appendices[text]
            self._say('    复用本阶段已创建的附录: %s' % text[:40])
        # 已有？
        if anchor is None and (num or head.get('title_kw')):
            try:
                ex = self._resolve_heading_loose({
                    'number': num, 'title_kw': head.get('title_kw'),
                    'style': head.get('style')})
                if ex is not None:
                    self._say('    复用已有附录标题: %s' % el_text(ex._p)[:40])
                    anchor = ex._p
            except LocateError:
                anchor = None
        if anchor is None:
            if not text:
                raise ConfigError('target_appendix.heading 缺 text 或 number，无法创建')
            tpl = head.get('clone_template')
            tpl_p = None
            if tpl:
                try:
                    tpl_p = self._resolve_heading_loose(tpl)
                except LocateError:
                    tpl_p = None
            if tpl_p is None:
                # 兜底：任意 Heading 1 段做模板
                cands = find_paragraphs(self.doc, style='Heading 1',
                                        all_occurrences=True)
                tpl_p = cands[0] if cands else None
            # 插到指定位置
            ib = tgt.get('insert_before')
            if ib:
                # 位置锚点解析失败必须报错——静默回退文末会把附录插错位置
                ref_p = self._resolve_heading_loose(ib)
                ref = ref_p._p
                # insert_before 语义 = 插到该章**之前**。
                # 注意：clone_paragraph_after 用的是 addnext（插到之后），
                # 语义与配置名相反，必须显式用 addprevious。
                new_p = clone_paragraph_before(ref, tpl_p, text,
                                              align=WD_ALIGN_PARAGRAPH.LEFT)
                if new_p is None:
                    raise ConfigError('无法创建附录标题（模板不可得）: %r' % text)
                anchor = new_p._p
                self._say('    创建附录标题: %s（插在 %s 之前）'
                          % (text[:36], el_text(ref_p._p)[:28]))
                self._record('move_tables', 'create_appendix', text)
                if text:
                    self._created_appendices[text] = anchor
            else:
                body = self.doc.element.body
                ref = body[-1] if len(body) else None
                if ref is not None and ref.tag.endswith('}sectPr'):
                    ref = ref.getprevious()
                new_p = clone_paragraph_after(ref, tpl_p, text,
                                              align=WD_ALIGN_PARAGRAPH.LEFT)
                if new_p is None:
                    raise ConfigError('无法创建附录标题（模板不可得）: %r' % text)
                anchor = new_p._p
                self._say('    创建附录标题: %s' % text[:44])
                self._record('move_tables', 'create_appendix', text)
                if text:
                    self._created_appendices[text] = anchor

        # 子标题骨架
        for sh in tgt.get('sub_headings') or []:
            ex = None
            if sh.get('number'):
                try:
                    ex = self._resolve_heading_loose(sh)
                except LocateError:
                    ex = None
            if ex is None:
                tpl_p = None
                # 用已有子标题做模板
                for cand in find_paragraphs(self.doc, style='Heading 2',
                                            all_occurrences=True):
                    tpl_p = cand
                    break
                if tpl_p is None:
                    for cand in find_paragraphs(self.doc, style='Normal',
                                                all_occurrences=True):
                        tpl_p = cand
                        break
                clone_paragraph_after(anchor, tpl_p, sh.get('text'),
                                      align=WD_ALIGN_PARAGRAPH.LEFT)
                self._record('move_tables', 'create_subheading', sh.get('text'))

        # 插入点：指定子标题之后
        place = tgt.get('place_after_subheading')
        if place:
            try:
                sp = self._resolve_heading_loose({'number': str(place)})
                anchor = sp._p
            except LocateError:
                pass
        return anchor

    # ---------- 阶段 5：压缩 ----------
    def _stage_compress(self, spec):
        for c in spec:
            kind = c.get('kind')
            if kind == 'paragraph':
                p = resolve_ref(self.doc, c['target'])
                old = el_text(p._p)
                # 保护：**标题段不得被正文覆盖**——整段替换会把小节标题
                # 变成正文，导致该小节「消失」（编号断档、目录失配）。
                # 要压缩小节正文请改用 kind='section_body'。
                if is_heading(p) and not c.get('allow_heading'):
                    raise ConfigError(
                        'compress.paragraph 的目标命中的是**标题段**（%r）：'
                        '整段替换会让该小节标题被正文吞掉、编号断档。'
                        '请改用 kind="section_body"（压缩该小节下的第一段正文），'
                        '或显式加 allow_heading=true。'
                        % old[:36])
                self._say('  压缩段落: %r → %r'
                          % (old[:36], (c.get('new_text') or '')[:36]))
                self._record('compress', 'paragraph',
                             '%d → %d 字' % (len(old), len(c.get('new_text') or '')))
                if not self.dry_run:
                    set_paragraph_text(p, c['new_text'])

            elif kind == 'section_body':
                # 压缩某小节的**正文首段**：定位标题 → 取其下第一个非 bullet 正文段
                h = resolve_ref(self.doc, c['target']['heading']
                                if 'heading' in c['target'] else c['target'])
                blocks = self._collect_section_blocks(h)
                tgt_p = None
                for kindb, obj in blocks:
                    if kindb != 'p' or is_bullet(obj):
                        continue
                    tgt_p = obj
                    break
                if tgt_p is None:
                    raise ConfigError(
                        'compress.section_body：小节 %r 下没有可压缩的正文段'
                        % el_text(h._p)[:36])
                old = el_text(tgt_p._p)
                self._say('  压缩小节正文: %r → %r'
                          % (old[:36], (c.get('new_text') or '')[:36]))
                self._record('compress', 'section_body',
                             '%d → %d 字' % (len(old), len(c.get('new_text') or '')))
                if not self.dry_run:
                    set_paragraph_text(tgt_p, c['new_text'])

            elif kind == 'bullets_merge':
                h = resolve_ref(self.doc, c['target']['heading'])
                blocks = self._collect_section_blocks(h)
                bullets = [o for k, o in blocks if k == 'p' and is_bullet(o)]
                keep = c.get('keep_leading_bullets', 0)
                new_items = c.get('new_items') or []
                self._say('  合并 bullet 组：原有 %d 条 → 保留 %d + 新写 %d'
                          % (len(bullets), keep, len(new_items)))
                self._record('compress', 'bullets_merge',
                             '%d → %d 条' % (len(bullets), keep + len(new_items)))
                if self.dry_run:
                    continue
                for b in bullets[keep:]:
                    if c.get('delete_remaining', True):
                        b._p.getparent().remove(b._p)
                # 新条目接在保留的 bullet 之后
                anchor = bullets[keep - 1]._p if keep else h._p
                tpl = bullets[0] if bullets else None
                for item in new_items:
                    np_ = clone_paragraph_after(anchor, tpl, item)
                    anchor = np_._p

            elif kind == 'table_cell':
                tbl = resolve_ref(self.doc, c['table'])
                row_kw = c.get('row_keyword')
                col = c.get('column', -1)
                hit = 0
                for r in tbl.rows:
                    if row_kw is not None and r.cells[0].text.strip() != row_kw:
                        continue
                    idx = col if col >= 0 else len(r.cells) + col
                    if 0 <= idx < len(r.cells):
                        old = r.cells[idx].text
                        self._say('    单元格改写: %r → %r'
                                  % (old[:28], (c.get('new_text') or '')[:28]))
                        if not self.dry_run:
                            set_cell_text_preserve(r.cells[idx], c['new_text'])
                        hit += 1
                self._record('compress', 'table_cell', '%d 格' % hit)

            elif kind == 'table_drop_column':
                tbl = resolve_ref(self.doc, c['table'])
                self._say('  删列: %s' % str(c['table'].get('header_keywords')))
                self._record('compress', 'table_drop_column', '待实现三同步')
                raise NotImplementedError(
                    '删列涉及 tblW/gridCol/tcW 三同步，请用 '
                    'patch_existing_manual.py 的列宽内核或手工处理；'
                    '本 runner 不做删列，避免踩「三同步」坑。')

            else:
                raise ConfigError('未知的 compress.kind: %r' % kind)

    # ---------- 阶段 6：尾部修复 ----------
    def _stage_tail_fix(self, spec):
        marker = spec.get('end_marker')
        # 6.1 把粘在 bullet 末尾的结尾标记摘出
        if marker and spec.get('detach_from_bullet', True):
            for p in find_paragraphs(self.doc, all_occurrences=True):
                t = el_text(p._p)
                if marker in t and el_text(p._p).strip() != marker:
                    self._say('  摘出粘在 %r 上的结尾标记'
                              % (style_name(p) or 'Normal'))
                    self._record('tail_fix', 'detach_marker', t.strip()[:50])
                    if not self.dry_run:
                        replace_in_paragraph(p, marker, '')
                        set_paragraph_text(p, t.replace(marker, '').rstrip())
            # 独立成段
            for p in find_paragraphs(self.doc, all_occurrences=True):
                if el_text(p._p).strip() == marker and is_bullet(p):
                    self._say('  结尾标记由 bullet 改为独立段落')
                    self._record('tail_fix', 'unbullet_marker', marker)
                    if not self.dry_run:
                        p.style = self.doc.styles['Normal']
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 6.2 错挂条目归还
        ret = spec.get('return_stray_to')
        if ret and not self.dry_run:
            src = resolve_ref(self.doc, {'prefix': ret['prefix']})
            dst = resolve_ref(self.doc, ret['heading'])
            self._say('  归还错挂条目到 %s' % el_text(dst._p)[:34])
            self._record('tail_fix', 'return_stray', ret['prefix'][:40])
            if ret.get('position', 'last_bullet') == 'last_bullet':
                blocks = self._collect_section_blocks(dst)
                bullets = [o for k, o in blocks if k == 'p' and is_bullet(o)]
                if bullets:
                    move_after(src._p, bullets[-1]._p)
                else:
                    move_after(src._p, dst._p)
            else:
                move_after(src._p, dst._p)

        # 6.3 重建增量说明正文
        rb = spec.get('rebuild_increment_body')
        if rb and not self.dry_run:
            h = find_heading(self.doc, str(rb.get('heading_number')),
                             title_kw=rb.get('heading_text'))
            if h is None:
                h = find_heading(self.doc, str(rb.get('heading_number')))
            if h is not None:
                set_paragraph_text(h, rb.get('heading_text') or el_text(h._p))
                blocks = self._collect_section_blocks(h)
                for kind, obj in blocks:
                    el = obj._p if kind == 'p' else obj._tbl
                    el.getparent().remove(el)
                anchor = h._p
                if rb.get('body_paragraph'):
                    tpl = None
                    for p in find_paragraphs(self.doc, style='Normal',
                                            all_occurrences=True):
                        if el_text(p._p).strip() and not is_bullet(p):
                            tpl = p
                            break
                    np_ = clone_paragraph_after(anchor, tpl, rb['body_paragraph'])
                    anchor = np_._p
                btpl = None
                for p in find_paragraphs(self.doc, style='List Bullet',
                                         all_occurrences=True):
                    btpl = p
                    break
                for b in rb.get('bullets') or []:
                    np_ = clone_paragraph_after(anchor, btpl, b)
                    anchor = np_._p
                self._say('  重建增量说明正文：概述 1 段 + bullet %d 条'
                          % len(rb.get('bullets') or []))
                self._record('tail_fix', 'rebuild_increment', rb.get('heading_text', ''))

    # ---------- 阶段 7：升版 ----------
    def _stage_release(self, spec):
        ver = spec.get('version')
        date = spec.get('date')

        # 封面
        cov = spec.get('cover')
        if cov:
            n = 0
            for p in find_paragraphs(self.doc, all_occurrences=True):
                if cov['match'] in el_text(p._p):
                    self._say('  封面版本替换: %r → %r'
                              % (cov['match'], cov['replace']))
                    self._record('release', 'cover_version', cov['replace'])
                    if not self.dry_run:
                        replace_in_paragraph(p, cov['match'], cov['replace'])
                    n += 1
                    if cov.get('in_paragraphs', 1) and n >= cov.get('in_paragraphs', 1):
                        break

        # 版本沿革表（按表头定位，禁索引）
        chg = spec.get('changelog_table')
        if chg:
            tbl = find_table(self.doc, chg.get('header_keywords', []),
                             match_mode=chg.get('match_mode', 'all'))
            if tbl is None:
                raise LocateError('未找到版本沿革表: %r' % chg.get('header_keywords'))
            row = chg.get('row') or {}
            headers = [c.text.strip() for c in tbl.rows[0].cells]
            values = [row.get(h, '') for h in headers]
            self._say('  版本沿革表新增行: %s' % ' | '.join(v[:18] for v in values))
            self._record('release', 'changelog_row', ' | '.join(v[:18] for v in values))
            if not self.dry_run:
                old_rows = [tuple(c.text for c in r.cells) for r in tbl.rows]
                append_table_row(tbl, values)
                new_rows = [tuple(c.text for c in r.cells) for r in tbl.rows]
                # 正向断言：只许**追加**一行，旧行必须逐格原样保留
                if len(new_rows) != len(old_rows) + 1 \
                        or new_rows[:len(old_rows)] != old_rows:
                    raise StructureDrift(
                        '版本沿革表新增行异常：行数 %d→%d（应 +1），或旧行未原样保留'
                        % (len(old_rows), len(new_rows)))

        # 增量说明节
        inc = spec.get('increment_section')
        if inc:
            self._check_increment_number(inc.get('heading_text'))
            self._say('  新增增量说明: %s' % inc.get('heading_text'))
            self._record('release', 'increment_section', inc.get('heading_text', ''))
            if not self.dry_run:
                self._insert_increment_section(inc)

        # 目录条目
        for te in spec.get('toc_entries') or []:
            if te.get('op') == 'delete':
                self._delete_toc_by_prefix(te.get('prefix'))
            elif te.get('op') == 'add':
                self._add_toc_entry(te)

    def _check_increment_number(self, heading_text):
        """新增增量说明节的编号**不得与现存标题编号重复**。

        真实坑：重编号把 6.11 改成 6.10 之后，再新增一节「6.10 增量说明」→
        文档里出现两个 6.10，目录/交叉引用全部对不上。此处提前拦截。
        """
        if not heading_text:
            return
        parsed = parse_number(heading_text)
        if not parsed:
            return
        num = parsed[1]
        for p in self.doc.paragraphs:
            t = el_text(p._p)
            if not t.startswith(num):
                continue
            if t.strip() == heading_text.strip():
                continue
            other = parse_number(t)
            if other and other[1] == num:
                raise ConfigError(
                    '增量说明编号 %r 与现存标题 %r 冲突。'
                    '新增增量说明必须用**当前未占用的**编号（如重编号后应接 6.11）。'
                    % (num, t[:40]))

    def _insert_increment_section(self, inc):
        """在**收尾标记之前**插入增量说明节，克隆模板保格式。

        顺序铁律：新增量说明节 → 概述段 → bullets → **收尾标记**（必须仍是最后一段）。
        用 `clone_paragraph_after(收尾标记)` 会把新节插到标记**之后**，
        导致收尾标记不再是末位——这是真实踩过的坑。
        """
        body = self.doc.element.body
        # 锚点：收尾标记段
        anchor = None
        for p in self.doc.paragraphs:
            if '本手册完' in el_text(p._p):
                anchor = p._p
        if anchor is None:
            anchor = body[-1]
            if anchor is not None and anchor.tag.endswith('}sectPr'):
                anchor = anchor.getprevious()
        tpl_ref = inc.get('clone_template') or {}
        tpl_h = None
        if tpl_ref.get('heading_number'):
            tpl_h = find_heading(self.doc, str(tpl_ref['heading_number']))
        if tpl_h is None:
            for p in find_paragraphs(self.doc, style='Heading 2',
                                     all_occurrences=True):
                tpl_h = p
                break
        # 插到收尾标记之前
        new_h = clone_paragraph_before(anchor, tpl_h, inc.get('heading_text'))
        if new_h is None:
            raise ConfigError('无法创建增量说明标题（模板不可得）')
        cur = new_h._p
        # 概述段模板：找非 bullet 的 Normal 段
        tpl_p = None
        for p in find_paragraphs(self.doc, style='Normal', all_occurrences=True):
            if el_text(p._p).strip() and not is_bullet(p):
                tpl_p = p
                break
        if inc.get('body_paragraph'):
            np_ = clone_paragraph_after(cur, tpl_p, inc['body_paragraph'])
            if np_ is None:
                raise ConfigError('增量说明概述段创建失败：模板段不可得')
            cur = np_._p
        btpl = None
        for p in find_paragraphs(self.doc, style='List Bullet',
                                 all_occurrences=True):
            btpl = p
            break
        if inc.get('bullets') and btpl is None:
            raise ConfigError('增量说明要点创建失败：文档中无 List Bullet 模板段')
        for b in inc.get('bullets') or []:
            np_ = clone_paragraph_after(cur, btpl, b)
            if np_ is None:
                raise ConfigError('增量说明要点创建失败：模板段不可得')
            cur = np_._p

    def _delete_toc_by_prefix(self, prefix):
        if not prefix:
            return
        for p in find_paragraphs(self.doc, style='List Bullet',
                                 prefix=prefix, all_occurrences=True):
            self._say('  删除目录条目: %s' % prefix[:44])
            self._record('release', 'delete_toc', prefix)
            if not self.dry_run:
                p._p.getparent().remove(p._p)

    def _resolve_toc_anchor(self, spec):
        """在**目录条目**（List Bullet）里定位锚点，返回段落元素或 None。

        spec 允许：
          - str                → 段落前缀匹配
          - {'number': 'H'}    → 依次尝试 前缀 / "附录 H" 子串 / 附录行正则
          - {'text': '5  附录 H …'} → 子串匹配
        目录条目文本形如 "5  附录 H  源文件清单与效力总览"，章级条目形如
        "6  编制说明、溯源声明与参考文献"，两者都要能定位。
        """
        if not spec:
            return None
        toc = find_paragraphs(self.doc, style='List Bullet',
                              all_occurrences=True)
        if isinstance(spec, str):
            for p in toc:
                if el_text(p._p).startswith(spec):
                    return p
            return None
        num = spec.get('number')
        if num:
            num = str(num)
            # 1) 附录条目：包含 "附录 H"
            for p in toc:
                if ('附录 ' + num) in el_text(p._p):
                    return p
            # 2) 章级条目：以编号开头（"6  编制说明…"）
            pat = re.compile(r'^\s*' + re.escape(num) + r'(?![0-9.])')
            for p in toc:
                if pat.match(el_text(p._p)):
                    return p
        txt = spec.get('text')
        if txt:
            for p in toc:
                if txt in el_text(p._p):
                    return p
        return None

    def _add_toc_entry(self, te):
        text = te.get('text')
        if not text:
            return
        # 模板：目录里的其它条目
        tpl = None
        for p in find_paragraphs(self.doc, style='List Bullet',
                                 all_occurrences=True):
            tpl = p
            break
        if tpl is None:
            raise ConfigError('新增目录条目失败：文档中无 List Bullet 模板段')
        anchor = tpl._p
        after = te.get('after')
        if after:
            ex = self._resolve_toc_anchor(after)
            if ex is None:
                raise ConfigError('新增目录条目失败：锚点 %r 未命中目录条目'
                                  % (after,))
            anchor = ex._p
        np_ = clone_paragraph_after(anchor, tpl, text)
        if np_ is None:
            raise ConfigError('新增目录条目失败：模板段不可得')
        self._say('  新增目录条目: %s' % text[:50])
        self._record('release', 'add_toc', text)

    # ---------- 收尾自检 ----------
    def final_self_check(self, tail_spec):
        """阶段 7 收尾自检。不满足直接抛错（不写文件）。"""
        problems = []
        if not tail_spec:
            return problems
        marker = tail_spec.get('end_marker')
        if marker:
            n = 0
            standalone = False
            for p in find_paragraphs(self.doc, all_occurrences=True):
                t = el_text(p._p)
                if marker in t:
                    n += 1
                    if t.strip() == marker:
                        standalone = True
            if tail_spec.get('end_marker_occurrences') and \
                    n != tail_spec['end_marker_occurrences']:
                problems.append('结尾标记出现 %d 次，应为 %d 次'
                                % (n, tail_spec['end_marker_occurrences']))
            if tail_spec.get('require_standalone_paragraph', True) and \
                    n > 0 and not standalone:
                problems.append('结尾标记未独立成段（粘在别的段尾）')

        if tail_spec.get('require_body_paragraph', True):
            import re as _re
            pat = tail_spec.get('increment_heading_regex', '增量说明')
            # **不要一刀切要求概述段**——真实文档里有「bullet 型」增量说明
            # （标题下直接跟若干条要点，本身就没有概述段，那是合法的）。
            # 只有 config 显式指定为「概述型」的节号才强制要求概述段。
            strict = tail_spec.get('strict_body_heading_numbers') or []
            if strict:
                for num in strict:
                    h = find_heading(self.doc, str(num))
                    if h is None:
                        problems.append('增量说明节 %s 未找到' % num)
                        continue
                    nxt = self._next_paragraph(h)
                    if nxt is None:
                        problems.append('增量说明「%s」下无正文段' % num)
                    elif is_bullet(nxt):
                        problems.append('增量说明「%s」下第一个段是 bullet，概述段缺失'
                                        % num)
            else:
                # 未指定时只做「不得错挂」检查：增量说明节的首条 bullet
                # 不得与它上方那一节的末条内容重复
                for p in find_paragraphs(self.doc, all_occurrences=True):
                    t = el_text(p._p).strip()
                    if is_bullet(p) or not t or not _re.search(pat, t):
                        continue
                    if parse_number(t) is None:
                        continue
                    prev_b = None
                    for q in find_paragraphs(self.doc, all_occurrences=True):
                        if q._p is p._p:
                            break
                        if is_bullet(q) and el_text(q._p).strip():
                            prev_b = el_text(q._p).strip()
                    nxt = self._next_paragraph(p)
                    if prev_b and nxt and is_bullet(nxt) \
                            and el_text(nxt._p).strip()[:24] == prev_b[:24]:
                        problems.append('增量说明「%s」首条与上一小节末条重复（错挂）'
                                        % t[:26])
                        break
        return problems

    def _next_paragraph(self, p):
        from docx_revise_toolkit import iter_block_items
        found = False
        for kind, obj in iter_block_items(self.doc):
            if kind != 'p':
                continue
            if found:
                return obj
            if obj._p is p._p:
                found = True
        return None


# =====================================================================
# CLI
# =====================================================================

def main(argv=None):
    ap = argparse.ArgumentParser(description='config 驱动的章节重构 runner')
    ap.add_argument('--input', required=True, help='输入 .docx（只读，绝不覆盖）')
    ap.add_argument('--output', required=True, help='输出 .docx（新文件）')
    ap.add_argument('--config', required=True, help='JSON 配置')
    ap.add_argument('--dry-run', action='store_true', help='只打印计划，不写文件')
    ap.add_argument('--report', default=None, help='执行报告 JSON 输出路径')
    ap.add_argument('--allow-overwrite-input', action='store_true',
                    help='允许 output == input（默认禁止）')
    args = ap.parse_args(argv)

    if not os.path.exists(args.input):
        print('输入文件不存在: %s' % args.input)
        return 2
    if not os.path.exists(args.config):
        print('配置文件不存在: %s' % args.config)
        return 2

    in_abs = os.path.abspath(args.input)
    out_abs = os.path.abspath(args.output)
    if in_abs == out_abs and not args.allow_overwrite_input:
        print('拒绝执行：输出与输入同路径（%s）。'
              '正式升版必须用新文件名。' % out_abs)
        return 2

    with open(args.config, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    doc = Document(in_abs)
    runner = RefactorRunner(doc, cfg, dry_run=args.dry_run)
    try:
        runner.run()
        problems = runner.final_self_check(cfg.get('tail_fix'))
        if problems:
            print('\n收尾自检未通过：')
            for p in problems:
                print('  - %s' % p)
            print('已中止，未写文件。')
            return 1
    except (ConfigError, LocateError, StructureDrift, NotImplementedError) as e:
        print('\n执行中止: %s: %s' % (type(e).__name__, e))
        print('未写文件。')
        return 1

    if args.dry_run:
        print('\n[dry-run] 未写文件。计划动作 %d 项：' % len(runner.actions))
        for a in runner.actions:
            print('  %-16s %-28s %s' % (a['stage'], a['action'], a['detail']))
        return 0

    doc.save(out_abs)
    print('\n已写出: %s' % out_abs)

    if args.report:
        with open(args.report, 'w', encoding='utf-8') as f:
            json.dump({'log': runner.log, 'actions': runner.actions},
                      f, ensure_ascii=False, indent=2)
        print('报告已写入: %s' % args.report)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
