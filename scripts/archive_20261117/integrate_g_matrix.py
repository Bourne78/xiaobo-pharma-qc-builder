# -*- coding: utf-8 -*-
"""
附录 G 整合：把物料明细总表[216]的名称/规格嵌入检验项×物料矩阵[219]单元格，
去掉单元格内代号，删除明细总表，改写 4 处悬空引用，另存 v8。

用法:
    python integrate_g_matrix.py --dry-run          # 只校验，不落盘
    python integrate_g_matrix.py                    # 执行并落盘 v8

设计铁律:
- 先用 python-docx .text 建映射（itertext 会因命名空间假重复，禁用）。
- 单元格写入走 set_cell_multiline：保留首 run rPr，格内每条物料一个 w:br。
- 悬空引用 4 段跨 run，用整段 set_paragraph_text 回写（保留首 run rPr）。
- 删表按表头定位，禁用 doc.tables[i]。
- [217]/[221] 空段承载 sectPr，绝不动。
"""
import argparse
import re
import sys
import os

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.table import Table

SRC = r"E:\Workbuddy\文献整理\注射用乳糖酸红霉素_全项检验准备与关键点控制手册_v7_20261113.docx"
DST = r"E:\Workbuddy\文献整理\注射用乳糖酸红霉素_全项检验准备与关键点控制手册_v8_20261115.docx"

CODE_RE = re.compile(r'(?:RP|RGT|SMP|COL|INS|CON)-\d{2}')

# 4 段改写：old 精确匹配 python-docx .text，new 为定稿文案
REWRITES = {
    213: (
        '本清单用于开检前备料核对，不是成本台账。六类物料编号前缀：RP- 对照品与菌种 / RGT- 试剂试液 / '
        'SMP- 供试品 / COL- 色谱柱 / INS- 仪器 / CON- 消耗品器皿。来源标记：SOP原文 · 质量标准(SOP未载) · '
        '需查。「√」列留空供现场打勾。',
        '本清单用于开检前备料核对，不是成本台账。六类物料编号前缀：RP- 对照品与菌种 / RGT- 试剂试液 / '
        'SMP- 供试品 / COL- 色谱柱 / INS- 仪器 / CON- 消耗品器皿。矩阵单元格按条目列示物料名称 / 规格 / '
        '关键属性，现行 SOP 与质量标准均未载明的信息填「需查」；现场如需逐项核对，可自行在矩阵相应行加注。',
    ),
    214: (
        '单一数据源：G.2 的检验项×物料矩阵由每条物料的 items 字段自动推导，与本节物料明细总表永远同步，'
        '不手工维护第二份清单。凡现行 SOP 与质量标准均未载明的信息（对照品证书、柱寿命、仪器资产编号、'
        '部分器皿规格）一律填「需查」。',
        '单一数据源：G.2 的检验项×物料矩阵由每条物料的 items 字段自动推导，矩阵单元格列示对应物料的'
        '名称 / 规格 / 关键属性；物料明细以源数据为准，不手工维护第二份清单。凡现行 SOP 与质量标准均未'
        '载明的信息（对照品证书、柱寿命、仪器资产编号、部分器皿规格）一律填「需查」。',
    ),
    215: (
        'G.2  物料明细总表与检验项矩阵（矩阵由物料 items 字段自动生成）',
        'G.2  检验项 × 物料矩阵（每格列示物料名称 / 规格 / 关键属性）',
    ),
    220: (
        '矩阵列示物料编号，完整名称 / 规格 / 需求量 / 来源依据见上方物料明细总表；「—」表示该类物料按'
        '现行 SOP 未用于该检验项。本矩阵由 items 字段自动生成；现场若实际使用清单外物料，须回写 items '
        '后重新生成，不得手工改表。',
        '矩阵单元格列示物料名称 / 规格 / 关键属性，每条物料独立成行；「—」表示该类物料按现行 SOP 未用于'
        '该检验项。本矩阵由 items 字段自动生成；现场若实际使用清单外物料，须回写 items 后重新生成，'
        '不得手工改表。',
    ),
}

XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


def set_paragraph_text_keep_rpr(p_el, text):
    """整段回写：保留首 run 的 rPr，其余 run 的 w:t 清空。p_el 是 w:p 元素。"""
    runs = p_el.findall(qn('w:r'))
    assert runs, '段落无 run'
    first = runs[0]
    ts = first.findall(qn('w:t'))
    if ts:
        t0 = ts[0]
    else:
        t0 = OxmlElement('w:t')
        first.append(t0)
    t0.text = text
    t0.set(XMLSPACE, 'preserve')
    for r in runs[1:]:
        for tt in r.findall(qn('w:t')):
            tt.text = ''
            tt.set(XMLSPACE, 'preserve')


def set_cell_multiline(cell, lines):
    """格内多行写入：保留首段首 run 的 rPr，每条物料一个 w:br。
    返回写入的 w:br 数量。禁用 cell.text=。"""
    ps = cell.paragraphs
    p0 = ps[0]._p
    runs = p0.findall(qn('w:r'))
    assert runs, '单元格首段无 run'
    proto_rpr = runs[0].find(qn('w:rPr'))
    import copy
    rpr_copy = copy.deepcopy(proto_rpr) if proto_rpr is not None else None

    # 清空首段所有 run
    for r in runs:
        p0.remove(r)
    # 逐行构造 run(w:t) + run(w:br)
    br_count = 0
    for i, line in enumerate(lines):
        r = OxmlElement('w:r')
        if rpr_copy is not None:
            r.append(copy.deepcopy(rpr_copy))
        t = OxmlElement('w:t')
        t.text = line
        t.set(XMLSPACE, 'preserve')
        r.append(t)
        p0.append(r)
        if i < len(lines) - 1:
            br = OxmlElement('w:r')
            if rpr_copy is not None:
                br.append(copy.deepcopy(rpr_copy))
            br_el = OxmlElement('w:br')
            br.append(br_el)
            p0.append(br)
            br_count += 1
    # 删除多余段
    for p in ps[1:]:
        p._p.getparent().remove(p._p)
    return br_count


def body_elements(doc):
    """只收集 w:p / w:tbl（与探测脚本口径一致），不含 body 级 sectPr。"""
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag == qn('w:p') or ch.tag == qn('w:tbl'):
            out.append(ch)
    return out


def find_table_by_header(doc, kws):
    for t in doc.tables:
        if not t.rows:
            continue
        head = ' '.join(c.text.strip() for c in t.rows[0].cells)
        if all(k in head for k in kws):
            return t
    return None


def main(dry_run):
    doc = Document(SRC)
    body = body_elements(doc)
    assert len(body) == 236, f'预期 body 236 块，实际 {len(body)}'

    # ---- Step 1: 建映射 ----
    det_tbl = find_table_by_header(doc, ['编号', '名称 / 规格 / 关键属性'])
    assert det_tbl is not None, '未找到明细总表'
    assert len(det_tbl.rows) == 65, f'明细表应 65 行，实际 {len(det_tbl.rows)}'
    id2name = {}
    for r in det_tbl.rows[1:]:
        c0 = r.cells[0].text.strip()
        c1 = r.cells[1].text.strip()
        assert c0 and c1, f'空映射行: {c0!r}/{c1!r}'
        assert c0 not in id2name, f'重复编号 {c0}'
        id2name[c0] = c1
    assert len(id2name) == 64, f'映射应 64 条，实际 {len(id2name)}'

    # 定位矩阵表
    mat_tbl = find_table_by_header(doc, ['检验项（卡号）', 'RP- 对照品'])
    assert mat_tbl is not None, '未找到矩阵表'
    assert len(mat_tbl.rows) == 16 and len(mat_tbl.columns) == 7, '矩阵维度异常'
    # 用元素身份确认是 [219]
    assert body[219] is mat_tbl._tbl, '矩阵表不是 body[219]'
    assert body[216] is det_tbl._tbl, '明细表不是 body[216]'

    # ---- 提取矩阵代号，校验完备性 ----
    codes_in_matrix = set()
    dash_count = 0
    plan = []  # (row_idx, col_idx, [names]) 或 (row_idx,col_idx,'—')
    for ri in range(1, 16):
        row = mat_tbl.rows[ri]
        for ci in range(1, 7):
            cell = row.cells[ci]
            txt = cell.text.strip()
            if txt == '—':
                dash_count += 1
                continue
            codes = CODE_RE.findall(txt)
            assert codes, f'R{ri}C{ci} 无代号也非破折号: {txt!r}'
            # 校验内容确实只由代号+顿号组成（防止混入其它文字被吞）
            stripped = CODE_RE.sub('', txt)
            assert stripped.replace('、', '').strip() == '', \
                f'R{ri}C{ci} 含非代号文本: {stripped!r}'
            missing = [c for c in codes if c not in id2name]
            assert not missing, f'R{ri}C{ci} 代号未命中映射: {missing}'
            codes_in_matrix.update(codes)
            plan.append((ri, ci, [id2name[c] for c in codes], len(codes)))
    assert len(codes_in_matrix) == 64, f'矩阵代号去重应 64，实际 {len(codes_in_matrix)}'
    assert codes_in_matrix == set(id2name), '矩阵代号集合与映射键不一致'
    print(f'[OK] 映射 {len(id2name)} 条；矩阵代号 {len(codes_in_matrix)} 个全覆盖；'
          f'待填格 {len(plan)} 个；破折号格 {dash_count} 个')

    if dry_run:
        print('[DRY-RUN] 校验通过，不落盘。')
        return

    # ---- Step 4: 改写 4 段悬空引用（先于删表，元素引用仍有效）----
    for idx, (old, new) in REWRITES.items():
        p_el = body[idx]
        para = Paragraph(p_el, doc)
        cur = para.text
        assert cur == old, f'[{idx}] 原文不匹配:\n 期望:{old!r}\n 实际:{cur!r}'
        set_paragraph_text_keep_rpr(p_el, new)
        assert Paragraph(p_el, doc).text == new, f'[{idx}] 回写后校验失败'
        print(f'[OK] 改写 [{idx}]: {new[:40]}...')

    # ---- Step 2: 矩阵单元格代号→名称+规格（格内 w:br 多行）----
    total_br = 0
    for ri, ci, names, ncode in plan:
        cell = mat_tbl.rows[ri].cells[ci]
        br = set_cell_multiline(cell, names)
        assert br == ncode - 1, f'R{ri}C{ci} w:br 数 {br} != {ncode-1}'
        total_br += br
        # 断言无裸代号残留
        assert not CODE_RE.search(cell.text), f'R{ri}C{ci} 残留代号: {cell.text[:60]!r}'
    print(f'[OK] 矩阵写入 {len(plan)} 格，插入 w:br {total_br} 个，裸代号残留 0')

    # ---- Step 3: 删除明细总表 ----
    det_tbl._tbl.getparent().remove(det_tbl._tbl)
    assert find_table_by_header(doc, ['编号', '名称 / 规格 / 关键属性']) is None, '明细表未删净'

    # ---- 结构断言 ----
    body2 = body_elements(doc)
    assert len(body2) == 235, f'删表后应 235 块，实际 {len(body2)}'
    # sectPr 仍 3 个，[217]/[221] 对应空段仍在（删表后索引 [216] 消失，原217→216,221→220）
    sect_count = len(doc.element.body.findall('.//' + qn('w:sectPr')))
    assert sect_count == 3, f'sectPr 应 3 个，实际 {sect_count}'
    # 横向节仍在
    sec_orients = [s.orientation for s in doc.sections]
    assert sec_orients == [0, 1, 0], f'分节方向异常: {sec_orients}'

    # ---- Step 5: 版面：确认数据行无 exact 行高 ----
    for ri in range(1, 16):
        tr = mat_tbl.rows[ri]._tr
        trPr = tr.find(qn('w:trPr'))
        if trPr is not None:
            h = trPr.find(qn('w:trHeight'))
            if h is not None:
                assert h.get(qn('w:hRule')) != 'exact', f'R{ri} 行高为 exact，会截断'
    print('[OK] 删表 + 结构断言通过；sectPr=3；分节[纵,横,纵]；无 exact 行高')

    # ---- Step 6: 保存 v8（同盘原子替换）----
    tmp = DST + '.tmp'
    doc.save(tmp)
    assert os.path.getsize(tmp) > 0
    os.replace(tmp, DST)
    print(f'[OK] 已保存: {DST}  size={os.path.getsize(DST)}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    main(a.dry_run)
