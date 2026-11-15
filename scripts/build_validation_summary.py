# -*- coding: utf-8 -*-
"""build_validation_summary.py — 把目标章节重写为「验证项目及指标汇总 + 稳定性结果矩阵」。

跨品种通用、配置驱动（JSON）。沉淀自《注射用乳糖酸红霉素》手册 v7 第 4 章重写实战。

设计要点（详见 references/validation_summary_playbook.md）：
  * 真实章节标题定位取「最后出现」者，避开目录(TOC)伪条目（chapter_refactor_playbook §3.3）。
  * 孤儿段防护：先把所有新建元素收集进列表，再统一 remove + insert 到标题之后；
    禁止用 addnext 移动最后子元素（会丢到文档末、「本手册完」之后）。
  * 列宽合计必须 = 9638 twips（A4 + 2cm 边距纵向版心 17cm）；横向放 landscape section。
  * OOXML 顺序：tblPr(borders→tblW→tblLayout→tblCellMar)→tblGrid→tr(trPr 含 cantSplit)
    →tc(tcPr 含 tcW + vAlign=center)；cantSplit 只能放 trPr，不可放 tcPr。
  * 数据中的不确定值标「OCR·需核」；效价不在稳定性考察项目清单内不得硬塞。

注意：本脚本不含任何产品专属数值（如分离度 2.28、各回收率区间、各时间点数值），
那些仅作为 references 下的示例 JSON / playbook 算例，禁止写进脚本默认值。

用法：
  python build_validation_summary.py --input SRC.docx --output OUT.docx --config CFG.json
  python build_validation_summary.py --config CFG.json --input SRC.docx --dry-run
  python build_validation_summary.py --selftest
"""
import argparse
import json
import os
import sys
import tempfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# A4 + 2cm 边距 → 纵向版心 17cm = 9638 twips。横向版心 25.7cm = 14552 twips。
DEFAULT_TOTAL_TWIPS = 9638

ASCII_FONT = "Times New Roman"
EASTASIA_FONT = "宋体"


def _set_run_fonts(rpr):
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ASCII_FONT)
    rfonts.set(qn("w:hAnsi"), ASCII_FONT)
    rfonts.set(qn("w:eastAsia"), EASTASIA_FONT)


def _make_para(text, size=9, bold=False, align=None, keep_next=False, style=None):
    """创建一个段落元素（不插入文档），返回 (w:p, 是否需加入 created 列表由调用方决定)。"""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    if keep_next:
        pPr.append(OxmlElement("w:keepNext"))
    if align is not None:
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), align)
        pPr.append(jc)
    rPr = OxmlElement("w:rPr")
    _set_run_fonts(rPr)
    if bold:
        rPr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(size * 2))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(size * 2))
    rPr.append(szCs)
    pPr.append(rPr)
    r = OxmlElement("w:r")
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(pPr)
    p.append(r)
    return p


def _make_fixed_table(rows, col_widths, size=9, header_bold=True):
    """构造定宽、居中、防跨页断裂的表格元素。rows[0] 为表头。"""
    assert len(col_widths) == len(rows[0]), "列宽数量必须与列数一致"
    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    bords = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "808080")
        bords.append(e)
    tblPr.append(bords)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(sum(col_widths)))
    tw.set(qn("w:type"), "dxa")
    tblPr.append(tw)
    tl = OxmlElement("w:tblLayout")
    tl.set(qn("w:type"), "fixed")
    tblPr.append(tl)
    cant_margin = OxmlElement("w:tblCellMar")
    for m in ("top", "left", "bottom", "right"):
        e = OxmlElement("w:" + m)
        e.set(qn("w:w"), "40")
        e.set(qn("w:type"), "dxa")
        cant_margin.append(e)
    tblPr.append(cant_margin)
    tbl.append(tblPr)
    grid = OxmlElement("w:tblGrid")
    for w in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    tbl.append(grid)
    for ri, row in enumerate(rows):
        tr = OxmlElement("w:tr")
        trPr = OxmlElement("w:trPr")
        trPr.append(OxmlElement("w:cantSplit"))  # 必须放 trPr
        tr.append(trPr)
        for ci, val in enumerate(row):
            tc = OxmlElement("w:tc")
            tcPr = OxmlElement("w:tcPr")
            wEl = OxmlElement("w:tcW")
            wEl.set(qn("w:w"), str(col_widths[ci]))
            wEl.set(qn("w:type"), "dxa")
            tcPr.append(wEl)
            va = OxmlElement("w:vAlign")  # 必须放 tcPr
            va.set(qn("w:val"), "center")
            tcPr.append(va)
            tc.append(tcPr)
            p = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center" if ci > 0 else "left")
            pPr.append(jc)
            p.append(pPr)
            r = OxmlElement("w:r")
            rPr = OxmlElement("w:rPr")
            _set_run_fonts(rPr)
            if header_bold and ri == 0:
                rPr.append(OxmlElement("w:b"))
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(size * 2))
            rPr.append(sz)
            szCs = OxmlElement("w:szCs")
            szCs.set(qn("w:val"), str(size * 2))
            rPr.append(szCs)
            r.append(rPr)
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = val
            r.append(t)
            p.append(r)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl


def _norm(s):
    """折叠连续空白为单空格，避免标题里 '4  '（两空格）与锚点 '4 '（单空格）失配。"""
    return " ".join((s or "").split())


def _body_text(body):
    out = []
    for el in body:
        if el.tag.split("}")[1] == "p":
            out.append("".join(n.text or "" for n in el.iter(qn("w:t"))))
        else:
            out.append(None)
    return out


def build(doc, cfg):
    """按 cfg 重写目标章节；返回 (created_elements, ch4_el, next_el)。"""
    body = doc.element.body
    texts = _body_text(body)

    chapter_anchor = _norm(cfg["chapter_anchor"])
    ch4_idxs = [i for i, t in enumerate(texts) if t is not None and chapter_anchor in _norm(t)]
    assert ch4_idxs, "未找到章节锚点：%s（请检查章节标题文本）" % cfg["chapter_anchor"]
    ch4_idx = ch4_idxs[-1]  # 取最后出现 = 真实正文章节（避开目录伪条目）
    ch4_el = list(body)[ch4_idx]

    next_anchor = cfg.get("next_anchor")
    if not next_anchor:
        raise ValueError("config 必须提供 next_anchor（下一章锚点）以界定删除范围")
    next_anchor_n = _norm(next_anchor)
    next_idx = None
    for i in range(ch4_idx + 1, len(texts)):
        if texts[i] is not None and next_anchor_n in _norm(texts[i]):
            next_idx = i
            break
    if next_idx is None:
        # 未找到下一章锚点：宁可报错也不静默删到文末（chapter_refactor_playbook §3.9）
        raise ValueError("未找到下一章锚点：%s" % next_anchor)

    # 删除 real ch4 与 next 之间的全部旧内容（不含两端）
    to_remove = []
    for i in range(ch4_idx + 1, next_idx):
        if texts[i] is not None or list(body)[i].tag.split("}")[1] == "tbl":
            to_remove.append(list(body)[i])
    for el in to_remove:
        el.getparent().remove(el)

    # 标题样式：4.1 / 4.2 子标题沿用章节标题样式
    ch4_para = None
    for p in doc.paragraphs:
        if p._p is ch4_el:
            ch4_para = p
            break
    ch4_style = ch4_para.style if ch4_para is not None else None

    required_total = int(cfg.get("required_total_twips", DEFAULT_TOTAL_TWIPS))
    created = []

    # 章首引导段
    if cfg.get("lead"):
        created.append(_make_para(cfg["lead"], size=9))

    # 4.1 验证项目及指标汇总表
    vt = cfg.get("validation_table")
    if vt:
        cw = vt["col_widths_twips"]
        assert sum(cw) == required_total, (
            "validation_table 列宽合计 %d ≠ 版心 %d twips" % (sum(cw), required_total)
        )
        h = doc.add_paragraph(style=ch4_style)
        r = h.add_run(vt.get("title", "4.1 验证项目及指标汇总"))
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.name = ASCII_FONT
        created.append(h._p)
        if vt.get("description"):
            created.append(_make_para(vt["description"], size=8))
        rows = [vt["headers"]] + vt["rows"]
        tbl = _make_fixed_table(rows, cw, size=9, header_bold=True)
        created.append(tbl)

    # 4.2 稳定性考察结果矩阵
    sm = cfg.get("stability_matrix")
    if sm:
        cw = sm["col_widths_twips"]
        assert sum(cw) == required_total, (
            "stability_matrix 列宽合计 %d ≠ 版心 %d twips" % (sum(cw), required_total)
        )
        h = doc.add_paragraph(style=ch4_style)
        r = h.add_run(sm.get("title", "4.2 稳定性考察结果"))
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.name = ASCII_FONT
        created.append(h._p)
        if sm.get("lead"):
            created.append(_make_para(sm["lead"], size=9, keep_next=True))
        rows = [sm["headers"]] + sm["rows"]
        tbl = _make_fixed_table(rows, cw, size=9, header_bold=True)
        created.append(tbl)
        if sm.get("accel_summary"):
            created.append(_make_para(sm["accel_summary"], size=9, keep_next=True))
        if sm.get("conclusion"):
            created.append(_make_para(sm["conclusion"], size=9, bold=True, keep_next=True))
        if sm.get("ocr_note"):
            created.append(_make_para(sm["ocr_note"], size=8, keep_next=True))

    return created, ch4_el, list(body)[next_idx]


def insert_created(body, created, ch4_el):
    """孤儿段防护：先全部移出（若有父节点）再统一插入到真实标题之后。"""
    for el in created:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    ins = list(body).index(ch4_el) + 1
    for el in created:
        body.insert(ins, el)
        ins += 1


def run(input_path, output_path, config_path, dry_run=False):
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    doc = Document(input_path)
    body = doc.element.body
    created, ch4_el, next_el = build(doc, cfg)

    plan = ["计划写入章节「%s」之下：" % cfg["chapter_anchor"]]
    for el in created:
        tag = el.tag.split("}")[1]
        if tag == "tbl":
            plan.append("  • 表格（%d 行 × %d 列）" % (
                len(el.findall(qn("w:tr"))),
                len(el.find(qn("w:tblGrid")).findall(qn("w:gridCol"))),
            ))
        else:
            txt = "".join(n.text or "" for n in el.iter(qn("w:t")))
            plan.append("  • 段落：%s" % (txt[:40] + ("…" if len(txt) > 40 else "")))
    plan.append("插入位置：真实章节标题（最后出现）之后、下一章「%s」之前。"
                % cfg.get("next_anchor"))
    print("\n".join(plan))

    if dry_run:
        print("\n[DRY-RUN] 未写入任何文件。")
        return

    insert_created(body, created, ch4_el)
    doc.save(output_path)
    print("\n已保存：%s（%d 字节）" % (output_path, os.path.getsize(output_path)))
    print("新插入元素数：%d" % len(created))


def selftest():
    tmpdir = tempfile.mkdtemp(prefix="bvs_selftest_")
    src = os.path.join(tmpdir, "t.docx")
    Document().save(src)

    doc = Document(src)
    # 模拟目录伪条目（List Bullet 样式，含锚点文本，位于真实章节之前）
    toc = doc.add_paragraph("4 方法转移 / 稳定性考察结论摘要", style="List Bullet")
    # 真实章节标题
    real = doc.add_paragraph("4 方法转移 / 稳定性考察结论摘要")
    # 中间夹带旧内容（段落 + 表格），应被删除
    doc.add_paragraph("旧的正文内容，应当被替换")
    doc.add_table(rows=1, cols=2)
    # 下一章锚点
    doc.add_paragraph("5 附录 A 源文件清单")
    doc.add_paragraph("—— 本手册完 ——")
    doc.save(src)

    cfg = {
        "chapter_anchor": "4 方法转移 / 稳定性考察结论摘要",
        "next_anchor": "附录 A",
        "lead": "本章汇总各分析方法的验证项目及关键指标。",
        "validation_table": {
            "title": "4.1 各分析方法验证项目及指标汇总",
            "description": "下表逐法列出验证项目与关键指标。",
            "headers": ["方法", "验证项目", "关键指标", "依据报告"],
            "rows": [["3.1 性状", "法定方法", "按 ChP 通则", "S04"]],
            "col_widths_twips": [1550, 2850, 4138, 1100],
        },
        "stability_matrix": {
            "title": "4.2 稳定性考察结果",
            "lead": "以下汇总长期条件下各考察项目结果。",
            "headers": ["考察项目", "0月", "3月", "6月", "9月", "12月*", "18月*"],
            "rows": [["性状", "类白色", "同左", "同左", "同左", "(OCR)类白色", "(OCR)类白色"]],
            "col_widths_twips": [2400, 1206, 1206, 1206, 1206, 1207, 1207],
            "accel_summary": "加速条件各项目均符合规定。",
            "conclusion": "结论：质量稳定。",
            "ocr_note": "备注：12/18 月标 (OCR·需核)。",
        },
    }
    # 写 config 到临时文件
    cfg_path = os.path.join(tmpdir, "cfg.json")
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)

    out = os.path.join(tmpdir, "out.docx")
    run(src, out, cfg_path)

    # 断言
    d2 = Document(out)
    body = d2.element.body
    texts = _body_text(body)
    ch4_idxs = [i for i, t in enumerate(texts) if t is not None and cfg["chapter_anchor"] in t]
    assert ch4_idxs, "self-test: 章节标题丢失"
    real_idx = ch4_idxs[-1]
    # 真实标题之后紧跟的应是章首引导段，而非目录伪条目
    after = texts[real_idx + 1]
    assert after is not None and after.startswith("本章汇总"), "self-test: 内容未插入到真实标题之后"
    # 旧内容已删除
    assert not any("旧的正文内容" in (t or "") for t in texts), "self-test: 旧内容未删除"
    # 表格数 = 2，且列宽合计 = 9638
    tables = body.findall(qn("w:tbl"))
    assert len(tables) == 2, "self-test: 表格数应为 2，实得 %d" % len(tables)
    for tb in tables:
        grid = tb.find(qn("w:tblGrid"))
        total = sum(int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol")))
        assert total == 9638, "self-test: 列宽合计 %d ≠ 9638" % total
        # cantSplit 必须在 trPr，不在 tcPr
        tr = tb.find(qn("w:tr"))
        assert tr.find(qn("w:trPr")).find(qn("w:cantSplit")) is not None, "self-test: cantSplit 不在 trPr"
        tc = tr.find(qn("w:tc"))
        assert tc.find(qn("w:tcPr")).find(qn("w:cantSplit")) is None, "self-test: cantSplit 误放 tcPr"
        assert tc.find(qn("w:tcPr")).find(qn("w:vAlign")) is not None, "self-test: tcPr 缺 vAlign"
    # 结尾标记仍存在且唯一
    assert texts.count("—— 本手册完 ——") == 1, "self-test: 结尾标记异常"
    print("\n[SELFTEST] PASS")


def main():
    ap = argparse.ArgumentParser(description="重写目标章节为「验证摘要 + 稳定性矩阵」")
    ap.add_argument("--input", help="源 DOCX 路径")
    ap.add_argument("--output", help="输出 DOCX 路径（不覆盖输入）")
    ap.add_argument("--config", help="JSON 配置路径")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划不写盘")
    ap.add_argument("--selftest", action="store_true", help="运行内置自检")
    args = ap.parse_args()

    if args.selftest:
        selftest()
        return
    if not (args.input and args.output and args.config):
        ap.error("非 selftest 模式需提供 --input / --output / --config")
    if os.path.abspath(args.input) == os.path.abspath(args.output):
        ap.error("output 必须与 input 不同（禁止覆盖源文件）")
    run(args.input, args.output, args.config, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
