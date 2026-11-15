# -*- coding: utf-8 -*-
"""delete_chapter.py — 删除手册中的指定章节/附录，并清理全部悬空交叉引用。

跨品种通用、配置驱动（JSON）。沉淀自《注射用乳糖酸红霉素》手册 v7 删除第 5 章
附录 G / 附录 H 的实战（尤其吸取「删 G 时只扫段落、漏了 2 处表格单元格引用导致回退」的教训）。

设计要点（详见 references/delete_chapter_playbook.md）：
  * 真实章节标题取「最后出现」者（避开目录 TOC 伪条目）；TOC 条目取「首次出现」者。
  * 删除范围：真实标题 → 结尾锚点（如「本手册完」）之前，标题本身一并删除，
    结尾锚点保留且全文唯一；相邻章节/附录不得误伤。
  * **悬空引用全扫描**：段落 + 表格单元格（iter w:t）都要扫；上次删 G 只扫段落，
    漏掉第 3 章操作卡单元格「见附录 G RP-02」与附录 E 表格行「附录 G 一律填需查」。
  * 引用改写是「改成不指向被删章节的等价表述」，不是简单删字：
    历史记述（版本沿革表）就地把「迁入新增附录 X」软化为「并入…（原附录 X，本版已移除）」；
    正文指针「见附录 X.y（...）」去掉指针保留括号内容。
  * run 级精确替换；若 find 命中元素拼接文本却不在任何单个 run 内（跨 run），
    报硬错而非静默跳过（chapter_refactor_playbook §3.6 跨 run token 教训）。
  * 验收：verify_absent 关键词零命中 + verify_present 关键词保留 + 结尾锚点唯一；
    任一不满足即抛错，禁止保存半成品。

注意：本脚本不含任何产品专属章节名/数值/文件编号；那些仅在 references 下的
示例 JSON / playbook 算例，禁止写进脚本默认值。

用法：
  python delete_chapter.py --input SRC.docx --output OUT.docx --config CFG.json
  python delete_chapter.py --config CFG.json --input SRC.docx --dry-run
  python delete_chapter.py --selftest
"""
import argparse
import json
import os
import sys
import tempfile

from docx import Document
from docx.oxml.ns import qn


def _norm(s):
    """折叠连续空白，避免标题里两空格与锚点单空格失配。"""
    return " ".join((s or "").split())


def _ptext(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t")))


def _body_elements(doc):
    return list(doc.element.body)


def _iter_blocks(body):
    """遍历 body 下所有含文本的块：段落(p)与表格单元格(tc)，returns (kind, element)。"""
    for el in body:
        tag = el.tag.split("}")[1]
        if tag == "p":
            yield ("para", el)
        elif tag == "tbl":
            for tr in el.findall(qn("w:tr")):
                for tc in tr.findall(qn("w:tc")):
                    yield ("cell", tc)


def _locate(body, anchor):
    """章节标题定位：返回 (first_idx, last_idx, indices)。仅匹配段落。"""
    hits = []
    for i, el in enumerate(body):
        if el.tag.split("}")[1] != "p":
            continue
        if _norm(anchor) in _norm(_ptext(el)):
            hits.append(i)
    if not hits:
        return None, None, []
    return hits[0], hits[-1], hits


def _find_end(body, start_idx, end_anchor):
    """从 start_idx 之后找结尾锚点（如「本手册完」）的段落索引。"""
    for i in range(start_idx + 1, len(body)):
        el = body[i]
        if el.tag.split("}")[1] == "p" and _norm(end_anchor) in _norm(_ptext(el)):
            return i
    return None


def _apply_rewrites(body, rewrites):
    """全扫描段落+表格单元格，逐条应用引用改写；返回变更日志。"""
    changes = []
    for rule in rewrites:
        find = rule["find"]
        repl = rule["replace"]
        scope = rule.get("scope", "any")  # para / cell / any
        locate = rule.get("locate")       # 可选：元素须含此文本
        expected = int(rule.get("expected_hits", 1))
        hit = 0
        for kind, el in _iter_blocks(body):
            if scope != "any" and scope != kind:
                continue
            full = _ptext(el)
            if locate and _norm(locate) not in _norm(full):
                continue
            if find not in full:
                continue
            # run 级精确替换：优先在单个 run 内完成
            runs = list(el.iter(qn("w:r")))
            done_in_run = False
            for r in runs:
                for t in r.findall(qn("w:t")):
                    if t.text and find in t.text:
                        t.text = t.text.replace(find, repl)
                        done_in_run = True
            # 校验：拼接文本里 find 必须已消失（否则是跨 run，拒绝静默跳过）
            if find in _ptext(el):
                raise ValueError(
                    "引用「%s」在元素中跨多个 run（无法安全替换，需人工拆 run 或改用 "
                    "docx_revise_toolkit 整段回写）。元素文本：%s" % (find, _ptext(el)[:120])
                )
            if not done_in_run:
                raise ValueError("引用改写未落地：%s" % find)
            hit += 1
            changes.append("  • [%s] %s → %s" % (kind, find, repl))
        if hit != expected:
            raise ValueError(
                "引用「%s」期望命中 %d 处，实得 %d 处——请检查 locate/scope 或文档是否已改"
                % (find, expected, hit)
            )
    return changes


def _fulltext(body):
    return "\n".join(_ptext(el) for _, el in _iter_blocks(body))


def _verify(body, cfg):
    """验收断言：残留词零命中 + 保留词存在 + 结尾锚点唯一。"""
    ft = _fulltext(body)
    for kw in cfg.get("verify_absent", []):
        n = ft.count(kw)
        if n != 0:
            raise ValueError("验收失败：残留「%s」%d 处（应为 0）" % (kw, n))
    for kw in cfg.get("verify_present", []):
        if kw not in ft:
            raise ValueError("验收失败：应保留的「%s」丢失" % kw)
    tail = cfg.get("tail_anchor")
    if tail:
        n = sum(1 for kind, el in _iter_blocks(body)
                if kind == "para" and _norm(tail) in _norm(_ptext(el)))
        if n != 1:
            raise ValueError("验收失败：结尾锚点「%s」出现 %d 次（应为 1）" % (tail, n))


def run(input_path, output_path, config_path, dry_run=False):
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    doc = Document(input_path)
    body = _body_elements(doc)

    chapter_anchor = cfg["chapter_anchor"]
    first, last, hits = _locate(body, chapter_anchor)
    if not hits:
        raise ValueError("未找到章节锚点：%s" % chapter_anchor)

    tail_anchor = cfg.get("tail_anchor")
    if not tail_anchor:
        raise ValueError("config 必须提供 tail_anchor（结尾锚点，如「本手册完」）")
    end_idx = _find_end(body, last, tail_anchor)
    if end_idx is None:
        raise ValueError("未找到结尾锚点：%s" % tail_anchor)
    if end_idx <= last:
        raise ValueError("结尾锚点位于章节标题之前，删除范围非法")

    # 删除集合：TOC 条目（首次出现，若与正文不同）+ 正文整块（含标题，到结尾锚点前）
    remove = set()
    if first != last and cfg.get("delete_toc_entry", True):
        remove.add(body[first])
    for i in range(last, end_idx):
        remove.add(body[i])
    n_removed = len(remove)

    plan = ["删除章节「%s」" % chapter_anchor]
    plan.append("  • TOC 条目：%s" % ("删除（索引 %d）" % first if first != last and cfg.get("delete_toc_entry", True) else "无/不删"))
    plan.append("  • 正文范围：索引 %d~%d（含标题，结尾锚点「%s」保留）" % (last, end_idx - 1, tail_anchor))
    plan.append("  • 删除元素总数：%d" % n_removed)
    if cfg.get("ref_rewrites"):
        plan.append("  • 悬空引用改写（段落+表格全扫描）：")
        plan.extend(_preview_rewrites(body, cfg["ref_rewrites"]))
    print("\n".join(plan))

    if dry_run:
        print("\n[DRY-RUN] 未写入任何文件。")
        return

    # 1) 先改写引用（删除会移除元素，故先改）
    changes = _apply_rewrites(body, cfg.get("ref_rewrites", []))
    # 2) 再删除章节块
    for el in remove:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    # 3) 验收（失败则抛错，不保存）
    _verify(_body_elements(doc), cfg)
    doc.save(output_path)

    print("\n已改写引用 %d 处：" % len(changes))
    for c in changes:
        print(c)
    print("已保存：%s（%d 字节）" % (output_path, os.path.getsize(output_path)))
    print("删除元素数：%d；验收通过。" % n_removed)


def _preview_rewrites(body, rewrites):
    """dry-run 预览：列出每条规则将命中的块数（不修改）。"""
    out = []
    for rule in rewrites:
        find = rule["find"]
        scope = rule.get("scope", "any")
        locate = rule.get("locate")
        n = 0
        for kind, el in _iter_blocks(body):
            if scope != "any" and scope != kind:
                continue
            full = _ptext(el)
            if locate and _norm(locate) not in _norm(full):
                continue
            if find in full:
                n += 1
        out.append("    - 「%s」命中 %d 处（scope=%s）" % (find, n, scope))
    return out


def selftest():
    tmpdir = tempfile.mkdtemp(prefix="delch_selftest_")
    src = os.path.join(tmpdir, "t.docx")
    doc = Document()
    # 目录伪条目（首次出现，需删除）
    doc.add_paragraph("5  附录 H  源文件清单与效力总览", style="List Bullet")
    # 保留章节 + 其中的悬空引用（真实场景：引用在其他章节，删除 H 后需清理）
    doc.add_paragraph("5  附录 A  限度表")
    p_ref = doc.add_paragraph()  # 正文段落引用（保留块内）
    p_ref.add_run("see 见附录 H.4（S17～S22） end")
    tbl = doc.add_table(rows=1, cols=1)  # 表格单元格引用（关键：必须扫到）
    tbl.cell(0, 0).text = "字段核对，见附录 H.4（S17～S22）"
    # 真实 H 章节标题（最后出现）+ 待删内容
    doc.add_paragraph("5  附录 H  源文件清单与效力总览")
    doc.add_paragraph("H.1 用途与口径")
    doc.add_paragraph("—— 本手册完 ——")
    doc.save(src)

    cfg = {
        "chapter_anchor": "附录 H  源文件清单与效力总览",
        "tail_anchor": "本手册完",
        "ref_rewrites": [
            {"find": "见附录 H.4（S17～S22）", "replace": "（S17～S22）", "expected_hits": 2}
        ],
        "verify_absent": ["附录 H", "H.4", "源文件清单与效力总览"],
        "verify_present": ["附录 A", "本手册完"],
    }
    cfg_path = os.path.join(tmpdir, "cfg.json")
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    out = os.path.join(tmpdir, "out.docx")
    run(src, out, cfg_path)

    d2 = Document(out)
    ft = _fulltext(_body_elements(d2))
    assert "附录 H" not in ft and "H.4" not in ft, "self-test: 残留未清"
    assert "附录 A" in ft and "本手册完" in ft, "self-test: 保留锚点丢失"
    assert "字段核对，（S17～S22）" in ft, "self-test: 表格单元格引用未改写"
    assert "see （S17～S22） end" in ft, "self-test: 段落引用未改写"
    assert d2.tables, "self-test: 表格不应被删"
    print("\n[SELFTEST] PASS")


def main():
    ap = argparse.ArgumentParser(description="删除章节并清理悬空引用")
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
