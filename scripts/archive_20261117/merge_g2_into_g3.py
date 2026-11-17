# -*- coding: utf-8 -*-
"""把附录 G 的 G.2.1~G.2.6 六表合并进 G.3，顺延重编号 G.3->G.2 / G.4->G.3。
lxml 手术式：固定元素引用、构造后插入、sectPr 原位不动。v7 原地改写。
"""
import sys
from copy import deepcopy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = sys.argv[1]
OUT = sys.argv[2]

doc = Document(SRC)
body = doc.element.body
E = list(body)  # 一次性固定引用
assert len(E) == 247, "body 元素数不符: %d" % len(E)

def pt(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t")))

def set_rt(run, text):
    """改 run 的 w:t.text，保留 rPr；无 w:t 则报错。"""
    t = run.find(qn("w:t"))
    assert t is not None, "run 无 w:t"
    t.text = text
    if text != text.strip():
        t.set(qn("xml:space"), "preserve")

def runs(el):
    return el.findall(qn("w:r"))

# ---------- 0. 身份断言 ----------
assert pt(E[211]).startswith("5  附录 G"), "211 非附录G标题"
assert pt(E[212]).startswith("G.1"), "212 非G.1标题"
assert pt(E[214]).startswith("单一数据源"), "214 非G.1正文2"
assert pt(E[228]).startswith("G.3"), "228 非原G.3标题"
assert pt(E[230]).startswith("矩阵列示物料编号"), "230 非矩阵说明段"
assert pt(E[232]).startswith("G.4"), "232 非原G.4标题"
assert pt(E[245]).strip() == "—— 本手册完 ——", "245 非本手册完"

# ---------- 1. 构造合并总表 65x6 ----------
TPL = E[218]  # G.2.2 表：表头带 tblHeader，列宽模板
W = [548, 2224, 1591, 2097, 2698, 480]  # 合计 9638
assert sum(W) == 9638

new_tbl = OxmlElement("w:tbl")
new_tbl.append(deepcopy(TPL.find(qn("w:tblPr"))))
new_tbl.append(deepcopy(TPL.find(qn("w:tblGrid"))))
new_tbl.append(deepcopy(TPL.findall(qn("w:tr"))[0]))  # 模板表头(带 tblHeader)

TABLE_IDX = [216, 218, 220, 222, 224, 226]  # RP/RGT/SMP/COL/INS/CON 顺序
data_count = {}
for idx in TABLE_IDX:
    src_trs = E[idx].findall(qn("w:tr"))
    data = src_trs[1:]  # 去表头
    data_count[idx] = len(data)
    for tr in data:
        ntr = deepcopy(tr)
        tcs = ntr.findall(qn("w:tc"))
        assert len(tcs) == 6, "列数非6: %d" % idx
        for tc, w in zip(tcs, W):
            tcPr = tc.find(qn("w:tcPr"))
            tcW = tcPr.find(qn("w:tcW"))
            assert tcW is not None, "tcW 缺失"
            tcW.set(qn("w:w"), str(w))
            tcW.set(qn("w:type"), "dxa")
        new_tbl.append(ntr)

assert data_count == {216: 4, 218: 25, 220: 9, 222: 1, 224: 12, 226: 13}, data_count
assert sum(data_count.values()) == 64
assert len(new_tbl.findall(qn("w:tr"))) == 65
print("[1] 合并总表构造完成: 65行x6列, 64数据行, 列宽9638")

# ---------- 2. 逐 run 改写编号/措辞 ----------
# E214: G.3->G.2 ; G.2六张分节 -> 本节物料明细总
r214 = runs(E[214])
assert r214[2].find(qn("w:t")).text == "G.3 "
assert r214[8].find(qn("w:t")).text == " G.2 "
assert r214[9].find(qn("w:t")).text == "六张分节"
set_rt(r214[2], "G.2 ")
set_rt(r214[8], "本节")
set_rt(r214[9], "物料明细总")

# E228 原G.3标题 -> 新G.2标题
r228 = runs(E[228])
assert [r.find(qn("w:t")).text for r in r228] == ["G.3  ", "检验项", " × ", "物料矩阵（由物料", " items ", "字段自动生成）"]
set_rt(r228[0], "G.2  ")
set_rt(r228[1], "物料明细总表")
set_rt(r228[2], "与")
set_rt(r228[3], "检验项矩阵（矩阵由物料")
# r4/r5 不变

# E230: 见 G.2.1～G.2.6 -> 见上方物料明细总表
r230 = runs(E[230])
assert r230[8].find(qn("w:t")).text == "G.2.1"
assert r230[9].find(qn("w:t")).text == "～"
assert r230[10].find(qn("w:t")).text == "G.2.6"
set_rt(r230[7], "")
set_rt(r230[8], "上方物料明细总表")
set_rt(r230[9], "")
set_rt(r230[10], "")

# E232 原G.4标题 -> 新G.3标题
r232 = runs(E[232])
assert r232[0].find(qn("w:t")).text == "G.4  "
set_rt(r232[0], "G.3  ")
print("[2] 编号/措辞改写完成")

# ---------- 3. 元素重排（sectPr 原位不动） ----------
title_g2 = E[228]
# 前移新G.2标题到 E214 之后（纵向节内，E227 之前）
E[214].addnext(title_g2)
# 合并总表紧跟新G.2标题
title_g2.addnext(new_tbl)
# 删除原 6 标题+6 表 = E[215..226]（12 元素）
for e in E[215:227]:
    e.getparent().remove(e)

# 横向节首加不编号加粗小标题（插在 E227 之后 = 横向节开头，矩阵表 E229 之前）
subtitle = OxmlElement("w:p")
pPr = OxmlElement("w:pPr")
rPr_p = OxmlElement("w:rPr")
b = OxmlElement("w:b")
rPr_p.append(b)
pPr.append(rPr_p)
subtitle.append(pPr)
run = OxmlElement("w:r")
rPr = OxmlElement("w:rPr")
b2 = OxmlElement("w:b")
rPr.append(b2)
run.append(rPr)
t = OxmlElement("w:t")
t.text = "检验项 × 物料矩阵（由物料 items 字段自动生成）"
run.append(t)
subtitle.append(run)
E[227].addnext(subtitle)
print("[3] 元素重排完成: 新G.2标题+总表在纵向节, 矩阵小标题+矩阵在横向节, G.3 在末纵向节")

doc.save(OUT)
print("[4] 已保存:", OUT)
print("    新G.2标题:", pt(title_g2))
print("    E214:", pt(E[214])[:100])
print("    E230:", pt(E[230])[:100])
print("    新G.3标题:", pt(E[232]))
