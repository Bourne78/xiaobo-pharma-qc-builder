# -*- coding: utf-8 -*-
"""把附录 G 的 G.3「首次开检前置条件（逐项打勾）」12 条清单并入 G.1。

口径（用户确认）：
  1) G.1 标题改为「G.1  用途与口径与首次开检前置条件」（就地改第二 run，保格式）；
  2) 12 条 List Bullet 清单整段搬移到 G.1 两段说明之后（直接移动 XML 元素，格式零损耗）；
  3) 删除 G.3 标题段；
  4) 承载分节的空段 [213]/[217]、矩阵块 [214][215][216] 一律不动。

验收：白名单外逐块 c14n 指纹零改动；G.3 悬空引用零残留；12 条清单全部命中且顺序正确；
      分节 [纵,横,纵] 未塌；尾部完整。

用法： python merge_g3_into_g1.py [--dry-run]
"""
import sys, os, hashlib, shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

SRC = Path(r"E:\Workbuddy\文献整理\注射用乳糖酸红霉素_全项检验准备与关键点控制手册_v8_20261115.docx")
TMP = Path(r"E:\Workbuddy\文献整理\_tmp_merge_g3.docx")

# 改前 body 索引（只计 w:p / w:tbl，见探查）
I_G1_TITLE = 209
I_G1_P1    = 210
I_G1_P2    = 211
I_SECT1    = 213   # 纵向→横向 分节空段（不可动）
I_MAT_SUB  = 214   # 矩阵小标题（不可动）
I_MAT_TBL  = 215   # 矩阵表（不可动）
I_MAT_NOTE = 216   # 矩阵说明（不可动）
I_SECT2    = 217   # 横向→纵向 分节空段（不可动）
I_G3_TITLE = 218
I_G3_FIRST = 219
I_G3_LAST  = 230
I_END      = 231   # —— 本手册完 ——

G1_TITLE_OLD = "G.1  用途与口径"
G1_TITLE_NEW = "G.1  用途与口径与首次开检前置条件"
G3_TITLE_TXT = "G.3  首次开检前置条件（逐项打勾）"

EXPECT_N = 232   # 改前 p/tbl 块数


def body_blocks(doc):
    """只收集 body 直属的 w:p / w:tbl（与探查索引口径一致）。"""
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag in (qn('w:p'), qn('w:tbl')):
            out.append(ch)
    return out


def fp(el):
    return hashlib.sha256(etree.tostring(el, method='c14n')).hexdigest()


def ptext(el):
    return Paragraph(el, None).text


def main():
    dry = '--dry-run' in sys.argv
    doc = Document(str(SRC))
    body = body_blocks(doc)
    print(f"改前 p/tbl 块数={len(body)}")
    assert len(body) == EXPECT_N, f"块数异常: {len(body)}"

    # ---------- 前置断言：锁定所有锚点 ----------
    assert ptext(body[I_G1_TITLE]).strip() == G1_TITLE_OLD, ptext(body[I_G1_TITLE])
    assert ptext(body[I_G3_TITLE]).strip() == G3_TITLE_TXT, ptext(body[I_G3_TITLE])
    assert '本手册完' in ptext(body[I_END]), ptext(body[I_END])
    # 分节空段
    for si in (I_SECT1, I_SECT2):
        assert ptext(body[si]).strip() == '', f"[{si}] 应为空段"
        assert body[si].find('.//'+qn('w:sectPr')) is not None, f"[{si}] 缺 sectPr"
    # 12 条清单编号连续 1..12
    for k, idx in enumerate(range(I_G3_FIRST, I_G3_LAST+1), start=1):
        p = body[idx]
        para = Paragraph(p, doc)
        assert para.style.name == 'List Bullet', (idx, para.style.name)
        assert para.text.strip().startswith(f"{k}. "), (idx, para.text.strip()[:20])

    # 改前快照（含分节/矩阵块）
    lock_idx = [I_SECT1, I_MAT_SUB, I_MAT_TBL, I_MAT_NOTE, I_SECT2]
    lock_fp = {i: fp(body[i]) for i in lock_idx}
    list_fp = [fp(body[i]) for i in range(I_G3_FIRST, I_G3_LAST+1)]
    list_txt = [ptext(body[i]) for i in range(I_G3_FIRST, I_G3_LAST+1)]

    if dry:
        print("[dry-run] 全部前置断言通过；将执行：")
        print(f"  - G.1 标题: {G1_TITLE_OLD!r} -> {G1_TITLE_NEW!r}")
        print(f"  - 搬移 12 条 List Bullet 到 [{I_G1_P2}] 之后")
        print(f"  - 删除 G.3 标题 [{I_G3_TITLE}]")
        print("  - 保留分节空段与矩阵块原样")
        return

    # ---------- 1) 改 G.1 标题第二 run（保格式） ----------
    g1t = Paragraph(body[I_G1_TITLE], None)
    assert len(g1t.runs) == 2, len(g1t.runs)
    assert g1t.runs[0].text.strip() == 'G.1', repr(g1t.runs[0].text)
    g1t.runs[1].text = "用途与口径与首次开检前置条件"
    print(f"[1] G.1 标题已改: {ptext(body[I_G1_TITLE]).strip()!r}")

    # ---------- 2) 搬移 12 条清单到 G.1 第二说明段之后 ----------
    anchor = body[I_G1_P2]
    for idx in range(I_G3_FIRST, I_G3_LAST+1):
        anchor.addnext(body[idx])   # addnext 会移动元素本身，逐个接在 anchor 后
        anchor = body[idx]          # 下一个接到上一个之后，保持顺序
    print(f"[2] 12 条清单已搬到 [{I_G1_P2}] 之后（顺序 1..12）")

    # ---------- 3) 删除 G.3 标题 ----------
    g3t_el = body[I_G3_TITLE]
    g3t_el.getparent().remove(g3t_el)
    print("[3] G.3 标题已删除")

    after = body_blocks(doc)
    print(f"改后 p/tbl 块数={len(after)} (预期 {EXPECT_N-1})")
    assert len(after) == EXPECT_N-1, len(after)

    # ---------- 验收 ----------
    # 新块索引（相对改后）：删了原 218，搬走 219~230 到 211 后
    # 改后顺序: ...210,211, [219..230], 212,213,214,215,216,217, 231
    # 锚点新位置: 213->216, 214->217, 215->218, 216->219, 217->220
    # 精确映射：old==218(G3标题) 删除；219..230 搬到 211 后变 212..223；
    # 212..217 因前面插入 12 条 => new = old+12；<=211 不变。
    def new_index(old):
        if old == I_G3_TITLE:
            return None
        if I_G3_FIRST <= old <= I_G3_LAST:
            return I_G1_P2 + 1 + (old - I_G3_FIRST)   # 212..223
        if old <= I_G1_P2:
            return old
        return old + 12   # 212..217 -> 224..229
    # 分节/矩阵块: 213->225, 214->226, 215->227, 216->228, 217->229
    for old in lock_idx:
        ni = new_index(old)
        assert fp(after[ni]) == lock_fp[old], f"锁块 [{old}]->[{ni}] 指纹变化！"
    print("[验收] 分节空段 + 矩阵块指纹全部不变 ✓")

    # 12 条清单在新位置内容逐条一致、顺序正确
    for k in range(12):
        ni = new_index(I_G3_FIRST + k)
        assert fp(after[ni]) == list_fp[k], f"清单第{k+1}条指纹变化"
        assert ptext(after[ni]) == list_txt[k], f"清单第{k+1}条文本变化"
    print("[验收] 12 条清单内容 + 顺序全部一致 ✓")

    # G.1 标题
    assert ptext(after[I_G1_TITLE]).strip() == G1_TITLE_NEW
    # G.3 标题残留
    allp = [ptext(x) for x in after if x.tag == qn('w:p')]
    assert not any(G3_TITLE_TXT in t for t in allp), "G.3 标题残留"
    assert not any(t.strip().startswith('G.3') for t in allp), "G.3 编号残留"
    print("[验收] G.1 标题已更新、G.3 标题零残留 ✓")

    # 尾部
    assert '本手册完' in ptext(after[-1]), ptext(after[-1])[:40]
    # 分节方向
    assert len(doc.sections) == 3, len(doc.sections)
    orients = [s.orientation for s in doc.sections]
    assert orients == [0, 1, 0], orients
    print(f"[验收] 文档尾完整、分节方向 {orients} 未塌 ✓")

    doc.save(str(TMP))
    print(f"\n已保存临时文件: {TMP}  size={os.path.getsize(TMP)}")
    print("请用覆盖脚本将 _tmp_merge_g3.docx 覆盖源文件（备份已就位）。")


if __name__ == '__main__':
    main()
