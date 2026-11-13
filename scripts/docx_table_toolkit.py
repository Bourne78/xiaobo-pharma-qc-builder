# -*- coding: utf-8 -*-
"""
Pharma QC Manual — python-docx 复用工具集
========================================
固定布局表格 / A4 版式 / OOXML 元素顺序 / 跨页重复表头 / SOP 引用块 等。

所有「会改动文档」的函数都把 python-docx 的 Document 实例作为第一个参数 `doc` 传入，
便于在一个脚本里多次复用、或在既有文档里追加内容。

依赖：python-docx>=1.0
用法：
    from docx_table_toolkit import (setup_page, add_grid_table_fixed,
                                    add_grid_table, add_kv_table,
                                    add_sop_quote_block,
                                    add_key_callout, add_header_footer,
                                    add_landscape_section, restore_portrait)
    doc = Document()
    setup_page(doc.sections[0], landscape=False)
    add_grid_table_fixed(doc, headers, data, widths, ...)
    doc.save("manual.docx")
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

# ---------- 配色 ----------
NAVY   = RGBColor(0x1F, 0x3A, 0x5F)
RED    = RGBColor(0xC0, 0x00, 0x00)
GREY   = RGBColor(0x55, 0x55, 0x55)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
HDR_FILL   = "1F3A5F"   # 表头深底
SUB_FILL   = "DDE6F0"   # 字段表 key 列
KEY_FILL   = "FBE4E4"   # 关键点红框底
ZEBRA_FILL = "F2F6FB"   # 斑马纹
SOP_FILL   = "F2F6FB"   # SOP 引用块底
SOP_EDGE   = "1F3A5F"   # SOP 引用块左竖线

# 字体族（可参数化）。默认配方：中文宋体 + 英文 Times New Roman。
# 既有「全宋体最小改动」成品可把 ASCII_FONT 设为 "宋体"（all-song 兼容模式）。
CN_FONT    = "宋体"                 # eastAsia 中文字体
BODY_FONT  = "宋体"                 # 正文默认
ASCII_FONT = "Times New Roman"      # ascii / hAnsi / cs

# 纵向版心 17.0cm；横向版心 25.7cm
PORTRAIT_TEXT_WIDTH_CM = 17.0
LANDSCAPE_TEXT_WIDTH_CM = 25.7


# =====================================================================
# 页面与字体
# =====================================================================
def setup_page(sec, landscape=False):
    """A4；纵向版心 17.0cm，横向版心 25.7cm。对 doc.sections[0] 调用一次。"""
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.0)
    return sec


def set_default_font(doc, name=BODY_FONT, size=10.5):
    """Normal 样式同步写 rFonts 四属性，避免复杂文种回退。"""
    style = doc.styles["Normal"]
    style.font.name = ASCII_FONT
    style.font.size = Pt(size)
    rFonts = style.element.rPr.rFonts
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), ASCII_FONT)
    rFonts.set(qn("w:eastAsia"), name)


def set_cjk(run, name=None, ascii_font=None):
    """设置 run 的 rFonts 四属性并让 w:sz / w:szCs 成对，避免 Word 回退字体。

    - name      : eastAsia（默认 CN_FONT，中文宋体）
    - ascii_font: ascii / hAnsi / cs（默认 ASCII_FONT，Times New Roman；
                  设为 "宋体" 即 all-song 全宋体兼容模式）
    - 空 run（无文本）跳过，避免撑高行距
    - w:szCs 必须用 sz.addnext(...) 紧跟 w:sz；勿用 rPr.append（OOXML 顺序会错）
    save 前仍须做全 run 后处理（含 Normal 样式），见 references/revision_playbook.md §3。
    """
    if not (run.text or "").strip():
        return
    ea = CN_FONT if name is None else name
    af = ASCII_FONT if ascii_font is None else ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), af)
    rFonts.set(qn("w:eastAsia"), ea)
    sz = rPr.find(qn("w:sz"))
    if sz is not None:
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            sz.addnext(szCs)
        szCs.set(qn("w:val"), sz.get(qn("w:val")) or "21")


# =====================================================================
# 单元格底层操作
# =====================================================================
def set_cell_borders(cell, **edges):
    """edges: top/left/bottom/right -> {'val':'single','sz':18,'color':'1F3A5F'}"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcB = tcPr.find(qn("w:tcBorders"))
    if tcB is None:
        tcB = OxmlElement("w:tcBorders")
        tcPr.append(tcB)
    for edge in ("top", "left", "bottom", "right"):
        if edge not in edges:
            continue
        ed = tcB.find(qn("w:" + edge))
        if ed is None:
            ed = OxmlElement("w:" + edge)
            tcB.append(ed)
        for k, v in edges[edge].items():
            ed.set(qn("w:" + k), str(v))


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def add_rich_runs(p, text, font=BODY_FONT, size=10.5, color=None, base_bold=False):
    """把带 **加粗** 标记的文本渲染到段落 p 中（** 成对出现时切换加粗）。
    base_bold=True 时全部加粗；文本中的 \\n 转为软换行。
    直接 add_run 会把 ** 原样显示，必须走本函数。"""
    segments = str(text).split("\n")
    for si, seg in enumerate(segments):
        if si > 0:
            p.add_run().add_break()
        for i, part in enumerate(seg.split("**")):
            if not part:
                continue
            r = p.add_run(part)
            set_cjk(r, font)
            r.font.size = Pt(size)
            r.font.bold = base_bold or (i % 2 == 1)
            if color is not None:
                r.font.color.rgb = color
    return p


def set_cell_text(cell, text, bold=False, color=None, size=9, font=BODY_FONT, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    add_rich_runs(p, text, font=font, size=size, color=color, base_bold=bold)
    return p


# =====================================================================
# 标题 / 段落 / 列表
# =====================================================================
def add_title(doc, text, sub=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_cjk(r, CN_FONT)
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = NAVY
    if sub:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(sub)
        set_cjk(r2, CN_FONT)
        r2.font.size = Pt(11); r2.font.color.rgb = GREY


def add_h1(doc, text):
    p = doc.add_heading(level=1)
    r = p.add_run(text); set_cjk(r, CN_FONT)
    r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = NAVY
    return p


def add_h2(doc, text):
    p = doc.add_heading(level=2)
    r = p.add_run(text); set_cjk(r, CN_FONT)
    r.font.size = Pt(12.5); r.font.bold = True; r.font.color.rgb = NAVY
    return p


def add_para(doc, text, size=10.5, bold=False, color=None, font=BODY_FONT, space=4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space)
    if align:
        p.alignment = align
    r = p.add_run(text); set_cjk(r, font)
    r.font.size = Pt(size); r.font.bold = bold
    if color is not None:
        r.font.color.rgb = color
    return p


def add_bullet(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    r = p.add_run(text); set_cjk(r, BODY_FONT)
    r.font.size = Pt(size)
    return p


# =====================================================================
# 普通表（流式布局，多表共用）
# =====================================================================
def add_kv_table(doc, rows, widths=None, header=None, key_w=3.2):
    """两列字段表。rows: [(k,v), ...]"""
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if header:
        hr = table.add_row().cells
        set_cell_text(hr[0], header[0], bold=True, color=WHITE, size=9.5, font=CN_FONT); shade(hr[0], HDR_FILL)
        set_cell_text(hr[1], header[1], bold=True, color=WHITE, size=9.5, font=CN_FONT); shade(hr[1], HDR_FILL)
    for i, (k, v) in enumerate(rows):
        cells = table.add_row().cells
        set_cell_text(cells[0], k, bold=True, color=NAVY, size=9.5, font=CN_FONT)
        shade(cells[0], SUB_FILL if i % 2 == 0 else ZEBRA_FILL)
        set_cell_text(cells[1], v, size=9.5)
    if widths:
        for row in table.rows:
            row.cells[0].width = Cm(widths[0]); row.cells[1].width = Cm(widths[1])
    return table


def add_grid_table(doc, headers, data, widths=None, font_size=9):
    """通用多列表（流式布局）。"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr[i], h, bold=True, color=WHITE, size=font_size, font=CN_FONT)
        shade(hdr[i], HDR_FILL)
    for ri, row in enumerate(data):
        cells = table.add_row().cells
        for ci, val in enumerate(row):
            set_cell_text(cells[ci], str(val), size=font_size)
            if ri % 2 == 1:
                shade(cells[ci], ZEBRA_FILL)
    if widths:
        for r in table.rows:
            for ci, w in enumerate(widths):
                r.cells[ci].width = Cm(w)
    return table


# =====================================================================
# 固定布局表（附录 G 等多列表专用）
# 关键：python-docx add_table 会把 tblGrid 均分为 8640 twips(15.24cm)，不跟随
# cell.width；必须手动同步 gridCol，并显式写 w:tblW(dxa)，否则 Word/WPS 渲染列宽
# 与设定不符，且 Word 的 PDF 导出对「首页第一张固定表」会丢表体。
# =====================================================================
def _set_grid_widths(table, widths_cm):
    total = int(round(sum(widths_cm) * 567))          # 1cm = 567 twips
    grid = table._tbl.find(qn('w:tblGrid'))
    for c, w in zip(grid.findall(qn('w:gridCol')), widths_cm):
        c.set(qn('w:w'), str(int(round(w * 567))))
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert_element_before(tblW, 'w:tblLayout', 'w:tblCellMar',
                                    'w:tblLook', 'w:tblCaption')
    tblW.set(qn('w:w'), str(total))
    tblW.set(qn('w:type'), 'dxa')


def _set_fixed_layout(table):
    """w:tblLayout type=fixed。OOXML 顺序：tblLayout 必须位于 tblLook 之前，否则 Word 报文档损坏。"""
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(old)
    el = OxmlElement('w:tblLayout')
    el.set(qn('w:type'), 'fixed')
    tblPr.insert_element_before(el, 'w:tblCellMar', 'w:tblLook',
                                'w:tblCaption', 'w:tblDescription', 'w:tblPrChange')


def _repeat_header(row):
    """跨页重复标题行。trPr 顺序：cantSplit -> trHeight -> tblHeader。"""
    trPr = row._tr.get_or_add_trPr()
    for old in trPr.findall(qn('w:tblHeader')):
        trPr.remove(old)
    el = OxmlElement('w:tblHeader')
    trPr.insert_element_before(el, 'w:tblCellSpacing', 'w:jc', 'w:hidden',
                                'w:ins', 'w:del', 'w:trPrChange')


def _no_split(row):
    """禁止该行跨页断开。cantSplit 必须排在 tblHeader 之前。"""
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:cantSplit')
    trPr.insert_element_before(el, 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing',
                                'w:jc', 'w:hidden', 'w:ins', 'w:del', 'w:trPrChange')


def add_grid_table_fixed(doc, headers, data, widths, font_size=8,
                         center_cols=(), min_row_cm=0.62, header_size=None):
    """附录 G 专用多列表格。
    widths: 各列宽(cm)，合计必须 <= 当前节版心（纵向 17.0 / 横向 25.7）
    center_cols: 需居中的列索引集合（编号列、勾选列等）"""
    assert len(widths) == len(headers), "widths 长度须等于 headers 长度"
    avail = LANDSCAPE_TEXT_WIDTH_CM if doc.sections[-1].orientation == WD_ORIENT.LANDSCAPE else PORTRAIT_TEXT_WIDTH_CM
    assert round(sum(widths), 4) <= avail + 1e-6, \
        "列宽合计 %.2fcm 超出版心 %.1fcm" % (sum(widths), avail)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_grid_widths(table, widths)
    _set_fixed_layout(table)
    hsize = header_size or font_size
    for i, h in enumerate(headers):
        c = table.rows[0].cells[i]
        set_cell_text(c, h, bold=True, color=WHITE, size=hsize, font=CN_FONT,
                      align=WD_ALIGN_PARAGRAPH.CENTER)
        shade(c, HDR_FILL)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # 仅对会跨页的较长表（>12 数据行）加重复表头；短表（如 8 行）不加。
    # 原因：固定布局表「位于页首 + 表头行带 tblHeader」会触发 Word PDF 导出丢表体。
    if len(data) > 12:
        _repeat_header(table.rows[0])
    # 表头行不加 cantSplit（固定布局表头若带 cantSplit 在页首易丢表体）
    for ri, row in enumerate(data):
        r = table.add_row()
        for ci, val in enumerate(row):
            c = r.cells[ci]
            al = WD_ALIGN_PARAGRAPH.CENTER if ci in center_cols else None
            set_cell_text(c, str(val), size=font_size, align=al)
            if ri % 2 == 1:
                shade(c, ZEBRA_FILL)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        r.height = Cm(min_row_cm)
        r.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        _no_split(r)
    for r in table.rows:
        for ci, w in enumerate(widths):
            r.cells[ci].width = Cm(w)
    return table


# =====================================================================
# 关键点框 / SOP 引用块
# =====================================================================
# 带圈数字（用于「关键点 / 易错点」内多序号逐条分行）
_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩"


def _split_numbered(text):
    """按开头带圈数字（①②③…）的位置把文本切成多条；无带圈序号则整体返回一条。
    序号本身保留在每条开头，便于阅读时逐条扫读。"""
    segs = re.split(r"(?=[①②③④⑤⑥⑦⑧⑨⑩])", text)
    return [s for s in segs if s.strip()]


def add_key_callout(doc, text, title="● 关键点 / 易错点"):
    """红色关键点提示框（单格表）。
    text 支持 **加粗** 标记；当含多个带圈序号（①②③…）时，每条独立成行，
    无序号则保持单段。整框加 w:cantSplit 防跨页断裂，左竖线加粗红色。
    （cantSplit 必须插在 w:trPr 首位，否则 Word 忽略。）
    标题默认用 ●：宋体下 🔴 / ⚠️ 不稳，成品文档内统一 ★ / ● / △；
    技能叙述里的「🔴 关键点框」仅为指代。验收用精确串「● 关键点 / 易错点」；
    既有脚本的 startswith("● 关键点") 计数与新默认串兼容；旧稿可显式传 title= 保持原样。"""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(title)
    set_cjk(r, CN_FONT); r.font.bold = True; r.font.size = Pt(10); r.font.color.rgb = RED
    # 多序号条目逐条分行；无序号则保持单段
    items = _split_numbered(text)
    if len(items) <= 1:
        p2 = cell.add_paragraph()
        add_rich_runs(p2, text, font=BODY_FONT, size=9.5)
    else:
        for it in items:
            pi = cell.add_paragraph()
            pi.paragraph_format.space_after = Pt(2)
            pi.paragraph_format.line_spacing = 1.05
            add_rich_runs(pi, it, font=BODY_FONT, size=9.5)
    shade(cell, KEY_FILL)
    # 左边框加粗红色，视觉上区别于普通引用块
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single"); left.set(qn("w:sz"), "18"); left.set(qn("w:color"), "C00000")
    borders.append(left)
    tcPr.append(borders)
    # 整框不跨页断开（cantSplit 须在 trPr 首位）
    trPr = table.rows[0]._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.insert(0, cant)
    return table


def _add_field(paragraph, field_code):
    """在段落内插入一个 Word 域（field），用于 PAGE / NUMPAGES 等动态字段。
    用 OxmlElement 构造 fldChar(begin)/instrText/fldChar(end) 三元组。
    注意：若环境允许 Word COM，导出 PDF 前必须 $doc.Fields.Update()，否则页码会显示 0；
    本机 COM 被沙箱拦截时 PDF 须由用户在 Word 手动导出（见 python_docx_pitfalls.md §10）。"""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = f" {field_code} "
    fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_end)
    return run


def add_header_footer(doc, header_text):
    """设置页眉 + 页脚，适用于单节或多节文档（如含横向附录的文档）。

    页眉：第 1 节写描述性文字（居中 8pt 灰字）；其余节链接沿用第 1 节。
    页脚：「第 X 页 / 共 Y 页」，X=PAGE 域，Y=NUMPAGES 域（居中 8pt 灰字）；
          其余节链接沿用第 1 节。

    顺序要点：必须先在 sections[0] 写好页眉页脚，再把后续节设为
    is_linked_to_previous=True，否则链接会覆盖已写内容。"""
    # ---- 页眉 ----
    hsec = doc.sections[0]
    hp = hsec.header.paragraphs[0]
    hp.text = ""
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = hp.add_run(header_text)
    set_cjk(hr, CN_FONT); hr.font.size = Pt(8); hr.font.color.rgb = GREY
    for s in doc.sections[1:]:
        s.header.is_linked_to_previous = True
    # ---- 页脚 ----
    fsec = doc.sections[0]
    fp = fsec.footer.paragraphs[0]
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("第 "); set_cjk(r, CN_FONT); r.font.size = Pt(8); r.font.color.rgb = GREY
    _add_field(fp, "PAGE")
    r = fp.add_run(" 页 / 共 "); set_cjk(r, CN_FONT); r.font.size = Pt(8); r.font.color.rgb = GREY
    _add_field(fp, "NUMPAGES")
    r = fp.add_run(" 页"); set_cjk(r, CN_FONT); r.font.size = Pt(8); r.font.color.rgb = GREY
    for s in doc.sections[1:]:
        s.footer.is_linked_to_previous = True
    return doc


def add_sop_quote_block(doc, sop_ref, body_text, size=9):
    """卡内浅底引用块：左竖线 + 浅蓝底 + 宋体小字，标注 SOP 章节号与页码。
    单元格内 \\n 换行由 add_run 自动转 <w:br/>。"""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(sop_ref)
    set_cjk(r, CN_FONT); r.font.bold = True; r.font.size = Pt(size + 0.5)
    r.font.color.rgb = NAVY
    p2 = cell.add_paragraph()
    r2 = p2.add_run(body_text)
    set_cjk(r2, BODY_FONT); r2.font.size = Pt(size)
    shade(cell, SOP_FILL)
    set_cell_borders(cell, left={'val': 'single', 'sz': 24, 'color': SOP_EDGE})
    return table


# =====================================================================
# 横向 section（宽表用）
# =====================================================================
def add_landscape_section(doc):
    """在文档末尾插入一个横向新节（版心 25.7cm），返回该节。
    宽表放此节；末尾调用 restore_portrait 切回纵向。"""
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.footer.is_linked_to_previous = True   # 继承第 1 节页脚
    return sec


def restore_portrait(doc):
    """横向宽表之后切回纵向 section。"""
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2.0)
    sec.footer.is_linked_to_previous = True
    return sec
