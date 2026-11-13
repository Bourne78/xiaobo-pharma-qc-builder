# -*- coding: utf-8 -*-
"""
附录 G  检验物料清单（备料核对清单）数据源模板
============================================

本文件是**模板**：给出六类物料的数据结构与自动派生「检验项 × 物料」矩阵的函数。
做新品种手册时，把 MATERIALS / CARD_INFO 替换为你产品的真实数据，字段结构不变，
build_matrix() 会自动算出矩阵，杜绝手工维护两份清单导致不同步。

六类物料（ID 前缀固定）
----------------------
对照品与工作对照品   RP-   (Reference Product)
试剂、试液与标准溶液 RGT-  (ReaGenT)
供试品与样品制备     SMP-  (SaMPle)
色谱柱与固定相       COL-  (COLumn)
仪器设备             INS-  (INstrumentS)
消耗品与器皿         CON-  (CONsumables)

单条物料字段
------------
id        物料编号（前缀见上）
name      名称
spec      规格      （SOP 未载明填 NEED）
lot       批号      （SOP 未载明填 NEED；批检验记录可回填，见下方分级口径）
supplier  供应商    （可选；批检验记录回填，来源多写法时以主表为准并留「待核对」）
expiry    有效期    （SOP 未载明填 NEED）
storage   保存条件  （SOP 未载明填 NEED）
prep      配制/使用说明（逐字取自 SOP 原文，不改写数字）
qty       需求量    （确定值 / 按SOP配方 / 需查）
use       用途说明
items     该物料服务的手册卡号列表，如 ["3.7", "3.8"]  —— build_matrix() 依赖此字段
src       来源标记（见下方 SRC_*）
src_ref   SOP/标准 章节或页码出处

来源标记口径
------------
SOP 原文载明          → SRC_SOP
质量标准载明、SOP 未载 → SRC_STD（须确认是否纳入 SOP 修订）
SOP 与标准均未载明    → SRC_CHECK（填「需查」；**禁止**用方法学验证报告批号充当对照品证书信息）
批检验记录扫描件      → 实际执行记录，可逐份汇总回填 lot / supplier，须标注来源 + 留「待核对」
                        （见 references/revision_playbook.md §4）
外部单位负责          → SRC_OUT（外包项）

用法
----
    from materials_template import MATERIALS, build_matrix, MATRIX_CAT_HEADERS, CARD_INFO
    MATRIX = build_matrix()                 # {卡号: [各类物料串]}，默认「完整名称（编号）」
    # MATRIX["3.7"] -> ["XX 对照品（RP-01）、XX 工作对照品（RP-02）", ...]
    # 列头用 MATRIX_CAT_HEADERS（中文全称（缩写））；裁剪列见 revision_playbook §6
"""

# ---------------------------------------------------------------- 来源标记
SRC_SOP   = "SOP原文"              # SOP 现行版明确载明
SRC_STD   = "质量标准(SOP未载)"     # 质量标准载明、SOP 未载（须确认是否纳入 SOP 修订）
SRC_CHECK = "需查"                 # 现行 SOP 与质量标准均未载明
SRC_OUT   = "外包"                 # 外部单位开发验证并负责检测

# ---------------------------------------------------------------- 需求量标注
Q_EXACT  = "按SOP可确定量"
Q_RECIPE = "按SOP配方"
Q_NONE   = "需查"

NEED = "需查"   # SOP 未载时的统一占位


# ================================================================
# 手册操作卡 ↔ SOP 章节号 对照（模板示例，替换为你的产品）
# 注意：手册第 3 章的卡号与 SOP 的章节编号**往往不一致**，核对物料时以本表换算。
# 每行： (卡号, 检验项, SOP章节号, SOP页码, 关键物料类别, 首次开检必齐项)
# ================================================================
CARD_INFO = [
    ("3.1",  "性状",           "§1.1–1.2",              "P2",     "无",                "无（目视检查，仅需取样）"),
    ("3.2",  "鉴别",           "§2.1–2.2",              "P2–3",   "试剂+对照品+柱",    "对照品、色谱柱、相应试剂"),
    ("3.3",  "pH 值",          "§3.1",                  "P3",     "缓冲液+pH计",       "pH计、pH标准缓冲液"),
    ("3.7",  "有关物质",       "§3.3",                  "P5",     "对照品+柱+HPLC",    "有关物质对照品、色谱柱、检测器"),
    ("3.8",  "含量测定",       "§4",                    "P12–13", "对照品+柱+HPLC",    "含量对照品、色谱柱、分析天平"),
    ("3.9",  "残留溶剂",       "§3.6",                  "P9",     "试剂+柱+GC",        "残留溶剂对照品、顶空瓶、气相色谱仪"),
]

CARD_SOP_MAP = {c[0]: c[2] for c in CARD_INFO}
CARD_NAME_MAP = {c[0]: c[1] for c in CARD_INFO}


# ================================================================
# 物料主数据（模板示例：少量样本，替换为你产品的完整清单）
# 六类以字典 key 区分；每类是一个 list[dict]。
# ================================================================
MATERIALS = {

    # ---------------------------------------------------------- 对照品与工作对照品
    "对照品与工作对照品": [
        dict(id="RP-01", name="XX 对照品",
             spec=NEED, lot=NEED, expiry=NEED, storage=NEED,
             prep="§x.x 对照品溶液：取 XX 对照品约 Xmg，精密称定，置 Xxml 量瓶中，加水溶解并稀释至刻度，摇匀。",
             qty="约 Xmg/份（平行 2 份）",
             use="系统适用性溶液、含量测定对照品溶液、鉴别",
             items=["3.2", "3.7", "3.8"], src=SRC_SOP, src_ref="§x.x、§x.x"),
        dict(id="RP-02", name="XX 工作对照品",
             spec=NEED, lot=NEED, expiry=NEED, storage=NEED,
             prep="§x.x 工作对照品溶液：取 XX 工作对照品……",
             qty="按SOP配方",
             use="有关物质系统适用性",
             items=["3.7"], src=SRC_SOP, src_ref="§x.x"),
    ],

    # ---------------------------------------------------------- 试剂、试液与标准溶液
    "试剂、试液与标准溶液": [
        dict(id="RGT-01", name="XX 试剂",
             spec="AR", lot=NEED, expiry=NEED, storage="室温",
             prep="直接取用；必要时临用新配。",
             qty=Q_EXACT,
             use="显色/反应体系",
             items=["3.2"], src=SRC_SOP, src_ref="§x.x"),
        dict(id="RGT-02", name="纯化水",
             spec="药典", lot=NEED, expiry=NEED, storage="室温",
             prep="直接取用。",
             qty=Q_EXACT,
             use="溶剂、稀释",
             items=["3.3", "3.7", "3.8"], src=SRC_SOP, src_ref="§x.x"),
    ],

    # ---------------------------------------------------------- 供试品与样品制备
    "供试品与样品制备": [
        dict(id="SMP-01", name="XX 注射液（规格A）",
             spec="2ml:10mg", lot=NEED, expiry=NEED, storage="阴凉",
             prep="原包装取样，按各检验项要求开瓶/稀释。",
             qty="每检验项按需",
             use="全部检验项供试品",
             items=["3.1", "3.2", "3.3", "3.7", "3.8", "3.9"], src=SRC_SOP, src_ref="§x.x"),
    ],

    # ---------------------------------------------------------- 色谱柱与固定相
    "色谱柱与固定相": [
        dict(id="COL-01", name="XX 色谱柱",
             spec="4.6×250mm, 5μm", lot=NEED, expiry=NEED, storage="室温",
             prep="按 SOP 平衡至基线稳定。",
             qty=Q_NONE,
             use="有关物质/含量 HPLC",
             items=["3.7", "3.8"], src=SRC_SOP, src_ref="§x.x"),
    ],

    # ---------------------------------------------------------- 仪器设备
    "仪器设备": [
        dict(id="INS-01", name="高效液相色谱仪（配 RID）",
             spec="—", lot=NEED, expiry=NEED, storage="实验室",
             prep="按 SOP 完成系统适用性与校验。",
             qty=Q_NONE,
             use="有关物质、含量测定",
             items=["3.7", "3.8"], src=SRC_SOP, src_ref="§x.x"),
        dict(id="INS-02", name="pH 计",
             spec="—", lot=NEED, expiry=NEED, storage="实验室",
             prep="用两种标准缓冲液校准。",
             qty=Q_NONE,
             use="pH 值",
             items=["3.3"], src=SRC_SOP, src_ref="§x.x"),
    ],

    # ---------------------------------------------------------- 消耗品与器皿
    "消耗品与器皿": [
        dict(id="CON-01", name="量瓶（10ml）",
             spec="A 级", lot=NEED, expiry=NEED, storage="器皿柜",
             prep="清洗、干燥后使用。",
             qty=Q_EXACT,
             use="溶液配制",
             items=["3.2", "3.7", "3.8"], src=SRC_SOP, src_ref="§x.x"),
    ],
}


# ================================================================
# 副表：检验项 × 物料 矩阵（由 MATERIALS 的 items 字段自动生成，杜绝手写两份不同步）
# ================================================================
MATRIX_CATS = ["对照品与工作对照品", "试剂、试液与标准溶液", "供试品与样品制备",
               "色谱柱与固定相", "仪器设备", "消耗品与器皿"]

# 列头：中文全称（缩写）。顺序与 MATRIX_CATS 一一对应。
# 裁剪某类别列时，须把该类别同时从 cats 与 headers 移除，并同步
# tblW / gridCol / 每行每格 tcW 三者，见 references/revision_playbook.md §6。
MATRIX_CAT_HEADERS = ["对照品与工作对照品（RP-）", "试剂、试液与标准溶液（RGT-）",
                      "供试品与样品制备（SMP-）", "色谱柱与固定相（COL-）",
                      "仪器设备（INS-）", "消耗品与器皿（CON-）"]
# 兼容别名：旧 import 仍可用，但**值已是中文全称（缩写）**——
# 不要照旧把它当「对照品 / 试剂·试液」这类简称列头使用，避免写出裸缩写列头。
MATRIX_CAT_SHORT = MATRIX_CAT_HEADERS


def build_matrix(cats=None, legacy=False):
    """返回 {手册卡号: [各类别的物料串]}。单一数据源仍是 MATERIALS[cat] 的 items 字段。

    默认（成品可读化口径）：每个物料写作「完整名称（编号）」，同类多个用顿号「、」连接，
    无物料填「—」；列头请用 MATRIX_CAT_HEADERS（中文全称（缩写））。
    cats: 可裁剪的类别列表，默认 MATRIX_CATS；去掉某类即少一列。
    legacy=True: 保留旧「名称 编号 + 换行」格式，仅供旧稿兼容。
    详见 references/revision_playbook.md §6。
    """
    cats = MATRIX_CATS if cats is None else cats
    out = {}
    for card in CARD_SOP_MAP:
        row = []
        for cat in cats:
            hit = [m for m in MATERIALS.get(cat, []) if card in m["items"]]
            if legacy:
                names = ["%s %s" % (m["name"], m["id"]) for m in hit]
                sep = "\n"
            else:
                names = ["%s（%s）" % (m["name"], m["id"]) for m in hit]
                sep = "、"
            row.append(sep.join(names) if names else "—")
        out[card] = row
    return out


MATRIX = build_matrix()
