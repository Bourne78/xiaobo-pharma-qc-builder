# 章节重构 SOP（跨品种通用）

`revision_playbook.md` 讲的是「改一份成稿」的总规则（写回安全、列宽、字体、验收、升版）。
本文件专讲其中**最容易出事故的一段**：**把一章合并 / 重编号 / 删减 / 精简 / 迁表 / 升版**。

> 分层：本文件全部内容为**第一层通用规则**，换任何品种都适用。
> 品种专项决策（某次手册的具体编号、具体限度、具体删哪一条）见
> `revision_playbook.md` §11，**禁止**升格为本文件的默认动作。
>
> 复用/指向：`revision_playbook.md` §1（写回安全）/ §5（手写代码路径）/ §8（验收）/
> §9（升版与收尾提醒）；`python_docx_pitfalls.md` §3/§6（元素顺序与 body 保序）；
> `manual_structure.md` §7（三层验收清单）。

---

## §1 「冗长」的量化与决策树

### §1.1 先量化，再动手

用户说「这一章内容乱 / 冗长，整理精简」时，**第一件事不是删字，是测量**：

```bash
python scripts/chapter_volume.py --docx X.docx --chapter 6
```

输出按小节给出：`正文` / `表内` / `合计` / `表数` / `行数` / `段数` / `bullet`。

**为什么必须把「表内」与「正文」分开统计**——这是本文件最重要的一条方法论：

| 口径 | 结论 | 结果 |
|---|---|---|
| 只看正文文字 | 「文字太多，删一删」 | 纯文字精简上限约 **8~25%** |
| 分开看 | 「体量大头在明细表」 | 搬表一次减 **50%+**，且表内事实一字不丢 |

> 真实案例：一章 8,371 字，其中 4 张明细表占 5,057 字（60%）。
> 纯文字精简到 7,691 字（−8.1%）就不动了；改为把明细表整体迁入新增附录后，
> 该章降到 3,833 字（−54.2%），**表内容一字未删**。

### §1.2 决策树

```
体量报告出来
   │
   ├─ 表内占比 ≥ 40%  →  【搬表】
   │     把明细表整体迁入新增附录（原章只留结论 + 指针）
   │     迁后必须改写指向旧位置的交叉引用
   │
   ├─ 表内占比 < 40%  →  【压缩文字】
   │     ① 合并冗长段落（一个密文段拆 2~3 段）
   │     ② 精简 bullet 组（N 条并 M 条，保留 ★ 级条目）
   │     ③ 删除纯重复小节（内容与别节 100% 重合）
   │
   └─ 内容重复（同一批信息在多处出现）
         →  【删重复节】+ 把唯一不重复的信息并入保留节的注记
```

**三条互斥的动作，不要叠加**：
- 搬表 → 该表所在小节自然变薄，**不要再顺手压该节的文字**（会误删有效信息）；
- 压缩文字 → 只动文字，**不要顺手搬表**（会把决策搅在一起，无法归因）；
- 删重复节 → 先确认重复的是**哪一份**，被删那份里的**独有信息**必须先并入保留处。

### §1.3 「删重复节」的唯一信息保全检查

删一节之前，逐条问：**这一节里有没有别处没有的信息？**

真实案例：删掉「受控文件与历史案例」节时，该节 95% 与别节重复，
但其中**签署链描述**与**配套报告编号**是独有的——必须先并入保留节的表后注记，
再删节。**先并后删，顺序不可反。**

---

## §2 「搬表而非删表」通则

### §2.1 三条铁律

1. **表内容一字不删，只挪位。** 迁表是**位置变更**，不是内容变更。
   迁后验收要断言**表维度不变**（`moved_tables.expect_dims`）。
2. **迁表后必须改写指向旧位置的交叉引用。** 原章里写着「见 6.7（见表）」，
   表搬到附录 H 之后，这句话就指向了空处——必须改成「见附录 H.4」。
   验收要**双向**断言：`must_contain` 新指向 + `must_not_contain` 旧指向。
3. **迁表顺序铁律：先移表，再把随附说明段接到表后。**
   反序（先插说明段再移表）会使说明段**滞留原位**——表走了，说明段还在原地，
   读起来像「表被抽走了」。

### §2.2 迁表后原章留什么

原章**只留**：
- 结论性判断（一两句，说清这批资料能不能用、效力如何）；
- 指向新位置的**指针**（"明细见附录 H.4"）；
- ⚠️ 警示条目（★ 级、风险项、待查项）——**这些不能搬走**，它们是结论不是明细。

原章**不留**：逐行清单、逐份文件的元数据表、重复的编号索引。

### §2.3 新附录的体例

新附录章标题用**与现有附录一致**的形态。真实文档的附录标题是
`5  附录 H  源文件清单与效力总览`——**字母不在行首**，
所以按 `number='H'` 匹配必然落空，定位时要用「标题文字含关键字」。

新附录内部建议结构：
```
附录 X  <名称>
  X.1  用途与口径          ← 一段话说明这批资料是什么、怎么读
  X.2  <分组>              ← 表 + 表后注记
  X.3  <分组>
  X.4  <分组> + 批号/委托方汇总表
```

**顺序**：新附录插在**现有附录之后、主章之前**（不要插到文档末尾）。

---

## §3 定位与匹配的 12 条硬约束

这一节每一条都对应一个真实踩过的坑。**改既有文档时逐条对照。**

### 3.1 身份比较只能用 `_p` / `._tbl`

```python
# ✗ 错：doc.paragraphs 每次访问返回新的 Paragraph 对象
if p in doc.paragraphs:      # 恒为 False

# ✓ 对：比元素
if p._p is q._p: ...
```

`doc.paragraphs` 每次访问都重新构造 Paragraph 列表，对象 `is` 比较必然失败。
用 `p._p is q._p` 才反映「是不是同一个段落元素」。

> 补充坑：用 `id(p._p)` 做去重 key 也不安全——Paragraph 是临时对象，
> 被回收后 `id` 会被复用，导致「明明存在的段被误判为重复而跳过」。
> **用元素对象本身做 key**（同时保持引用存活）。

### 3.2 一律用 `style.name`，禁用 styleId

真实文档的 Heading 2 的 styleId 是 `'000006'` 或 `'21'`，**不是** `'Heading2'`。
不同文档、不同 Word 版本都不一样。

```python
p.style.name == 'Heading 2'     # ✓
p.style.style_id == 'Heading2'  # ✗
```

### 3.3 目录条目与正文标题同名 → 按 style.name 消歧

目录里的「1  产品基本信息」和正文的章标题「1  产品基本信息」**文本完全相同**，
只是样式不同：目录 = `List Bullet`，正文 = `Heading 1`。

```python
# 找正文标题
find_paragraphs(doc, prefix='1  产品基本信息', exclude_bullet=True)   # 1 段
# 连目录一起找（做零回归核对时）
find_paragraphs(doc, prefix='1  产品基本信息', exclude_bullet=False)  # 2 段
```

**默认必须 `exclude_bullet=True`**，否则会误中目录那一份。
但要改正文项目符号段时，必须显式设 `False`——`find_by_prefix` 类的
默认跳过 bullet 的写法会**连正文 bullet 一起漏改**。

### 3.4 同一章内子标题样式不统一 → 不限定 Heading 层级

实测一章之内：`6.5 / 6.6 / 6.7` 是 `Heading 2`，而 `6.1 / 6.2 / 6.4 / 6.9 / 6.10` 是 `Normal`。

```python
# ✗ 错：只查 Heading 2 → 漏掉 Normal 样式的那一半
find_heading(doc, '6.9', style='Heading 2')      # None

# ✓ 对：不限层级，用编号前缀定位
find_heading(doc, '6.9')                          # 命中（style = Normal）
```

### 3.5 编号 token 与标题文字常分属不同 run

`"6.1"` | `" 效力口径"` 是两个 run；单 run 正则匹配不到。
而且**空格数不统一**（`"6.2 "` 单空格 vs `"6.5  "` 双空格）。

```python
# ✗ 错
re.match(r'^6\.2\s+', '6.2  数据字段权威来源')   # 勉强，但 '6.2 ' 单空格版本会漏

# ✓ 对：只匹配行首编号 token，容忍任意空格数
re.match(r'^\s*6\.2(?![0-9.])', text)
```

改写时同理——只重写**行首编号 token**，其余 run 与格式原样保留。

### 3.6 交叉引用也可能跨 run

`"份见"` | `" 0.6"` 分属两个 run。
解法：**整段拼接 → 在拼接结果上替换 → 回写首 run → 清空其余 run**。

```python
full = ''.join(t.text or '' for t in p._p.iter(qn('w:t')))
# 在 full 上做替换，再按字符区间切回各 run
```

### 3.7 表格内段落不在 `doc.paragraphs` 里

实测正文段 226、**表内段 1612**。全文扫描必须「段落 + 所有表格单元格」一起扫，
否则会漏掉 87% 的内容。

```python
for p in all_paragraphs(doc, include_cells=True): ...
```

### 3.8 定位表一律用表头关键字，禁硬编码索引

搬表会让 `doc.tables` 索引**整体前移**——搬一次后，下标 65~70 的内容全变了。

```python
# ✗ 错
tbl = doc.tables[58]

# ✓ 对
tbl = find_table(doc, ['版本', '生效日期', '修订内容', '来源'])
```

### 3.9 定位失败必须抛错，不能静默返回 None

静默 None 会让「没找到」变成「什么都没做」，看起来成功了。
`resolve_ref()` 找不到就抛 `LocateError`。

### 3.10 位置锚点解析失败不能静默回退

`insert_before` 找不到锚点时**必须报错**。静默回退到「文末」会把章节插到
错误位置，而且**看起来是成功的**。

> 真实案例：`insert_before={"number": "6"}` 若解析到**目录里**的「6」条目
> （目录与正文同名，见 3.3），新附录会被插到目录后面——
> 症状是「附录出现在文末而不在正文中间」。
> **位置锚点一律 `exclude_bullet=True`。**

### 3.11 点名 bullet 样式时，「排除 bullet」过滤必须自动失效

`find_paragraphs()` 默认 `exclude_bullet=True`（用于避开目录条目）。
若实现里**先做 bullet 排除、再做 style 匹配**，那么
`style='List Bullet'` 会**恒返回空**——点名要 bullet，却永远查不到。

正确顺序：**先判 style，再判 bullet 排除**；并且当 `style` 本身就是 bullet 样式时，
排除过滤直接失效。

```python
# ✗ 错：style='List Bullet' 永远空
if exclude_bullet and is_bullet(p): continue
if style is not None and style_name(p) != style: continue

# ✓ 对
if style is not None and style_name(p) != style: continue
if exclude_bullet and style not in _BULLET_STYLES and is_bullet(p): continue
```

> 真实案例：增量说明的要点模板、目录条目的插入/删除锚点，
> 全部依赖「按 List Bullet 样式定位」，一旦恒空就表现为
> 「模板取不到 → clone 返回 None → 后续 `np_._p` 抛
> `AttributeError: 'NoneType' object has no attribute '_p'`」——
> 症状离病根很远，排查成本高。

### 3.12 python-docx 的 `_Row` 构造：parent 必须是带 `.table` 的对象

`docx.table._Row(tr, parent)` 的 `parent` 会被 `Row.table` 反查：

```python
@property
def table(self):
    return self._parent.table
```

所以 `parent` 必须是 `_Rows` 或 `Table`（它们才有 `.table`），
**不能直接传 `CT_Tbl`**。

```python
# ✗ 错：AttributeError: 'CT_Tbl' object has no attribute 'table'
row = _Row(tr, tbl._tbl)

# ✓ 对
row = _Row(tr, tbl.rows)
```

> 典型场景：克隆表格末行后逐格写值（版本沿革表追加行）。

---

## §4 改写与写入的 8 条硬约束

### 4.1 `xml:space` 必须用 Clark 记法

```python
# ✗ 错：ValueError: Invalid attribute name
t.set('xml:space', 'preserve')

# ✓ 对
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
t.set(XMLSPACE, 'preserve')
```

任何写入 `w:t` 文本的地方都要带——否则首尾空格会被吃掉。

### 4.2 禁用 `cell.text = `

`cell.text = "..."` 会清空全部 run，**rPr 全丢**（字体/字号/加粗都没了）。
修订期改单元格内容要用 `set_cell_text_preserve()`（保留首 run 的 rPr）。

### 4.3 整段回写只保留首 run 的 rPr

`set_paragraph_text()` 把新文本写进首 run、其余 run 清空——**段内局部加粗会丢**。
这是**已知取舍**：压缩正文前先确认该段没有局部强调格式。
需要保留局部格式时，改用 `replace_in_paragraph()` 做局部替换而非整段回写。

### 4.4 克隆保格式，禁裸建

新增标题段 / 概述段 / bullet 条目时，**克隆**上一版同类型段落；
版本沿革表新增行时**克隆末行**。

```python
# ✗ 错：格式会漂
doc.add_paragraph('新标题')

# ✓ 对
new_p = clone_paragraph_after(ref_el, template_el, '新标题')
```

### 4.5 「插到之前」与「插到之后」必须分清

```python
clone_paragraph_after(ref, tpl, text)    # → ref.addnext()  插到 ref 之后
clone_paragraph_before(ref, tpl, text)   # → ref.addprevious()  插到 ref 之前
```

用 `addnext` 实现「插到之前」是常见错误——新章节会跑到锚点章节的**后面**。

### 4.6 重编号必须按源编号降序执行

```python
# ✗ 错：6.11→6.10 完成后，紧随其后的 6.10→6.9 会把刚改好的 6.10 又改一次
for old in mapping:            # 顺序执行
    renumber(old, mapping[old])

# ✓ 对：按源编号从大到小
for old in sorted(mapping, key=number_sort_key, reverse=True):
    renumber(old, mapping[old])
```

**另一条约束**：目标号不能与文档中**已存在**的号相撞。
把 `6.10→6.9` 而文档里本来就有一个 6.9，结果会出现**三个 6.9**。
重编号映射要设计成**连续区间**（`6.9→6.10、6.10→6.11、6.11→6.12`），
而不是交换。runner 会做**冲突预检**，冲突直接报错中止。

> **降序仍然不够**：只按源编号降序执行，实测仍会出现 7 个标题**全变成同一个号**
> （因为「占位符 → 目标号」的第二阶段会把新号再次当作源号命中）。
> 正确做法是**两阶段**：阶段一把所有源号改写成不可能冲突的占位符（`§§TMPn§§`），
> 阶段二再统一把占位符替换为目标号。见 `docx_revise_toolkit.renumber_headings()`。

### 4.7 整段回写**不得命中标题段**

「把这一段压缩成一句话」如果 `target` 恰好命中的是**小节标题**，
整段回写会把标题替换成正文 → **该小节标题消失、编号断档、目录失配**，
而文档仍能正常打开，极难察觉。

正确做法：
- `compress` 命中标题段时**直接报错**（要求改用「压缩小节正文」的写法）；
- 提供专门的 kind：`section_body`——定位小节标题，取其下**第一个非 bullet 正文段**替换。

```python
# ✗ 错：6.1 效力口径 → 整段被正文吞掉
p = resolve_ref(doc, {'prefix': '6.1  效力口径'})
set_paragraph_text(p, '……')          # 标题没了

# ✓ 对：压缩该小节下的正文首段
h = resolve_ref(doc, {'number': '6.1', 'title_kw': '效力口径'})
body = first_non_bullet_paragraph_under(h)
set_paragraph_text(body, '……')
```

### 4.8 「配置了但一条没命中」必须报错，不能静默 0

凡是 config 里**点名了要改的串**（交叉引用改指、章节标题、目录锚点），
若替换命中数为 0，必须报错——0 命中通常意味着**串写错了**
（含全角括号与否、子节号是否存在、是否跨 run 拆分），而不是「不需要改」。

```python
n = replace(old, new)
if n == 0 and not ur.get('allow_zero'):
    raise ConfigError('引用改指 0 处命中：%r → %r，请核对真实写法' % (old, new))
```

> 真实案例：迁表后要求把「（见附录 H.2）」改指到新附录，但文档里真实写法是
> 「见附录 H.3」——**没有 H.2 这个引用**，静默 0 命中会让旧位置的引用**留在原地**，
> 正是「迁表三铁律」第三条要防的事。

---

## §5 尾部结构完整性契约（升版的出口契约）

合并 / 升版脚本最容易在**文末**出事，且症状隐蔽（能打开、能读、只是内容错位）。
以下三条**必须自动断言**，不能靠眼看。

### 5.1 结尾标记必须独立成段且全文唯一

```
✗ 「……开检前须现场查证（E-01 仍为 ★高）。  —— 本手册完 ——」   ← 粘在 bullet 末尾
✓ 「—— 本手册完 ——」                                          ← 独立段落，居中
```

断言：
- 全文出现次数 == 期望值（通常 1）；
- 存在至少一处，该段文本 `.strip()` 后**等于**标记本身。

### 5.2 增量说明节不得错挂别节条目

事故形态：增量说明标题下**正文为空**，而本该属于**上一小节**的条目被挂了上去。

判定口径（**不要一刀切要求概述段**）：

| 增量说明类型 | 标题下应是什么 | 判定 |
|---|---|---|
| bullet 型 | 直接跟若干条要点 | 合法，不要求概述段 |
| 概述型 | 先一段 Normal 概述，再跟要点 | 必须有概述段 |

因此：
- **通用断言**：增量说明节的首条 bullet 不得与上一小节的末条**内容重复**（重复即错挂）；
- **精确断言**：在 config 里用 `strict_body_heading_numbers: ["6.11"]`
  指定「本应概述型」的节号，断言其下第一个非空段**不是** bullet。

> 只写通用断言会误报：bullet 型增量说明（标题下直接跟要点）本身就没有概述段，
> 一刀切要求概述段会把合法文档判成 FAIL。

### 5.3 条目不得跨小节错挂

判据：带编号的 bullet，其编号**不得大于**其上方最近的小节标题编号。

### 5.4 增量说明的编号不得与现存标题冲突

升版流程里最容易自相矛盾的一步：**先把 6.11 重编号成 6.10，
再新增一节「6.10 v8 增量说明」** —— 结果文档里有两个 6.10。

正确顺序与口径：
1. 重编号后，先看**当前最大小节号**；
2. 新增增量说明的编号 = 当前最大号 + 1；
3. runner 在插入增量说明前做**冲突预检**：
   解析 `heading_text` 的编号，若与现存标题编号相同 → 直接报错中止。

> 症状隐蔽：文档能打开、能读，只是目录里两个 6.10、
> 交叉引用指向谁都不确定。

---

## §6 用户指令 → 动作映射

| 用户说 | 含义 | 该做什么 | 该问什么 |
|---|---|---|---|
| 「回退」 | 回到上一版成品 | 覆盖式**拷回**上一版（**不删除**），把作废版本隔离到 `_回退_xxx\` 目录，记录 sha | 若有多个候选版本，问清回退到哪一版 |
| 「把 A 和 B 合并，放到 B 位置」 | 章节合并 | 合并 + 重编号 + 同步目录/交叉引用 | **必须先问**：A 的内容是「并入 B」还是「保留在原位」？这决定编号怎么排 |
| 「这一章内容乱 / 冗长，整理精简」 | 章节重构 | **先量化**（`chapter_volume.py`）→ 按决策树选搬表 / 压缩 / 删重复 | 若搬表，需确认新附录叫什么、放哪 |
| 「打开我看」 | 预览 | 用系统默认程序打开 | 注意**工作区文件名可能与 Z 盘规范名不同**（Z 盘带 `8.` 前缀，工作区不带） |
| 「删掉 X」 | 删元素 | 先确认 X 的唯一信息是否已并入别处 | 信息在别处没有 → 先并后删 |
| 「升版 / 出 v8」 | 正式升版 | 新文件名 + 就地更新当前版本字段 + 新增沿革行 + 新增增量说明 | 升哪个版本号、是否新增增量说明 |


---

## §7 `refactor_chapter.py --config` schema

### §7.1 固定阶段流水线

```
merge_sections → delete_sections → renumber → move_tables
→ compress → tail_fix → release
```

**为什么是固定阶段而不是自由 steps 数组**：章节重构有**内在因果序**——
先合并/删除才有稳定的编号空间；先重编号，迁表时才知道旧号是什么；
迁表会改变文档序，压缩必须在其后；尾部修复依赖最终结构；升版必须最后。
自由数组极易写出「重编号先于合并」这类顺序错。

**每个阶段结束后都跑一次结构漂移检测**：表内容指纹的多重集合必须不变。
但有两类**合法**的内容变化要放行，否则会被误报成「表被改坏」：

| 阶段 | 允许变动的表 | 依据 |
|---|---|---|
| `release` | `changelog_table` 命中的版本沿革表 | 配置里点了名；正向断言「只追加一行、旧行逐格保留」 |
| `compress` | `kind='table_cell'` 的目标表 | 配置里点了名 |

其余任何表的内容变化 → 报 `StructureDrift` 并**中止、不写文件**。

### §7.2 config 顶层键

```jsonc
{
  "$schema_version": "1.0",      // 必填；不认识直接报错（防旧配置静默跑偏）
  "meta":  { "product": "", "from_version": "", "to_version": "", "date": "" },
  "io":    { "input": "", "output": "", "baseline": "", "dry_run": false,
             "allow_overwrite_input": false },

  "merge_sections": [ ... ],     // 见 §7.3
  "delete_sections":[ ... ],     // 见 §7.4
  "renumber":        { ... },    // 见 §7.5
  "move_tables":     [ ... ],    // 见 §7.6
  "compress":        [ ... ],    // 见 §7.7
  "tail_fix":        { ... },    // 见 §7.8
  "release":         { ... }     // 见 §7.9
}
```

### §7.3 `merge_sections[]`

```jsonc
{
  "id": "m1",
  "source_heading": { "number": "0", "title_kw": "编制说明与溯源声明",
                      "style": "Heading 1", "occurrence": 0 },
  "target_heading": { "number": "6", "style": "Heading 1", "occurrence": 0 },
  "keep_position": "target",              // 'target' = 放到目标位置；'source' = 留在原位
  "rename_source_to": "编制说明、溯源声明与参考文献",
  "retitle_target": null,
  "subsection_numbering": {
    "map": { "0.1": "6.9.1", "0.2": "6.9.2" },   // 源小节号 → 目标号
    "apply_to_toc": true,                        // 目录条目同步
    "apply_to_cross_refs": true                  // 正文交叉引用同步
  },
  "delete_source_toc_entries": true,
  "confirm_required": true,                     // A 的处置方式必须先问用户
  "on_conflict": "abort"                        // 定位不到即中止
}
```

### §7.4 `delete_sections[]`

```jsonc
{ "heading": { "number": "6.3", "style": "Heading 2" },
  "title_kw": null,
  "with_trailing_bullets": 0,     // 连带删除其后 N 个 bullet
  "reason": "内容已迁往附录 H" }  // 审计用（进报告，不进文档）
```

### §7.5 `renumber`

```jsonc
{ "map": { "6.9": "6.10", "6.10": "6.11" },   // 旧号 → 新号
  "scope": { "chapter": "6" },                 // 限定章号；null = 全文
  "include_toc": true,
  "include_body": true,
  "update_cross_references": true,
  "cross_reference_patterns": [ "(?<![\\d.])(\\d+\\.\\d+)(?![\\d])" ],
  "numeric_guard": ["0.5μg", "0.81", "5.0%"], // 命中这些字面量的段落整段跳过
  "dry_run": false }
```

### §7.6 `move_tables[]`

```jsonc
{
  "id": "mv-h2",
  "table": { "header_keywords": ["编号", "源文件", "读取状态"],  // 按表头定位
             "match_mode": "all",          // 'all' | 'any'
             "row_keyword": null,
             "occurrence": 0 },
  "target_appendix": {
    "heading": { "number": "H",
                 "text": "5  附录 H  源文件清单与效力总览",
                 "style": "Heading 1",
                 "clone_template": { "number": "G", "style": "Heading 1" } },
    "insert_before": { "number": "6", "style": "Heading 1" },   // 插到该章之前
    "sub_headings": [ { "number": "H.1", "text": "H.1  用途与口径" },
                      { "number": "H.2", "text": "H.2  v1 基础源文件" } ],
    "place_after_subheading": "H.2"      // 表插到哪个子标题之后
  },
  "carry_notes": [ { "prefix": "受控文件签署链", "after": "table" } ],
  "update_references": [ { "from": "（见 6.7）", "to": "（见附录 H.4）",
                           "scope": "paragraph_and_cells",
                           "occurrence": -1 } ],   // -1 = 全部
  "delete_source_leftovers": true,
  "note": "表内容一字不删，仅挪位"
}
```

> `update_references` **命中 0 处会直接报错**（见 4.8）。
> `from` 必须与文档里的**真实写法**逐字一致：文档中常写成
> 「见附录 H.3」而非「（见附录 H.3）」，且并不存在 `H.2` 这种子节号——
> 写错会被守卫当场拦下，不会静默留下指向旧位置的引用。
> 确实无需改指时显式加 `"allow_zero": true`。

### §7.7 `compress[]`

```jsonc
{ "kind": "paragraph",      "target": { "prefix": "本手册以 STP-SP-XXX" },
  "new_text": "主依据为 …" },
{ "kind": "section_body",   "target": { "heading": { "number": "6.1",
                                                     "title_kw": "效力口径" } },
  "new_text": "脱敏口径：…" },          // 压缩该小节下的第一段正文
{ "kind": "bullets_merge",  "target": { "heading": { "number": "6.5" } },
  "new_items": ["…", "…"], "keep_leading_bullets": 3, "delete_remaining": true },
{ "kind": "table_cell",     "table": { "header_keywords": ["新增文件（源目录原名）"] },
  "row_keyword": "<行首列值>", "column": 3, "new_text": "§7.6 对照品溶液 …" }
```

> `kind: "paragraph"` 命中**标题段**时**直接报错**（见 4.7）——
> 整段回写会吞掉小节标题。压缩小节正文请用 `kind: "section_body"`。
>
> `kind: "table_drop_column"` **故意不实现**——删列涉及
> `tblW` / `gridCol` / `tcW` 三同步，走 `patch_existing_manual.py` 的列宽内核更安全。

### §7.8 `tail_fix`

```jsonc
{ "end_marker": "—— 本手册完 ——",
  "require_standalone_paragraph": true,
  "end_marker_occurrences": 1,
  "increment_heading_regex": "增量说明",
  "require_body_paragraph": true,
  "min_body_paragraphs": 1,
  "detach_from_bullet": true,
  "return_stray_to": { "heading": { "number": "6.10" },
                        "position": "last_bullet",
                        "prefix": "对照品证书、色谱柱柱效记录" },
  "rebuild_increment_body": { "heading_number": "6.11",
                              "heading_text": "6.11  v5 / v6 / v7 增量说明",
                              "body_paragraph": "v5（2026-11-13）…",
                              "bullets": ["…", "…"] } }
```

### §7.9 `release`

```jsonc
{ "version": "v7", "date": "2026-11-13",
  "cover": { "match": "v6 · 2026-11-13", "replace": "v7 · 2026-11-13",
             "in_paragraphs": 1 },
  "changelog_table": { "header_keywords": ["版本", "生效日期", "修订内容", "来源"],
                       "match_mode": "all",
                       "row": { "版本": "v7", "生效日期": "2026-11-13",
                                "修订内容": "…", "来源": "Z 盘 v6 为输入，另存 v7" } },
  "increment_section": { "heading_text": "6.11  v5 / v6 / v7 增量说明",
                         "body_paragraph": "…", "bullets": ["…"],
                         "clone_template": { "heading_number": "6.8" } },
  "toc_entries": [ { "op": "add", "after": { "number": "G" },
                     "text": "5  附录 H  源文件清单与效力总览" },
                    { "op": "delete", "prefix": "6  参考文献与待核对项" } ] }
```

---

## §8 `verify_revision.py` 新增字段

全部新字段用 `cfg.get(...)` **整块取**，块缺失 → 该层打印一行 SKIP、
**绝不产生 FAIL**。老 config 是别的品种的既有资产，不能因为脚本升级就变红。

### §8.1 `tail_integrity`

```jsonc
{ "end_marker": "—— 本手册完 ——",
  "end_marker_standalone": true,
  "end_marker_occurrences": 1,
  "increment_heading_regex": "增量说明",
  "require_body_paragraph": true,
  "forbid_cross_section_attach": true,
  "strict_body_heading_numbers": ["6.11"] }
```

### §8.2 `section_volume`

```jsonc
{ "chapters": ["6"], "min_body_chars": 1,
  "assert_no_empty_sections": true, "report_top_n": 5 }
```

### §8.3 `cross_reference_checks`

```jsonc
{ "must_contain": ["见附录 H.4"], "must_not_contain": ["见 6.7"],
  "scan_cells": true }
```

### §8.4 `moved_tables[]`

```jsonc
[ { "header_keywords": ["编号", "源文件"], "match_mode": "all",
    "expect_dims": [6, 5] } ]
```

> **必须按表头关键字定位**，绝不能用索引——迁表后索引会漂移。

### §8.5 `allowed_table_changes[]`（baseline 比对的表级白名单）

baseline 逐格比对原本是**按 `doc.tables` 索引**做的。中间一旦发生迁表/删节，
索引会整体漂移 → 逐格比对必然假阳性（表现为「t52 行数不同」这类与真实改动无关的差异）。

因此 verify 脚本现在做两级判定：

| 情况 | 判定方式 |
|---|---|
| 两版**表头序列相同**（表序未变） | 沿用**索引式**逐格比对（老行为，逐字不变） |
| 表头序列**不同**（表序已变） | 自动降级为**表内容指纹的多重集合**比对 |

`allowed_table_changes` 是**表级**白名单——声明「这张表本轮允许改内容」：

```jsonc
[ { "header_keywords": ["版本", "生效日期", "修订内容", "来源"],
    "match_mode": "all",
    "reason": "release 阶段合法追加 v8 行（旧行必须逐格保留）" } ]
```

两张版本里匹配到该表头的表，其指纹从「丢失/新增」集合中扣除；
同时对白名单表做**正向断言**：行数只许**末尾追加**，旧行必须逐格原样保留。

> 与 `allowed_cells`（精确键 `t{i}.r{j}.c{k}`）和
> `allowed_cell_patterns`（描述式正则）并存，三者是三个粒度：
> 单元格精确键 / 单元格描述式 / 整表级。**优先级**：表序已变时只走表级。

---

## §9 算例（**非通用常量，禁止写入脚本默认值**）

以下数字来自一次真实改版，仅用于说明各口径的**量级关系**：

| 口径 | 改版前 | 改版后 |
|---|---|---|
| 目标章正文+表内合计 | 8,371 字 | 3,833 字 |
| 其中表内文字 | 3,948 字（47.2%） | 460 字（12.0%） |
| 该章表数 | 6 | 1 |
| 明细表去向 | — | 迁入新增附录（5 张表，**内容一字未删**） |
| 文档总表数 | 71 | 71（**不变**） |
| 纯文字精简能达到的减量 | ≈ −8% | — |
| 搬表方案达到的减量 | — | ≈ −54% |

> ⚠️ 这些数字是**算例**，不是通用常量。任何品种的默认值都应从
> `chapter_volume.py` 的实测报告得出，不得把上表数字写进脚本。

---

## 附录 A 引用索引

| 本文件 | 现有资料 |
|---|---|
| §1 体量决策树 | `scripts/chapter_volume.py`；`revision_playbook.md` §1 |
| §2 搬表通则 | `scripts/refactor_chapter.py`（阶段 move_tables）；`revision_playbook.md` §8 |
| §3 定位约束 | `scripts/docx_revise_toolkit.py`；`python_docx_pitfalls.md` §3/§6 |
| §4 写入约束 | `scripts/docx_revise_toolkit.py`；`python_docx_pitfalls.md` §5 |
| §5 尾部契约 | `scripts/verify_revision.py`（尾部结构级）；`revision_playbook.md` §8 |
| §6 指令映射 | `revision_playbook.md` §1/§9 |
| §6.5 编号冲突 | `scripts/refactor_chapter.py`（`_check_increment_number`） |
| §7 config schema | `scripts/refactor_chapter.py` |
| §8 verify 字段 | `scripts/verify_revision.py`；`revision_playbook.md` 附录 C |
| 收尾提醒话术 | `revision_playbook.md` §9 / 附录 B（**不在本文件重复**） |
