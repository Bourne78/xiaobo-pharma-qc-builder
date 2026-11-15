# -*- coding: utf-8 -*-
"""
chapter_volume —— 章节体量报告（只读）

用途
----
用户判定某章「冗长 / 内容乱」时，**先量化、再动手**。本脚本回答三个问题：
  1. 这一章一共多少字？分摊到各小节是多少？
  2. 体量大头在**正文文字**还是**明细表**？
  3. 该走「压缩文字」还是「搬表」？

为什么必须分开统计表内文字
--------------------------
真实案例：一章总 8,371 字，其中 4 张明细表占 5,057 字（60%）。
  * 只看正文文字 → 误判为「文字太多，删一删」→ 纯文字精简最多只减 8~25%
  * 分开看     → 结论是「把明细表整体搬进新附录，原章只留结论与指针」
                   → 一次减 54%，且表内事实一字不丢

本脚本**只读**，绝不写文档。核心计算在 `docx_revise_toolkit.measure_sections`。

用法
----
    python chapter_volume.py --docx X.docx                       # 全文章节体量
    python chapter_volume.py --docx X.docx --chapter 6           # 只看第 6 章
    python chapter_volume.py --docx X.docx --format json --out r.json
    python chapter_volume.py --docx X.docx --top 5               # 只看体量 TOP5
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                                  # noqa: E402
from docx_revise_toolkit import measure_sections           # noqa: E402


def _fmt_row(s):
    return {
        'number': s['number'] or '-',
        'title': (s['title'] or '')[:34],
        'style': s['style'] or '-',
        'chars': s['chars'],
        'table_chars': s['table_chars'],
        'total_chars': s['total_chars'],
        'tables': s['tables'],
        'rows': s['table_rows'],
        'paras': s['paras'],
        'bullets': s['bullets'],
        'first_item': (s['first_item'] or '')[:50],
    }


def render_text(secs, chapter=None, top=None):
    """文本报告。"""
    lines = []
    if chapter is not None:
        lines.append('=== 章节体量：第 %s 章 ===' % chapter)
    else:
        lines.append('=== 全文章节体量 ===')
    header = ('%-7s %-34s %-10s %7s %7s %8s %5s %6s %5s %5s'
              % ('编号', '标题', '样式', '正文', '表内', '合计', '表数', '行数', '段数', 'bullet'))
    lines.append(header)
    lines.append('-' * len(header))

    shown = secs
    truncated = False
    if top and top > 0 and len(secs) > top:
        shown = sorted(secs, key=lambda s: -s['total_chars'])[:top]
        shown = sorted(shown, key=lambda s: (s['level'] or 0, s['number'] or ''))
        truncated = True

    for s in shown:
        lines.append('%-7s %-34s %-10s %7d %7d %8d %5d %6d %5d %5d'
                     % (s['number'] or '-', (s['title'] or '')[:34],
                        s['style'] or '-', s['chars'], s['table_chars'],
                        s['total_chars'], s['tables'], s['table_rows'],
                        s['paras'], s['bullets']))

    lines.append('-' * len(header))
    tot_c = sum(s['chars'] for s in secs)
    tot_t = sum(s['table_chars'] for s in secs)
    lines.append('%-7s %-34s %-10s %7d %7d %8d %5d %6d %5d %5d'
                 % ('合计', '(全部 %d 节)' % len(secs), '', tot_c, tot_t,
                    tot_c + tot_t, sum(s['tables'] for s in secs),
                    sum(s['table_rows'] for s in secs),
                    sum(s['paras'] for s in secs),
                    sum(s['bullets'] for s in secs)))

    # 决策提示
    lines.append('')
    lines.append('--- 精简决策提示 ---')
    if tot_c + tot_t == 0:
        lines.append('  无体量数据。')
    else:
        ratio = tot_t / float(tot_c + tot_t) * 100
        if ratio >= 40:
            lines.append('  表内文字占 %.1f%%（%d / %d 字）——体量大头在明细表。' % (ratio, tot_t, tot_c + tot_t))
            lines.append('  → 优先「搬表而非删表」：把明细表整体迁入新增附录（表内一字不删），')
            lines.append('    原章只留结论与指针；迁后必须改写指向旧位置的交叉引用。')
        else:
            lines.append('  表内文字占 %.1f%%（%d / %d 字）——体量大头在正文文字。' % (ratio, tot_t, tot_c + tot_t))
            lines.append('  → 优先「压缩文字」：合并冗长段落、精简 bullet 组、删除纯重复小节。')
        biggest = max(secs, key=lambda s: s['total_chars']) if secs else None
        if biggest:
            lines.append('  最大单节：%s %s（%d 字，其中表内 %d 字 / %d 表）'
                         % (biggest['number'] or '-', biggest['title'][:26],
                            biggest['total_chars'], biggest['table_chars'],
                            biggest['tables']))
    if truncated:
        lines.append('')
        lines.append('（已按 --top %d 截断显示，合计为全量统计）' % top)
    return '\n'.join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description='章节体量报告（只读）')
    ap.add_argument('--docx', required=True, help='目标 .docx（只读打开）')
    ap.add_argument('--chapter', default=None, help='只统计指定章号，如 6')
    ap.add_argument('--format', choices=['text', 'json'], default='text',
                    help='输出格式（默认 text）')
    ap.add_argument('--top', type=int, default=0,
                    help='只显示体量 TOP N 节（0 = 全部）')
    ap.add_argument('--out', default=None, help='输出文件（JSON 格式时可用）')
    args = ap.parse_args(argv)

    if not os.path.exists(args.docx):
        print('文件不存在: %s' % args.docx)
        return 2

    doc = Document(args.docx)
    secs = measure_sections(doc, chapter=args.chapter)

    if args.format == 'json':
        payload = {
            'docx': os.path.abspath(args.docx),
            'chapter': args.chapter,
            'sections': [_fmt_row(s) for s in secs],
            'totals': {
                'chars': sum(s['chars'] for s in secs),
                'table_chars': sum(s['table_chars'] for s in secs),
                'total_chars': sum(s['total_chars'] for s in secs),
                'tables': sum(s['tables'] for s in secs),
                'sections': len(secs),
            },
        }
        text = json.dumps(payload, ensure_ascii=False, indent=2)
    else:
        text = render_text(secs, chapter=args.chapter, top=args.top)

    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(text)
        print('报告已写入: %s' % args.out)
    else:
        print(text)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
