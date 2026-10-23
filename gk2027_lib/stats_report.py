# -*- coding: utf-8 -*-
"""根据工作簿里的快照列生成独立统计表（支持单日/两日对比）。"""
from __future__ import annotations

import collections
import re
import statistics
from pathlib import Path

import xlrd
import xlwt
from xlwt import easyxf

from .util import backup, ensure_dir, fmt_int, log

B = ('borders: left thin, right thin, top thin, bottom thin, '
     'left_colour gray50, right_colour gray50, top_colour gray50, bottom_colour gray50;')
STYLE_TITLE = easyxf('font: name 黑体, height 340, bold on; align: horiz center, vert center;')
STYLE_NOTE = easyxf('font: name 宋体, height 190; align: horiz left, vert center, wrap on;')
STYLE_HEADER = easyxf('font: name 宋体, height 210, bold on, colour white; pattern: pattern solid, fore_colour dark_red; align: horiz center, vert center, wrap on; ' + B)
STYLE_DATA = easyxf('font: name 宋体, height 190; align: horiz left, vert center; ' + B)
STYLE_DATA_C = easyxf('font: name 宋体, height 190; align: horiz center, vert center; ' + B)
STYLE_NUM = easyxf('font: name 宋体, height 190; align: horiz center, vert center; ' + B, num_format_str='#,##0')
STYLE_NUM_RED = easyxf('font: name 宋体, height 190, bold on, colour dark_red; align: horiz center, vert center; ' + B, num_format_str='#,##0')
STYLE_NUM_ORANGE = easyxf('font: name 宋体, height 190, bold on, colour orange; align: horiz center, vert center; ' + B, num_format_str='#,##0')
STYLE_NUM_GREEN = easyxf('font: name 宋体, height 190, bold on, colour green; align: horiz center, vert center; ' + B, num_format_str='#,##0')
STYLE_DELTA = easyxf('font: name 宋体, height 190, bold on, colour dark_red; align: horiz center, vert center; ' + B, num_format_str='+#,##0;-#,##0;0')
STYLE_TIER_G = easyxf('font: name 宋体, height 190, bold on, colour dark_green; pattern: pattern solid, fore_colour light_green; align: horiz center, vert center; ' + B)
STYLE_TIER_Y = easyxf('font: name 宋体, height 190, bold on, colour dark_yellow; pattern: pattern solid, fore_colour light_yellow; align: horiz center, vert center; ' + B)
STYLE_SEC = easyxf('font: name 宋体, height 250, bold on; align: horiz left, vert center;')
STYLE_LBL = easyxf('font: name 宋体, height 200, bold on; align: horiz left, vert center; ' + B)
STYLE_VAL = easyxf('font: name 宋体, height 200; align: horiz left, vert center; ' + B)
STYLE_TBL_H = easyxf('font: name 宋体, height 200, bold on, colour white; pattern: pattern solid, fore_colour dark_red; align: horiz center, vert center, wrap on; ' + B)
STYLE_TBL = easyxf('font: name 宋体, height 200; align: horiz center, vert center; ' + B)
STYLE_TBL_L = easyxf('font: name 宋体, height 200; align: horiz left, vert center; ' + B)


def _snapshots(header):
    cols = {}
    for c, h in enumerate(header):
        h = str(h).strip()
        m = re.match(r'报名人数\((.+?)\)$', h)
        if m:
            cols.setdefault(m.group(1), {})['count'] = c
        m = re.match(r'报录比\((.+?)\)$', h)
        if m:
            cols.setdefault(m.group(1), {})['ratio'] = c
        m = re.match(r'更新时间\((.+?)\)$', h)
        if m:
            cols.setdefault(m.group(1), {})['update'] = c
    return [{'label': k, **v} for k, v in cols.items() if 'count' in v]


def build_report(workbook: Path, out: Path, labels=None, force: bool = False) -> dict:
    if out.exists() and not force:
        raise FileExistsError(f'输出已存在：{out}\n加 --force 覆盖，或换 --out。')
    rb = xlrd.open_workbook(str(workbook), formatting_info=True)
    s = rb.sheet_by_name('符合条件岗位')
    header = [s.cell_value(1, c) for c in range(s.ncols)]
    snaps = _snapshots(header)
    if labels:
        snaps = [x for x in snaps if x['label'] in labels]
    if not snaps:
        raise ValueError('工作簿里没有报名人数快照列，请先运行 snapshot。')
    snaps = snaps[-2:] if len(snaps) >= 2 else snaps
    one = len(snaps) == 1
    last = snaps[-1]
    prev = snaps[0] if not one else None

    records = []
    for r in range(2, s.nrows):
        code = str(s.cell_value(r, 8)).strip()
        if not code:
            continue
        xf = rb.xf_list[s.cell_xf_index(r, 0)]
        tier = '绿色' if xf.background.pattern_colour_index == 42 else '黄色'
        try:
            kaolu = int(float(s.cell_value(r, 11))) or 1
        except Exception:
            kaolu = 1
        rec = {'dept': str(s.cell_value(r, 0)), 'name': str(s.cell_value(r, 1)),
               'bureau': str(s.cell_value(r, 2)), 'title': str(s.cell_value(r, 4)),
               'code': code, 'kaolu': kaolu, 'loc': str(s.cell_value(r, 20)), 'tier': tier}
        for sn in snaps:
            try:
                rec[sn['label']] = int(float(s.cell_value(r, sn['count'])))
            except Exception:
                rec[sn['label']] = None
        if sn.get('update') is not None:
            rec[sn['label'] + '_up'] = str(s.cell_value(r, sn['update']))
        else:
            rec[sn['label'] + '_up'] = ''
        records.append(rec)
    records = [x for x in records if x[last['label']] is not None]
    records.sort(key=lambda x: (-x[last['label']], -(x[last['label']] / x['kaolu']), x['name']))

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('统计总表')
    if one:
        headers = ['序号', '部门代码', '部门名称', '用人司局', '招考职位', '职位代码', '招考人数', '工作地点',
                   f"报名人数({last['label']})", f"报录比({last['label']})", f"更新时间({last['label']})", '匹配等级']
        ncol = len(headers)
    else:
        headers = ['序号', '部门代码', '部门名称', '用人司局', '招考职位', '职位代码', '招考人数', '工作地点',
                   f"报名人数({prev['label']})", f"报录比({prev['label']})",
                   f"报名人数({last['label']})", f"报录比({last['label']})", '较昨日增加',
                   f"更新时间({last['label']})", '匹配等级']
        ncol = len(headers)
    ws.write_merge(0, 0, 0, ncol - 1, f'2027国考符合条件岗位报名人数统计（{", ".join(x["label"] for x in snaps)}）', STYLE_TITLE)
    ws.row(0).height = 520; ws.row(0).height_mismatch = True
    ws.write_merge(1, 1, 0, ncol - 1, '数据来源：登录系统职位详情接口；报名人数为系统报名总数；按最新快照报名人数从高到低排序。', STYLE_NOTE)
    ws.row(1).height = 340; ws.row(1).height_mismatch = True
    for c, h in enumerate(headers):
        ws.write(2, c, h, STYLE_HEADER)
    ws.row(2).height = 460; ws.row(2).height_mismatch = True
    widths = [6, 10, 24, 30, 30, 14, 8, 20] + ([11, 10, 18] if one else [11, 10, 11, 10, 10, 18]) + [8]
    for c, w in enumerate(widths):
        ws.col(c).width = 256 * w
    for i, x in enumerate(records, 1):
        r = i + 2
        ws.write(r, 0, i, STYLE_DATA_C)
        ws.write(r, 1, x['dept'], STYLE_DATA_C)
        ws.write(r, 2, x['name'], STYLE_DATA)
        ws.write(r, 3, x['bureau'], STYLE_DATA)
        ws.write(r, 4, x['title'], STYLE_DATA)
        ws.write(r, 5, x['code'], STYLE_DATA_C)
        ws.write(r, 6, x['kaolu'], STYLE_NUM)
        ws.write(r, 7, x['loc'], STYLE_DATA)
        if one:
            v = x[last['label']]
            st = STYLE_NUM_RED if v >= 1000 else (STYLE_NUM_ORANGE if v >= 500 else (STYLE_NUM_GREEN if v < 100 else STYLE_NUM))
            ws.write(r, 8, v, st)
            ws.write(r, 9, '%.1f:1' % (v / x['kaolu']), STYLE_DATA_C)
            ws.write(r, 10, x.get(last['label'] + '_up', ''), STYLE_DATA_C)
            ws.write(r, 11, x['tier'], STYLE_TIER_G if x['tier'] == '绿色' else STYLE_TIER_Y)
        else:
            v1, v2 = x[prev['label']], x[last['label']]
            ws.write(r, 8, v1, STYLE_NUM)
            ws.write(r, 9, '%.1f:1' % (v1 / x['kaolu']), STYLE_DATA_C)
            st = STYLE_NUM_RED if v2 >= 1000 else (STYLE_NUM_ORANGE if v2 >= 500 else (STYLE_NUM_GREEN if v2 < 100 else STYLE_NUM))
            ws.write(r, 10, v2, st)
            ws.write(r, 11, '%.1f:1' % (v2 / x['kaolu']), STYLE_DATA_C)
            ws.write(r, 12, v2 - v1, STYLE_DELTA)
            ws.write(r, 13, x.get(last['label'] + '_up', ''), STYLE_DATA_C)
            ws.write(r, 14, x['tier'], STYLE_TIER_G if x['tier'] == '绿色' else STYLE_TIER_Y)
        ws.row(r).height = 340; ws.row(r).height_mismatch = True
    try:
        ws.set_panes_frozen(True); ws.set_horz_split_pos(3); ws.set_vert_split_pos(0)
    except Exception:
        pass

    # 汇总
    ws2 = wb.add_sheet('汇总')
    for c, w in enumerate([20, 14, 14, 14, 38]):
        ws2.col(c).width = 256 * w
    r = 0
    def sec(text):
        nonlocal r
        ws2.write_merge(r, r, 0, 4, text, STYLE_SEC); ws2.row(r).height = 400; ws2.row(r).height_mismatch = True; r += 1
    def kv(k, v):
        nonlocal r
        ws2.write(r, 0, k, STYLE_LBL); ws2.write(r, 1, v, STYLE_VAL); ws2.row(r).height = 330; ws2.row(r).height_mismatch = True; r += 1
    def table(headers, rows, left=()):
        nonlocal r
        for c, h in enumerate(headers):
            ws2.write(r, c, h, STYLE_TBL_H)
        ws2.row(r).height = 340; ws2.row(r).height_mismatch = True; r += 1
        for row in rows:
            for c, v in enumerate(row):
                ws2.write(r, c, v, STYLE_TBL_L if c in left else STYLE_TBL)
            ws2.row(r).height = 320; ws2.row(r).height_mismatch = True; r += 1
        r += 1
    ws2.write_merge(0, 0, 0, 4, f'2027国考符合条件岗位报名人数统计 · 汇总', STYLE_TITLE); ws2.row(0).height = 520; ws2.row(0).height_mismatch = True; r = 1
    ws2.write_merge(1, 1, 0, 4, '数据来源：登录系统职位详情接口；报名人数为报名总数。', STYLE_NOTE); ws2.row(1).height = 320; ws2.row(1).height_mismatch = True; r = 3
    vals_last = [x[last['label']] for x in records]
    sec(f'一、总体（{last["label"]}）')
    kv('岗位数', len(records)); kv('招录人数合计', sum(x['kaolu'] for x in records))
    kv(f'{last["label"]}报名合计', fmt_int(sum(vals_last)))
    kv(f'{last["label"]}平均报名', round(sum(vals_last) / len(vals_last), 1))
    kv(f'{last["label"]}中位数', statistics.median(vals_last))
    kv(f'{last["label"]}最高', '%s（%s）' % (fmt_int(max(vals_last)), records[0]['bureau']))
    if not one:
        vals_prev = [x[prev['label']] for x in records]
        d = sum(vals_last) - sum(vals_prev)
        kv('较上一快照增加', '%s（+%.1f%%）' % (fmt_int(d), 100 * d / sum(vals_prev)))
    sec('二、按匹配等级')
    rows = []
    for t in ['绿色', '黄色']:
        lst = [x for x in records if x['tier'] == t]
        if lst:
            rows.append([t, len(lst), sum(x[last['label']] for x in lst), round(sum(x[last['label']] for x in lst) / len(lst), 1)])
    table(['匹配等级', '岗位数', f'{last["label"]}合计', f'{last["label"]}平均'], rows, left=(0,))
    sec(f'三、{last["label"]}报名人数区间分布')
    bins = [('1-10人', 1, 10), ('11-50人', 11, 50), ('51-100人', 51, 100), ('101-300人', 101, 300),
            ('301-500人', 301, 500), ('501-1000人', 501, 1000), ('1000人以上', 1001, 10 ** 9)]
    rows = []
    for label, lo, hi in bins:
        n = sum(1 for v in vals_last if lo <= v <= hi)
        rows.append([label, n, '%.1f%%' % (100 * n / len(vals_last))])
    table(['区间', '岗位数', '占比'], rows, left=(0,))
    sec(f'四、{last["label"]}报名人数前20')
    rows = []
    for i, x in enumerate(records[:20], 1):
        rows.append([i, x['name'], x['bureau'], x['title'], x['kaolu'], x[last['label']]])
    table(['排名', '部门', '用人司局', '职位', '招考', last['label']], rows, left=(1, 2, 3))
    if not one:
        sec('五、较上一快照增加前15')
        rows = []
        for x in sorted(records, key=lambda y: -(y[last['label']] - y[prev['label']]))[:15]:
            rows.append([x['name'], x['bureau'], x[prev['label']], x[last['label']], '%+d' % (x[last['label']] - x[prev['label']])])
        table(['部门', '用人司局', prev['label'], last['label'], '增加'], rows, left=(0, 1))
    sec('六、按部门统计（岗位数前20）')
    c = collections.Counter(); sm = collections.Counter()
    for x in records:
        c[x['name']] += 1; sm[x['name']] += x[last['label']]
    rows = [[k, v, sm[k], round(sm[k] / v, 1)] for k, v in c.most_common(20)]
    table(['部门', '岗位数', f'{last["label"]}合计', f'{last["label"]}平均'], rows, left=(0,))
    sec('七、按地区统计（岗位数前15）')
    reg = collections.Counter(); rsm = collections.Counter()
    for x in records:
        loc = x['loc']; p = loc[:2]
        if loc.startswith('内蒙古'):
            p = '内蒙古'
        if loc.startswith('黑龙江'):
            p = '黑龙江'
        reg[p] += 1; rsm[p] += x[last['label']]
    rows = [[k, v, rsm[k], round(rsm[k] / v, 1)] for k, v in reg.most_common(15)]
    table(['地区', '岗位数', f'{last["label"]}合计', f'{last["label"]}平均'], rows, left=(0,))

    # 说明
    ws3 = wb.add_sheet('说明')
    ws3.col(0).width = 256 * 18; ws3.col(1).width = 256 * 90
    ws3.write_merge(0, 0, 0, 1, '2027国考符合条件岗位报名人数统计 · 说明', STYLE_TITLE)
    ws3.row(0).height = 520; ws3.row(0).height_mismatch = True
    notes = [
        ('数据来源', '登录系统职位详情接口返回的报名人数（系统字段 num）。'),
        ('统计口径', '报名人数=系统报名总数；报录比=报名人数÷招考人数。'),
        ('排序规则', f'统计总表按{last["label"]}报名人数从高到低排序。'),
        ('动态提示', '报名截止后人数不再变化；此前抓取的数据会继续上涨，以系统实时数据为准。'),
    ]
    for i, (k, v) in enumerate(notes, 1):
        ws3.write(i, 0, k, STYLE_LBL); ws3.write(i, 1, v, STYLE_VAL)
        ws3.row(i).height = 420; ws3.row(i).height_mismatch = True

    ensure_dir(out)
    backup(out)
    wb.save(str(out))
    return {'out': str(out), 'snapshots': [x['label'] for x in snaps], 'records': len(records)}
