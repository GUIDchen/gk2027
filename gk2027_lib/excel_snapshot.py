# -*- coding: utf-8 -*-
"""把某一天的报名人数快照追加到工作簿：与上一天并列，计算较昨日增加。"""
from __future__ import annotations

import datetime
import json
import re
from pathlib import Path

import xlrd
import xlwt
from xlutils.copy import copy

from .util import backup, ensure_dir, fmt_int, log

STYLE_HEADER = xlwt.easyxf('font: name 宋体, height 220; align: horiz center, vert center, wrap on;')
STYLE_DATA = xlwt.easyxf('font: name Arial, height 200; align: horiz general, vert bottom;')
STYLE_GREEN = xlwt.easyxf('font: name Arial, height 200; align: horiz general, vert bottom; pattern: pattern solid, fore_colour light_green;')
STYLE_YELLOW = xlwt.easyxf('font: name Arial, height 200; align: horiz general, vert bottom; pattern: pattern solid, fore_colour light_yellow;')
STYLE_SEC = xlwt.easyxf('font: name 宋体, height 220, bold on; align: horiz left, vert center;')
STYLE_LBL = xlwt.easyxf('font: name 宋体, height 200, bold on; align: horiz left, vert top, wrap on;')
STYLE_TXT = xlwt.easyxf('font: name 宋体, height 200; align: horiz left, vert top, wrap on;')

TARGET_SHEETS = ['符合条件岗位', '陕西省岗位', '甘肃省岗位', '西安市岗位']


def _row_style(rb, sheet, r):
    xf = rb.xf_list[sheet.cell_xf_index(r, 0)]
    if xf.background.fill_pattern == 1 and xf.background.pattern_colour_index == 42:
        return STYLE_GREEN
    if xf.background.fill_pattern == 1 and xf.background.pattern_colour_index == 43:
        return STYLE_YELLOW
    return STYLE_DATA


def _fmt_up(ms):
    if not ms:
        return ''
    try:
        return datetime.datetime.fromtimestamp(int(ms) / 1000).strftime('%Y-%m-%d %H:%M')
    except Exception:
        return ''


def add_snapshot(workbook: Path, counts_json: Path, label: str, force: bool = False) -> dict:
    counts = json.load(open(counts_json, encoding='utf-8'))
    cmap = {}
    for x in counts:
        if x.get('book') is None:
            continue
        cmap[(str(x.get('dept', '')).strip(), str(x.get('code', '')).strip())] = x
    if not cmap:
        raise ValueError('counts 文件里没有有效记录（book 全为空）。')

    rb = xlrd.open_workbook(str(workbook), formatting_info=True)
    wb = copy(rb)
    names = rb.sheet_names()
    base = rb.sheet_by_name('符合条件岗位')
    header = [base.cell_value(1, c) for c in range(base.ncols)]

    # 解析已有快照标签
    labels = []
    label_cols = {}
    for c, h in enumerate(header):
        m = re.match(r'报名人数\((.+?)\)$', str(h).strip())
        if m:
            lab = m.group(1)
            labels.append(lab)
            label_cols[lab] = c
    if label in label_cols and not force:
        raise ValueError(f'工作簿里已经有“{label}”快照；加 --force 覆盖（会新增一套同名列，不建议）。')
    prev = labels[-1] if labels else None
    last_col = max(c for c, h in enumerate(header) if str(h).strip()) if any(str(h).strip() for h in header) else 26
    start = last_col + 1
    new_headers = [f'报名人数({label})', f'报录比({label})']
    if prev:
        new_headers.append(f'较昨日增加({label})')
    new_headers.append(f'更新时间({label})')

    # 写表头 + 数据
    for name in TARGET_SHEETS:
        if name not in names:
            continue
        s = rb.sheet_by_name(name)
        ws = wb.get_sheet(names.index(name))
        for i, h in enumerate(new_headers):
            ws.write(1, start + i, h, STYLE_HEADER)
        widths = [6000, 5000]
        if prev:
            widths.append(5000)
        widths.append(7000)
        for i, w in enumerate(widths):
            ws.col(start + i).width = w
        if s.nrows <= 2:
            continue
        for r in range(2, s.nrows):
            code = str(s.cell_value(r, 8)).strip()
            dept = str(s.cell_value(r, 0)).strip()
            if not code:
                continue
            st = _row_style(rb, s, r)
            x = cmap.get((dept, code))
            if not x or x.get('book') is None:
                for i in range(len(new_headers)):
                    ws.write(r, start + i, '', st)
                continue
            book = int(float(x['book']))
            try:
                kaolu = int(float(s.cell_value(r, 11))) or 1
            except Exception:
                kaolu = 1
            ratio = book / kaolu
            ws.write(r, start, book, st)
            ws.write(r, start + 1, '%.1f:1' % ratio, st)
            idx = start + 2
            if prev:
                pcol = label_cols.get(prev)
                old = None
                if pcol is not None:
                    v = s.cell_value(r, pcol)
                    try:
                        old = int(float(v))
                    except Exception:
                        old = None
                ws.write(r, idx, (book - old) if old is not None else '', st)
                idx += 1
            ws.write(r, idx, _fmt_up(x.get('up')), st)

    # 写筛选说明
    si = names.index('筛选说明')
    wse = wb.get_sheet(si)
    r = rb.sheet_by_name('筛选说明').nrows
    wse.write(r, 0, '十、%s报名人数更新' % label, STYLE_SEC)
    r += 1

    def put(k, v):
        nonlocal r
        wse.write(r, 0, k, STYLE_LBL)
        wse.write(r, 1, v, STYLE_TXT)
        n = max(len(str(k)) * 2, len(str(v)))
        h = 300 if n <= 110 else (450 if n <= 220 else 600)
        wse.row(r).height = h
        wse.row(r).height_mismatch = True
        r += 1

    books = [int(float(x['book'])) for x in counts if x.get('book') is not None]
    put('覆盖范围', f'共 {len(books)} 个职位抓到 {label} 报名人数；总报名 {fmt_int(sum(books))} 人，平均 {sum(books)/len(books):.1f} 人。')
    top = sorted([x for x in counts if x.get('book') is not None], key=lambda x: -int(float(x['book'])))[:3]
    if top:
        put('最高职位', '；'.join(f"{x.get('bureau','')} {fmt_int(x['book'])} 人" for x in top))
    if prev:
        put('对比', f'新增“报名人数({label})/报录比({label})/较昨日增加/更新时间({label})”，上一快照为“{prev}”。')
    put('提示', '报名人数会随报名进程变化，以系统实时数据为准；黄色岗位仍需电话确认专业、学位、英语、体能等条件。')

    # 清理 xlutils 可能生成的空 Sheet2
    wslist = wb._Workbook__worksheets
    wslist[:] = [s for s in wslist if s.name != 'Sheet2']
    for i, s in enumerate(wslist):
        try:
            s.index = i
        except Exception:
            pass

    ensure_dir(workbook)
    backup(workbook)
    wb.save(str(workbook))
    return {'label': label, 'records': len(books), 'total': sum(books), 'workbook': str(workbook)}
