# -*- coding: utf-8 -*-
"""生成个人筛选工作簿：原4表标色 + 符合条件岗位 + 地区表 + 筛选说明。"""
from __future__ import annotations

import datetime
import shutil
from pathlib import Path

import xlrd
import xlwt
from xlutils.copy import copy

from .matcher import evaluate
from .util import CFG, backup, ensure_dir, log

PROFILE = CFG.get('profile', {})

STYLE_NOTE = xlwt.easyxf('font: name 黑体, height 320; align: horiz center, vert center;')
STYLE_HEADER = xlwt.easyxf('font: name 宋体, height 220; align: horiz center, vert center, wrap on;')
STYLE_DATA = xlwt.easyxf('font: name Arial, height 200; align: horiz general, vert bottom;')
STYLE_GREEN = xlwt.easyxf('font: name Arial, height 200; align: horiz general, vert bottom; pattern: pattern solid, fore_colour light_green;')
STYLE_YELLOW = xlwt.easyxf('font: name Arial, height 200; align: horiz general, vert bottom; pattern: pattern solid, fore_colour light_yellow;')
STYLE_TITLE = xlwt.easyxf('font: name 黑体, height 300, bold on; align: horiz left, vert center;')
STYLE_SEC = xlwt.easyxf('font: name 宋体, height 220, bold on; align: horiz left, vert center;')
STYLE_LBL = xlwt.easyxf('font: name 宋体, height 200, bold on; align: horiz left, vert top, wrap on;')
STYLE_TXT = xlwt.easyxf('font: name 宋体, height 200; align: horiz left, vert top, wrap on;')
STYLE_WARN = xlwt.easyxf('font: name 宋体, height 220, bold on, colour red; align: horiz left, vert center, wrap on;')

LEVEL_ORDER = {'中央': 0, '省（副省）级': 1, '市（地）级': 2, '县（区）级及以下': 3}


def _row_style(sheet, r):
    xf = sheet.cell_xf_index(r, 0)
    return STYLE_DATA


def _write_sheet(wb, name, header, items, note, empty_note=None):
    ws = wb.add_sheet(name, cell_overwrite_ok=True)
    ws.write_merge(0, 0, 0, 6, note, STYLE_NOTE)
    ws.row(0).height = 600
    ws.row(0).height_mismatch = True
    for c, v in enumerate(header):
        ws.write(1, c, v, STYLE_HEADER)
    ws.row(1).height = 360
    ws.row(1).height_mismatch = True
    for i, (vals, tier) in enumerate(items, start=2):
        st = STYLE_GREEN if tier == 'green' else (STYLE_YELLOW if tier == 'yellow' else STYLE_DATA)
        for c in range(27):
            ws.write(i, c, vals[c], st)
        ws.row(i).height = 360
        ws.row(i).height_mismatch = True
    ws.col(0).width = 3840
    for c in range(1, 27):
        ws.col(c).width = 7680
    if empty_note:
        ws.write_merge(2, 2, 0, 6, empty_note, STYLE_WARN)
        ws.row(2).height = 700
        ws.row(2).height_mismatch = True
    return ws


def build(official_xls: Path, out: Path, add_regions: bool = True, force: bool = False) -> dict:
    if out.exists() and not force:
        raise FileExistsError(f'输出已存在：{out}\n加 --force 覆盖，或换 --out。')
    rb = xlrd.open_workbook(str(official_xls), formatting_info=True)
    wb = copy(rb)
    header = [rb.sheets()[0].cell_value(1, c) for c in range(27)]
    note = str(rb.sheets()[0].cell_value(0, 0))
    idx = {h: i for i, h in enumerate(header)}

    matches = []
    all_rows = []
    total_people = 0
    for si, s in enumerate(rb.sheets()):
        if s.nrows <= 2:
            continue
        ws = wb.get_sheet(si)
        hdr = [s.cell_value(1, c) for c in range(s.ncols)]
        ix = {h: i for i, h in enumerate(hdr)}
        for r in range(2, s.nrows):
            row = [s.cell_value(r, c) for c in range(s.ncols)]
            if not str(row[ix['职位代码']]).strip():
                continue
            tier, major_tier, reasons = evaluate(row, ix)
            all_rows.append((row, tier, ix))
            try:
                total_people += int(float(row[ix['招考人数']]))
            except Exception:
                pass
            if tier in ('green', 'yellow'):
                st = STYLE_GREEN if tier == 'green' else STYLE_YELLOW
                for c in range(s.ncols):
                    ws.write(r, c, row[c], st)
                matches.append((row, tier, ix))
    log(f'筛选完成：共 {len(all_rows)} 个职位，符合 {len(matches)} 个（绿 {sum(1 for x in matches if x[1]=="green")} / 黄 {sum(1 for x in matches if x[1]=="yellow")}）')

    # 符合条件岗位
    level = {'中央': 0, '省（副省）级': 1, '市（地）级': 2, '县（区）级及以下': 3}
    matches.sort(key=lambda m: (0 if m[1] == 'green' else 1,
                                level.get(str(m[0][m[2]['机构层级']]), 9),
                                str(m[0][m[2]['部门名称']]),
                                str(m[0][m[2]['职位代码']])))
    ws = wb.add_sheet('符合条件岗位', cell_overwrite_ok=True)
    ws.write_merge(0, 0, 0, 6, note, STYLE_NOTE)
    ws.row(0).height = 600; ws.row(0).height_mismatch = True
    for c, v in enumerate(header):
        ws.write(1, c, v, STYLE_HEADER)
    ws.row(1).height = 360; ws.row(1).height_mismatch = True
    for i, (row, tier, ix) in enumerate(matches, start=2):
        st = STYLE_GREEN if tier == 'green' else STYLE_YELLOW
        for c in range(27):
            ws.write(i, c, row[c], st)
        ws.row(i).height = 360; ws.row(i).height_mismatch = True
    ws.col(0).width = 3840
    for c in range(1, 27):
        ws.col(c).width = 7680

    # 地区表
    region_sheets = []
    if add_regions:
        for region in CFG.get('regions', ['陕西', '甘肃', '西安']):
            sel = [(row, tier) for row, tier, ix in matches
                   if region in str(row[idx['工作地点']])]
            sel.sort(key=lambda it: (0 if it[1] == 'green' else 1,
                                     level.get(str(it[0][idx['机构层级']]), 9),
                                     str(it[0][idx['部门名称']]),
                                     str(it[0][idx['职位代码']])))
            name = region + '市岗位' if region == '西安' else region + '省岗位'
            empty = None
            if region == '西安' and not sel:
                empty = '符合条件岗位中暂无西安市岗位；全表西安岗位均未通过当前筛选，详见“西安市全部岗位（参考）”。'
            _write_sheet(wb, name, header, sel, note, empty_note=empty)
            region_sheets.append(name)
        # 西安全量参考
        ref = []
        for s in rb.sheets():
            if s.nrows <= 2:
                continue
            hdr = [s.cell_value(1, c) for c in range(s.ncols)]
            ix = {h: i for i, h in enumerate(hdr)}
            for r in range(2, s.nrows):
                row = [s.cell_value(r, c) for c in range(s.ncols)]
                if not str(row[ix['职位代码']]).strip():
                    continue
                if '西安' in str(row[ix['工作地点']]):
                    ref.append((row, None))
        ref.sort(key=lambda it: (str(it[0][1]), str(it[0][8])))
        ref_name = '西安市全部岗位（参考）'
        _write_sheet(wb, ref_name, header, ref,
                     '参考表：全表西安市岗位，均未通过本人条件筛选，仅供了解，请勿直接作为可报岗位。')
        region_sheets.append(ref_name)

    # 筛选说明
    ws = wb.add_sheet('筛选说明', cell_overwrite_ok=True)
    ws.col(0).width = 5000
    ws.col(1).width = 30000
    r = 0
    ws.write_merge(0, 0, 0, 1, '2027国考岗位筛选说明（按本人条件）', STYLE_TITLE)
    ws.row(0).height = 520; ws.row(0).height_mismatch = True
    r = 1
    def put(k, v):
        nonlocal r
        ws.write(r, 0, k, STYLE_LBL)
        ws.write(r, 1, v, STYLE_TXT)
        n = max(len(str(k)) * 2, len(str(v)))
        h = 300 if n <= 110 else (450 if n <= 220 else 600)
        ws.row(r).height = h; ws.row(r).height_mismatch = True
        r += 1
    put('生成日期', datetime.date.today().isoformat())
    put('原始岗位表', f'{official_xls.name}（{len(all_rows)} 个职位 / {total_people} 人）')
    put('筛选结果', f'绿色（明确符合）{sum(1 for x in matches if x[1]=="green")} 个；黄色（基本符合、需确认）{sum(1 for x in matches if x[1]=="yellow")} 个；合计 {len(matches)} 个')
    p = PROFILE
    put('本人条件', '%s、%s岁、%s、%s（%s %s）、%s、基层%s年、CET4 %s、CET6 %s、%s、户籍%s%s%s。' % (
        p.get('gender', ''), p.get('age', ''), p.get('party', ''), p.get('education_level', ''),
        p.get('major_code', ''), p.get('major_name', ''), p.get('degree_level', ''),
        p.get('grassroots_years', 0), p.get('cet4', 0), p.get('cet6_score', 0),
        p.get('fresh_label', '') or '往届', p.get('household_province', ''), p.get('household_city', ''), p.get('household_county', '')))
    put('筛选口径', '按本人学历/学位/政治面貌自动匹配职位要求；基层工作年限、服务基层项目按本人年限/项目匹配。')
    put('专业口径', '以最高学历阶段为准：优先 %s / %s；相近大类：%s；同时接受工学、理工、不限专业。' % (
        '/'.join(str(x) for x in (p.get('major_codes') or [])), p.get('major_name', ''),
        '/'.join(str(x) for x in (p.get('broad_keywords') or []))))
    put('其他口径', '按本人年龄、应届身份、性别、英语成绩、证书、户籍、退役军人/残疾人/少数民族/运动员等条件过滤；不满足的职位剔除。')
    put('颜色图例', '绿色=明确符合，建议优先；黄色=基本符合但需电话确认（0835学术学位、计算机/电子信息大类、学位学士、英语熟练、体能测评等）。')
    put('重要提示', '本表按“%s”身份假设；若实际身份不同，凡备注“限应届/非应届须2年工作经历”的职位需重新核对。专业、学位、户籍等最终以招录机关答复为准。' % (p.get('fresh_label', '') or '往届'))
    put('数据来源', '国家公务员局《中央机关及其直属机构2027年度考试录用公务员招考简章》原始 .xls；原 4 个工作表内容、列宽、行高、合并单元格均未改动。')
    put('免责声明', '筛选结果仅供参考，不替代官方资格审查。')

    # 调整工作表顺序
    order = ['中央党群机关', '中央国家行政机关（本级）', '中央国家行政机关省级以下直属机构',
             '中央国家行政机关参照公务员法管理事业单位', '符合条件岗位'] + region_sheets + ['筛选说明', 'Sheet1']
    wslist = wb._Workbook__worksheets
    wslist.sort(key=lambda s: order.index(s.name) if s.name in order else 99)
    wslist[:] = [s for s in wslist if s.name != 'Sheet2']
    for i, s in enumerate(wslist):
        try:
            s.index = i
        except Exception:
            pass

    ensure_dir(out)
    backup(out)
    wb.save(str(out))
    return {'total': len(all_rows), 'match': len(matches),
            'green': sum(1 for x in matches if x[1] == 'green'),
            'yellow': sum(1 for x in matches if x[1] == 'yellow'),
            'out': str(out)}
