# -*- coding: utf-8 -*-
"""gk2027 命令行入口：环境检查、筛选、抓取、合并、统计、菜单。"""
from __future__ import annotations

import argparse
import datetime
import json
import sys
from pathlib import Path

from . import downloader, excel_build, excel_snapshot, fetch, stats_report
from .util import CFG, CONFIG_PATH, P, ROOT, ChromeError, ask, ask_yes, backup, banner, check_chrome, chrome_js, chrome_tabs, fmt_int, load_config, log, sha256


def _default_label() -> str:
    return datetime.date.today().strftime('%m-%d')


def _default_counts_path(label: str) -> Path:
    return P('counts_dir') / f"counts_{datetime.date.today().isoformat()}.json"



def _ask_int(prompt, default):
    while True:
        v = ask(prompt, str(default))
        try:
            return int(float(v))
        except ValueError:
            log('请输入数字。')


def _reload_profile():
    """init 后让同进程内的 matcher/excel_build 用上新配置（菜单连续操作时需要）。"""
    import importlib
    from . import util, matcher, excel_build
    util.CFG = load_config()
    global CFG
    CFG = util.CFG
    importlib.reload(matcher)
    importlib.reload(excel_build)


def cmd_init(args):
    banner('设置本人条件')
    cfg = load_config()
    p = cfg.get('profile', {})
    from_json = getattr(args, 'from_json', None)
    if from_json:
        inc = json.load(open(from_json, encoding='utf-8'))
        p.update(inc.get('profile', inc))
        log(f'已从 {from_json} 导入条件。')
    else:
        p['gender'] = ask('性别（男/女）', p.get('gender', '男'))
        p['age'] = _ask_int('年龄', p.get('age', 31))
        p['party'] = ask('政治面貌（中共党员/共青团员/群众/民主党派）', p.get('party', '中共党员'))
        edu = ask('学历（大专/本科/硕士/博士）', p.get('education_level', '硕士'))
        p['education_level'] = edu
        p['education'] = {'大专': '大专', '本科': '本科', '硕士': '硕士研究生', '博士': '博士研究生'}.get(edu, edu)
        p['degree_level'] = ask('学位（无/学士/硕士/博士）', p.get('degree_level', '硕士'))
        if edu in ('硕士', '博士'):
            p['major_code'] = ask('研究生专业代码', p.get('major_code', ''))
            p['major_name'] = ask('研究生专业名称', p.get('major_name', ''))
            codes = ask('专业代码（逗号分隔）', ','.join(p.get('major_codes') or [p['major_code']]))
            p['major_codes'] = [x.strip() for x in codes.split(',') if x.strip()]
            kws = ask('专业关键词（逗号分隔）', ','.join(p.get('major_keywords') or [p['major_name']]))
            p['major_keywords'] = [x.strip() for x in kws.split(',') if x.strip()]
            eq = ask('学术学位近似代码（可空，逗号分隔）', ','.join(p.get('academic_equiv') or []))
            p['academic_equiv'] = [x.strip() for x in eq.split(',') if x.strip()]
        else:
            p['bachelor_major'] = ask('本科专业名称', p.get('bachelor_major', ''))
            p['major_name'] = p.get('major_name', '') or p['bachelor_major']
            p['academic_equiv'] = []
            codes = ask('专业代码（逗号分隔）', ','.join(p.get('major_codes') or []))
            p['major_codes'] = [x.strip() for x in codes.split(',') if x.strip()]
        bkw = ask('专业大类关键词（逗号分隔）', ','.join(p.get('broad_keywords') or []))
        p['broad_keywords'] = [x.strip() for x in bkw.split(',') if x.strip()]
        p['grassroots_years'] = _ask_int('基层工作年限（年）', p.get('grassroots_years', 0))
        p['work_years'] = _ask_int('累计工作年限（年）', p.get('work_years', p.get('grassroots_years', 0)))
        sp = ask('服务基层项目（无/大学生村官/三支一扶/西部计划/特岗教师/退役士兵）', p.get('service_project', '') or '无')
        p['service_project'] = '' if sp in ('无', 'none', '') else sp
        p['cet4'] = _ask_int('CET-4 分数（没有填0）', p.get('cet4', 0))
        p['cet6_score'] = _ask_int('CET-6 分数（没有填0）', p.get('cet6_score', 0))
        p['cet6'] = p['cet6_score'] > 0
        p['household_province'] = ask('户籍省', p.get('household_province', ''))
        p['household_city'] = ask('户籍市', p.get('household_city', ''))
        p['household_county'] = ask('户籍县/区', p.get('household_county', ''))
        fl = ask('应届身份（2027届/2026届/2025届/往届）', p.get('fresh_label', ''))
        p['fresh_label'] = '' if fl in ('往届', '无', '') else fl
        certs = ask('已取得证书（逗号分隔，可空；如 法律职业资格,CPA）', ','.join(p.get('certificates') or []))
        p['certificates'] = [x.strip() for x in certs.split(',') if x.strip()]
        p['is_veteran'] = ask_yes('是否退役军人？', p.get('is_veteran', False))
        p['is_disabled'] = ask_yes('是否残疾人？', p.get('is_disabled', False))
        p['is_minority'] = ask_yes('是否少数民族？', p.get('is_minority', False))
        p['is_athlete'] = ask_yes('是否有运动员等级/赛事名次？', p.get('is_athlete', False))
    cfg['profile'] = p
    backup(CONFIG_PATH)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    log(f"已保存：{CONFIG_PATH}")
    log(f"  {p.get('gender')} / {p.get('age')}岁 / {p.get('party')} / {p.get('education_level')} / {p.get('major_code')} {p.get('major_name')} / 基层{p.get('grassroots_years')}年 / CET4 {p.get('cet4')} / CET6 {p.get('cet6_score')} / 户籍 {p.get('household_province')}{p.get('household_city')}{p.get('household_county')} / {p.get('fresh_label') or '往届'}")
    _reload_profile()
    return 0


def cmd_download(args):
    banner('下载官方资料')
    out = Path(args.out).expanduser() if getattr(args, 'out', None) else P('official_xls').parent.parent
    log(f'输出目录：{out}')
    log('将下载：招考公告/大纲、政策法规、常见问题、相关附件、岗位表。')
    if not ask_yes('开始下载？', True):
        log('已取消。')
        return 1
    try:
        res = downloader.download_all(out, force=getattr(args, 'force', False), skip_pdf=getattr(args, 'skip_pdf', False), only=getattr(args, 'only', 'all'))
    except Exception as e:
        log('下载失败：' + str(e))
        return 1
    log(f"完成：附件 {res.get('resources', 0)} 个；岗位表：{res.get('position') or '未下载'}")
    return 0

def cmd_doctor(args):
    banner('环境检查')
    ok = True
    fetch_ok = True
    for mod in ('xlrd', 'xlwt', 'xlutils'):
        try:
            m = __import__(mod)
            log(f'[OK] {mod} {getattr(m, "__version__", "")}')
        except Exception as e:
            ok = False
            log(f'[X ] {mod} 未安装：{e}')
    for key in ('official_xls', 'workbook', 'stats_xls'):
        p = P(key)
        if p.exists():
            log(f'[OK] {key}: {p} ({p.stat().st_size/1024/1024:.1f} MB)')
        else:
            ok = False
            log(f'[X ] {key} 不存在：{p}')
    log('\n[Chrome]')
    if not check_chrome():
        log('[X ] Chrome 未运行')
        ok = False
    else:
        tabs = chrome_tabs()
        if not tabs:
            log('[!] 没有打开 bm.scs.gov.cn 页面')
        for u, t in tabs:
            log('  - ' + u)
        try:
            r = chrome_js("'JS_OK '+(1+1)", 'person_position.html', timeout=20)
            log(f'[OK] 页面脚本：{r}')
        except Exception as e:
            fetch_ok = False
            log(f'[!] 页面脚本不可用：{e}')
    log('\n[配置]')
    prof = CFG.get('profile', {})
    log(f"  {prof.get('gender')} / {prof.get('age')}岁 / {prof.get('party')} / {prof.get('education')} / {prof.get('major_code')} {prof.get('major_name')} / CET4 {prof.get('cet4')} / 户籍 {prof.get('household_province')}{prof.get('household_city')}{prof.get('household_county')}")
    if ok and not fetch_ok:
        log('\n结论：基础环境正常；抓取功能暂不可用，请按上面提示开启 Chrome 的 JS 自动化。')
    else:
        log('\n结论：' + ('环境正常，可以开始。' if ok else '有缺失项，请按上面提示处理。'))
    return 0 if ok else 1


def cmd_build(args):
    banner('生成个人筛选工作簿')
    official = Path(args.official).expanduser() if args.official else P('official_xls')
    out = Path(args.out).expanduser() if args.out else P('workbook')
    if not official.exists():
        raise SystemExit(f'官方岗位表不存在：{official}')
    if out.exists() and not args.force:
        if out.resolve() == P('workbook').resolve():
            if not ask_yes(f'这是当前工作簿：{out}\nbuild 会重建，并丢失已有的报名人数快照列。确定覆盖？', False):
                log('已取消。')
                return 1
        elif not ask_yes(f'输出已存在：{out}\n是否覆盖？', False):
            log('已取消。')
            return 1
    res = excel_build.build(official, out, add_regions=not args.no_regions, force=True)
    log(f"完成：{res['out']}")
    log(f"符合 {res['match']} 个（绿 {res['green']} / 黄 {res['yellow']}）")
    log('原 4 个工作表内容未改，新增“符合条件岗位”' + ('、地区表' if not args.no_regions else '') + '和“筛选说明”。')
    return 0


def cmd_fetch(args):
    banner('抓取报名人数')
    wb = Path(args.workbook).expanduser() if args.workbook else P('workbook')
    label = args.label or _default_label()
    out = Path(args.out).expanduser() if args.out else _default_counts_path(label)
    log(f'工作簿：{wb}')
    log(f'标签：{label}')
    log(f'输出：{out}')
    log('要求：Chrome 已登录，已打开 个人中心 → 职位查询 → 按代码选择职位，并保持页面打开。')
    if out.exists() and not ask_yes(f'counts 已存在：{out}\n是否覆盖？', False):
        log('已取消。')
        return 1
    if not ask_yes('是否继续？', True):
        log('已取消。')
        return 1
    try:
        res = fetch.fetch_counts(wb, label, out, workers=args.workers, retries=args.retries)
    except ChromeError as e:
        log('抓取失败：' + str(e))
        return 1
    log(f"完成：{res['records']} 条，失败 {res['errors']} 条")
    return 0


def cmd_snapshot(args):
    banner('合并报名人数到工作簿')
    wb = Path(args.workbook).expanduser() if args.workbook else P('workbook')
    counts = Path(args.counts).expanduser()
    label = args.label or _default_label()
    if not counts.exists():
        raise SystemExit(f'counts 文件不存在：{counts}')
    if not wb.exists():
        raise SystemExit(f'工作簿不存在：{wb}')
    try:
        res = excel_snapshot.add_snapshot(wb, counts, label, force=args.force)
    except ValueError as e:
        log(str(e))
        return 1
    log(f"完成：{res['records']} 个职位，总报名 {fmt_int(res['total'])} 人")
    return 0


def cmd_report(args):
    banner('生成报名人数统计表')
    wb = Path(args.workbook).expanduser() if args.workbook else P('workbook')
    out = Path(args.out).expanduser() if args.out else P('stats_xls')
    labels = args.labels.split(',') if args.labels else None
    if out.exists() and not args.force:
        if not ask_yes(f'输出已存在：{out}\n是否覆盖？', False):
            log('已取消。')
            return 1
    try:
        res = stats_report.build_report(wb, out, labels=labels, force=True)
    except Exception as e:
        log('生成失败：' + str(e))
        return 1
    log(f"完成：{res['out']}（快照：{','.join(res['snapshots'])}，共 {res['records']} 条）")
    return 0


def cmd_all(args):
    banner('一键全流程')
    label = args.label or _default_label()
    wb = P('workbook')
    counts = Path(args.counts).expanduser() if args.counts else _default_counts_path(label)
    log(f'标签：{label}')
    log(f'工作簿：{wb}')
    log(f'counts：{counts}')
    if not wb.exists():
        log(f'工作簿不存在：{wb}\n请先运行 build 生成筛选表。')
        return 1
    if not args.skip_fetch:
        log('步骤 1/4：抓取报名人数')
        try:
            res = fetch.fetch_counts(wb, label, counts, workers=args.workers, retries=args.retries)
            log(f"  抓取 {res['records']} 条，失败 {res['errors']} 条")
        except ChromeError as e:
            log('抓取失败：' + str(e))
            if not ask_yes('是否跳过抓取，使用已有 counts 继续？', False):
                return 1
    if not counts.exists():
        log(f'counts 不存在：{counts}，无法继续。')
        return 1
    log('步骤 2/4：合并到工作簿')
    try:
        excel_snapshot.add_snapshot(wb, counts, label, force=args.force)
    except ValueError as e:
        log(str(e))
        if not ask_yes('是否跳过合并，直接生成统计表？', False):
            return 1
    log('步骤 3/4：生成统计表')
    try:
        stats_report.build_report(wb, P('stats_xls'), force=True)
    except Exception as e:
        log('统计表生成失败：' + str(e))
        return 1
    log('步骤 4/4：完成。')
    return 0


def cmd_guide(args=None):
    banner('使用说明')
    log('''这是一个 2027 国考岗位筛选 + 报名人数统计的小工具。

常用命令：
  ./run.command doctor                  检查环境和 Chrome 登录状态
  ./run.command build                   生成带颜色/新sheet的筛选表
  ./run.command fetch --label 10-23     抓取今天的报名人数（需Chrome登录）
  ./run.command snapshot --label 10-23 --counts data/counts/xxx.json
                                          把人数合并进Excel（与上一天并列）
  ./run.command report                  生成独立统计表（总表/汇总/说明）
  ./run.command all --label 10-24       一键：抓取→合并→统计
  ./run.command                        进入交互菜单

目录：
  2027国考资料/05_岗位表/                 官方原始岗位表
  2027国考资料/06_个人筛选/               生成的筛选表/统计表
  data/counts/                            每次抓取的原始JSON（可追溯）
  gk2027_config.json                      路径与本人条件配置

抓取前的准备：
  1. Chrome 打开 https://bm.scs.gov.cn/kl2027 并登录；
  2. 进入 个人中心 → 职位查询 → 按代码选择职位，保持页面打开；
  3. Chrome 菜单：显示 → 开发者 → 勾选“允许 Apple 事件中的 JavaScript”；
  4. 运行 fetch，按提示操作。

注意：
  - 报名人数会随报名进程变化，以系统实时数据为准；
  - 黄色岗位仍需电话确认专业/学位/英语/体能等条件；
  - 黄色/绿色判断是辅助筛选，不替代官方资格审查。
''')
    return 0


def _menu():
    while True:
        banner('2027国考岗位工具')
        log('1. 环境检查（doctor）')
        log('2. 设置本人条件（init）')
        log('3. 下载官方资料（download）')
        log('4. 生成筛选表（build）')
        log('5. 抓取今天报名人数（fetch）')
        log('6. 合并人数到Excel（snapshot）')
        log('7. 生成统计表（report）')
        log('8. 一键全流程（all）')
        log('9. 使用说明（guide）')
        log('0. 退出')
        choice = ask('请选择', '1')
        try:
            if choice == '1':
                cmd_doctor(None)
            elif choice == '2':
                cmd_init(argparse.Namespace(from_json=None))
            elif choice == '3':
                cmd_download(argparse.Namespace(out=None, force=False, skip_pdf=False, only='all'))
            elif choice == '4':
                cmd_build(argparse.Namespace(official=None, out=None, no_regions=False, force=False))
            elif choice == '5':
                cmd_fetch(argparse.Namespace(workbook=None, label=None, out=None, workers=3, retries=2))
            elif choice == '6':
                counts = ask('counts JSON 路径', str(_default_counts_path(_default_label())))
                label = ask('标签（如 10-23）', _default_label())
                cmd_snapshot(argparse.Namespace(workbook=None, counts=counts, label=label, force=False))
            elif choice == '7':
                cmd_report(argparse.Namespace(workbook=None, out=None, labels=None, force=False))
            elif choice == '8':
                cmd_all(argparse.Namespace(label=None, counts=None, skip_fetch=False, workers=3, retries=2, force=False))
            elif choice == '9':
                cmd_guide(None)
            elif choice == '0':
                log('再见。')
                return 0
            else:
                log('无效选项。')
        except SystemExit:
            pass
        except Exception as e:
            log('出错：' + str(e))
        if choice != '0':
            ask('按回车继续', '')


def build_parser():
    p = argparse.ArgumentParser(
        prog='gk2027',
        description='2027国考岗位筛选与报名人数统计工具',
        epilog='不带子命令运行时进入交互菜单；各子命令加 -h 查看参数。',
    )
    sub = p.add_subparsers(dest='cmd')

    sp = sub.add_parser('doctor', help='检查 Python 依赖、文件、Chrome 登录状态')
    sp.set_defaults(func=cmd_doctor)
    sp = sub.add_parser('init', help='设置本人条件（交互式）')
    sp.add_argument('--from-json', help='从 JSON 文件导入 profile')
    sp.set_defaults(func=cmd_init)

    sp = sub.add_parser('download', help='下载官方公告/大纲/政策/问答/附件/岗位表')
    sp.add_argument('--out', help='输出目录，默认 2027国考资料')
    sp.add_argument('--force', action='store_true', help='覆盖已存在文件')
    sp.add_argument('--skip-pdf', action='store_true', help='只存 HTML，不生成 PDF')
    sp.add_argument('--only', choices=['all', 'position', 'articles', 'resources'], default='all', help='只下载某一部分')
    sp.set_defaults(func=cmd_download)


    sp = sub.add_parser('build', help='从官方岗位表生成个人筛选工作簿')
    sp.add_argument('--official', help='官方岗位表 .xls 路径')
    sp.add_argument('--out', help='输出工作簿路径')
    sp.add_argument('--no-regions', action='store_true', help='不生成陕西/甘肃/西安地区表')
    sp.add_argument('--force', action='store_true', help='覆盖已存在文件')
    sp.set_defaults(func=cmd_build)

    sp = sub.add_parser('fetch', help='从已登录的 Chrome 抓取报名人数')
    sp.add_argument('--workbook', help='筛选工作簿路径')
    sp.add_argument('--label', help='快照标签，如 10-23')
    sp.add_argument('--out', help='counts JSON 输出路径')
    sp.add_argument('--workers', type=int, default=3, help='并发数（默认3）')
    sp.add_argument('--retries', type=int, default=2, help='失败重试次数（默认2）')
    sp.set_defaults(func=cmd_fetch)

    sp = sub.add_parser('snapshot', help='把 counts JSON 合并到工作簿，与上一天并列')
    sp.add_argument('--workbook', help='筛选工作簿路径')
    sp.add_argument('--counts', required=True, help='fetch 生成的 counts JSON')
    sp.add_argument('--label', help='快照标签，如 10-23')
    sp.add_argument('--force', action='store_true', help='允许重复标签')
    sp.set_defaults(func=cmd_snapshot)

    sp = sub.add_parser('report', help='生成独立报名人数统计表')
    sp.add_argument('--workbook', help='筛选工作簿路径')
    sp.add_argument('--out', help='统计表输出路径')
    sp.add_argument('--labels', help='指定两个快照标签，逗号分隔，如 10-22,10-23')
    sp.add_argument('--force', action='store_true', help='覆盖已存在文件')
    sp.set_defaults(func=cmd_report)

    sp = sub.add_parser('all', help='一键：抓取 → 合并 → 统计')
    sp.add_argument('--label', help='快照标签')
    sp.add_argument('--counts', help='已有 counts JSON（配合 --skip-fetch）')
    sp.add_argument('--skip-fetch', action='store_true', help='跳过抓取，直接用已有 counts')
    sp.add_argument('--workers', type=int, default=3)
    sp.add_argument('--retries', type=int, default=2)
    sp.add_argument('--force', action='store_true')
    sp.set_defaults(func=cmd_all)

    sp = sub.add_parser('guide', help='查看使用说明')
    sp.set_defaults(func=cmd_guide)
    return p


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    if not argv:
        return _menu()
    parser = build_parser()
    args = parser.parse_args(argv)
    if not hasattr(args, 'func'):
        parser.print_help()
        return 0
    return args.func(args)


if __name__ == '__main__':
    raise SystemExit(main())
