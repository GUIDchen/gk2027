# -*- coding: utf-8 -*-
"""下载 2027 国考专题网站的公告/大纲/政策/问答/附件/岗位表。"""
from __future__ import annotations

import base64
import datetime
import json
import re
import subprocess
import urllib.request
import zipfile
from pathlib import Path

from .util import CFG, ensure_dir, log

API = 'http://dl.scs.gov.cn'
UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'
CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
HB01 = CFG.get('hb01_id', '8a81f6eca0221f1401a0242759fc0000')

GROUPS = {
    '招考公告': ('0000000062b7b2b60162bccd55ec0006', '01_招考公告与考试大纲'),
    '政策法规': ('0000000062b7b2b60162bcd13002000c', '02_政策法规'),
    '常见问题': ('0000000062b7b2b60162bcd208c5000e', '03_常见问题'),
}
POSITION_NAME = '中央机关及其直属机构2027年度考试录用公务员招考简章.xls'


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    return urllib.request.urlopen(req, timeout=90).read()


def fetch_json(url: str):
    return json.loads(fetch(url).decode('utf-8'))


def sanitize(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', '_', str(name)).strip()
    return name[:100]


def _inline_images(content: str) -> str:
    def repl(m):
        rid = m.group(1)
        try:
            data = fetch(f'{API}/download/{rid}')
            if data[:4] == b'\x89PNG':
                mime = 'image/png'
            elif data[:3] == b'\xff\xd8\xff':
                mime = 'image/jpeg'
            else:
                mime = 'image/png'
            return 'src="data:%s;base64,%s"' % (mime, base64.b64encode(data).decode())
        except Exception:
            return m.group(0)
    return re.sub(r'src="downloadUrl([0-9a-zA-Z]+)"', repl, content)


def _make_html(article: dict, url: str) -> str:
    title = article.get('articleTitle', '')
    ctime = article.get('ctime')
    dt = datetime.datetime.fromtimestamp(int(ctime) / 1000).strftime('%Y-%m-%d') if ctime else ''
    content = _inline_images(article.get('content', ''))
    return f'''<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>{title}</title>
<style>
@page {{ size: A4; margin: 18mm 16mm; }}
body {{ font-family:"PingFang SC","Songti SC","SimSun",serif; color:#1a1a1a; }}
.doc {{ max-width: 820px; margin: 0 auto; padding: 28px 18px 40px; }}
h1.doc-title {{ font-size:24px; line-height:1.5; text-align:center; margin:4px 0 10px; }}
.meta {{ text-align:center; color:#666; font-size:12px; padding-bottom:10px; border-bottom:1px solid #ddd; margin-bottom:18px; }}
.content {{ font-size:16px; line-height:1.9; }}
.content p {{ margin:0 0 6px; }}
.content img {{ max-width:100%; height:auto; }}
.content table {{ border-collapse:collapse; max-width:100%; }}
.content td, .content th {{ border:1px solid #999; padding:4px 6px; }}
.source {{ margin-top:28px; padding-top:8px; border-top:1px dashed #ccc; font-size:11px; color:#888; word-break:break-all; }}
</style></head><body><div class="doc">
<h1 class="doc-title">{title}</h1>
<div class="meta">发布日期：{dt}　|　来源：中央机关及其直属机构2027年度考试录用公务员专题网站</div>
<div class="content">{content}</div>
<div class="source">原文链接：<a href="{url}">{url}</a></div>
</div></body></html>'''


def _render_pdf(html_path: Path, pdf_path: Path) -> bool:
    if not Path(CHROME).exists():
        log('  [跳过PDF] 未找到 Google Chrome')
        return False
    cmd = [CHROME, '--headless=new', '--disable-gpu', '--no-sandbox', '--no-pdf-header-footer',
           f'--print-to-pdf={pdf_path}', '--virtual-time-budget=20000', html_path.as_uri()]
    try:
        subprocess.run(cmd, capture_output=True, timeout=300)
        return pdf_path.exists() and pdf_path.stat().st_size > 5000
    except Exception as e:
        log(f'  [PDF失败] {e}')
        return False


def _article_url(aid: str, col: str, parent: str) -> str:
    return (f'http://bm.scs.gov.cn/pp/gkweb/core/web/ui/business/article/articledetail.html'
            f'?ArticleId={aid}&id={parent}&eid={col}')


def download_group(out: Path, group_name: str, force: bool = False, skip_pdf: bool = False) -> list:
    group_id, folder = GROUPS[group_name]
    dest = out / folder
    dest.mkdir(parents=True, exist_ok=True)
    data = fetch_json(f'{API}/api/article/{HB01}/{group_id}')
    groups = data.get('groupList', [])
    order = 0
    resources = []
    for g in groups:
        for a in g.get('articleList', []):
            order += 1
            aid = a.get('id')
            col = a.get('cmsArticleColumnId') or group_id
            parent = a.get('parentColumnId') or '0000000062b7b2b60162bccd55ec0006'
            try:
                detail = fetch_json(f'{API}/api/article/{aid}')
            except Exception as e:
                log(f'  [失败] {a.get("articleTitle")}: {e}')
                continue
            article = detail.get('article', {})
            title = str(article.get('articleTitle') or a.get('articleTitle') or aid)
            stem = f'{order:02d}_{sanitize(title)}'
            hp = dest / (stem + '.html')
            pp = dest / (stem + '.pdf')
            if hp.exists() and not force:
                log(f'  [跳过] {stem}.html')
            else:
                hp.write_text(_make_html(article, _article_url(aid, col, parent)), encoding='utf-8')
                log(f'  [HTML] {stem}.html')
            if not skip_pdf and (force or not pp.exists()):
                _render_pdf(hp, pp)
            for r in detail.get('resourceList') or []:
                resources.append(r)
    return resources


def download_resources(out: Path, force: bool = False) -> dict:
    dl_dir = out / '04_相关下载'
    dl_dir.mkdir(parents=True, exist_ok=True)
    data = fetch_json(f'{API}/api/res/{HB01}/1110')
    res = data.get('resList', [])
    saved = []
    zip_path = None
    for r in res:
        name = str(r.get('resourceName') or r.get('resResourceId'))
        if name.lower().endswith('.zip') or '招考简章' in name:
            zip_path = dl_dir / '中央机关及其直属机构2027年度考试录用公务员招考简章.zip'
            target = zip_path
        else:
            target = dl_dir / sanitize(name)
        if target.exists() and not force:
            log(f'  [跳过] {target.name}')
        else:
            target.write_bytes(fetch(f'{API}/download/{r.get("resResourceId")}'))
            log(f'  [附件] {target.name}')
        saved.append(target)
    return {'files': saved, 'zip': zip_path}


def extract_position(zip_path: Path, out: Path, force: bool = False) -> Path:
    pos_dir = out / '05_岗位表'
    pos_dir.mkdir(parents=True, exist_ok=True)
    target = pos_dir / POSITION_NAME
    if target.exists() and not force:
        log(f'  [跳过] {target.name}')
        return target
    with zipfile.ZipFile(zip_path) as z:
        member = z.infolist()[0]
        data = z.read(member)
    target.write_bytes(data)
    log(f'  [岗位表] {target.name} ({len(data)/1024/1024:.1f} MB)')
    return target


def download_all(out: Path, force: bool = False, skip_pdf: bool = False, only: str = 'all') -> dict:
    out = Path(out)
    ensure_dir(out)
    result = {'articles': 0, 'resources': 0, 'position': None}
    if only in ('all', 'position', 'resources'):
        # 先下载附件/岗位表
        res = download_resources(out, force=force)
        result['resources'] = len(res['files'])
        if res['zip']:
            result['position'] = str(extract_position(res['zip'], out, force=force))
    if only in ('all', 'articles'):
        for name in GROUPS:
            log(f'== {name}')
            rs = download_group(out, name, force=force, skip_pdf=skip_pdf)
            # 证监会大纲等文章附件
            if rs:
                zdir = out / '04_相关下载' / '证监会专业科目大纲'
                other = out / '04_相关下载'
                zdir.mkdir(parents=True, exist_ok=True)
                other.mkdir(parents=True, exist_ok=True)
                for r in rs:
                    nm = str(r.get('resourceName') or '')
                    if not nm:
                        continue
                    is_doc = nm.lower().endswith(('.doc', '.docx'))
                    target = zdir / sanitize(nm) if is_doc else other / sanitize(nm)
                    if target.exists() and not force:
                        continue
                    target.write_bytes(fetch(f'{API}/download/{r.get("resResourceId")}'))
                    log(f'  [附件] {target.relative_to(out)}')
    return result
