# -*- coding: utf-8 -*-
"""从已登录的 Chrome 职位查询页抓取报名人数。"""
from __future__ import annotations

import base64
import json
import time
from pathlib import Path

import xlrd

from .util import CFG, ChromeError, chrome_js, ensure_dir, fmt_int, log

PAGE = 'person_position.html'


def _js_b64(js: str) -> str:
    b64 = base64.b64encode(js.encode('utf-8')).decode()
    return ("(function(){var s=document.createElement('script');"
            "s.textContent=new TextDecoder().decode(Uint8Array.from(atob('" + b64 +
            "'),function(c){return c.charCodeAt(0);}));"
            "document.body.appendChild(s);return 'ok';})()")


def build_jobs(workbook: Path):
    rb = xlrd.open_workbook(str(workbook), formatting_info=True)
    s = rb.sheet_by_name('符合条件岗位')
    header = [s.cell_value(1, c) for c in range(s.ncols)]
    ix = {h: i for i, h in enumerate(header)}
    jobs = []
    for r in range(2, s.nrows):
        code = str(s.cell_value(r, ix['职位代码'])).strip()
        if not code:
            continue
        xf = rb.xf_list[s.cell_xf_index(r, 0)]
        tier = 'green' if xf.background.pattern_colour_index == 42 else 'yellow'
        jobs.append({
            'code': code,
            'dept': str(s.cell_value(r, ix['部门代码'])).strip(),
            'name': str(s.cell_value(r, ix['部门名称'])),
            'bureau': str(s.cell_value(r, ix['用人司局'])),
            'title': str(s.cell_value(r, ix['招考职位'])),
            'num': s.cell_value(r, ix['招考人数']),
            'loc': str(s.cell_value(r, ix['工作地点'])),
            'tier': tier,
        })
    return jobs


def _check_js():
    try:
        r = chrome_js("'JS_OK '+(1+1)", PAGE, timeout=30)
    except ChromeError:
        raise
    if 'JS_OK 2' not in r:
        raise ChromeError('Chrome 页面脚本执行异常：' + r)


def _probe_api(job):
    js = f"""(function(){{
      var ta=document.getElementById('__api_probe');
      if(!ta){{ta=document.createElement('textarea');ta.id='__api_probe';ta.style.display='none';document.body.appendChild(ta);}}
      ta.textContent='PENDING';
      $.ajax({{type:'GET',url:neu.personBaseUrl+'/api/search/detail/positionCode/{job['code']}/{job['dept']}/1110/0/0',dataType:'json',timeout:10000,
        success:function(r){{var d=r.searchHb11Dto||{{}};ta.textContent='OK '+(d.num===undefined?'':d.num)+' '+(d.booknum===undefined?'':d.booknum);}},
        error:function(x){{ta.textContent='ERR '+x.status;}}}});
      return 'sent';
    }})()"""
    chrome_js(_js_b64(js), PAGE, timeout=30)
    for _ in range(15):
        time.sleep(1)
        val = chrome_js("(function(){var e=document.getElementById('__api_probe');return e?e.textContent:'NONE';})()", PAGE, timeout=30)
        if val.startswith('OK'):
            return val
        if val.startswith('ERR'):
            if '401' in val:
                raise ChromeError('登录会话已失效（接口返回 401）。请重新登录，打开 职位查询 → 按代码选择职位 页面后重试。')
            raise ChromeError('职位详情接口返回错误：' + val)
    raise ChromeError('职位详情接口超时，请检查页面是否正常。')


def _inject_jobs(jobs):
    # 先清空页面里的旧缓存，避免重复注入导致 2 倍岗位
    chrome_js(_js_b64("window.__jobs=[];document.title='JOBS 0';"), PAGE, timeout=30)
    time.sleep(0.3)
    log(f'正在向页面注入 {len(jobs)} 个岗位（分块）...')
    for start in range(0, len(jobs), 150):
        chunk = jobs[start:start + 150]
        js = ("window.__jobs=(window.__jobs||[]).concat(" + json.dumps(chunk, ensure_ascii=False) +
              ");document.title='JOBS '+(window.__jobs?window.__jobs.length:0);")
        chrome_js(_js_b64(js), PAGE, timeout=60)
        time.sleep(0.25)
    n = chrome_js('document.title', PAGE, timeout=30)
    if n != f'JOBS {len(jobs)}':
        raise ChromeError(f'岗位注入失败，页面显示：{n}')


def _run_fetch(workers: int, retries: int):
    js = """document.title='FETCH_START';
window.__res={}; window.__pending=(window.__jobs||[]).length; window.__idx=0;
var jobs=window.__jobs||[];
function finish(){
  var ta=document.getElementById('__out_fetch');
  if(!ta){ta=document.createElement('textarea');ta.id='__out_fetch';ta.style.display='none';document.body.appendChild(ta);}
  ta.textContent=JSON.stringify(Object.keys(window.__res).map(function(k){return window.__res[k];}));
  document.title='FETCH_DONE '+Object.keys(window.__res).length;
}
function done(j,rec){
  var key=j.dept+'|'+j.code;
  if(!window.__res[key]) window.__pending--;
  window.__res[key]=rec;
  if(window.__pending%25===0) document.title='FETCH '+window.__pending+' left';
  if(window.__pending<=0) finish();
}
function fetchOne(j,attempt){
  $.ajax({type:'GET',url:neu.personBaseUrl+'/api/search/detail/positionCode/'+j.code+'/'+j.dept+'/1110/0/0',dataType:'json',timeout:12000,
    success:function(r){var d=r.searchHb11Dto||{};done(j,{code:j.code,dept:j.dept,name:j.name,bureau:j.bureau,title:j.title,num:j.num,book:d.num,up:d.upTime,rawbook:d.booknum,rawfirst:d.firstnum});work();},
    error:function(x){if(attempt<__RETRIES__){setTimeout(function(){fetchOne(j,attempt+1);},300);}else{done(j,{code:j.code,dept:j.dept,name:j.name,bureau:j.bureau,title:j.title,num:j.num,book:null,up:null,err:'HTTP '+x.status});work();}}});
}
function work(){ if(window.__idx>=jobs.length) return; var j=jobs[window.__idx++]; fetchOne(j,0); }
for(var k=0;k<__WORKERS__;k++) work();
"""
    js = js.replace('__WORKERS__', str(workers)).replace('__RETRIES__', str(retries))
    chrome_js(_js_b64(js), PAGE, timeout=60)


def _read_textarea(workers: int):
    length = int(chrome_js("(function(){var e=document.getElementById('__out_fetch');return e?e.textContent.length:-1;})()", PAGE, timeout=30))
    if length < 0:
        raise ChromeError('未找到抓取结果，请重试。')
    parts = []
    for start in range(0, length, 40000):
        end = min(start + 40000, length)
        part = chrome_js(f"document.getElementById('__out_fetch').textContent.substring({start},{end})", PAGE, timeout=60)
        parts.append(part)
    raw = ''.join(parts)
    return json.loads(raw)


def fetch_counts(workbook: Path, label: str, out: Path, workers: int = 3, retries: int = 2) -> dict:
    jobs = build_jobs(workbook)
    if not jobs:
        raise ValueError('工作簿里没有“符合条件岗位”数据。')
    log(f'待抓取岗位：{len(jobs)} 个')
    _check_js()
    log('Chrome 页面脚本可用。')
    probe = _probe_api(jobs[0])
    log(f'登录与接口正常（{probe}）。')
    _inject_jobs(jobs)
    log(f'开始抓取（{workers} 路并发，失败重试 {retries} 次）...')
    _run_fetch(workers, retries)
    last = ''
    start = time.time()
    while time.time() - start < 900:
        time.sleep(3)
        title = chrome_js('document.title', PAGE, timeout=30)
        if title != last:
            log('  ' + title)
            last = title
        if title.startswith('FETCH_DONE'):
            break
    else:
        raise ChromeError('抓取超时（15 分钟），请重试。')
    data = _read_textarea(workers)
    errors = [x for x in data if x.get('book') is None]
    ensure_dir(out)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    log(f'抓取完成：{len(data)} 条，失败 {len(errors)} 条；已保存 {out}')
    # 清理页面临时元素
    try:
        chrome_js("(function(){['__out_fetch','__api_probe'].forEach(function(id){var e=document.getElementById(id);if(e)e.remove();});document.title='职位查询';return 'cleaned';})()", PAGE, timeout=30)
    except Exception:
        pass
    return {'records': len(data), 'errors': len(errors), 'out': str(out), 'label': label}
