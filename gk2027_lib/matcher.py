# -*- coding: utf-8 -*-
"""岗位匹配规则（按 profile 配置，可用于不同考生）。"""
from __future__ import annotations

import re

from .util import CFG

PROVINCES = ['北京','天津','河北','山西','内蒙古','辽宁','吉林','黑龙江','上海','江苏','浙江','安徽','福建','江西','山东','河南','湖北','湖南','广东','广西','海南','重庆','四川','贵州','云南','西藏','陕西','甘肃','青海','宁夏','新疆','香港','澳门','台湾']

P = CFG.get('profile', {})
EDU = P.get('education_level') or ('硕士' if '硕士' in str(P.get('education','')) else ('本科' if '本科' in str(P.get('education','')) else ('博士' if '博士' in str(P.get('education','')) else '大专')))
DEG = P.get('degree_level') or P.get('degree') or '无'
PARTY = P.get('party', '')
GENDER = P.get('gender', '')
AGE = int(P.get('age', 31))
FRESH = P.get('fresh_label', '')
GRASS = int(P.get('grassroots_years', 0) or 0)
WORK_YEARS = int(P.get('work_years', GRASS) or 0)
SERVICE = P.get('service_project', '') or ''
CERTIFICATES = set(P.get('certificates', []) or [])
VETERAN = bool(P.get('is_veteran', False))
DISABLED = bool(P.get('is_disabled', False))
MINORITY = bool(P.get('is_minority', False))
ATHLETE = bool(P.get('is_athlete', False))
CET4 = int(P.get('cet4', 0) or 0)
CET6 = int(P.get('cet6_score', 0) or (425 if P.get('cet6') else 0))
IELTS = float(P.get('ielts', 0) or 0)
TOEFL = int(P.get('toefl', 0) or 0)
HOME_PROV = P.get('household_province', '')
HOME_CITY = P.get('household_city', '')
HOME_COUNTY = P.get('household_county', '')
MAJOR_CODES = [str(x) for x in (P.get('major_codes') or [P.get('major_code', '')]) if str(x).strip()]
MAJOR_NAME = str(P.get('major_name', '') or '')
ACADEMIC_EQUIV = [str(x) for x in (P.get('academic_equiv') or []) if str(x).strip()]
BROAD = list(P.get('broad_keywords') or [])


def edu_allowed(v: str) -> bool:
    v = str(v)
    if EDU == '大专':
        return v in {'大专及以上', '仅限大专', '大专或本科'}
    if EDU == '本科':
        return v in {'本科及以上', '仅限本科', '本科或硕士研究生', '大专及以上', '大专或本科'}
    if EDU == '硕士':
        return v in {'本科及以上', '硕士研究生及以上', '仅限硕士研究生', '本科或硕士研究生', '大专及以上'}
    if EDU == '博士':
        return v in {'仅限博士研究生', '硕士研究生及以上', '本科及以上', '本科或硕士研究生'}
    return v in {'本科及以上', '硕士研究生及以上', '仅限硕士研究生', '本科或硕士研究生', '大专及以上'}


def deg_allowed(v: str) -> bool:
    v = str(v)
    if DEG in ('无', '', '无学位'):
        return v == '无要求'
    if DEG == '学士':
        if EDU == '本科':
            return v in {'学士', '无要求', '与最高学历相对应的学位'}
        return v in {'学士', '无要求'}
    if DEG == '硕士':
        return v in {'硕士', '学士', '无要求', '与最高学历相对应的学位'}
    if DEG == '博士':
        return v in {'博士', '硕士', '学士', '无要求', '与最高学历相对应的学位'}
    return v in {'硕士', '学士', '无要求', '与最高学历相对应的学位'}


def pol_allowed(v: str) -> bool:
    v = str(v)
    if PARTY == '中共党员':
        return v in {'中共党员', '中共党员或共青团员', '不限'}
    if PARTY == '共青团员':
        return v in {'共青团员', '中共党员或共青团员', '不限'}
    if PARTY == '民主党派':
        return v in {'不限', '民主党派'}
    return v == '不限'


def grassroots_ok(v: str) -> bool:
    need = {'一年': 1, '二年': 2, '三年': 3, '五年以上': 5}
    return v == '无限制' or GRASS >= need.get(v, 99)


def service_ok(v: str) -> bool:
    if v == '无限制':
        return True
    return bool(SERVICE) and SERVICE in v


def major_tier(text: str) -> str:
    t = str(text or '').strip()
    if not t:
        return 'D'
    if '不限' in t:
        return 'U'
    if '核' in t and ('理工' in t or '工学' in t) and not any(k in t for k in ([MAJOR_NAME] if MAJOR_NAME else []) + BROAD):
        return 'D'
    if any(c and c in t for c in MAJOR_CODES):
        return 'A'
    if any(c and c in t for c in ACADEMIC_EQUIV):
        return 'A2'
    if MAJOR_NAME and MAJOR_NAME in t:
        return 'A'
    if any(k and k in t for k in BROAD):
        return 'B'
    if any(k in t for k in ['工学', '理工']):
        return 'C'
    return 'D'


def split_stages(text: str):
    t = str(text or '')
    yan_m = re.search(r'(?:以)?(?:研究生|硕士研究生|硕士)(?:专业)?(?:报考)?(?:阶段|学历)?(?:为|：|:)\s*(.*?)(?=[；;。\n]|$)', t)
    ben_m = re.search(r'(?:以)?(?:大学)?本科(?:专业)?(?:报考)?(?:阶段|学历)?(?:为|：|:)\s*(.*?)(?=[；;。\n]|$)', t)
    yan = yan_m.group(1).strip() if yan_m else ''
    ben = ben_m.group(1).strip() if ben_m else ''
    if ben and '研究生' in ben:
        ben = ben.split('研究生')[0].strip().rstrip('，, ')
    gen = t
    if ben:
        gen = gen.replace(ben, '')
    if yan:
        gen = gen.replace(yan, '')
    return ben, yan, gen.strip(' ；;。\n')


def major_tier_for_edu(text: str, edu: str) -> str:
    ben, yan, gen = split_stages(text)
    if EDU in ('硕士', '博士') and yan:
        return major_tier(yan)
    if EDU == '本科' and ben:
        return major_tier(ben)
    return major_tier(gen or ben or yan or text)


def household_check(note: str):
    n = str(note)
    for phrase in ['可为符合条件人员办理集体户口', '可为符合条件人员提供集体宿舍、办理集体户口',
                   '单位无法办理集体户口', '单位不办理集体户口', '办理集体户口', '集体户口',
                   '办理集体户籍', '集体户籍']:
        n = n.replace(phrase, '')
    if not any(k in n for k in ['户籍', '户口', '生源']):
        return True, ''
    if re.search(r'不限户籍|户籍不限|不限生源|生源不限', n):
        return True, ''
    regions = re.findall(r'([^，。；;\n]{1,40}?)(?:户籍|户口|生源)', n)
    joined = ' '.join(regions)
    if not joined:
        return False, '户籍/生源限制表述不明确'
    if (HOME_CITY and HOME_CITY in joined) or (HOME_COUNTY and HOME_COUNTY in joined):
        return True, ''
    if HOME_PROV and HOME_PROV in joined:
        rest = joined.replace(HOME_PROV, '').strip(' ，,')
        if not rest or len(rest) <= 1:
            return True, ''
        return False, '限' + joined[:24] + '户籍/生源'
    for p in PROVINCES:
        if p in joined:
            return False, '限' + p + '户籍/生源'
    if any(k in joined for k in ['当地', '本县', '本市', '本州', '本省']):
        return False, '限当地户籍/生源，需确认'
    return False, '限' + joined[:24] + '户籍/生源'


def english_check(note: str):
    if not any(k in note for k in ['英语', '四级', '六级', '雅思', '托福', '专八', '专业八级', 'CATTI', 'PETS', '外语']):
        return True, ''
    if '英语专业' in note and '英语' not in MAJOR_NAME:
        return False, '限英语专业'
    if '六级' in note:
        m = re.search(r'六级[^0-9]{0,12}(\d{3})', note)
        need = int(m.group(1)) if m else 425
        if CET6 >= need:
            return True, ''
        return False, '英语要求六级%d分' % need
    if any(k in note for k in ['专业八级', '专八', '雅思', '托福', 'CATTI', '全国英语等级考试5级', 'PETS5']):
        return False, '英语要求专八/雅思/托福等'
    m = re.search(r'四级[^0-9]{0,12}(\d{3})', note)
    if m:
        need = int(m.group(1))
        if CET4 >= need:
            return True, ''
        return False, '四级要求%d分' % need
    if '四级' in note:
        return (CET4 > 0), ('' if CET4 > 0 else '四级未通过')
    if any(k in note for k in ['英语熟练', '英语良好', '英语流利', '英语水平较高', '较强的英语']):
        return True, '英语水平要求需确认'
    if '英语' in note:
        return True, '英语要求需确认'
    return True, ''


CERT_WORDS = ['法律职业资格', '司法考试', '注册会计师', 'CPA', '税务师', '律师执业', '计算机等级', '软考',
              '计算机技术与软件专业技术资格', '执业医师', '教师资格', '护士执业', '翻译专业资格']


def cert_check(note: str):
    required = [w for w in CERT_WORDS if w in note]
    if not required:
        return True, ''
    if any(w in CERTIFICATES for w in required):
        return True, ''
    # 优先/加分不算硬条件
    for w in required:
        idx = note.find(w)
        seg = note[max(0, idx - 20):idx + 40]
        if '优先' in seg and not re.search(r'须|必须|要求|应取得|需取得|取得.*资格', seg):
            return True, ''
    return False, '要求' + required[0]


def fresh_check(note: str):
    years = re.findall(r'(20\d\d)届', note) + re.findall(r'(20\d\d)年应届', note)
    if years and FRESH not in years:
        return False, '限' + '/'.join(sorted(set(years))) + '届'
    if '限应届' in note and not FRESH:
        return False, '限应届高校毕业生'
    if '非应届' in note and not FRESH:
        if WORK_YEARS < 2:
            return False, '非应届须2年以上工作经历'
    return True, ''


def gender_check(note: str):
    if GENDER == '男':
        if re.search(r'限女性|仅限女性|女性报考|，女性|女性。|女性（', note) and '男女不限' not in note:
            if '适合女性' in note and not re.search(r'限女性|仅限女性', note):
                return True, '适合女性'
            return False, '限女性'
    elif GENDER == '女':
        if re.search(r'限男性|仅限男性|男性报考|，男性|男性。|男性（', note) and '男女不限' not in note:
            if '适合男性' in note and not re.search(r'限男性|仅限男性', note):
                return True, '适合男性'
            return False, '限男性'
    return True, ''


def workexp_check(note: str):
    m = re.search(r'(\d+|一|二|三|五)年(以上|及以上)?[^。；;\n]{0,30}工作经历', note)
    if not m:
        return True, ''
    if '非应届' in note and FRESH:
        return True, ''
    num = m.group(1)
    need = {'一': 1, '二': 2, '三': 3, '五': 5}.get(num, int(num) if num.isdigit() else 99)
    if WORK_YEARS >= need:
        return True, ''
    return False, '要求%d年工作经历' % need


def age_general_ok():
    if AGE < 18 or AGE > 43:
        return False
    if AGE > 38:
        if not (FRESH == '2027届' and EDU in ('硕士', '博士')):
            return False
    return True


def age_check(note: str, attr: str, level: str, dept: str, title: str):
    if '25周岁' in note and AGE > 25:
        return False, '年龄要求25周岁以下'
    if re.search(r'30周岁以下|不超过30周岁|30周岁及以下|年龄在30', note) and AGE > 30:
        return False, '年龄要求30周岁以下'
    if re.search(r'35周岁以下|35周岁及以下', note) and AGE > 35:
        return False, '年龄要求35周岁以下'
    if '特警' in str(title) and AGE > 25:
        return False, '特警职位年龄25周岁以下'
    relax = FRESH == '2027届' and EDU in ('硕士', '博士') and AGE <= 35
    if str(attr) == '公安机关人民警察职位':
        if str(level) in ('市（地）级', '县（区）级及以下') and AGE > 30 and not relax and '法医' not in str(title):
            return False, '公安市（地）级及以下年龄30周岁以下'
    if ('监狱' in str(dept) or '戒毒' in str(dept)) and '一线' in note and AGE > 30 and not relax:
        return False, '监狱戒毒一线干警年龄30周岁以下'
    return True, ''


def other_check(note: str):
    if re.search(r'限退役军人|仅限退役军人', note) and not VETERAN:
        return False, '限退役军人'
    if '残疾人' in note and '不限' not in note and not DISABLED:
        return False, '残疾人专门职位'
    if '少数民族' in note and '不限' not in note and not MINORITY:
        return False, '限少数民族'
    if '定向' in note and '招录' in note:
        return False, '定向招录'
    if ('运动员' in note or ('前8名' in note and '赛事' in note)) and not ATHLETE:
        return False, '要求运动员等级/赛事名次'
    return True, ''


def evaluate(row, idx):
    g = lambda k: row[idx[k]]
    edu = str(g('学历'))
    deg = str(g('学位'))
    pol = str(g('政治面貌'))
    base_reasons = []
    if not edu_allowed(edu):
        base_reasons.append('学历不符:' + edu)
    if not deg_allowed(deg):
        base_reasons.append('学位不符:' + deg)
    if not pol_allowed(pol):
        base_reasons.append('政治面貌不符:' + pol)
    if not grassroots_ok(str(g('基层工作最低年限'))):
        base_reasons.append('基层年限:' + str(g('基层工作最低年限')))
    if not service_ok(str(g('服务基层项目工作经历'))):
        base_reasons.append('服务基层项目限制')
    attr = str(g('职位属性'))
    if '非通用语' in attr:
        base_reasons.append('非通用语职位')
    if not age_general_ok():
        base_reasons.append('年龄不符(18-38，应届硕博43)')
    note = str(g('备注'))
    dept = str(g('部门名称'))
    title = str(g('招考职位'))
    level = str(g('机构层级'))
    tier = major_tier_for_edu(g('专业'), edu)
    if tier == 'D':
        base_reasons.append('专业不符')
    checks = [
        age_check(note, attr, level, dept, title),
        fresh_check(note),
        gender_check(note),
        english_check(note),
        cert_check(note),
        household_check(note),
        workexp_check(note),
        other_check(note),
    ]
    reason_texts = []
    for ok, reason in checks:
        if not ok:
            base_reasons.append(reason)
        elif reason:
            reason_texts.append(reason)
    if base_reasons:
        return 'red', tier, base_reasons
    flags = []
    if tier == 'A2':
        flags.append('专业为学术学位代码，专业学位需确认')
    if tier == 'B':
        flags.append('专业为相关大类，需确认')
    if DEG in ('硕士', '博士') and deg == '学士':
        flags.append('学位要求学士，硕/博能否报考需确认')
    for r in reason_texts:
        if r and r != '适合女性' and r != '适合男性':
            flags.append(r)
    if '适合女性' in reason_texts or '适合男性' in reason_texts:
        flags.append('性别倾向需确认')
    if '体能测评' in note:
        flags.append('需体能测评')
    if '特殊标准' in note:
        flags.append('体检特殊标准')
    if str(g('职位属性')) == '公安机关人民警察职位':
        flags.append('公安人民警察职位（需体能测评/特殊体检）')
    if flags:
        return 'yellow', tier, flags
    if tier in ('A', 'U', 'C'):
        return 'green', tier, []
    return 'yellow', tier, ['专业匹配需确认']
