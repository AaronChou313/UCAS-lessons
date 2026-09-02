# -*- coding: utf-8 -*-
"""
国科大 2026-2027 学年(秋)第一学期 课程大纲采集脚本
- 官方入口: https://jwba.ucas.ac.cn/sc/public/coursePublic
- 大纲页:   https://jwba.ucas.ac.cn/sc/course/courseplan/{id}
- 特点: 缓存 / 断点续跑 / 低频请求 / 限流退避
"""
import urllib.request, urllib.parse, re, json, os, time, sys, random, csv
from datetime import datetime

BASE = "https://jwba.ucas.ac.cn"
RAW_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "raw")
os.makedirs(RAW_DIR, exist_ok=True)

UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120 Safari/537.36"}

# ---------------- 解析单个大纲页 ----------------
def parse_syllabus(html):
    html = re.sub(r'<script.*?</script>', '', html, flags=re.S)
    html = re.sub(r'<style.*?</style>', '', html, flags=re.S)
    matches = list(re.finditer(r'<strong[^>]*>(.*?)</strong>', html, re.S))
    labels = []
    for m in matches:
        lbl = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        if lbl:
            labels.append((lbl, m.start(), m.end()))
    def text_between(start, end):
        seg = html[start:end]
        seg = re.sub(r'<br\s*/?>', '\n', seg)
        seg = re.sub(r'</p>|</div>|</tr>|</li>', '\n', seg)
        seg = re.sub(r'<[^>]+>', '', seg)
        lines = [re.sub(r'[ \t ]+', ' ', l).strip() for l in seg.split('\n')]
        return '\n'.join(l for l in lines if l)
    result = {}
    header_map = {'课程编码：': 'code', '英文名称：': 'en', '课时：': 'hours',
                  '学分：': 'credits', '课程属性：': 'category', '主讲教师：': 'teacher'}
    section_map = {'教学目的要求': 'objectives', '预修课程': 'prerequisites',
                   '大纲内容': 'outline', '教材信息': 'textbook', '参考书': 'references',
                   '课程教师信息': 'teacher_info'}
    name_cand = None
    for i, (lbl, s, e) in enumerate(labels):
        nxt = labels[i+1][1] if i+1 < len(labels) else len(html)
        if name_cand is None and lbl not in header_map and not lbl.endswith('：') and lbl not in section_map:
            name_cand = lbl
        if lbl in header_map:
            result[header_map[lbl]] = text_between(e, nxt).split('\n')[0].strip()
        elif lbl in section_map:
            result[section_map[lbl]] = text_between(e, nxt)
    if name_cand:
        result['name'] = name_cand
    return result

# ---------------- 下载（缓存 + 断点 + 退避） ----------------
def fetch(plan_id, retries=3):
    path = os.path.join(RAW_DIR, f"{plan_id}.html")
    if os.path.exists(path) and os.path.getsize(path) > 500:
        return open(path, encoding='utf-8').read(), True  # cached
    url = f"{BASE}/sc/course/courseplan/{plan_id}"
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=20) as r:
                html = r.read().decode('utf-8', 'ignore')
            if len(html) > 500 and '课程大纲' in html:
                open(path, 'w', encoding='utf-8').write(html)
                return html, False
            # 内容过小/异常，视为失败
            print(f"  [warn] {plan_id} 内容异常 size={len(html)}", flush=True)
        except Exception as e:
            wait = min(30, 2 ** (attempt + 1)) + random.random()
            print(f"  [retry {attempt+1}] {plan_id}: {str(e)[:60]} 等待{wait:.0f}s", flush=True)
            time.sleep(wait)
    return None, False

# ---------------- 主流程 ----------------
def main():
    mres = json.load(open('/tmp/match_result.json', encoding='utf-8'))
    matched = mres['matched']  # [{code,name,plan_id,catalog_name}]
    print(f"待下载大纲: {len(matched)}", flush=True)

    crawl_ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    results = []
    failures = []
    done = 0
    for idx, item in enumerate(matched, 1):
        html, cached = fetch(item['plan_id'])
        if html is None:
            failures.append({'code': item['code'], 'name': item['name'],
                             'plan_id': item['plan_id'], 'reason': 'download_failed'})
            print(f"[{idx}/{len(matched)}] FAIL {item['code']}", flush=True)
        else:
            data = parse_syllabus(html)
            data['plan_id'] = item['plan_id']
            data['syllabus_url'] = f"{BASE}/sc/course/courseplan/{item['plan_id']}"
            data['match_code'] = item['code']          # Excel 侧课程编码
            data['match_name'] = item['name']          # Excel 侧课程名
            data['catalog_name'] = item['catalog_name']  # 目录侧课程名
            data['crawl_time'] = crawl_ts
            results.append(data)
            done += 1
            if idx % 50 == 0 or idx == len(matched):
                print(f"[{idx}/{len(matched)}] 成功 {done} 失败 {len(failures)}", flush=True)
        # 低频：0.3s + 随机抖动
        time.sleep(0.3 + random.random() * 0.2)

    # ---------------- 输出 ----------------
    base = os.path.dirname(os.path.abspath(__file__))
    meta = {
        'source': f"{BASE}/sc/public/coursePublic",
        'term': '2026—2027学年(秋)第一学期',
        'termId': '89576',
        'crawl_time': crawl_ts,
        'match_method': 'exact_full_code (Excel完整课程编码精确匹配目录课程编号)',
        'total_matched': len(matched),
        'downloaded_ok': len(results),
        'failed': len(failures),
    }
    json.dump({'meta': meta, 'syllabi': results}, open(os.path.join(base, 'course_syllabi.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)

    # CSV 索引
    by_plan = {c['plan_id']: c for c in json.load(open('/tmp/catalog.json', encoding='utf-8'))}
    with open(os.path.join(base, 'course_syllabi.csv'), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['课程编码', '课程名称', '目录课程名', '开课院系', '校区', '大纲链接', '课时', '学分', '课程属性', '主讲教师', '采集时间'])
        for d in results:
            cat = by_plan.get(d['plan_id'], {})
            w.writerow([d['match_code'], d['match_name'], d['catalog_name'], cat.get('dept', ''),
                        cat.get('campus', ''), d['syllabus_url'], d.get('hours', ''), d.get('credits', ''),
                        d.get('category', ''), d.get('teacher', ''), d['crawl_time']])

    # 未匹配/失败
    with open(os.path.join(base, 'unmatched.csv'), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['课程编码', '课程名称', '原因'])
        for u in mres['unmatched']:
            w.writerow([u['code'], u['name'], u['reason']])
        for fl in failures:
            w.writerow([fl['code'], fl['name'], fl['reason']])

    print(f"\n完成: 成功 {len(results)} / {len(matched)}，失败 {len(failures)}")
    print(f"输出: course_syllabi.json / course_syllabi.csv / unmatched.csv")

if __name__ == '__main__':
    main()
