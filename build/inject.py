# -*- coding: utf-8 -*-
"""Inject merged COURSES JSON into the HTML template."""
import json, re, sys, os

tpl = open('template.html', encoding='utf-8').read()
courses = json.load(open('courses_merged.json', encoding='utf-8'))

assert '/*__COURSES__*/' in tpl, 'placeholder missing'
blob = json.dumps(courses, ensure_ascii=False, separators=(',', ':'))
out = tpl.replace('/*__COURSES__*/', 'const COURSES=' + blob + ';')

# 注入课程大纲链接映射（编码 -> syllabus_url），来自 syllabi/course_syllabi.json
syl_path = '../syllabi/course_syllabi.json'
if os.path.exists(syl_path):
    syl = json.load(open(syl_path, encoding='utf-8'))
    urlmap = {s['match_code']: s['syllabus_url'] for s in syl['syllabi'] if s.get('match_code') and s.get('syllabus_url')}
    syl_blob = json.dumps(urlmap, ensure_ascii=False, separators=(',', ':'))
    print(f'injected syllabus links: {len(urlmap)}')
else:
    syl_blob = '{}'
out = out.replace('/*__SYLLABI__*/{}', syl_blob)

# update hero stats from actual data
n_college = len(set(c['college'] for c in courses))
n_first = len(set(c['first'] for c in courses))
n_cat = len(set(c['category'] for c in courses))
n_total = len(courses)
out = re.sub(r'<div class="stat"><strong>\d+</strong><span>课程记录</span>', f'<div class="stat"><strong>{n_total}</strong><span>课程记录</span>', out)
out = re.sub(r'<div class="stat"><strong>\d+</strong><span>开课院系</span>', f'<div class="stat"><strong>{n_college}</strong><span>开课院系</span>', out)
out = re.sub(r'<div class="stat"><strong>\d+</strong><span>一级学科分类</span>', f'<div class="stat"><strong>{n_first}</strong><span>一级学科分类</span>', out)
out = re.sub(r'<div class="stat"><strong>\d+</strong><span>课程类别</span>', f'<div class="stat"><strong>{n_cat}</strong><span>课程类别</span>', out)

dest = sys.argv[1] if len(sys.argv) > 1 else '../index.html'
open(dest, 'w', encoding='utf-8').write(out)
print(f'OK: {n_total} courses, {n_college} colleges, {n_first} firsts, {n_cat} categories -> {dest} ({len(out)/1024:.0f} KB)')
