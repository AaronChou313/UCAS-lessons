# -*- coding: utf-8 -*-
"""Inject merged COURSES JSON into the HTML template."""
import json, re, sys

tpl = open('template.html', encoding='utf-8').read()
courses = json.load(open('courses_merged.json', encoding='utf-8'))

assert '/*__COURSES__*/' in tpl, 'placeholder missing'
blob = json.dumps(courses, ensure_ascii=False, separators=(',', ':'))
out = tpl.replace('/*__COURSES__*/', 'const COURSES=' + blob + ';')

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
