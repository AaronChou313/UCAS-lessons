# -*- coding: utf-8 -*-
"""Extract official Hangzhou courses (code prefix 280216) from a UCAS export."""
import argparse
import json
import re
from collections import Counter, defaultdict

import openpyxl


PREFIX = '280216'
HTML_FILE = '../index.html'
OUT = 'hangzhou_courses.json'


def parse_slot(value):
    match = re.match(r'周([一二三四五六日天])\((.+)\)', str(value or '').strip())
    if not match:
        return []
    day = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '日': 7, '天': 7}[match.group(1)]
    slots = []
    for part in re.split(r'[,，]', match.group(2)):
        nums = [int(n) for n in re.findall(r'\d+', part)]
        if not nums:
            continue
        start, end = nums[0], nums[-1]
        slots.append((day, start, end))
    return slots


def subject_map():
    html = open(HTML_FILE, encoding='utf-8').read()
    match = re.search(r'const COURSES=(\[.*?\]);', html, re.S)
    courses = json.loads(match.group(1))
    counts = defaultdict(lambda: defaultdict(Counter))
    for course in courses:
        counts[course['subjectCode']][course['first']][course['second']] += 1
    result = {}
    for code, firsts in counts.items():
        first = max(firsts, key=lambda value: sum(firsts[value].values()))
        result[code] = (first, firsts[first].most_common(1)[0][0])
    return result


def main():
    parser = argparse.ArgumentParser(description='Extract official Hangzhou courses from a UCAS export')
    parser.add_argument('xlsx', help='path to course_plan_yjs.xlsx')
    parser.add_argument('--semester', choices=['秋季', '春季'], default='秋季')
    args = parser.parse_args()

    workbook = openpyxl.load_workbook(args.xlsx, read_only=True, data_only=True)
    mappings = subject_map()
    groups = []
    current = None
    for worksheet in workbook.worksheets:
        for values in worksheet.iter_rows(min_row=2, values_only=True):
            code = str(values[2] or '').strip()
            if values[0] is not None:
                current = None
                if code.startswith(PREFIX):
                    current = {'anchor': values, 'slots': []}
                    groups.append(current)
            elif current is not None:
                current['slots'].append(values)

    courses = []
    for group in groups:
        values = group['anchor']
        code = str(values[2]).strip()
        discipline = str(values[8] or '').strip()
        subject_code = code[6:12]
        first, second = mappings.get(subject_code, (discipline or '其他 / 自设学科', discipline or '自设课程'))
        hours, credits = [float(value) for value in str(values[9]).split('/')]
        slots = []
        rooms = []
        for row in [values] + group['slots']:
            weeks = str(row[12] or '').strip()
            for day, start, end in parse_slot(row[13]):
                slots.append({'day': day, 'p': str(start) if start == end else f'{start}-{end}',
                              'start': start, 'end': end, 'w': weeks})
            room = str(row[14] or '').strip()
            if room and room not in rooms:
                rooms.append(room)

        courses.append({
            'id': 0,
            'code': code, 'name': str(values[3]).strip(), 'en': str(values[4] or '').strip(),
            'college': str(values[1]).strip(), 'campus': '杭州', 'semester': args.semester,
            'category': str(values[6]).strip(), 'discipline': discipline, 'subjectCode': subject_code,
            'first': first, 'second': second, 'level': str(values[7] or '').strip(),
            'hours': hours, 'credits': credits,
            'capacity': float(values[10]) if values[10] not in (None, '', '/') else None,
            'enrolled': int(float(values[11])) if values[11] not in (None, '', '/') else 0,
            'exam': str(values[16] or '').strip(), 'teach': str(values[15] or '').strip(),
            'chief': str(values[17] or '').strip(), 'chiefUnit': str(values[18] or '').strip(),
            'main': str(values[19] or '').strip(), 'mainUnit': str(values[20] or '').strip(),
            'ta': str(values[21] or '').strip(), 'taUnit': str(values[22] or '').strip(),
            'convener': str(values[23] or '').strip(), 'room': ' / '.join(rooms), 'slots': slots,
        })

    if not courses:
        raise RuntimeError('no Hangzhou courses found')
    existing = []
    try:
        existing = json.load(open(OUT, encoding='utf-8'))
    except FileNotFoundError:
        pass
    merged = [course for course in existing if course.get('semester') != args.semester] + courses
    json.dump(merged, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print(f'written {OUT}: {len(courses)} {args.semester} courses, {len(merged)} total')


if __name__ == '__main__':
    main()
