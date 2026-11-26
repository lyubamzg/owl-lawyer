# -*- coding: utf-8 -*-
"""Обновить базу судебных участков России для «Совы».

Собираются два открытых источника:
  * ГАС РФ «Правосудие» (sudrf.ru) — классификационный код (ВНКО) и официальное
    название каждого мирового судебного участка;
  * mir-sud.ru — описание территории участка (поселения, улицы, границы).

Результат: data/sudebnye_uchastki.json — его читает sudrf_base.py, когда живая
ГАС «Правосудие» не отвечает и суд в шапке иначе пришлось бы оставить тот,
что указан вручную (один участок на весь район).

Запуск:  python3 update_sudebnye_uchastki.py
"""
from __future__ import annotations

import difflib
import json
import re
import sys
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA = BASE / 'data'
OUT = DATA / 'sudebnye_uchastki.json'

UA = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/124.0.0.0 Safari/537.36 OwlLawyer/1.0')
MIR = 'https://mir-sud.ru'
SUDRF = 'https://sudrf.ru'
NUM_RE = re.compile(r'№\s*(\d+)')
WORD_RE = re.compile(r'[а-яёa-z]{4,}')


def get(url, tries=3, encoding='utf-8'):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': UA,
                                                       'Referer': SUDRF + '/index.php?id=300'})
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.read().decode(encoding, 'replace')
        except Exception as ex:  # noqa: BLE001
            if i == tries - 1:
                print('  !! не удалось:', url, ex)
                return ''
            time.sleep(1.2 * (i + 1))


def norm(s):
    s = re.sub(r'\(.*?\)', ' ', str(s or ''))
    s = s.replace('ё', 'е')
    s = re.sub(r'[^а-яa-z0-9]+', ' ', s.lower())
    return re.sub(r'\s+', ' ', s).strip()


def region_key(name):
    n = norm(name)
    n = re.sub(r'\b(республика|республике|область|области|край|края|округ|автономный)\b', ' ', n)
    n = re.sub(r'\s+-\s+.*$', ' ', n)
    return re.sub(r'\s+', ' ', n).strip()


def tokens(s):
    return set(WORD_RE.findall(norm(s)))


def plain(html):
    t = re.sub(r'<(script|style).*?</\1>', ' ', html, flags=re.S | re.I)
    t = re.sub(r'<[^>]+>', ' ', t)
    return re.sub(r'\s+', ' ', t)


# ------------------------------------------------------ 1. ГАС «Правосудие» ----
def collect_sudrf():
    form = get(f'{SUDRF}/index.php?act=go_ms_search&id=300&ms_type=ms&searchtype=ms&var=true',
               encoding='cp1251')
    regions = [(c, n.strip()) for c, n in re.findall(r"<option value='(\d+)'[^>]*>([^<]+)</option>", form)
               if c.isdigit() and c != '0']
    print('регионов в ГАС:', len(regions))
    out = {}

    def work(item):
        code, name = item
        h = get(f'{SUDRF}/index.php?act=go_ms_search&court_subj={code}&id=300'
                f'&ms_type=ms&searchtype=ms&var=true', encoding='cp1251')
        rows = []
        for b in re.split(r'(?=listcontrol\(\d+,"\d{2}MS\d{4}")', h):
            m = re.match(r'listcontrol\(\d+,"(\d{2}MS\d{4})"\);\'>\s*([^<]+?)\s*</a>', b)
            if not m:
                continue
            site = re.search(r"href='(http[^']+)'", b)
            rows.append({'code': m.group(1), 'name': m.group(2).strip(),
                         'site': site.group(1) if site else ''})
        return code, name, rows

    with ThreadPoolExecutor(max_workers=6) as ex:
        for code, name, rows in ex.map(work, regions):
            out[code] = {'region': name, 'participants': rows}
    print('участков в ГАС:', sum(len(v['participants']) for v in out.values()))
    return out


# ------------------------------------------------------ 2. mir-sud: территории ----
def collect_mirsud():
    root = get(MIR + '/regiony-rf/')
    slugs = sorted(set(re.findall(r'href="(/regiony-rf/[^"/]+/)"', root)))
    print('регионов на mir-sud:', len(slugs))
    courts = []
    for slug in slugs:
        h = get(MIR + slug)
        links = sorted(set(re.findall(r'href="(/regiony-rf/[^"#]+/district/[^"#]+/)"', h)))
        m = re.search(r'<h1[^>]*>([^<]{3,160})</h1>', h)
        name = re.sub(r'\s+', ' ', m.group(1)).strip() if m else slug
        for c in links:
            courts.append((c, name, slug))
    print('страниц судов:', len(courts))

    def parse_court(html, region, court_name, url):
        t = plain(html)
        i = t.find('Территориальная подсудность')
        if i < 0:
            return []
        j = t.find('Источник импортированного', i)
        if j < 0:
            j = t.find('Навигация', i)
        seg = t[i:j if j > i else i + 40000]
        out = []
        for p in re.split(r'(?=Мировой судебный участок\s*№\s*\d+)', seg):
            m = re.match(r'Мировой судебный участок\s*№\s*(\d+)\s*(.*)', p)
            if not m:
                continue
            out.append({'num': int(m.group(1)), 'desc': p.strip(), 'region': region,
                        'court': court_name, 'url': url})
        return out

    rows = []
    done = [0]

    def work(t):
        url, region, court_slug = t
        html = get(MIR + url)
        res = parse_court(html, region, court_slug, MIR + url)
        if not res:
            html = get(MIR + url)
            res = parse_court(html, region, court_slug, MIR + url)
        done[0] += 1
        if done[0] % 200 == 0:
            print(f'  {done[0]}/{len(courts)}', flush=True)
        return res

    with ThreadPoolExecutor(max_workers=8) as ex:
        for res in ex.map(work, courts):
            rows.extend(res)
    print('участков с описанием территории:', len(rows))
    return rows


# --------------------------------------------------------------- сборка ----
def build():
    sudrf = collect_sudrf()
    mirsud = collect_mirsud()

    regions, by_key, parts = {}, {}, {}
    for code, info in sudrf.items():
        rk = region_key(info['region'])
        by_key.setdefault(rk, code)
        regions.setdefault(code, info['region'])
        for p in info['participants']:
            m = NUM_RE.search(p['name'])
            if m and p['code']:
                parts[p['code']] = {'code': p['code'], 'num': int(m.group(1)), 'name': p['name'],
                                    'region': info['region'], 'region_code': code,
                                    'site': p.get('site', ''), 'desc': '', 'court': ''}

    def find_region(name):
        rk = region_key(name)
        if rk in by_key:
            return by_key[rk]
        cand = difflib.get_close_matches(rk, list(by_key), n=1, cutoff=0.72)
        return by_key[cand[0]] if cand else None

    by_region_num = {}
    for p in parts.values():
        by_region_num.setdefault((p['region_code'], p['num']), []).append(p)

    used = set()
    for row in mirsud:
        code = find_region(row['region'])
        if not code:
            continue
        cands = by_region_num.get((code, row['num']), [])
        if not cands:
            continue
        pick = cands[0] if len(cands) == 1 else None
        if pick is None:
            ct = tokens(row.get('court') or '')
            best, best_score = None, 0
            for c in cands:
                sc = len(ct & tokens(c['name']))
                if sc > best_score:
                    best, best_score = c, sc
            pick = best if best_score else None
        if not pick or pick['code'] in used:
            continue
        used.add(pick['code'])
        pick['desc'] = row.get('desc') or ''
        pick['court'] = row.get('court') or ''

    participants = sorted(parts.values(), key=lambda p: (p['region_code'], p['code']))
    DATA.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({
        'built': time.strftime('%Y-%m-%d'),
        'sources': [
            'ГАС РФ «Правосудие» (sudrf.ru) — классификационный код, официальное название, сайт',
            'mir-sud.ru — описание территории участка (поселения, улицы, границы) и районный суд',
        ],
        'note': 'Живая ГАС «Правосудие» — основной источник; эта база — резерв, когда сайт не отвечает.',
        'participants': participants,
    }, ensure_ascii=False, indent=1), encoding='utf-8')
    print('участков в базе:', len(participants),
          '| с описанием территории:', sum(1 for p in participants if p['desc']))
    print('файл:', OUT, '|', round(OUT.stat().st_size / 1024), 'КБ')


if __name__ == '__main__':
    build()
