# Doc cac file .xlsx trong thu muc data/ va tao www/data.js
import glob, json, re, openpyxl
def s(v): return '' if v is None else str(v).strip()
D = {}
for f in sorted(glob.glob('data/*.xlsx')):
    wb = openpyxl.load_workbook(f, data_only=True)
    for ws in wb:
        m = re.search(r'N([1-5])', ws.title)
        if not m: continue
        lv = 'n' + m.group(1)
        rows = list(ws.iter_rows(values_only=True))
        if not rows: continue
        head = [s(c) for c in rows[0]]
        body = [[s(c) for c in r] for r in rows[1:]]
        body = [r + [''] * 8 for r in body if any(r[1:])]
        t = ws.title
        if 'Kanji' in t:
            key, kind, items = lv+'k', 'kanji', [dict(k=r[1], v=r[2], on=r[3], kun=r[4].replace('\n', ' / '), ex=r[5].split('\n'), m=r[6].split('\n')) for r in body if r[1]]
            title = 'Kanji ' + lv.upper()
        elif 'vựng' in t:
            key, kind, title = lv+'v', 'vocab', 'Từ vựng ' + lv.upper()
            if 'Phiên âm' in head: items = [dict(w=r[1], r=r[2], m=r[3], t=r[4]) for r in body]
            elif 'フリガナ' in head:   # N3 tro len: moi so thu tu la 1 the, cac dong duoi la tu di kem
                kind, items = 'vocabg', []
                for r in body:
                    if not r[1]: continue
                    if r[0] or not items: items.append(dict(n=r[0], w=r[1], r=r[2], m=r[3], c=[]))
                    else: items[-1]['c'].append(dict(w=r[1], r=r[2], m=r[3]))
            else: items = [dict(w=r[1], r='', m=r[2], t=r[3]) for r in body]
            items = [i for i in items if i['w']]
        elif 'pháp' in t:
            key, kind, title = lv+'g', 'grammar', 'Ngữ pháp ' + lv.upper()
            if '使う方' in head:   # kieu N3: moi muc gom nhieu dong
                items = []
                for r in body:
                    if r[0]: items.append(dict(g=r[1], m=r[2], n='', d=[]))
                    elif items and (r[1] or r[2]): items[-1]['d'].append(dict(f=r[1], t=r[2]))
            else: items = [dict(g=r[1], m=r[2], n=r[3], d=[]) for r in body if r[1]]
        else: continue
        if items: D[key] = dict(t=title, k=kind, i=items)
        print(key, len(items))
open('www/data.js', 'w', encoding='utf-8').write('window.DATA=' + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + ';')
