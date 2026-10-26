"""健保署特約診所、藥局名單（政府開放資料，每天更新）→ 只留北屯區、潭子區還在特約的，存成 JSON 給 LINE 小幫手「全全」查。

里長 10/26「醫療部分有辦法連動嗎」：Google（Apps Script）的伺服器連不到健保署網站，改由 GitHub Actions 每天抓一次，
放在 medical-data 分支，Apps Script 再讀 raw.githubusercontent.com 上的這個檔案。
資料來源：data.gov.tw 資料集 39283（診所）、39284（藥局），政府資料開放授權條款第 1 版。
"""
import csv
import io
import json
import sys
import urllib.request
from datetime import datetime, timedelta, timezone

SOURCES = {
    'clinics': 'https://info.nhi.gov.tw/api/iode0000s01/Dataset?rId=A21030000I-D21004-009',
    'pharmacies': 'https://info.nhi.gov.tw/api/iode0000s01/Dataset?rId=A21030000I-D21005-001',
}
AREAS = ('北屯區', '潭子區')


def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (beitun-buzi line bot; open data)'})
    with urllib.request.urlopen(req, timeout=180) as res:
        raw = res.read()
    for enc in ('utf-8-sig', 'big5', 'cp950'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    raise RuntimeError('看不懂的編碼：' + url)


def rows(text):
    return list(csv.DictReader(io.StringIO(text.lstrip('\ufeff'))))


def keep(r):
    addr = (r.get('地址') or '').replace('台中市', '臺中市')
    return addr.startswith('臺中市') and any(a in addr for a in AREAS) and not (r.get('終止合約或歇業日期') or '').strip()


def slim(r):
    return {
        'id': (r.get('醫事機構代碼') or '').strip(),
        'name': (r.get('醫事機構名稱') or '').strip(),
        'kind': (r.get('醫事機構種類') or '').strip(),
        'tel': (r.get('電話') or '').strip(),
        'addr': (r.get('地址') or '').strip().replace('台中市', '臺中市'),
        'services': (r.get('服務項目') or '').strip(),
        'depts': (r.get('診療科別') or '').strip(),
        'hours': (r.get('固定看診時段') or '').strip(),
        'note': (r.get('備註') or '').strip(),
    }


def main(out_path):
    out = {'source': '衛生福利部中央健康保險署 健保特約醫事機構（政府資料開放平臺 39283、39284）',
           'updated': datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M'), 'areas': list(AREAS)}
    for key, url in SOURCES.items():
        data = rows(fetch(url))
        if not data or '醫事機構名稱' not in data[0]:
            raise RuntimeError(f'{key} 欄位不對：{list(data[0].keys()) if data else "沒有資料"}')
        out[key] = [slim(r) for r in data if keep(r)]
        print(f'{key}: 全國 {len(data)} 筆，{"、".join(AREAS)} {len(out[key])} 筆')
        if out[key]:
            print('  例：', json.dumps(out[key][0], ensure_ascii=False)[:600])
    if len(out['clinics']) < 50:
        raise RuntimeError('診所太少，可能抓錯資料，不更新')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'beitun-medical.json')
