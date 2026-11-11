#!/usr/bin/env python3
"""
สร้างไฟล์หน้าเว็บใหม่จากไฟล์ Excel ของ AppSheet
ไม่เก็บชื่อสถานที่และพิกัดลงในไฟล์

ติดตั้งครั้งเดียว:   pip install pandas openpyxl

วิธีใช้:
    python make_site.py ไฟล์ข้อมูล.xlsx index.html ผลลัพธ์.html

    ไฟล์ข้อมูล.xlsx  ไฟล์ Excel ข้อมูลเคสล่าสุด
    index.html       ไฟล์หน้าเว็บปัจจุบัน ใช้เป็นต้นแบบ โค้ดและหน้าตาจะไม่เปลี่ยน
    ผลลัพธ์.html     ไฟล์ใหม่ที่ใช้อัปโหลดขึ้น GitHub แทนไฟล์เดิม
"""
import sys
import re
import json
import gzip
import base64
import pandas as pd

SEV = {'แดง': 'r', 'เหลือง': 'y', 'เขียว': 'g'}
BLOB = re.compile(r'(<div id="data-blob" style="display:none">)([A-Za-z0-9+/=]+)(</div>)')


def clean(s):
    if not isinstance(s, str):
        return ''
    s = s.replace('\\n', '\n').replace('\r\n', '\n').replace('\r', '\n')
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()


def dic(vals):
    items, index, out = [], {}, []
    for v in vals:
        if v not in index:
            index[v] = len(items)
            items.append(v)
        out.append(index[v])
    return items, out


def build(xlsx):
    df = pd.read_excel(xlsx, dtype=str)
    total = len(df)
    df = df.dropna(subset=['ระดับความรุนแรง', 'หมวดหลัก', 'รายละเอียดของเหตุการณ์'], how='all')
    junk = (df['หมวดหลัก'] == 'หมวดหลัก') | (df['ระดับความรุนแรง'] == 'รอยืนยันสถานะ')
    df = df[~junk].reset_index(drop=True)

    unknown = set(df['ระดับความรุนแรง']) - set(SEV)
    if unknown:
        raise SystemExit(f'พบระดับความรุนแรงที่ไม่รู้จัก: {unknown}')

    cats, cat_idx = dic(df['หมวดหลัก'])
    subs, sub_idx = dic(df['หมวดรอง'].fillna(''))

    rows = []
    for k in range(len(df)):
        rows.append([
            SEV[df.loc[k, 'ระดับความรุนแรง']],
            cat_idx[k],
            sub_idx[k],
            0,  # place index 0 = no place (locations are not stored)
            1 if str(df.loc[k, 'สถานะดำเนินการ']).startswith('2') else 0,
            clean(df.loc[k, 'รายละเอียดของเหตุการณ์']),
            clean(df.loc[k, 'แนวทาง และผลดำเนินการแก้ไข']),
            '',  # no row-level location
        ])

    print(f'อ่านได้ {len(rows)} เคส (ตัดแถวว่างหรือแถวหัวตารางออก {total - len(rows)} แถว)')
    data = {'c': cats, 's': subs, 'p': [['ALL', '', '']], 'r': rows}
    raw = json.dumps(data, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    return base64.b64encode(gzip.compress(raw, 9)).decode('ascii')


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(1)
    xlsx, template, out = sys.argv[1:]
    blob = build(xlsx)
    html = open(template, encoding='utf-8').read()
    m = BLOB.search(html)
    if not m:
        raise SystemExit('ไม่พบส่วนข้อมูลในไฟล์ index.html ที่ใช้เป็นต้นแบบ')
    html = html[:m.start(2)] + blob + html[m.end(2):]
    open(out, 'w', encoding='utf-8').write(html)
    print('เขียนไฟล์แล้ว:', out)


if __name__ == '__main__':
    main()
