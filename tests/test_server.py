# -*- coding: utf-8 -*-
"""
검증 3 — 백엔드 HTTP 종단 테스트
서버를 직접 띄우고 /api/config, /api/calc, /api/generate 를 호출해
웹 UI 가 하는 것과 동일한 경로로 .xls 가 내려오는지 확인한다.
"""
import io
import os
import sys
import json
import time
import struct
import threading
import urllib.request
import urllib.error
import urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'backend'))

import server as srvmod                                  # noqa: E402
from http.server import ThreadingHTTPServer              # noqa: E402
from xlsbiff import Cfb                                  # noqa: E402

FAILS = []


def check(label, cond, detail=''):
    print('  [%s] %s%s' % ('OK' if cond else '실패', label,
                           ('  — ' + detail) if detail else ''))
    if not cond:
        FAILS.append(label)


PAYLOAD = {
    "quoteNo": "DSS-2026-0823-99",
    "quoteDate": "2026-08-23",
    "customer": "㈜서버테스트",
    "contact": "이영희 대리",
    "subject": "HTTP 종단 테스트 견적",
    "remark": "테스트",
    "vatRate": 0.1,
    "items": [
        {"name": "품목 A", "spec": "SPEC-1", "unit": "ea", "qty": 3, "price": 100000},
        {"name": "품목 B", "spec": "SPEC-2", "unit": "set", "qty": 1, "price": 250000},
    ],
}


def post(url, obj):
    req = urllib.request.Request(
        url, data=json.dumps(obj).encode('utf-8'),
        headers={'Content-Type': 'application/json'})
    return urllib.request.urlopen(req, timeout=20)


def main():
    port = srvmod.free_port(8791)
    httpd = ThreadingHTTPServer(('127.0.0.1', port), srvmod.Handler)
    t = threading.Thread(target=httpd.serve_forever, daemon=True)
    t.start()
    base = 'http://127.0.0.1:%d' % port
    print('테스트 서버: %s' % base)
    time.sleep(0.3)

    try:
        # /api/config
        cfg = json.loads(urllib.request.urlopen(base + '/api/config', timeout=10)
                         .read().decode('utf-8'))
        check('GET /api/config', cfg['maxItems'] == 36 and cfg['itemRows'] == [23, 58],
              'maxItems=%s itemRows=%s' % (cfg['maxItems'], cfg['itemRows']))

        # 정적 파일
        for p in ('/', '/app.js', '/style.css'):
            r = urllib.request.urlopen(base + p, timeout=10)
            check('GET %s' % p, r.status == 200 and len(r.read()) > 100)

        # /api/calc
        r = post(base + '/api/calc', PAYLOAD)
        c = json.loads(r.read().decode('utf-8'))
        check('POST /api/calc 합계 계산',
              c['supply'] == 550000 and c['vat'] == 55000 and c['total'] == 605000,
              json.dumps(c['supply']) + '/' + json.dumps(c['vat']) + '/' + json.dumps(c['total']))

        # /api/generate
        r = post(base + '/api/generate', PAYLOAD)
        blob = r.read()
        cd = r.headers.get('Content-Disposition') or ''
        check('POST /api/generate 200 + xls 바이트', r.status == 200 and len(blob) > 50000,
              '%d bytes' % len(blob))
        check('OLE2 시그니처', blob[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
        check('Content-Type = application/vnd.ms-excel',
              r.headers.get('Content-Type') == 'application/vnd.ms-excel')
        m = urllib.parse.unquote(cd.split("filename*=UTF-8''")[-1]) if "UTF-8''" in cd else ''
        check('한글 파일명 헤더', m.startswith('견적서_20260823_'), m)
        check('합계 헤더', r.headers.get('X-Quote-Total') == '605000',
              str(r.headers.get('X-Quote-Total')))

        # 내려온 바이트가 실제로 열리는지
        wbdata = Cfb(blob).streams['Workbook'][1]
        check('내려받은 파일에서 Workbook 스트림 추출', len(wbdata) > 60000,
              '%d bytes' % len(wbdata))
        import xlrd
        tmp = os.path.join(HERE, '..', 'out', '_server_test.xls')
        open(tmp, 'wb').write(blob)
        sh = xlrd.open_workbook(tmp, formatting_info=True).sheet_by_index(0)
        check('xlrd 로 열기 + C14 공급처', sh.cell_value(13, 2) == PAYLOAD['customer'],
              repr(sh.cell_value(13, 2)))
        check('H62 총액', abs(sh.cell_value(61, 7) - 605000) < 0.5,
              str(sh.cell_value(61, 7)))
        os.remove(tmp)

        # 빈 행이 섞여도 기재한 품목은 전부 나와야 한다
        mixed = dict(PAYLOAD)
        mixed['items'] = [
            {"name": "품목1", "spec": "S1", "unit": "ea", "qty": 1, "price": 1000},
            {"name": "", "spec": "", "unit": "", "qty": "", "price": ""},      # 빈 행
            {"name": "품목2", "spec": "S2", "unit": "ea", "qty": 2, "price": 2000},
            {"name": "", "spec": "", "unit": "", "qty": 0, "price": 0},        # 빈 행
            {"name": "", "spec": "규격만 있는 행", "unit": "", "qty": 1, "price": 500},
        ]
        r = post(base + '/api/generate', mixed)
        blob2 = r.read()
        check('빈 행이 섞여도 200 (다운로드가 막히지 않음)', r.status == 200)
        check('기재한 3행만 기록됨', r.headers.get('X-Quote-Items') == '3',
              str(r.headers.get('X-Quote-Items')))
        tmp2 = os.path.join(HERE, '..', 'out', '_server_mixed.xls')
        open(tmp2, 'wb').write(blob2)
        sh2 = xlrd.open_workbook(tmp2).sheet_by_index(0)
        rows = [(sh2.cell_value(r0, 2), sh2.cell_value(r0, 3)) for r0 in (22, 23, 24)]
        check('빈 행이 빠지고 23·24·25행에 연속 배치',
              rows == [('품목1', 'S1'), ('품목2', 'S2'), ('', '규격만 있는 행')], str(rows))
        check('26행부터는 비어 있음', sh2.cell_value(25, 2) == '')
        check('합계 = 1000 + 4000 + 500', abs(sh2.cell_value(59, 7) - 5500) < 0.5,
              str(sh2.cell_value(59, 7)))
        os.remove(tmp2)

        # 공급처가 비어도 다운로드는 가능해야 한다
        nocust = dict(PAYLOAD); nocust['customer'] = ''
        r = post(base + '/api/generate', nocust)
        check('공급처 미입력이어도 200 (더 이상 필수 아님)', r.status == 200 and len(r.read()) > 50000)

        # 유효성 검사 (에러 경로)
        bad = dict(PAYLOAD); bad['items'] = []
        try:
            post(base + '/api/generate', bad)
            check('품목 0개 → 400 반환', False)
        except urllib.error.HTTPError as e:
            body = json.loads(e.read().decode('utf-8'))
            check('품목 0개 → 400 반환', e.code == 400 and '품목' in body['error'],
                  body['error'])

        bad2 = dict(PAYLOAD)
        bad2['items'] = [{'name': 'x', 'qty': 1, 'price': 1}] * 37
        try:
            post(base + '/api/generate', bad2)
            check('품목 37개 → 400 반환', False)
        except urllib.error.HTTPError as e:
            body = json.loads(e.read().decode('utf-8'))
            check('품목 37개(최대 36 초과) → 400 반환', e.code == 400, body['error'])

    finally:
        httpd.shutdown()

    print()
    if FAILS:
        print('결과: FAIL — %d건' % len(FAILS))
        for f in FAILS:
            print('   - ' + f)
        return 1
    print('결과: PASS — 전 항목 통과')
    return 0


if __name__ == '__main__':
    sys.exit(main())
