# -*- coding: utf-8 -*-
"""
검증 4 — 한계/예외 상황

  A. 최대 36행 × 긴 한글 문자열  → SST 가 BIFF 레코드 한도(8224B)를 넘겨
     CONTINUE 분할 경로가 실제로 동작하는지 (평상시엔 타지 않는 코드 경로)
  B. 품목 1행만
  C. 수량 0 / 소수 / 큰 금액 / 빈 규격·단위 / 특수문자
  D. 같은 QuoteWriter 로 연속 생성해도 이전 결과가 누적되지 않는지 (원본 무오염)
"""
import os
import sys
import struct
import hashlib

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'backend'))

from xlsbiff import Cfb, Workbook, parse_ref     # noqa: E402
from quote_writer import QuoteWriter             # noqa: E402
import xlrd                                      # noqa: E402

TPL = os.path.join(HERE, '..', 'template', '견적서.xls')
OUT = os.path.join(HERE, '..', 'out')
FAILS = []


def check(label, cond, detail=''):
    print('  [%s] %s%s' % ('OK' if cond else '실패', label,
                           ('  — ' + detail) if detail else ''))
    if not cond:
        FAILS.append(label)


def sst_stats(blob):
    data = Cfb(blob).streams['Workbook'][1]
    p, n_sst, n_cont, total = 0, 0, 0, 0
    in_sst = False
    while p + 4 <= len(data):
        code, ln = struct.unpack('<HH', data[p:p + 4])
        if code == 0x00FC:
            n_sst += 1; in_sst = True; total += ln
        elif code == 0x003C and in_sst:
            n_cont += 1; total += ln
        elif code != 0x003C:
            in_sst = False
        p += 4 + ln
    return n_sst, n_cont, total


def base(items, **kw):
    d = {"quoteNo": "T-1", "quoteDate": "2026-08-23", "customer": "㈜한계시험",
         "subject": "한계 시험", "items": items, "vatRate": 0.1}
    d.update(kw)
    return d


def main():
    if not os.path.isdir(OUT):
        os.makedirs(OUT)
    w = QuoteWriter()
    tpl_hash_before = hashlib.sha256(open(TPL, 'rb').read()).hexdigest()

    # ---- A. 최대 행 + 긴 문자열 -------------------------------------
    print('A) 36행 × 긴 한글 문자열 (SST CONTINUE 분할 경로)')
    long_items = []
    for i in range(36):
        long_items.append({
            "name": "고주파 정합기용 특수 실드 케이블 어셈블리 제%02d형 (커넥터 일체형)" % (i + 1),
            "spec": "KR16840YA%03d-특수사양-내열형-40M-양단커넥터-%02d" % (i, i),
            "unit": "ea", "qty": i + 1, "price": 123456 + i * 1000,
        })
    d = base(long_items, remark="가" * 180)
    blob, res = w.generate(d)
    p = os.path.join(OUT, '_stress_36.xls')
    open(p, 'wb').write(blob)

    n_sst, n_cont, tot = sst_stats(blob)
    check('SST 총 바이트가 단일 레코드 한도(8224)를 초과', tot > 8224, '%d bytes' % tot)
    check('CONTINUE 분할 발생 (SST %d + CONTINUE %d)' % (n_sst, n_cont), n_cont >= 1)
    try:
        Workbook(Cfb(blob).streams['Workbook'][1])
        check('자체 파서 재파싱', True)
    except Exception as e:
        check('자체 파서 재파싱', False, repr(e))
    sh = xlrd.open_workbook(p, formatting_info=True).sheet_by_index(0)
    ok = all(sh.cell_value(23 - 1 + i, 2) == long_items[i]['name'] for i in range(36))
    check('36행 품명이 xlrd 로 전부 정확히 읽힘', ok)
    check('36행 마지막(58행) 값', sh.cell_value(57, 2) == long_items[35]['name'],
          repr(sh.cell_value(57, 2))[:50])
    exp = sum((i + 1) * (123456 + i * 1000) for i in range(36))
    check('공급가액 = %s' % format(exp, ','), abs(sh.cell_value(59, 7) - exp) < 0.5,
          str(sh.cell_value(59, 7)))
    check('특이사항 180자 보존', len(sh.cell_value(20, 2)) >= 180,
          '%d자' % len(sh.cell_value(20, 2)))
    # 이미지가 이 극단 케이스에서도 그대로인지
    def jpegs(b):
        data = Cfb(b).streams['Workbook'][1]
        out, pp = [], 0
        blobs = b''
        grab = False
        while pp + 4 <= len(data):
            code, ln = struct.unpack('<HH', data[pp:pp + 4])
            pay = data[pp + 4:pp + 4 + ln]
            if code == 0x00EB:
                blobs += pay; grab = True
            elif code == 0x003C and grab:
                blobs += pay
            else:
                grab = False
            pp += 4 + ln
        i = 0
        while True:
            s = blobs.find(b'\xff\xd8\xff', i)
            if s < 0: break
            e = blobs.find(b'\xff\xd9', s)
            if e < 0: break
            out.append(hashlib.sha256(blobs[s:e + 2]).hexdigest()); i = e + 2
        return out
    check('로고/직인 JPEG 2개 해시 원본과 동일',
          jpegs(blob) == jpegs(open(TPL, 'rb').read()), str(len(jpegs(blob))) + '개')
    # CONTINUE 로 넘어간 문자열까지 EXTSST 가 정확히 가리키는지
    import test_generate as tg
    before = len(tg.FAILS)
    tg.pointer_audit(blob, '36행/CONTINUE 분할본')
    FAILS.extend(tg.FAILS[before:])

    # ---- B. 1행만 ---------------------------------------------------
    print()
    print('B) 품목 1행')
    blob, res = w.generate(base([{"name": "단품", "spec": "", "unit": "", "qty": 1, "price": 1}]))
    p = os.path.join(OUT, '_stress_1.xls')
    open(p, 'wb').write(blob)
    sh = xlrd.open_workbook(p).sheet_by_index(0)
    check('C23 품명', sh.cell_value(22, 2) == '단품')
    check('H23 금액 = 1', abs(sh.cell_value(22, 7) - 1) < 1e-9)
    check('24행~58행 비어 있음',
          all(sh.cell_value(r, c) in ('', 0) for r in range(23, 58) for c in range(1, 8)))
    check('H60/H61/H62', (abs(sh.cell_value(59, 7) - 1) < 1e-9
                          and abs(sh.cell_value(60, 7) - 0.1) < 1e-9
                          and abs(sh.cell_value(61, 7) - 1.1) < 1e-9),
          '%s / %s / %s' % (sh.cell_value(59, 7), sh.cell_value(60, 7), sh.cell_value(61, 7)))

    # ---- C. 값 경계 -------------------------------------------------
    print()
    print('C) 값 경계 / 특수문자')
    items = [
        {"name": "수량0", "spec": "-", "unit": "ea", "qty": 0, "price": 500000},
        {"name": "소수 수량", "spec": "1.5m", "unit": "m", "qty": 2.5, "price": 12345.67},
        {"name": "대형금액", "spec": "＄&<>\"'%", "unit": "식", "qty": 1, "price": 987654321},
        {"name": "①②③ ㈜·㎡ 특수문자", "spec": "", "unit": "", "qty": 1, "price": 1000},
    ]
    # address 는 더 이상 쓰지 않는 키 — 남아 있어도 무시되어야 한다
    blob, res = w.generate(base(items, contact="담당자<&>", remark="비고 &<>",
                                address="쓰지 않는 키"))
    p = os.path.join(OUT, '_stress_edge.xls')
    open(p, 'wb').write(blob)
    sh = xlrd.open_workbook(p).sheet_by_index(0)
    check('사용하지 않는 address 키는 무시됨 (특이사항에 섞이지 않음)',
          sh.cell_value(20, 2) == '비고 &<>', repr(sh.cell_value(20, 2)))
    check('수량 0 → 금액 0', sh.cell_value(22, 7) == 0)
    check('소수 수량 2.5 × 12345.67', abs(sh.cell_value(23, 7) - 2.5 * 12345.67) < 1e-6,
          str(sh.cell_value(23, 7)))
    check('대형 금액 987,654,321', abs(sh.cell_value(24, 7) - 987654321) < 0.5)
    check('특수문자 품명 보존', sh.cell_value(25, 2) == '①②③ ㈜·㎡ 특수문자',
          repr(sh.cell_value(25, 2)))
    check('빈 규격/단위 → 빈 문자열', sh.cell_value(25, 3) == '' and sh.cell_value(25, 4) == '')
    exp = 0 + 2.5 * 12345.67 + 987654321 + 1000
    check('합계 일치', abs(sh.cell_value(59, 7) - exp) < 1e-6, str(sh.cell_value(59, 7)))

    # ---- D. 연속 생성 무오염 ------------------------------------------
    print()
    print('D) 연속 생성 시 상태 누적 없음 / 원본 무오염')
    a, _ = w.generate(base([{"name": "AAA", "spec": "S", "unit": "ea", "qty": 1, "price": 100}]))
    b, _ = w.generate(base([{"name": "BBB", "spec": "T", "unit": "ea", "qty": 2, "price": 200}]))
    a2, _ = w.generate(base([{"name": "AAA", "spec": "S", "unit": "ea", "qty": 1, "price": 100}]))
    check('동일 입력 → 동일 바이트 (결정적 출력)', a == a2, '%d vs %d' % (len(a), len(a2)))
    check('다른 입력 → 다른 바이트', a != b)
    pa = os.path.join(OUT, '_stress_a.xls'); open(pa, 'wb').write(a)
    sa = xlrd.open_workbook(pa).sheet_by_index(0)
    check('두번째 생성물이 첫번째 데이터를 남기지 않음', sa.cell_value(22, 2) == 'AAA'
          and sa.cell_value(23, 2) == '')
    check('원본 템플릿 파일 무변경',
          hashlib.sha256(open(TPL, 'rb').read()).hexdigest() == tpl_hash_before)

    # ---- E. 미입력 항목에 원본 잔여 내용이 남지 않는지 ------------------
    print()
    print('E) 입력하지 않은 항목은 원본 내용이 남지 않고 비워지는가')
    blob, res = w.generate({
        "quoteNo": "T-E", "quoteDate": "2026-08-23", "customer": "(주)공란시험",
        "subject": "", "contact": "", "remark": "", "vatRate": 0.1,
        "items": [{"name": "AAA", "spec": "S", "unit": "ea", "qty": 1, "price": 1000}]})
    p = os.path.join(OUT, '_stress_blank.xls')
    open(p, 'wb').write(blob)
    sh = xlrd.open_workbook(p).sheet_by_index(0)
    # 원본 C15 에는 '부속케이블 60kW (2세트:Sourse+Bias)' 가 들어 있다
    tpl_sh = xlrd.open_workbook(TPL).sheet_by_index(0)
    check('원본 C15 에는 예전 품명이 들어 있음(전제 확인)',
          '부속케이블' in tpl_sh.cell_value(14, 2), repr(tpl_sh.cell_value(14, 2)))
    check('품명 미입력 → C15 가 비워짐 (원본 잔여 내용 제거)',
          sh.cell_value(14, 2) == '', repr(sh.cell_value(14, 2)))
    check('담당자 미입력 → D14 비어 있음', sh.cell_value(13, 3) == '')
    check('특이사항 미입력 → C21 비어 있음', sh.cell_value(20, 2) == '')
    check('기본값 있는 항목은 그대로 채워짐',
          sh.cell_value(16, 2) == '발행일로 부터 4주'
          and sh.cell_value(17, 2) == '발주로 부터 2개월'
          and sh.cell_value(18, 2) == '납입 후 결제 조건')
    check('입력한 항목은 정상', sh.cell_value(12, 2) == 'T-E'
          and sh.cell_value(13, 2) == '(주)공란시험')
    os.remove(p)

    for f in ('_stress_36.xls', '_stress_1.xls', '_stress_edge.xls', '_stress_a.xls'):
        try:
            os.remove(os.path.join(OUT, f))
        except OSError:
            pass

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
