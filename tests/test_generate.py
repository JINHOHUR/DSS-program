# -*- coding: utf-8 -*-
"""
검증 2 — 실제 데이터로 견적서 1건 생성 후 원본 대비 전수 검증

확인 항목
  1  파일이 정상적으로 열리는가            (자체 파서 + 제3자 라이브러리 xlrd)
  2  회사 로고 유지                        (MSODRAWINGGROUP BLIP / JPEG 바이트 비교)
  3  법인 직인 유지                        (동일)
  4  원본 셀 서식 유지                     (XF/FONT/FORMAT/STYLE/PALETTE 레코드 비교)
  5  병합 셀 유지                          (MERGEDCELLS 레코드 비교)
  6  테두리 유지                           (XF 레코드에 포함 - 4번과 동일 근거)
  7  인쇄 레이아웃 유지                    (SETUP/PLS/HEADER/FOOTER/여백/COLINFO)
  8  입력 데이터가 정확한 위치에 들어갔는가 (xlrd 로 셀 단위 확인)
  9  기존 수식 유지                        (FORMULA rgce 바이트 비교)
 10  전체 양식 차이 요약                   (레코드 단위 diff)
"""
import os
import sys
import struct
import hashlib
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'backend'))

from xlsbiff import Cfb, Workbook, parse_ref, col_letter   # noqa: E402
from quote_writer import QuoteWriter, suggest_filename   # noqa: E402

ROOT = os.path.join(HERE, '..')
TPL = os.path.join(ROOT, 'template', '견적서.xls')
OUT_DIR = os.path.join(ROOT, 'out')

FAILS = []


def check(label, cond, detail=''):
    print('  [%s] %s%s' % ('OK' if cond else '실패', label,
                           ('  — ' + detail) if detail else ''))
    if not cond:
        FAILS.append(label)
    return cond


SAMPLE = {
    "quoteNo": "DSS-2026-0823-01",
    "quoteDate": "2026-08-23",
    "customer": "㈜에스케이하이닉스 청주",
    "contact": "김철수 책임",
    "subject": "RF Generator 부속 케이블 SET (60kW)",
    "validity": "발행일로 부터 4주",
    "delivery": "발주 후 6주",
    "payment": "납품 후 익월 말 현금 결제",
    "remark": "설치 및 시운전 비용 별도 / 납품 장소: 충북 청주시 흥덕구 신흥로 123",
    "vatRate": 0.1,
    "items": [
        {"name": "INPUT CABLE (RFG)", "spec": "KR16840YA427(40M)", "unit": "ea", "qty": 2, "price": 1250000},
        {"name": "INPUT CABLE (T/C)", "spec": "KR16840YA419(40M)", "unit": "ea", "qty": 1, "price": 1180000},
        {"name": "INTERFACE",         "spec": "KR16840YA620(20M)", "unit": "ea", "qty": 1, "price": 640000},
        {"name": "DEVICE NET(RFG)",   "spec": "KR16840YA599",      "unit": "ea", "qty": 2, "price": 315000},
        {"name": "MB-SIGNAL",         "spec": "KR16840YA531(2.5M)", "unit": "ea", "qty": 4, "price": 187500},
        {"name": "CONTROL CABLE",     "spec": "KP16840YA515(1.5M)", "unit": "ea", "qty": 1, "price": 96000},
        {"name": "CAN-DUMMY",         "spec": "DRS2-2",            "unit": "ea", "qty": 3, "price": 42000},
    ],
}


# --- 레코드 유틸 -------------------------------------------------------------

def records_of(path_or_bytes):
    raw = path_or_bytes if isinstance(path_or_bytes, bytes) else open(path_or_bytes, 'rb').read()
    data = Cfb(raw).streams['Workbook'][1]
    out, p = [], 0
    while p + 4 <= len(data):
        code, ln = struct.unpack('<HH', data[p:p + 4])
        out.append((code, data[p + 4:p + 4 + ln]))
        p += 4 + ln
    return out


def by_code(recs, *codes):
    return [pay for code, pay in recs if code in codes]


def extract_jpegs(recs):
    """MSODRAWINGGROUP(+CONTINUE) 에서 JPEG 바이트 추출"""
    blob = b''
    grab = False
    for code, pay in recs:
        if code == 0x00EB:
            blob += pay
            grab = True
        elif code == 0x003C and grab:
            blob += pay
        else:
            grab = False
    imgs = []
    i = 0
    while True:
        s = blob.find(b'\xff\xd8\xff', i)
        if s < 0:
            break
        e = blob.find(b'\xff\xd9', s)
        if e < 0:
            break
        imgs.append(blob[s:e + 2])
        i = e + 2
    return imgs


CELL_CODES = {0x0006, 0x0201, 0x0203, 0x0204, 0x0205, 0x027E,
              0x00BD, 0x00BE, 0x00FD}


def pointer_audit(raw, tag):
    """BOUNDSHEET / INDEX / DBCELL / EXTSST 의 절대·상대 오프셋이
    실제 레코드 경계에 정확히 안착하는지 스트림을 직접 걸어서 확인."""
    stream = Cfb(raw).streams['Workbook'][1]
    pos, codes, p = [], [], 0
    while p + 4 <= len(stream):
        code, ln = struct.unpack('<HH', stream[p:p + 4])
        pos.append(p)
        codes.append((code, stream[p + 4:p + 4 + ln]))
        p += 4 + ln
    at = dict(zip(pos, range(len(pos))))

    bs = [i for i, (c, _) in enumerate(codes) if c == 0x0085][0]
    ply = struct.unpack('<I', codes[bs][1][:4])[0]
    check('[%s] BOUNDSHEET.lbPlyPos → 워크시트 BOF' % tag,
          ply in at and codes[at[ply]][0] == 0x0809, 'offset=%d' % ply)

    ix = [i for i, (c, _) in enumerate(codes) if c == 0x020B][0]
    ixp = codes[ix][1]
    n = (len(ixp) - 16) // 4
    ptrs = struct.unpack('<%dI' % n, ixp[16:])
    check('[%s] INDEX.rgibRw %d개 → 실제 DBCELL' % (tag, n),
          all(q in at and codes[at[q]][0] == 0x00D7 for q in ptrs), str(list(ptrs)))
    dcw = struct.unpack('<I', ixp[12:16])[0]
    check('[%s] INDEX → DEFCOLWIDTH' % tag,
          dcw in at and codes[at[dcw]][0] == 0x0055, 'offset=%d' % dcw)

    dbok, msg, nrow = True, [], 0
    for b, q in enumerate(ptrs):
        pay = codes[at[q]][1]
        frp = q - struct.unpack('<I', pay[:4])[0]
        if frp not in at or codes[at[frp]][0] != 0x0208:
            dbok = False
            msg.append('블록%d dbRtrw→ROW 아님' % b)
            continue
        if struct.unpack('<H', codes[at[frp]][1][:2])[0] != b * 32:
            dbok = False
            msg.append('블록%d 첫 ROW 번호 불일치' % b)
        rg = struct.unpack('<%dH' % ((len(pay) - 4) // 2), pay[4:])
        cur = frp + 20
        for k, off in enumerate(rg):
            cur += off
            nrow += 1
            if cur == q:            # 뒤쪽 빈 행이 블록 끝을 가리키는 정상 케이스
                continue
            if cur not in at or codes[at[cur]][0] not in CELL_CODES:
                dbok = False
                msg.append('블록%d 행%d 포인터 이상' % (b, b * 32 + k))
    check('[%s] DBCELL 행 포인터 %d개 → 실제 셀 레코드 경계' % (tag, nrow),
          dbok, '; '.join(msg[:5]))

    ex = [i for i, (c, _) in enumerate(codes) if c == 0x00FF][0]
    exp = codes[ex][1]
    nb = (len(exp) - 2) // 8
    exok = True
    for b in range(nb):
        ib, cb, _r = struct.unpack('<IHH', exp[2 + b * 8:10 + b * 8])
        rp = ib - cb
        if rp not in at or codes[at[rp]][0] not in (0x00FC, 0x003C):
            exok = False
    check('[%s] EXTSST 버킷 %d개 → SST 내 문자열' % (tag, nb), exok)


def main():
    if not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)

    w = QuoteWriter()
    data, result = w.generate(SAMPLE)
    fname = suggest_filename(SAMPLE)
    out_path = os.path.join(OUT_DIR, fname)
    with open(out_path, 'wb') as f:
        f.write(data)

    print('생성: %s (%d bytes)' % (out_path, len(data)))
    print('공급가액 %s / 부가세 %s / 합계 %s'
          % (format(result.supply, ',.0f'), format(result.vat, ',.0f'),
             format(result.total, ',.0f')))
    print()
    print('생성기 보고:')
    for n in result.notes:
        print('  * ' + n)
    print()

    o = records_of(TPL)
    g = records_of(data)

    # 1 --------------------------------------------------------------
    print('1) 파일이 정상적으로 열리는가')
    try:
        Workbook(Cfb(data).streams['Workbook'][1])
        check('자체 BIFF 파서로 재파싱', True)
    except Exception as e:
        check('자체 BIFF 파서로 재파싱', False, repr(e))
    try:
        import xlrd
        bk = xlrd.open_workbook(out_path, formatting_info=True)
        sh = bk.sheet_by_index(0)
        check('제3자 라이브러리 xlrd 로 열기', True,
              'sheet=%r %d행 x %d열' % (sh.name, sh.nrows, sh.ncols))
    except Exception as e:
        check('제3자 라이브러리 xlrd 로 열기', False, repr(e))
        bk = sh = None

    # 2,3 ------------------------------------------------------------
    print()
    print('2·3) 회사 로고 / 법인 직인 (JPEG) 유지')
    oi, gi = extract_jpegs(o), extract_jpegs(g)
    check('이미지 개수 동일', len(oi) == len(gi) == 2,
          '원본 %d개 / 생성 %d개' % (len(oi), len(gi)))
    for k in range(min(len(oi), len(gi))):
        ho = hashlib.sha256(oi[k]).hexdigest()[:16]
        hg = hashlib.sha256(gi[k]).hexdigest()[:16]
        check('이미지#%d 바이트 동일 (%d bytes)' % (k, len(oi[k])), ho == hg,
              'sha256 %s vs %s' % (ho, hg))
    check('MSODRAWINGGROUP(이미지 데이터) 레코드 동일',
          by_code(o, 0x00EB) == by_code(g, 0x00EB))
    check('OBJ(그림 객체) 레코드 동일', by_code(o, 0x005D) == by_code(g, 0x005D))

    # 도형 배치: 로고는 그대로, 직인만 의도적으로 이동
    ow = Workbook(Cfb(open(TPL, 'rb').read()).streams['Workbook'][1])
    gw = Workbook(Cfb(data).streams['Workbook'][1])
    op, gp = ow.pictures(), gw.pictures()
    check('그림 2개 유지', len(op) == len(gp) == 2, '%d / %d' % (len(op), len(gp)))
    check('그림0(회사 로고) 앵커 변경 없음', op[0][2] == gp[0][2], str(gp[0][2]))
    tl = w.m['stamp']['topLeft']
    check('그림1(법인 직인) 앵커가 설정값으로 이동',
          gp[1][2][1:5] == (tl['col'], tl['dx'], tl['row'], tl['dy']),
          '원본 %s → 생성 %s' % (op[1][2][1:5], gp[1][2][1:5]))
    # 크기 보존 확인
    def size(wb_, a):
        x1, y1 = wb_.anchor_to_abs(a[1], a[2], a[3], a[4])
        x2, y2 = wb_.anchor_to_abs(a[5], a[6], a[7], a[8])
        return x2 - x1, y2 - y1
    so, sg = size(ow, op[1][2]), size(gw, gp[1][2])
    check('직인 크기 유지 (가로·세로 오차 1% 이내)',
          abs(so[0] - sg[0]) < so[0] * 0.01 and abs(so[1] - sg[1]) < so[1] * 0.01,
          '원본 %.0fx%.0f → 생성 %.0fx%.0f' % (so[0], so[1], sg[0], sg[1]))

    # 4,6 ------------------------------------------------------------
    print()
    print('4·6) 셀 서식 / 글꼴 / 숫자서식 / 테두리 유지')
    for code, name in [(0x00E0, 'XF(셀 서식·테두리)'), (0x0031, 'FONT'),
                       (0x041E, 'FORMAT(숫자서식)'), (0x0293, 'STYLE'),
                       (0x0092, 'PALETTE'), (0x087D, 'XFEXT'), (0x0892, 'STYLEEXT')]:
        a, b = by_code(o, code), by_code(g, code)
        check('%s %d개 전부 동일' % (name, len(a)), a == b,
              '' if a == b else '원본 %d / 생성 %d' % (len(a), len(b)))

    # 5 --------------------------------------------------------------
    print()
    print('5) 병합 셀 — 양식 병합은 유지, 품목영역 잔재만 해제')
    om, gm = set(ow.merged_ranges()), set(gw.merged_ranges())
    check('원본 병합 4건', len(om) == 4,
          ', '.join('%s%d:%s%d' % (chr(65 + c1), r1 + 1, chr(65 + c2), r2 + 1)
                    for r1, c1, r2, c2 in sorted(om)))
    keep = {r for r in om if not (22 <= r[0] and r[2] <= 57)}
    drop = om - keep
    check('양식 병합(A1:H1, C12:D12)은 그대로 유지', keep <= gm,
          ', '.join('%s%d:%s%d' % (chr(65 + c1), r1 + 1, chr(65 + c2), r2 + 1)
                    for r1, c1, r2, c2 in sorted(keep)))
    check('품목영역(23~58행) 병합만 해제됨', gm == keep and len(drop) == 2,
          '해제: ' + ', '.join('%s%d:%s%d' % (chr(65 + c1), r1 + 1, chr(65 + c2), r2 + 1)
                              for r1, c1, r2, c2 in sorted(drop)))
    if sh:
        check('xlrd 확인 병합 2건', len(sh.merged_cells) == 2, str(sorted(sh.merged_cells)))

    # 7 --------------------------------------------------------------
    print()
    print('7) 인쇄 레이아웃 유지')
    for code, name in [(0x00A1, 'SETUP(용지·배율)'), (0x004D, 'PLS(프린터설정)'),
                       (0x0014, 'HEADER'), (0x0015, 'FOOTER'),
                       (0x0026, 'LEFTMARGIN'), (0x0027, 'RIGHTMARGIN'),
                       (0x0028, 'TOPMARGIN'), (0x0029, 'BOTTOMMARGIN'),
                       (0x0083, 'HCENTER'), (0x0084, 'VCENTER'),
                       (0x007D, 'COLINFO(열 너비)'), (0x0055, 'DEFCOLWIDTH'),
                       (0x0225, 'DEFAULTROWHEIGHT'), (0x0081, 'WSBOOL'),
                       (0x089C, 'HEADERFOOTER')]:
        a, b = by_code(o, code), by_code(g, code)
        check('%s 동일' % name, a == b)

    # ROW 레코드: 품목영역 높이만 변경되었는지
    ro = {struct.unpack('<H', p[:2])[0]: p for p in by_code(o, 0x0208)}
    rg = {struct.unpack('<H', p[:2])[0]: p for p in by_code(g, 0x0208)}
    check('ROW 레코드 개수 동일', len(ro) == len(rg), '%d개' % len(ro))
    hchanged, other, unsynced, unhidden = [], [], [], []
    for r in ro:
        a, b = ro[r], rg[r]
        if a == b:
            continue
        ha = struct.unpack('<H', a[6:8])[0] & 0x7FFF
        hb = struct.unpack('<H', b[6:8])[0] & 0x7FFF
        ga = struct.unpack('<H', a[12:14])[0]
        gb = struct.unpack('<H', b[12:14])[0]
        # 허용되는 변경: 행 높이(miyRw) + fUnsynced(0x0040) 켜기 + fDyZero(0x0020) 끄기
        ok = (a[:6] == b[:6] and a[8:12] == b[8:12] and a[14:] == b[14:]
              and gb == ((ga | 0x0040) & ~0x0020))
        if ok:
            hchanged.append((r + 1, ha, hb))
            if not (ga & 0x0040):
                unsynced.append(r + 1)
            if ga & 0x0020:
                unhidden.append(r + 1)
        else:
            other.append(r + 1)
    check('행 높이·fUnsynced·숨김해제 외 ROW 속성 변경 없음', not other, str(other))
    print('       (fUnsynced 세팅 행: %s / 숨김 해제 행: %s)'
          % (unsynced or '없음', unhidden or '없음'))
    inarea = all(23 <= r <= 58 for r, _, _ in hchanged)
    check('행 높이 변경은 품목영역(23~58) 안에서만', inarea,
          ', '.join('%d행 %d→%d' % t for t in hchanged) or '변경 없음')

    # 품목영역 행 간격이 실제로 균일한지 (+ 엑셀이 재계산하지 못하게 fUnsynced 켜짐)
    hs = {}
    off = []
    for r in range(22, 58):
        pay = rg[r]
        hs.setdefault(struct.unpack('<H', pay[6:8])[0] & 0x7FFF, []).append(r + 1)
        if not (struct.unpack('<H', pay[12:14])[0] & 0x0040):
            off.append(r + 1)
    check('품목영역 36개 행의 높이가 모두 동일', len(hs) == 1,
          '높이 %s' % ', '.join('%d(%d개 행)' % (k, len(v)) for k, v in hs.items()))
    check('품목영역 전 행에 fUnsynced 켜짐 (엑셀 자동 높이 재계산 차단)', not off,
          '꺼진 행: %s' % off)

    # 숨김 행/열 — 값이 들어 있어도 엑셀에서 안 보이면 의미가 없다.
    # (원본 24·25행이 fDyZero 숨김이라 2·3번 품목이 보이지 않던 결함의 회귀 테스트)
    o_hidden = [r + 1 for r in range(22, 58)
                if struct.unpack('<H', ro[r][12:14])[0] & 0x0020]
    g_hidden = [r + 1 for r in range(22, 58)
                if struct.unpack('<H', rg[r][12:14])[0] & 0x0020]
    check('원본 품목영역에 숨김 행이 있었음(전제 확인)', o_hidden == [24, 25], str(o_hidden))
    check('생성물 품목영역에 숨김(fDyZero) 행 없음 — 모든 품목이 화면·인쇄에 보임',
          not g_hidden, '숨김 행: %s' % g_hidden)
    ghc = [col_letter(c) for c in range(9) if gw.col_hidden(c)]
    check('품목 열 A~H 중 숨김 열 없음', not ghc, '숨김 열: %s' % ghc)

    # 9 --------------------------------------------------------------
    print()
    print('9) 기존 수식 유지')

    def formulas(recs):
        out = {}
        for code, pay in recs:
            if code == 0x0006:
                r, c = struct.unpack('<HH', pay[:4])
                cce = struct.unpack('<H', pay[20:22])[0]
                out[(r, c)] = pay[22:22 + cce]
        return out
    fo, fg = formulas(o), formulas(g)
    for ref in ('H61', 'H62'):
        rc = parse_ref(ref)
        check('%s 수식(rgce) 원본과 바이트 동일' % ref,
              rc in fo and rc in fg and fo[rc] == fg[rc],
              (fo.get(rc, b'').hex() + ' vs ' + fg.get(rc, b'').hex()))
    rc = parse_ref('H60')
    check('H60 공급가액 수식은 SUM 으로 의도적 교체',
          rc in fg and fg[rc] != fo.get(rc),
          '원본 %s → 생성 %s' % (fo.get(rc, b'').hex(), fg.get(rc, b'').hex()))

    # 8 --------------------------------------------------------------
    print()
    print('8) 입력 데이터가 정확한 위치에 들어갔는가')
    if sh:
        def cv(ref):
            r, c = parse_ref(ref)
            return sh.cell_value(r, c)

        import xlrd
        d = xlrd.xldate_as_tuple(cv('C12'), bk.datemode)
        check('C12 발행일자', d[:3] == (2026, 8, 23), str(d[:3]))
        check('C13 발행번호', cv('C13') == SAMPLE['quoteNo'], repr(cv('C13')))
        check('C14 공급처',   cv('C14') == SAMPLE['customer'], repr(cv('C14')))
        check('D14 담당자',   cv('D14') == SAMPLE['contact'], repr(cv('D14')))
        check('C15 품명',     cv('C15') == SAMPLE['subject'], repr(cv('C15')))
        check('C17 유효기간', cv('C17') == SAMPLE['validity'], repr(cv('C17')))
        check('C18 납기',     cv('C18') == SAMPLE['delivery'], repr(cv('C18')))
        check('C19 결재조건', cv('C19') == SAMPLE['payment'], repr(cv('C19')))
        check('C21 특이사항', cv('C21') == SAMPLE['remark'], repr(cv('C21')))

        ok = True
        detail = []
        for i, it in enumerate(SAMPLE['items']):
            r = 23 + i
            row = [cv('B%d' % r), cv('C%d' % r), cv('D%d' % r), cv('E%d' % r),
                   cv('F%d' % r), cv('G%d' % r), cv('H%d' % r)]
            want = [str(i + 1), it['name'], it['spec'], it['unit'],
                    float(it['qty']), float(it['price']),
                    float(it['qty']) * float(it['price'])]
            if row != want:
                ok = False
                detail.append('%d행 %r != %r' % (r, row, want))
        check('품목 %d행 B~H 값 일치' % len(SAMPLE['items']), ok, ' | '.join(detail))

        empty = True
        for r in range(23 + len(SAMPLE['items']), 59):
            for cl in 'BCDEFGH':
                if cv('%s%d' % (cl, r)) not in ('', 0, None):
                    empty = False
        check('남은 품목행(%d~58)은 비어 있음' % (23 + len(SAMPLE['items'])), empty)

        check('H60 공급가액 캐시값', abs(cv('H60') - result.supply) < 0.5,
              '%s (기대 %s)' % (cv('H60'), result.supply))
        check('H61 부가세 캐시값', abs(cv('H61') - result.vat) < 0.5, str(cv('H61')))
        check('H62 합계 캐시값', abs(cv('H62') - result.total) < 0.5, str(cv('H62')))
        check('C16 금액(=H62)', abs(cv('C16') - result.total) < 0.5, str(cv('C16')))

    # 9b -------------------------------------------------------------
    print()
    print('9b) 오프셋 의존 레코드 포인터 무결성 (독립 검증)')
    # 같은 검사를 Excel 이 만든 원본에도 돌려서, 검증기 자체가 옳음을 먼저 보인다.
    pointer_audit(open(TPL, 'rb').read(), '원본(Excel 제작)')
    pointer_audit(data, '생성물')

    # 10 -------------------------------------------------------------
    print()
    print('10) 원본 대비 레코드 단위 차이 요약')
    from collections import Counter
    co, cg = Counter(c for c, _ in o), Counter(c for c, _ in g)
    diff = []
    for code in sorted(set(co) | set(cg)):
        if co[code] != cg[code]:
            diff.append('0x%04X: %d -> %d' % (code, co[code], cg[code]))
    print('   레코드 개수 변화: ' + (', '.join(diff) if diff else '없음'))
    same = sum(1 for code in set(co) | set(cg)
               if by_code(o, code) == by_code(g, code))
    print('   전체 %d종 레코드 중 %d종이 원본과 완전히 동일'
          % (len(set(co) | set(cg)), same))
    print('   Workbook 스트림 크기: %d -> %d bytes'
          % (len(Cfb(open(TPL, 'rb').read()).streams['Workbook'][1]),
             len(Cfb(data).streams['Workbook'][1])))

    print()
    if FAILS:
        print('결과: FAIL — %d건' % len(FAILS))
        for f in FAILS:
            print('   - ' + f)
        return 1
    print('결과: PASS — 전 항목 통과')
    print('생성 파일: %s' % out_path)
    return 0


if __name__ == '__main__':
    sys.exit(main())
