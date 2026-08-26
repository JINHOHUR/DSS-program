# -*- coding: utf-8 -*-
"""
quote_writer.py — 원본 견적서.xls 템플릿에 입력 데이터만 채워 새 .xls 를 만든다.

매번 config/mapping.json 과 원본 템플릿을 새로 읽어 시작하므로
이전 생성 결과가 누적되지 않는다. 원본 파일은 절대 수정하지 않는다.
"""

import io
import os
import json
import datetime

from xlsbiff import (XlsTemplate, parse_ref, col_letter,
                     formula_mul, formula_sum_col, formula_ref, BiffError)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAPPING_PATH = os.path.join(ROOT, 'config', 'mapping.json')


def load_mapping(path=MAPPING_PATH):
    with io.open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _num(v, default=0.0):
    if v is None or v == '':
        return default
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace(',', '').strip())


def _text(v):
    return '' if v is None else str(v)


def _is_empty_item(it):
    """품명·규격·단위가 모두 비어 있고 수량·단가도 없는 '빈 행'인가."""
    for k in ('name', 'spec', 'unit'):
        if _text(it.get(k)).strip():
            return False
    for k in ('qty', 'price'):
        v = it.get(k)
        if v is None:
            continue
        if _text(v).strip() in ('', '0'):
            continue
        try:
            if _num(v) != 0:
                return False
        except ValueError:
            return False           # 숫자가 아닌 값도 '내용 있음'으로 본다
    return True


def _parse_date(v):
    if isinstance(v, datetime.date):
        return v
    s = _text(v).strip()
    if not s:
        return datetime.date.today()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError('날짜 형식을 해석할 수 없습니다: %r' % s)


class QuoteResult(object):
    def __init__(self, data, supply, vat, total, notes):
        self.data = data
        self.supply = supply
        self.vat = vat
        self.total = total
        self.notes = notes


class QuoteWriter(object):

    def __init__(self, mapping=None, template_path=None):
        self.m = mapping or load_mapping()
        self.template_path = template_path or os.path.join(
            ROOT, *self.m['template'].split('/'))

    # ------------------------------------------------------------------
    def max_items(self):
        it = self.m['items']
        return it['lastRow'] - it['firstRow'] + 1

    @staticmethod
    def clean_items(d):
        """완전히 빈 행만 걸러낸다. 내용이 조금이라도 있으면 그대로 내보낸다.
        (품명이 비어 있어도 규격·수량 등이 있으면 견적서에 기재한다)"""
        return [it for it in (d.get('items') or []) if not _is_empty_item(it)]

    def validate(self, d):
        errs = []
        items = self.clean_items(d)
        if not items:
            errs.append('내용이 입력된 품목이 하나도 없습니다.')
        if len(items) > self.max_items():
            errs.append('품목은 최대 %d개까지 가능합니다 (입력 %d개).'
                        % (self.max_items(), len(items)))
        for i, it in enumerate(items, 1):
            try:
                _num(it.get('qty'))
                _num(it.get('price'))
            except ValueError:
                errs.append('%d번 품목의 수량 또는 단가가 숫자가 아닙니다.' % i)
        try:
            _parse_date(d.get('quoteDate'))
        except ValueError as e:
            errs.append(str(e))
        return errs

    # ------------------------------------------------------------------
    def calculate(self, d):
        """웹 UI 와 서버가 동일한 규칙을 쓰도록 계산을 한 곳에 둔다."""
        rate = _num(d.get('vatRate'), self.m['defaults']['vatRate'])
        rows = []
        supply = 0.0
        for it in self.clean_items(d):
            q = _num(it.get('qty'))
            p = _num(it.get('price'))
            amt = q * p
            supply += amt
            rows.append({'qty': q, 'price': p, 'amount': amt})
        vat = supply * rate
        return supply, vat, supply + vat, rows

    # ------------------------------------------------------------------
    def generate(self, d):
        """입력 dict -> (xls bytes, QuoteResult)"""
        errs = self.validate(d)
        if errs:
            raise ValueError('\n'.join(errs))

        m = self.m
        notes = []
        tpl = XlsTemplate(self.template_path)
        wb = tpl.wb

        supply, vat, total, rowcalc = self.calculate(d)
        items = self.clean_items(d)
        skipped = len(d.get('items') or []) - len(items)
        if skipped:
            notes.append('아무 내용도 없는 빈 행 %d개는 건너뛰었습니다.' % skipped)
        notes.append('품목 %d행을 %d~%d행에 기재했습니다.'
                     % (len(items), m['items']['firstRow'],
                        m['items']['firstRow'] + len(items) - 1))
        it_cfg = m['items']
        first, last = it_cfg['firstRow'], it_cfg['lastRow']
        cols = it_cfg['columns']

        # 1) 품목 영역 초기화 (값이 있는 셀만 BLANK 로 되돌림)
        cleared = 0
        for r in range(first, last + 1):
            for cl in it_cfg['clearColumns']:
                ref = '%s%d' % (cl, r)
                xf = None
                for key, cc in cols.items():
                    if cc['col'] == cl:
                        xf = cc['xf']
                        break
                if wb.clear_content(ref, xf):
                    cleared += 1
        notes.append('품목영역 %s%d:%s%d 초기화 — 값이 있던 셀 %d개를 비움'
                     % (it_cfg['clearColumns'][0], first,
                        it_cfg['clearColumns'][-1], last, cleared))

        # 2) 품목 영역: 숨김 해제가 먼저다.
        # 원본은 24·25행이 fDyZero(숨김)로 되어 있어, 값이 들어가도 엑셀에서 보이지 않는다.
        if it_cfg.get('unhideItemRows', True):
            unhid = [r for r in range(first, last + 1) if wb.set_row_hidden(r - 1, False)]
            if unhid:
                notes.append('숨겨져 있던 품목 행 숨김 해제: %s행 (원본에서 숨김 상태여서 '
                             '값을 넣어도 엑셀 화면·인쇄에 나오지 않았음)'
                             % ', '.join(str(r) for r in unhid))
            hcols = [c for c in range(0, 9) if wb.col_hidden(c)]
            for c in hcols:
                wb.set_col_hidden(c, False)
            if hcols:
                notes.append('숨겨져 있던 열 숨김 해제: %s'
                             % ', '.join(col_letter(c) for c in hcols))

        # 2-a) 품목 영역 행 높이 정규화 (전 행에 fUnsynced 를 켜서 엑셀이 재계산 못 하게)
        h = it_cfg.get('normalizeRowHeight')
        if h:
            changed = []
            for r in range(first, last + 1):
                before = wb.row_height(r - 1)
                if wb.set_row_height(r - 1, h) and before != h:
                    changed.append('%d행 %d→%d' % (r, before, h))
            notes.append('품목영역 %d~%d행 행 높이를 %d(=%.2fpt)로 통일'
                         % (first, last, h, h / 20.0)
                         + (' — 실제 변경: ' + ', '.join(changed) if changed else ''))

        # 2-b) 품목 영역 안의 병합 셀 해제
        # 원본은 27·43행에 C:D 병합이 남아 있어(옛 그룹 제목행) 그 두 행만
        # 품명 칸이 규격 칸까지 합쳐져 보이고 규격이 가려진다.
        if it_cfg.get('unmergeItemRows', True):
            gone = wb.remove_merges_in_rows(first - 1, last - 1)
            if gone:
                notes.append('품목영역 병합 셀 해제: '
                             + ', '.join('%s%d:%s%d' % (col_letter(c1), r1 + 1,
                                                        col_letter(c2), r2 + 1)
                                         for r1, c1, r2, c2 in gone)
                             + ' (원본 그룹 제목행 잔재 — 규격 칸을 가림)')

        # 3) 상단 정보
        hdr = m['header']
        defaults = m['defaults']
        # 입력이 없는 항목은 반드시 '비운다'.
        # 그냥 건너뛰면 원본 템플릿의 예전 내용(예: C15 품명)이 그대로 남는다.
        blanked = []
        for key, cfg in hdr.items():
            if key == 'remark':
                continue
            val = d.get(key)
            if val in (None, ''):
                val = defaults.get(key, '')
            if cfg['type'] == 'date':
                wb.set_date(cfg['cell'], _parse_date(val))
                continue
            text = _text(val)
            if text:
                wb.set_string(cfg['cell'], text)
            elif cfg.get('keepIfEmpty'):
                pass                      # 템플릿 원본 값을 그대로 둔다 (예: 은행계좌)
            elif wb.clear_content(cfg['cell']):
                blanked.append('%s(%s)' % (cfg['label'], cfg['cell']))
        if blanked:
            notes.append('입력이 없어 원본의 기존 내용을 비운 칸: ' + ', '.join(blanked))

        # 특이사항 + 주소
        remark = _text(d.get('remark')).strip()
        addr = _text(d.get('address')).strip()
        acfg = m.get('address', {})
        if addr and acfg.get('mode') == 'remark':
            piece = acfg.get('prefix', '주소: ') + addr
            remark = piece + (acfg.get('sep', '  /  ') + remark if remark else '')
            notes.append('원본 양식에 주소 전용 칸이 없어 특이사항(%s)에 병기'
                         % hdr['remark']['cell'])
        if remark:
            wb.set_string(hdr['remark']['cell'], remark)
        else:
            wb.clear_content(hdr['remark']['cell'])

        # 4) 품목
        c_no, c_nm = cols['no'], cols['name']
        c_sp, c_un = cols['spec'], cols['unit']
        c_qt, c_pr, c_am = cols['qty'], cols['price'], cols['amount']
        qcol = parse_ref(c_qt['col'] + '1')[1]
        pcol = parse_ref(c_pr['col'] + '1')[1]
        for i, it in enumerate(items):
            r = first + i
            calc = rowcalc[i]
            wb.set_string('%s%d' % (c_no['col'], r), str(i + 1), c_no['xf'])
            wb.set_string('%s%d' % (c_nm['col'], r), _text(it.get('name')), c_nm['xf'])
            wb.set_string('%s%d' % (c_sp['col'], r), _text(it.get('spec')), c_sp['xf'])
            wb.set_string('%s%d' % (c_un['col'], r), _text(it.get('unit')), c_un['xf'])
            wb.set_number('%s%d' % (c_qt['col'], r), calc['qty'], c_qt['xf'])
            wb.set_number('%s%d' % (c_pr['col'], r), calc['price'], c_pr['xf'])
            wb.set_formula('%s%d' % (c_am['col'], r),
                           formula_mul(r - 1, qcol, pcol),
                           calc['amount'], c_am['xf'])

        # 4-b) 법인 직인 위치
        st = m.get('stamp') or {}
        if st.get('enabled'):
            tl = st['topLeft']
            old, new = wb.move_picture(st.get('pictureIndex', 1),
                                       tl['col'], tl['dx'], tl['row'], tl['dy'])
            def _ref(a):
                return '%s%d(+%d/1024, +%d/1024)' % (col_letter(a[1]), a[3] + 1, a[2], a[4])
            notes.append('%s 위치 이동: %s → %s (크기·이미지 데이터는 그대로)'
                         % (st.get('label', '직인'), _ref(old), _ref(new)))

        # 5) 합계 영역
        t = m['totals']
        acol = parse_ref(c_am['col'] + '1')[1]
        sup_cell = t['supply']['cell']
        if t['supply']['mode'] == 'sum':
            old = wb.get_formula_rgce(sup_cell)
            wb.set_formula(sup_cell,
                           formula_sum_col(acol, first - 1, last - 1),
                           supply)
            notes.append('%s 공급가액 수식 교체: 원본 =H27+H43 (구 그룹소계 2칸 합)'
                         ' → =SUM(%s%d:%s%d) [품목 전 구간]'
                         % (sup_cell, c_am['col'], first, c_am['col'], last))
        else:
            wb.update_formula_cache(sup_cell, supply)

        rate = _num(d.get('vatRate'), m['defaults']['vatRate'])
        vat_cell = t['vat']['cell']
        if abs(rate - m['defaults']['vatRate']) < 1e-12:
            wb.update_formula_cache(vat_cell, vat)      # 원본 수식 레코드 그대로
        else:
            from xlsbiff import ptg_ref, ptg_num, PTG_MUL
            sr, sc = parse_ref(sup_cell)
            wb.set_formula(vat_cell, ptg_ref(sr, sc) + ptg_num(rate) + PTG_MUL, vat)
            notes.append('부가세율이 기본(10%%)과 달라 %s 수식을 =%s*%s 로 변경'
                         % (vat_cell, sup_cell, rate))

        wb.update_formula_cache(t['total']['cell'], total)

        # 상단 '금 액' 칸 = 합계 참조
        ha = m.get('headerAmount')
        if ha and ha.get('mode') == 'formula_ref_total':
            tr, tc = parse_ref(t['total']['cell'])
            wb.set_formula(ha['cell'], formula_ref(tr, tc), total)

        return tpl.to_bytes(), QuoteResult(d, supply, vat, total, notes)


def suggest_filename(d):
    date = _parse_date(d.get('quoteDate')).strftime('%Y%m%d')
    cust = _text(d.get('customer')).strip() or '견적서'
    no = _text(d.get('quoteNo')).strip()
    bad = '\\/:*?"<>|\r\n\t'
    cust = ''.join(ch for ch in cust if ch not in bad)[:30]
    parts = ['견적서', date, cust]
    if no:
        parts.append(''.join(ch for ch in no if ch not in bad)[:30])
    return '_'.join(parts) + '.xls'
