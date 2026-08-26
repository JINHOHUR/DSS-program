# -*- coding: utf-8 -*-
"""
xlsbiff.py — BIFF8(.xls) 원본 보존 편집 엔진

경로 A 구현체.
원본 .xls 의 OLE2(CFB) 컨테이너와 Workbook 스트림의 BIFF 레코드를 그대로 유지한 채
"필요한 셀 레코드만" 교체/삽입한다.

보존 대상 (전혀 건드리지 않음)
  - MSODRAWINGGROUP / MSODRAWING / OBJ  → 회사 로고, 법인 직인 (JPEG 2개)
  - XF / FONT / FORMAT / STYLE / PALETTE → 셀 서식, 글꼴, 숫자서식
  - MERGEDCELLS                         → 병합 셀
  - ROW / COLINFO / DEFCOLWIDTH         → 행 높이, 열 너비
  - SETUP / PLS / HEADER / FOOTER / 여백 → 인쇄 설정
  - 기존 FORMULA 레코드                 → 수식 (캐시값만 갱신 가능)

재계산이 필요한 오프셋 의존 레코드 (원본 규칙을 실측 검증 후 동일하게 재생성)
  - BOUNDSHEET.lbPlyPos
  - INDEX (rgibRw, DEFCOLWIDTH 위치)
  - DBCELL (dbRtrw, rgdb)
  - EXTSST (ib, cbOffset)

무편집 재작성 시 원본과 바이트 단위로 100% 동일한 파일이 나오는 것으로 검증한다.
(tests/test_roundtrip.py)
"""

import struct
import datetime
from collections import OrderedDict

# ----------------------------------------------------------------------------
# OLE2 / CFB (Compound File Binary) v3, 512-byte sector
# ----------------------------------------------------------------------------

CFB_SIG = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
SECTOR = 512
DIR_ENTRY = 128


class CfbError(Exception):
    pass


class Cfb(object):
    """OLE2 컨테이너. 디렉터리 엔트리를 원본 바이트 그대로 보존하고
    start-sector / size 필드만 갱신해서 다시 쓴다."""

    def __init__(self, raw):
        if raw[:8] != CFB_SIG:
            raise CfbError('OLE2/CFB 시그니처가 아닙니다.')
        self.header = bytearray(raw[:SECTOR])
        ssz, mssz = struct.unpack('<HH', raw[30:34])
        if ssz != 9:
            raise CfbError('512바이트 섹터(v3)만 지원합니다. (sector shift=%d)' % ssz)
        self.mini_shift = mssz
        (self.n_dir_sect, self.n_fat_sect, self.dir_start, self.tx_sig,
         self.mini_cutoff, self.mini_start, self.n_minifat,
         self.difat_start, self.n_difat) = struct.unpack('<IIIIIIIII', raw[40:76])
        difat = list(struct.unpack('<109I', raw[76:512]))
        if self.n_difat:
            raise CfbError('DIFAT 확장 섹터를 쓰는 대용량 파일은 지원하지 않습니다.')
        self.raw = raw

        def sect(n):
            o = SECTOR + n * SECTOR
            return raw[o:o + SECTOR]

        fat = []
        for i in range(self.n_fat_sect):
            fat += list(struct.unpack('<128I', sect(difat[i])))
        self.fat = fat

        def chain(start):
            out, s, guard = [], start, 0
            while s < FATSECT:
                out.append(s)
                s = fat[s]
                guard += 1
                if guard > len(fat) + 8:
                    raise CfbError('FAT 체인 순환')
            return out

        dirdata = b''.join(sect(s) for s in chain(self.dir_start))
        self.dir_entries = [bytearray(dirdata[i:i + DIR_ENTRY])
                            for i in range(0, len(dirdata), DIR_ENTRY)]

        # 미니 스트림
        root = self.dir_entries[0]
        mini_start_sect = struct.unpack('<I', root[116:120])[0]
        mini_size = struct.unpack('<I', root[120:124])[0]
        ministream = b''
        if mini_start_sect < FATSECT:
            ministream = b''.join(sect(s) for s in chain(mini_start_sect))[:mini_size]
        minifat = []
        if self.mini_start < FATSECT:
            md = b''.join(sect(s) for s in chain(self.mini_start))
            minifat = list(struct.unpack('<%dI' % (len(md) // 4), md))

        # 스트림 읽기 (디렉터리 순서 유지)
        self.streams = OrderedDict()   # name -> (dir_index, bytes)
        for idx, e in enumerate(self.dir_entries):
            if e[66] != 2:             # 2 = stream
                continue
            nlen = struct.unpack('<H', e[64:66])[0]
            name = bytes(e[:max(0, nlen - 2)]).decode('utf-16-le')
            start = struct.unpack('<I', e[116:120])[0]
            size = struct.unpack('<I', e[120:124])[0]
            if size < self.mini_cutoff:
                mss = 1 << self.mini_shift
                buf, s = [], start
                while s < FATSECT:
                    buf.append(ministream[s * mss:(s + 1) * mss])
                    s = minifat[s]
                data = b''.join(buf)[:size]
            else:
                data = b''.join(sect(s) for s in chain(start))[:size]
            self.streams[name] = (idx, data)

    def write(self, streams):
        """streams: {name: bytes} 로 교체한 새 CFB 바이트 반환."""
        order = []
        for name, (idx, data) in self.streams.items():
            new = streams.get(name, data)
            if len(new) < self.mini_cutoff:
                raise CfbError(
                    '스트림 %r 이 미니스트림 임계치(%d) 미만입니다. 이 템플릿에서는 발생하지 않아야 합니다.'
                    % (name, self.mini_cutoff))
            order.append((idx, new))

        # 1) 레이아웃: [스트림 데이터][FAT 섹터][디렉터리 섹터]
        def nsect(n):
            return (n + SECTOR - 1) // SECTOR

        place = []
        cur = 0
        for idx, data in order:
            cnt = nsect(len(data))
            place.append((idx, data, cur, cnt))
            cur += cnt
        data_sectors = cur
        n_dir_sect = nsect(len(self.dir_entries) * DIR_ENTRY)

        n_fat = 1
        while True:
            total = data_sectors + n_fat + n_dir_sect
            need = max(1, (total + 127) // 128)
            if need == n_fat:
                break
            n_fat = need
        fat_first = data_sectors
        dir_first = data_sectors + n_fat
        total = dir_first + n_dir_sect

        # 2) FAT 구성
        fat = [FREESECT] * (n_fat * 128)
        for idx, data, first, cnt in place:
            for i in range(cnt):
                fat[first + i] = ENDOFCHAIN if i == cnt - 1 else first + i + 1
        for i in range(n_fat):
            fat[fat_first + i] = FATSECT
        for i in range(n_dir_sect):
            fat[dir_first + i] = ENDOFCHAIN if i == n_dir_sect - 1 else dir_first + i + 1

        # 3) 디렉터리 엔트리 갱신 (start sector / size 만)
        dirs = [bytearray(e) for e in self.dir_entries]
        for idx, data, first, cnt in place:
            dirs[idx][116:120] = struct.pack('<I', first)
            dirs[idx][120:124] = struct.pack('<I', len(data))
            dirs[idx][124:128] = b'\x00\x00\x00\x00'

        # 4) 헤더 갱신
        h = bytearray(self.header)
        h[40:44] = struct.pack('<I', 0)              # v3: 디렉터리 섹터 수 미사용
        h[44:48] = struct.pack('<I', n_fat)
        h[48:52] = struct.pack('<I', dir_first)
        h[68:72] = struct.pack('<I', ENDOFCHAIN)     # DIFAT 확장 없음
        h[72:76] = struct.pack('<I', 0)
        difat = [FREESECT] * 109
        for i in range(n_fat):
            difat[i] = fat_first + i
        h[76:512] = struct.pack('<109I', *difat)

        out = bytearray(h)
        for idx, data, first, cnt in place:
            out += data + b'\x00' * (cnt * SECTOR - len(data))
        for i in range(n_fat):
            out += struct.pack('<128I', *fat[i * 128:(i + 1) * 128])
        dd = b''.join(bytes(e) for e in dirs)
        out += dd + b'\x00' * (n_dir_sect * SECTOR - len(dd))
        assert len(out) == SECTOR + total * SECTOR
        return bytes(out)


# ----------------------------------------------------------------------------
# BIFF8 레코드
# ----------------------------------------------------------------------------

BOF = 0x0809
EOF_ = 0x000A
CONTINUE = 0x003C
BOUNDSHEET = 0x0085
INDEX = 0x020B
DBCELL = 0x00D7
SST = 0x00FC
EXTSST = 0x00FF
DEFCOLWIDTH = 0x0055
ROW = 0x0208
DIMENSIONS = 0x0200
DATEMODE = 0x0022
MERGEDCELLS = 0x00E5

LABELSST = 0x00FD
NUMBER = 0x0203
RK = 0x027E
BLANK = 0x0201
BOOLERR = 0x0205
FORMULA = 0x0006
MULRK = 0x00BD
MULBLANK = 0x00BE

SINGLE_CELL = (LABELSST, NUMBER, RK, BLANK, BOOLERR, FORMULA)
MULTI_CELL = (MULRK, MULBLANK)
CELL_RECS = SINGLE_CELL + MULTI_CELL

MAX_PAYLOAD = 8224
ROWBLOCK = 32
SST_BUCKET = 8


class BiffError(Exception):
    pass


def col_letter(c):
    s = ''
    while True:
        s = chr(ord('A') + c % 26) + s
        c = c // 26 - 1
        if c < 0:
            return s


def cellref(r, c):
    return '%s%d' % (col_letter(c), r + 1)


def parse_ref(ref):
    """'H62' -> (61, 7)"""
    ref = ref.strip().upper().replace('$', '')
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    if i == 0 or i == len(ref):
        raise ValueError('셀 주소 형식 오류: %r' % ref)
    col = 0
    for ch in ref[:i]:
        col = col * 26 + (ord(ch) - 64)
    return int(ref[i:]) - 1, col - 1


# --- ptg (수식 토큰) 빌더 -----------------------------------------------------

def ptg_ref(row, col, val_class=True):
    return struct.pack('<BHH', 0x44 if val_class else 0x24, row, col | 0xC000)


def ptg_area(r1, c1, r2, c2, val_class=True):
    return struct.pack('<BHHHH', 0x45 if val_class else 0x25,
                       r1, r2, c1 | 0xC000, c2 | 0xC000)


def ptg_num(v):
    return b'\x1f' + struct.pack('<d', float(v))


PTG_ADD = b'\x03'
PTG_MUL = b'\x05'


def ptg_func_var(iftab, nargs):
    return struct.pack('<BBH', 0x42, nargs, iftab)


FN_SUM = 4


def formula_mul(r, c1, c2):
    """=<c1><r> * <c2><r>"""
    return ptg_ref(r, c1) + ptg_ref(r, c2) + PTG_MUL


def formula_sum_col(col, r1, r2):
    """=SUM(<col><r1>:<col><r2>)"""
    return ptg_area(r1, col, r2, col) + ptg_func_var(FN_SUM, 1)


def formula_ref(r, c):
    return ptg_ref(r, c)


# --- SST 문자열 인코딩 --------------------------------------------------------

def encode_sst_string(text):
    """XLUnicodeRichExtendedString (rich/ext 없음)"""
    if len(text) > 32767:
        raise BiffError('문자열이 너무 깁니다.')
    if all(ord(ch) < 256 for ch in text):
        return struct.pack('<HB', len(text), 0x00) + text.encode('latin-1')
    return struct.pack('<HB', len(text), 0x01) + text.encode('utf-16-le')


def sst_string_len(buf, o):
    """헤더만 읽어 문자열이 차지하는 전체 바이트 수를 계산 (경계 검사용)."""
    if o + 3 > len(buf):
        return None
    cch = struct.unpack('<H', buf[o:o + 2])[0]
    grbit = buf[o + 2]
    n = 3
    crun = cbext = 0
    if grbit & 0x08:
        if o + n + 2 > len(buf):
            return None
        crun = struct.unpack('<H', buf[o + n:o + n + 2])[0]
        n += 2
    if grbit & 0x04:
        if o + n + 4 > len(buf):
            return None
        cbext = struct.unpack('<i', buf[o + n:o + n + 4])[0]
        n += 4
    n += cch * (2 if grbit & 0x01 else 1)
    n += crun * 4 + max(0, cbext)
    return n


def decode_sst_string(buf, o):
    """(text, total_bytes) 반환. rich/ext 포함 전체 길이."""
    total = sst_string_len(buf, o)
    if total is None or o + total > len(buf):
        raise BiffError('SST 문자열이 레코드 경계를 넘습니다 (offset=%d)' % o)
    cch = struct.unpack('<H', buf[o:o + 2])[0]
    grbit = buf[o + 2]
    p = o + 3
    if grbit & 0x08:
        p += 2
    if grbit & 0x04:
        p += 4
    if grbit & 0x01:
        text = buf[p:p + cch * 2].decode('utf-16-le')
    else:
        text = buf[p:p + cch].decode('latin-1')
    return text, total


# ----------------------------------------------------------------------------
# Workbook 스트림
# ----------------------------------------------------------------------------

class Workbook(object):

    def __init__(self, data):
        self.records = []          # [code, bytearray]
        p, n = 0, len(data)
        while p + 4 <= n:
            code, ln = struct.unpack('<HH', data[p:p + 4])
            if p + 4 + ln > n:
                raise BiffError('레코드가 스트림 경계를 넘습니다 (pos=%d)' % p)
            self.records.append([code, bytearray(data[p + 4:p + 4 + ln])])
            p += 4 + ln
        if p != n:
            raise BiffError('스트림 끝에 %d 바이트가 남았습니다.' % (n - p))

        self._locate()
        self._parse_sst()

    # -- 구조 파악 -----------------------------------------------------------

    def _locate(self):
        self.i_boundsheet = None
        self.i_sst = None
        self.i_extsst = None
        self.i_index = None
        self.i_defcolwidth = None
        self.sheet_bof = None
        self.datemode_1904 = False

        bofs = [i for i, (c, _) in enumerate(self.records) if c == BOF]
        if len(bofs) < 2:
            raise BiffError('워크시트 substream 을 찾지 못했습니다.')
        self.sheet_bof = bofs[1]
        self.sheet_end = len(self.records)
        for i in range(self.sheet_bof, len(self.records)):
            if self.records[i][0] == EOF_:
                self.sheet_end = i
                break

        for i, (c, pay) in enumerate(self.records):
            if c == BOUNDSHEET and self.i_boundsheet is None:
                self.i_boundsheet = i
            elif c == SST:
                self.i_sst = i
            elif c == EXTSST:
                self.i_extsst = i
            elif c == DATEMODE:
                self.datemode_1904 = bool(struct.unpack('<H', bytes(pay[:2]))[0])
            elif i > self.sheet_bof:
                if c == INDEX and self.i_index is None:
                    self.i_index = i
                elif c == DEFCOLWIDTH and self.i_defcolwidth is None:
                    self.i_defcolwidth = i
        if self.i_sst is None:
            raise BiffError('SST 레코드가 없습니다.')

    def _parse_sst(self):
        """SST + 후속 CONTINUE 를 문자열 단위 raw 바이트로 분해."""
        i = self.i_sst
        run = [i]
        j = i + 1
        while j < len(self.records) and self.records[j][0] == CONTINUE:
            run.append(j)
            j += 1
        self.sst_run_len = len(run)
        bufs = [bytes(self.records[k][1]) for k in run]
        self.sst_total, cst_unique = struct.unpack('<ii', bufs[0][:8])
        self.sst_raw = []      # 문자열별 raw 바이트
        self.sst_text = []
        bi, o = 0, 8
        for k in range(cst_unique):
            # 이 레코드에 문자열 헤더가 더 들어갈 자리가 없으면 다음 CONTINUE 로
            while bi < len(bufs) and (o >= len(bufs[bi])
                                      or sst_string_len(bufs[bi], o) is None):
                bi += 1
                o = 0
            if bi >= len(bufs):
                raise BiffError('SST 문자열 %d개를 다 읽기 전에 레코드가 끝났습니다.' % cst_unique)
            ln = sst_string_len(bufs[bi], o)
            if o + ln > len(bufs[bi]):
                raise BiffError(
                    'SST 문자열이 CONTINUE 경계에 걸쳐 분할된 파일은 지원하지 않습니다 '
                    '(문자열 #%d). 이 도구가 만든 파일은 문자열을 분할하지 않습니다.' % k)
            text, _ = decode_sst_string(bufs[bi], o)
            self.sst_raw.append(bufs[bi][o:o + ln])
            self.sst_text.append(text)
            o += ln
        rest = len(bufs[bi]) - o + sum(len(b) for b in bufs[bi + 1:])
        if rest:
            raise BiffError('SST 파싱 잔여 바이트 %d' % rest)
        self.sst_index = {}
        for k, t in enumerate(self.sst_text):
            self.sst_index.setdefault(t, k)
        self.sst_dirty = False

    def sst_add(self, text):
        if text in self.sst_index:
            return self.sst_index[text]
        raw = encode_sst_string(text)
        if len(raw) + 4 > MAX_PAYLOAD:
            raise BiffError('단일 문자열이 BIFF 레코드 한도를 넘습니다.')
        k = len(self.sst_raw)
        self.sst_raw.append(raw)
        self.sst_text.append(text)
        self.sst_index[text] = k
        self.sst_dirty = True
        return k

    # -- 셀 접근 -------------------------------------------------------------

    @staticmethod
    def _cell_row(code, pay):
        return struct.unpack('<H', bytes(pay[:2]))[0]

    @staticmethod
    def _cell_cols(code, pay):
        """(colFirst, colLast) 포함 범위"""
        if code in SINGLE_CELL:
            c = struct.unpack('<H', bytes(pay[2:4]))[0]
            return c, c
        c1 = struct.unpack('<H', bytes(pay[2:4]))[0]
        c2 = struct.unpack('<H', bytes(pay[-2:]))[0]
        return c1, c2

    def _xf_of(self, code, pay, col):
        if code in SINGLE_CELL:
            return struct.unpack('<H', bytes(pay[4:6]))[0]
        c1 = struct.unpack('<H', bytes(pay[2:4]))[0]
        k = col - c1
        if code == MULBLANK:
            return struct.unpack('<H', bytes(pay[4 + k * 2:6 + k * 2]))[0]
        return struct.unpack('<H', bytes(pay[4 + k * 6:6 + k * 6]))[0]

    def _split_multi(self, code, pay, col):
        """MULBLANK/MULRK 에서 col 을 제거하고 좌/우 조각 레코드 리스트 반환."""
        row = struct.unpack('<H', bytes(pay[:2]))[0]
        c1 = struct.unpack('<H', bytes(pay[2:4]))[0]
        c2 = struct.unpack('<H', bytes(pay[-2:]))[0]
        unit = 2 if code == MULBLANK else 6
        body = bytes(pay[4:-2])
        out = []

        def make(a, b):
            if a > b:
                return
            seg = body[(a - c1) * unit:(b - c1 + 1) * unit]
            if a == b:
                if code == MULBLANK:
                    out.append([BLANK, bytearray(struct.pack('<HHH', row, a,
                                struct.unpack('<H', seg[:2])[0]))])
                else:
                    xf = struct.unpack('<H', seg[:2])[0]
                    out.append([RK, bytearray(struct.pack('<HHH', row, a, xf) + seg[2:6])])
            else:
                out.append([code, bytearray(struct.pack('<HH', row, a) + seg
                                            + struct.pack('<H', b))])
        make(c1, col - 1)
        make(col + 1, c2)
        return out

    def _row_default_xf(self, row):
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW and struct.unpack('<H', bytes(pay[:2]))[0] == row:
                flags = struct.unpack('<H', bytes(pay[12:14]))[0]
                ixfe = struct.unpack('<H', bytes(pay[14:16]))[0]
                if flags & 0x0080:
                    return ixfe & 0x0FFF
        return 15

    def _find_cell(self, row, col):
        """(record_index, code, xf) 또는 (insert_index, None, None)"""
        first_gt = None
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c not in CELL_RECS:
                continue
            r = self._cell_row(c, pay)
            if r < row:
                continue
            if r > row:
                if first_gt is None:
                    first_gt = i
                continue
            c1, c2 = self._cell_cols(c, pay)
            if c1 <= col <= c2:
                return i, c, self._xf_of(c, pay, col)
            if c1 > col:
                return i, None, None       # 같은 행, 이 위치 앞에 삽입
        if first_gt is not None:
            return first_gt, None, None
        # 이 행 이후 셀 레코드가 없음 → 해당 블록의 DBCELL 앞
        for i in range(self.sheet_bof, self.sheet_end):
            if self.records[i][0] == DBCELL:
                if i > self._last_cell_index_before(row):
                    return i, None, None
        raise BiffError('삽입 위치를 찾지 못했습니다: %s' % cellref(row, col))

    def _last_cell_index_before(self, row):
        last = self.sheet_bof
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c in CELL_RECS and self._cell_row(c, pay) <= row:
                last = i
        return last

    def _put(self, row, col, make_record, xf_override=None):
        """make_record(xf) -> [code, payload]"""
        i, code, xf = self._find_cell(row, col)
        if xf_override is not None:
            xf = xf_override
        if code is None:                      # 신규 삽입
            if xf is None:
                xf = self._row_default_xf(row)
            self.records.insert(i, make_record(xf))
            self._shift(i, 1)
        elif code in SINGLE_CELL:
            self.records[i] = make_record(xf)
        else:                                 # MULBLANK / MULRK 분할
            pieces = self._split_multi(code, self.records[i][1], col)
            new = make_record(xf)
            left = [p for p in pieces if self._cell_cols(p[0], p[1])[1] < col]
            right = [p for p in pieces if self._cell_cols(p[0], p[1])[0] > col]
            repl = left + [new] + right
            self.records[i:i + 1] = repl
            self._shift(i, len(repl) - 1)
        self._extend_row(row, col)

    def _shift(self, at, delta):
        if delta == 0:
            return
        if self.sheet_bof >= at:
            self.sheet_bof += delta
        if self.sheet_end >= at:
            self.sheet_end += delta
        for name in ('i_boundsheet', 'i_sst', 'i_extsst', 'i_index', 'i_defcolwidth'):
            v = getattr(self, name)
            if v is not None and v >= at:
                setattr(self, name, v + delta)

    def _extend_row(self, row, col):
        """ROW 레코드의 colMic/colMac 확장 + DIMENSIONS 확장"""
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW and struct.unpack('<H', bytes(pay[:2]))[0] == row:
                cm, cM = struct.unpack('<HH', bytes(pay[2:6]))
                ncm, ncM = min(cm, col), max(cM, col + 1)
                if (ncm, ncM) != (cm, cM):
                    pay[2:6] = struct.pack('<HH', ncm, ncM)
                break
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == DIMENSIONS:
                rf, rl, cf, cl = struct.unpack('<iiHH', bytes(pay[:12]))
                n = (max(rf, 0) if row >= rf else row, max(rl, row + 1),
                     min(cf, col), max(cl, col + 1))
                pay[:12] = struct.pack('<iiHH', min(rf, row), max(rl, row + 1),
                                       min(cf, col), max(cl, col + 1))
                break

    # -- 공개 API ------------------------------------------------------------

    def set_string(self, ref, text, xf=None):
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        text = u'' if text is None else str(text)
        idx = self.sst_add(text)
        self.sst_total += 0   # cstTotal 은 저장 시 재계산
        self._put(row, col,
                  lambda x: [LABELSST, bytearray(struct.pack('<HHHi', row, col, x, idx))],
                  xf)

    def set_number(self, ref, value, xf=None):
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        self._put(row, col,
                  lambda x: [NUMBER, bytearray(struct.pack('<HHH', row, col, x)
                                               + struct.pack('<d', float(value)))],
                  xf)

    def set_date(self, ref, d, xf=None):
        """datetime.date -> Excel 일련번호"""
        base = datetime.date(1904, 1, 1) if self.datemode_1904 else datetime.date(1899, 12, 30)
        serial = (d - base).days
        if not self.datemode_1904 and serial > 59:
            pass   # 1900 윤년 버그: 1900-03-01 이후는 1899-12-30 기준이 이미 정확
        self.set_number(ref, serial, xf)

    def set_blank(self, ref, xf=None):
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        self._put(row, col,
                  lambda x: [BLANK, bytearray(struct.pack('<HHH', row, col, x))],
                  xf)

    def clear_content(self, ref, xf=None):
        """값이 들어있는 셀만 BLANK 로 되돌린다.
        이미 비어 있는 셀(BLANK/MULBLANK)은 건드리지 않는다 → 원본 레코드 최대 보존."""
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        i, code, cur_xf = self._find_cell(row, col)
        if code is None or code in (BLANK, MULBLANK):
            return False
        self._put(row, col,
                  lambda x: [BLANK, bytearray(struct.pack('<HHH', row, col, x))],
                  xf if xf is not None else cur_xf)
        return True

    def set_row_height(self, row, height):
        """ROW 레코드의 miyRw(행 높이, 1/20 pt)를 지정하고 fUnsynced 를 켠다.
        이미 같은 높이여도 fUnsynced 를 반드시 켜야 한다. 이 플래그가 꺼져 있으면
        Excel 이 글꼴 크기에 맞춰 행 높이를 다시 계산해 버려서 행 간격이 들쭉날쭉해진다."""
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW and struct.unpack('<H', bytes(pay[:2]))[0] == row:
                old = struct.unpack('<H', bytes(pay[6:8]))[0]
                g = struct.unpack('<H', bytes(pay[12:14]))[0]
                changed = (old & 0x7FFF) != height or not (g & 0x0040)
                pay[6:8] = struct.pack('<H', (old & 0x8000) | (height & 0x7FFF))
                pay[12:14] = struct.pack('<H', g | 0x0040)   # fUnsynced
                return changed
        return False

    def row_hidden(self, row):
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW and struct.unpack('<H', bytes(pay[:2]))[0] == row:
                return bool(struct.unpack('<H', bytes(pay[12:14]))[0] & 0x0020)
        return False

    def set_row_hidden(self, row, hidden):
        """ROW 레코드의 fDyZero(행 숨김) 비트를 켜고 끈다.
        이 비트가 켜져 있으면 셀 값이 들어 있어도 Excel 화면·인쇄에서 행이 보이지 않는다."""
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW and struct.unpack('<H', bytes(pay[:2]))[0] == row:
                g = struct.unpack('<H', bytes(pay[12:14]))[0]
                new = (g | 0x0020) if hidden else (g & ~0x0020)
                if new == g:
                    return False
                pay[12:14] = struct.pack('<H', new)
                return True
        return False

    def col_hidden(self, col):
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == self.COLINFO:
                c1, c2, w, xf, g = struct.unpack('<HHHHH', bytes(pay[:10]))
                if c1 <= col <= c2:
                    return bool(g & 0x0001)
        return False

    def set_col_hidden(self, col, hidden):
        """COLINFO 의 fHidden 비트 조작. 해당 열만 떼어내야 하면 COLINFO 를 분할한다."""
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c != self.COLINFO:
                continue
            c1, c2, w, xf, g = struct.unpack('<HHHHH', bytes(pay[:10]))
            if not (c1 <= col <= c2):
                continue
            new = (g | 0x0001) if hidden else (g & ~0x0001)
            if new == g:
                return False
            tail = bytes(pay[10:])
            if c1 == c2:
                pay[:10] = struct.pack('<HHHHH', c1, c2, w, xf, new)
                return True
            parts = []
            if c1 < col:
                parts.append(struct.pack('<HHHHH', c1, col - 1, w, xf, g) + tail)
            parts.append(struct.pack('<HHHHH', col, col, w, xf, new) + tail)
            if col < c2:
                parts.append(struct.pack('<HHHHH', col + 1, c2, w, xf, g) + tail)
            repl = [[self.COLINFO, bytearray(b)] for b in parts]
            self.records[i:i + 1] = repl
            self._shift(i, len(repl) - 1)
            self._geo = None
            return True
        return False

    def merged_ranges(self):
        """[(rowFirst, colFirst, rowLast, colLast), ...] (0-기준)"""
        out = []
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c != MERGEDCELLS:
                continue
            n = struct.unpack('<H', bytes(pay[:2]))[0]
            for k in range(n):
                r1, r2, c1, c2 = struct.unpack('<HHHH', bytes(pay[2 + k * 8:10 + k * 8]))
                out.append((r1, c1, r2, c2))
        return out

    def remove_merges_in_rows(self, row_first, row_last):
        """지정한 행 범위 안에 완전히 들어가는 병합을 해제한다.
        반환: 해제한 범위 목록"""
        removed = []
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c != MERGEDCELLS:
                continue
            n = struct.unpack('<H', bytes(pay[:2]))[0]
            keep = []
            for k in range(n):
                item = struct.unpack('<HHHH', bytes(pay[2 + k * 8:10 + k * 8]))
                r1, r2, c1, c2 = item
                if row_first <= r1 and r2 <= row_last:
                    removed.append((r1, c1, r2, c2))
                else:
                    keep.append(item)
            if len(keep) != n:
                new = struct.pack('<H', len(keep))
                for it in keep:
                    new += struct.pack('<HHHH', *it)
                self.records[i][1] = bytearray(new)
        return removed

    def row_height(self, row):
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW and struct.unpack('<H', bytes(pay[:2]))[0] == row:
                return struct.unpack('<H', bytes(pay[6:8]))[0] & 0x7FFF
        return None

    def set_formula(self, ref, rgce, cached, xf=None, grbit=None):
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        g = self.default_formula_grbit() if grbit is None else grbit

        def mk(x):
            pay = (struct.pack('<HHH', row, col, x)
                   + struct.pack('<d', float(cached))
                   + struct.pack('<HIH', g, 0, len(rgce)) + rgce)
            return [FORMULA, bytearray(pay)]
        self._put(row, col, mk, xf)

    def default_formula_grbit(self):
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == FORMULA:
                return struct.unpack('<H', bytes(pay[14:16]))[0]
        return 0x0002

    def update_formula_cache(self, ref, value):
        """기존 FORMULA 레코드의 수식(rgce)은 그대로 두고 캐시 결과값만 갱신."""
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c != FORMULA:
                continue
            r, cc = struct.unpack('<HH', bytes(pay[:4]))
            if (r, cc) == (row, col):
                pay[6:14] = struct.pack('<d', float(value))
                return True
        return False

    def get_formula_rgce(self, ref):
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c != FORMULA:
                continue
            r, cc = struct.unpack('<HH', bytes(pay[:4]))
            if (r, cc) == (row, col):
                cce = struct.unpack('<H', bytes(pay[20:22]))[0]
                return bytes(pay[22:22 + cce])
        return None

    def cell_xf(self, ref):
        row, col = parse_ref(ref) if isinstance(ref, str) else ref
        i, code, xf = self._find_cell(row, col)
        return xf

    # -- 시트 기하 / 그림 앵커 ------------------------------------------------

    COLINFO = 0x007D
    MSODRAWING = 0x00EC
    DEFAULTROWHEIGHT = 0x0225
    ESCHER_CLIENT_ANCHOR = 0xF010

    def _geometry(self):
        """열 너비(1/256 문자) / 행 높이(twip) 누적 경계"""
        if getattr(self, '_geo', None):
            return self._geo
        defw = 2048
        for i in range(self.sheet_bof, self.sheet_end):
            if self.records[i][0] == DEFCOLWIDTH:
                defw = struct.unpack('<H', bytes(self.records[i][1][:2]))[0] * 256
                break
        widths = [defw] * 256
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == self.COLINFO:
                c1, c2, w = struct.unpack('<HHH', bytes(pay[:6]))
                for k in range(c1, min(c2, 255) + 1):
                    widths[k] = w
        defh = 255
        for i in range(self.sheet_bof, self.sheet_end):
            if self.records[i][0] == self.DEFAULTROWHEIGHT:
                defh = struct.unpack('<H', bytes(self.records[i][1][2:4]))[0] & 0x7FFF
                break
        heights = [defh] * 1024
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW:
                r = struct.unpack('<H', bytes(pay[:2]))[0]
                if r < 1024:
                    heights[r] = struct.unpack('<H', bytes(pay[6:8]))[0] & 0x7FFF
        cx, cy = [0], [0]
        for w in widths:
            cx.append(cx[-1] + w)
        for h in heights:
            cy.append(cy[-1] + h)
        self._geo = (widths, cx, heights, cy)
        return self._geo

    def anchor_to_abs(self, col, dx, row, dy):
        widths, cx, heights, cy = self._geometry()
        return (cx[col] + widths[col] * dx / 1024.0,
                cy[row] + heights[row] * dy / 1024.0)

    def abs_to_anchor(self, x, y):
        widths, cx, heights, cy = self._geometry()
        col = 0
        while col < 255 and cx[col + 1] <= x:
            col += 1
        row = 0
        while row < 1023 and cy[row + 1] <= y:
            row += 1
        dx = int(round((x - cx[col]) / float(widths[col]) * 1024)) if widths[col] else 0
        dy = int(round((y - cy[row]) / float(heights[row]) * 1024)) if heights[row] else 0
        return col, max(0, min(1023, dx)), row, max(0, min(1023, dy))

    @staticmethod
    def _escher_scan(buf, base=0, out=None):
        """Escher 레코드 트리에서 (type, 데이터 시작 오프셋, 길이) 수집"""
        if out is None:
            out = []
        p = 0
        while p + 8 <= len(buf):
            vi, typ, ln = struct.unpack('<HHI', bytes(buf[p:p + 8]))
            out.append((typ, base + p + 8, ln))
            if (vi & 0x0F) == 0x0F:
                Workbook._escher_scan(buf[p + 8:p + 8 + ln], base + p + 8, out)
            p += 8 + ln
        return out

    def pictures(self):
        """시트의 그림 목록. [(레코드 인덱스, 앵커 데이터 오프셋, 앵커 9튜플), ...]"""
        out = []
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c != self.MSODRAWING:
                continue
            for typ, off, ln in self._escher_scan(pay):
                if typ == self.ESCHER_CLIENT_ANCHOR and ln >= 18:
                    out.append((i, off, struct.unpack('<9H', bytes(pay[off:off + 18]))))
        return out

    def move_picture(self, index, col1, dx1, row1, dy1):
        """그림의 좌상단 앵커를 옮긴다. 크기는 그대로 유지(우하단을 같은 양만큼 이동).
        반환: (이전 앵커, 새 앵커)"""
        pics = self.pictures()
        if index >= len(pics):
            raise BiffError('그림 #%d 가 없습니다 (총 %d개).' % (index, len(pics)))
        ri, off, a = pics[index]
        x1, y1 = self.anchor_to_abs(a[1], a[2], a[3], a[4])
        x2, y2 = self.anchor_to_abs(a[5], a[6], a[7], a[8])
        nx1, ny1 = self.anchor_to_abs(col1, dx1, row1, dy1)
        nc2, ndx2, nr2, ndy2 = self.abs_to_anchor(nx1 + (x2 - x1), ny1 + (y2 - y1))
        new = (a[0], col1, dx1, row1, dy1, nc2, ndx2, nr2, ndy2)
        self.records[ri][1][off:off + 18] = struct.pack('<9H', *new)
        return a, new

    # -- 직렬화 --------------------------------------------------------------

    def _rebuild_sst(self):
        """SST(+CONTINUE) 와 EXTSST 를 재작성. 문자열은 레코드 경계를 넘지 않게 배치."""
        cst_total = sum(1 for i in range(self.sheet_bof, self.sheet_end)
                        if self.records[i][0] == LABELSST)
        if cst_total == 0:
            cst_total = self.sst_total
        payloads = [bytearray(struct.pack('<ii', cst_total, len(self.sst_raw)))]
        layout = []                      # (run_ordinal, offset_in_payload)
        for raw in self.sst_raw:
            if len(payloads[-1]) + len(raw) > MAX_PAYLOAD:
                payloads.append(bytearray())
            layout.append((len(payloads) - 1, len(payloads[-1])))
            payloads[-1] += raw
        newrun = [[SST, payloads[0]]]
        for p in payloads[1:]:
            newrun.append([CONTINUE, p])

        nb = (len(self.sst_raw) + SST_BUCKET - 1) // SST_BUCKET
        ext = bytearray(struct.pack('<H', SST_BUCKET) + b'\x00' * (nb * 8))

        i = self.i_sst
        old = self.sst_run_len
        self.records[i:i + old] = newrun
        self._shift(i + 1, len(newrun) - old)
        self.sst_run_len = len(newrun)
        if self.i_extsst is None:
            raise BiffError('EXTSST 레코드가 없습니다.')
        self.records[self.i_extsst] = [EXTSST, ext]
        self._sst_layout = layout

    def _sst_layout_from_current(self):
        layout = []
        off = 8
        for raw in self.sst_raw:
            layout.append((0, off))
            off += len(raw)
        return layout

    def build(self):
        # SST 는 항상 재작성 (문자열 추가 여부와 무관하게 동일 바이트가 나오도록 설계)
        self._rebuild_sst()

        pos = []
        p = 0
        for code, pay in self.records:
            pos.append(p)
            p += 4 + len(pay)
        total = p

        # BOUNDSHEET.lbPlyPos
        self.records[self.i_boundsheet][1][0:4] = struct.pack('<I', pos[self.sheet_bof])

        # EXTSST
        ext = self.records[self.i_extsst][1]
        nb = (len(self.sst_raw) + SST_BUCKET - 1) // SST_BUCKET
        for b in range(nb):
            si = b * SST_BUCKET
            ordinal, off = self._sst_layout[si]
            recpos = pos[self.i_sst + ordinal]
            ib = recpos + 4 + off
            ext[2 + b * 8:10 + b * 8] = struct.pack('<IHH', ib, ib - recpos, 0)

        # INDEX / DBCELL
        self._fix_index_dbcell(pos)

        out = bytearray(total)
        p = 0
        for code, pay in self.records:
            out[p:p + 4] = struct.pack('<HH', code, len(pay))
            out[p + 4:p + 4 + len(pay)] = pay
            p += 4 + len(pay)
        return bytes(out)

    def _fix_index_dbcell(self, pos):
        rowpos = {}
        firstcell = {}
        dbcells = []
        for i in range(self.sheet_bof, self.sheet_end):
            c, pay = self.records[i]
            if c == ROW:
                rowpos[struct.unpack('<H', bytes(pay[:2]))[0]] = pos[i]
            elif c == DBCELL:
                dbcells.append(i)
            elif c in CELL_RECS:
                r = self._cell_row(c, pay)
                if r not in firstcell:
                    firstcell[r] = pos[i]

        for b, di in enumerate(dbcells):
            base = b * ROWBLOCK
            if base not in rowpos:
                raise BiffError('행 블록 %d 의 첫 ROW 레코드가 없습니다.' % b)
            first_row_pos = rowpos[base]
            rows = [base + k for k in range(ROWBLOCK) if (base + k) in rowpos]
            if not rows:
                continue
            block_end = pos[di]
            # 뒤에서 앞으로 스캔: 빈 행은 다음 행의 셀 시작 위치를 갖는다
            p_of = {}
            cur = block_end
            for r in reversed(rows):
                if r in firstcell:
                    cur = firstcell[r]
                p_of[r] = cur
            rgdb = []
            prev = first_row_pos + 20
            for r in rows:
                rgdb.append(p_of[r] - prev)
                prev = p_of[r]
            pay = self.records[di][1]
            want = 4 + len(rgdb) * 2
            if len(pay) != want:
                raise BiffError('DBCELL 크기 불일치 (%d != %d)' % (len(pay), want))
            for v in rgdb:
                if v < 0 or v > 0xFFFF:
                    raise BiffError('DBCELL 오프셋 범위 초과: %d' % v)
            pay[0:4] = struct.pack('<I', pos[di] - first_row_pos)
            pay[4:] = struct.pack('<%dH' % len(rgdb), *rgdb)

        if self.i_index is not None:
            pay = self.records[self.i_index][1]
            n = (len(pay) - 16) // 4
            if n != len(dbcells):
                raise BiffError('INDEX/DBCELL 개수 불일치')
            if self.i_defcolwidth is not None:
                pay[12:16] = struct.pack('<I', pos[self.i_defcolwidth])
            pay[16:] = struct.pack('<%dI' % n, *[pos[i] for i in dbcells])


# ----------------------------------------------------------------------------
# 상위 진입점
# ----------------------------------------------------------------------------

class XlsTemplate(object):
    """원본 .xls 를 열어 셀만 바꾸고 새 .xls 바이트를 만든다."""

    STREAM = 'Workbook'

    def __init__(self, path):
        with open(path, 'rb') as f:
            raw = f.read()
        self.cfb = Cfb(raw)
        if self.STREAM not in self.cfb.streams:
            raise BiffError('Workbook 스트림이 없습니다.')
        self.wb = Workbook(self.cfb.streams[self.STREAM][1])

    def to_bytes(self):
        return self.cfb.write({self.STREAM: self.wb.build()})

    def save(self, path):
        data = self.to_bytes()
        with open(path, 'wb') as f:
            f.write(data)
        return len(data)
