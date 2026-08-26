# -*- coding: utf-8 -*-
"""
검증 1 — 무편집 라운드트립
원본 .xls 를 읽어 그대로 다시 쓴 결과가 원본과 바이트 단위로 동일한지 확인한다.
통과하면 다음이 증명된다.
  * CFB(OLE2) 컨테이너 재작성 로직이 원본과 동일한 파일을 만든다
  * BIFF 레코드 파서/직렬화기가 무손실이다
  * BOUNDSHEET / INDEX / DBCELL / EXTSST 오프셋 재계산 규칙이 원본 규칙과 일치한다
  * 이미지(MSODRAWINGGROUP/MSODRAWING/OBJ), XF, 병합셀, 인쇄설정이 원본 그대로다
"""
import os
import sys
import struct

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'backend'))

from xlsbiff import XlsTemplate, Cfb   # noqa: E402

TPL = os.path.join(HERE, '..', 'template', '견적서.xls')


def main():
    orig = open(TPL, 'rb').read()
    t = XlsTemplate(TPL)
    out = t.to_bytes()

    print('원본     : %d bytes' % len(orig))
    print('재작성   : %d bytes' % len(out))

    ok = True

    # 1) 스트림 단위 비교
    a = Cfb(orig).streams
    b = Cfb(out).streams
    print()
    print('--- 스트림 비교 ---')
    for name in a:
        da = a[name][1]
        db = b.get(name, (None, b''))[1]
        same = da == db
        ok &= same
        print('  %-32r %6d -> %6d  %s' % (name, len(da), len(db), 'OK' if same else '불일치'))

    # 2) 전체 파일 바이트 비교
    print()
    if out == orig:
        print('*** 전체 파일 바이트 100%% 동일 ***')
    else:
        ok = False
        print('!!! 전체 파일 불일치 !!!')
        n = min(len(out), len(orig))
        diffs = [i for i in range(n) if out[i] != orig[i]]
        print('  길이 %d vs %d, 다른 바이트 %d개' % (len(orig), len(out), len(diffs)))
        for i in diffs[:40]:
            print('    off %6d (sector %3d+%3d): %02X -> %02X'
                  % (i, (i - 512) // 512, (i - 512) % 512, orig[i], out[i]))
    print()
    print('결과:', 'PASS' if ok else 'FAIL')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
