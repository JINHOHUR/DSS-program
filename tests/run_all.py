# -*- coding: utf-8 -*-
"""전체 검증 실행:  python tests/run_all.py"""
import os
import sys
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
TESTS = [
    ('1. 무편집 라운드트립 (원본 바이트 동일성)', 'test_roundtrip.py'),
    ('2. 실데이터 견적서 생성 + 원본 대비 전수 검증', 'test_generate.py'),
    ('3. 백엔드 HTTP 종단 테스트', 'test_server.py'),
    ('4. 한계/예외 상황', 'test_stress.py'),
]

env = dict(os.environ, PYTHONIOENCODING='utf-8')
results = []
for title, f in TESTS:
    print('\n' + '=' * 70)
    print(' ' + title)
    print('=' * 70)
    r = subprocess.run([sys.executable, os.path.join(HERE, f)], env=env)
    results.append((title, r.returncode == 0))

print('\n' + '=' * 70)
print(' 전체 결과')
print('=' * 70)
for title, ok in results:
    print('  %-4s %s' % ('PASS' if ok else 'FAIL', title))
sys.exit(0 if all(ok for _, ok in results) else 1)
