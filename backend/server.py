# -*- coding: utf-8 -*-
"""
server.py — 견적서 프로그램 로컬 백엔드

외부 라이브러리 없이 파이썬 표준 라이브러리만 사용한다.
  GET  /                 웹 UI
  GET  /api/config       매핑 라벨 / 기본값 / 품목 최대 행 수
  POST /api/calc         합계 계산 (서버 규칙으로 재확인용)
  POST /api/generate     입력 JSON -> 원본 양식 기반 .xls 다운로드

실행:  python backend/server.py   또는  start.bat
"""

import io
import os
import sys
import json
import socket
import traceback
import webbrowser
from urllib.parse import quote, urlparse
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
WEB = os.path.join(ROOT, 'web')
OUT = os.path.join(ROOT, 'out')
sys.path.insert(0, HERE)

from quote_writer import QuoteWriter, suggest_filename, load_mapping   # noqa: E402

MIME = {'.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
        '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
        '.svg': 'image/svg+xml', '.ico': 'image/x-icon'}


class Handler(BaseHTTPRequestHandler):
    server_version = 'QuoteApp/1.0'

    def log_message(self, fmt, *args):
        sys.stderr.write('  %s\n' % (fmt % args))

    # -- helpers ---------------------------------------------------------
    def _send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False))

    def _body(self):
        n = int(self.headers.get('Content-Length') or 0)
        return json.loads(self.rfile.read(n).decode('utf-8')) if n else {}

    # -- routes ----------------------------------------------------------
    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/':
            path = '/index.html'
        if path == '/api/config':
            m = load_mapping()
            w = QuoteWriter(m)
            return self._json(200, {
                'maxItems': w.max_items(),
                'defaults': m['defaults'],
                'itemRows': [m['items']['firstRow'], m['items']['lastRow']],
                'headerCells': {k: v['cell'] for k, v in m['header'].items()},
                'itemColumns': {k: v['col'] for k, v in m['items']['columns'].items()},
                'totalCells': {k: v['cell'] for k, v in m['totals'].items()},
                'addressMode': m.get('address', {}).get('mode'),
            })

        fp = os.path.normpath(os.path.join(WEB, path.lstrip('/')))
        if not fp.startswith(WEB) or not os.path.isfile(fp):
            return self._send(404, '404 Not Found', 'text/plain; charset=utf-8')
        ext = os.path.splitext(fp)[1].lower()
        with open(fp, 'rb') as f:
            return self._send(200, f.read(), MIME.get(ext, 'application/octet-stream'))

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            data = self._body()
        except Exception as e:
            return self._json(400, {'error': 'JSON 파싱 실패: %s' % e})

        try:
            if path == '/api/calc':
                w = QuoteWriter()
                s, v, t, rows = w.calculate(data)
                return self._json(200, {'supply': s, 'vat': v, 'total': t, 'rows': rows})

            if path == '/api/generate':
                w = QuoteWriter()
                errs = w.validate(data)
                if errs:
                    return self._json(400, {'error': '\n'.join(errs)})
                blob, res = w.generate(data)
                n_items = len(w.clean_items(data))
                name = suggest_filename(data)
                if not os.path.isdir(OUT):
                    os.makedirs(OUT)
                with open(os.path.join(OUT, name), 'wb') as f:   # 서버에도 사본 보관
                    f.write(blob)
                q = quote(name.encode('utf-8'))
                return self._send(200, blob, 'application/vnd.ms-excel', {
                    'Content-Disposition': "attachment; filename=quote.xls; "
                                           "filename*=UTF-8''%s" % q,
                    'X-Quote-Supply': str(int(res.supply)),
                    'X-Quote-Vat': str(int(res.vat)),
                    'X-Quote-Total': str(int(res.total)),
                    'X-Quote-Items': str(n_items),
                })
        except ValueError as e:
            return self._json(400, {'error': str(e)})
        except Exception:
            traceback.print_exc()
            return self._json(500, {'error': traceback.format_exc(limit=3)})

        return self._json(404, {'error': 'unknown endpoint'})


def free_port(start=8765):
    for p in range(start, start + 40):
        s = socket.socket()
        try:
            s.bind(('127.0.0.1', p))
            return p
        except OSError:
            continue
        finally:
            s.close()
    raise RuntimeError('사용 가능한 포트를 찾지 못했습니다.')


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    no_browser = '--no-browser' in sys.argv
    port = int(args[0]) if args else free_port()
    srv = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    url = 'http://127.0.0.1:%d/' % port
    print('=' * 64)
    print(' 견적서 프로그램')
    print(' 템플릿 : %s' % QuoteWriter().template_path)
    print(' 주소   : %s' % url)
    print(' 종료   : Ctrl+C')
    print('=' * 64)
    if not no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\n종료합니다.')


if __name__ == '__main__':
    main()
