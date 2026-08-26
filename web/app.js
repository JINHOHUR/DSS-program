'use strict';

const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
const HEADER_KEYS = ['quoteDate', 'quoteNo', 'customer', 'contact',
                     'subject', 'validity', 'delivery', 'payment', 'remark'];
const STORE = 'dss-quote-draft-v1';

let CFG = { maxItems: 36, defaults: {} };
let items = [];

/* ---------- utils ---------- */
const won = n => '₩' + Math.round(n || 0).toLocaleString('ko-KR');
const num = v => {
  const n = parseFloat(String(v == null ? '' : v).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
};
function today() {
  const d = new Date(), p = x => String(x).padStart(2, '0');
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}
function banner(kind, msg) {
  const b = $('#banner');
  b.className = 'banner ' + kind;
  b.textContent = msg;
  b.hidden = !msg;
}

/* ---------- items table ---------- */
function blankItem() { return { name: '', spec: '', unit: 'ea', qty: 1, price: 0 }; }

function renderItems() {
  const tb = $('#items tbody');
  tb.innerHTML = '';
  items.forEach((it, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td class="idx">' + (i + 1) + '</td>' +
      '<td><input type="text" data-f="name"  placeholder="품명"></td>' +
      '<td><input type="text" data-f="spec"  placeholder="규격 / 모델"></td>' +
      '<td><input type="text" data-f="unit"  placeholder="ea"></td>' +
      '<td><input type="text" data-f="qty"   class="num"></td>' +
      '<td><input type="text" data-f="price" class="num"></td>' +
      '<td class="amt"></td>' +
      '<td class="act">' +
        '<button class="btn icon" data-a="up"   title="위로">▲</button>' +
        '<button class="btn icon" data-a="down" title="아래로">▼</button>' +
        '<button class="btn icon" data-a="del"  title="삭제">✕</button>' +
      '</td>';
    $$('input', tr).forEach(inp => {
      const f = inp.dataset.f;
      inp.value = (f === 'qty' || f === 'price')
        ? (it[f] === '' ? '' : Number(it[f]).toLocaleString('ko-KR'))
        : it[f];
      inp.addEventListener('input', () => {
        it[f] = (f === 'qty' || f === 'price') ? num(inp.value) : inp.value;
        refreshRow(tr, it);
        refreshTotals();
        save();
      });
      if (f === 'qty' || f === 'price') {
        inp.addEventListener('focus', () => { inp.value = it[f] || ''; });
        inp.addEventListener('blur', () => {
          inp.value = it[f] === '' ? '' : Number(it[f]).toLocaleString('ko-KR');
        });
      }
      inp.addEventListener('keydown', e => {
        if (e.key === 'Enter' && i === items.length - 1) { e.preventDefault(); addItem(); }
      });
    });
    $$('button', tr).forEach(b => b.addEventListener('click', () => rowAction(i, b.dataset.a)));
    refreshRow(tr, it);
    tb.appendChild(tr);
  });
  $('#empty-msg').hidden = items.length > 0;
  $('#item-count').textContent = items.length;
  refreshTotals();
}

function refreshRow(tr, it) {
  $('.amt', tr).textContent = won(num(it.qty) * num(it.price));
}

function rowAction(i, a) {
  if (a === 'del') items.splice(i, 1);
  else if (a === 'up' && i > 0) items.splice(i - 1, 0, items.splice(i, 1)[0]);
  else if (a === 'down' && i < items.length - 1) items.splice(i + 1, 0, items.splice(i, 1)[0]);
  renderItems(); save();
}

function addItem() {
  if (items.length >= CFG.maxItems) {
    banner('err', '품목은 최대 ' + CFG.maxItems + '행까지 가능합니다. (원본 양식의 품목 영역 '
                  + CFG.itemRows[0] + '~' + CFG.itemRows[1] + '행)');
    return;
  }
  banner('', '');
  items.push(blankItem());
  renderItems();
  const rows = $$('#items tbody tr');
  const last = rows[rows.length - 1];
  if (last) $('input[data-f="name"]', last).focus();
  save();
}

/* ---------- totals ---------- */
function refreshTotals() {
  const rate = num($('#vatRate').value);
  const supply = items.reduce((s, it) => s + num(it.qty) * num(it.price), 0);
  const vat = supply * rate;
  $('#t-supply').textContent = won(supply);
  $('#t-vat').textContent = won(vat);
  $('#t-total').textContent = won(supply + vat);
}

/* ---------- payload ---------- */
/* 품명·규격·단위가 모두 비고 수량·단가도 없는 '완전히 빈 행'인가 */
function isEmptyItem(it) {
  if (String(it.name || '').trim()) return false;
  if (String(it.spec || '').trim()) return false;
  if (String(it.unit || '').trim()) return false;
  return num(it.qty) === 0 && num(it.price) === 0;
}

function pack(list) {
  const d = { items: list.map(it => ({
    name: it.name, spec: it.spec, unit: it.unit,
    qty: num(it.qty), price: num(it.price)
  })), vatRate: num($('#vatRate').value) };
  HEADER_KEYS.forEach(k => { d[k] = $('#' + k).value; });
  return d;
}

/* 서버로 보낼 내용 — 완전히 빈 행만 제외하고 기재한 것은 모두 보낸다 */
function payload() { return pack(items.filter(it => !isEmptyItem(it))); }
/* 브라우저 임시저장 — 작성 중인 빈 행까지 그대로 보관 */
function draft() { return pack(items); }

function save() { try { localStorage.setItem(STORE, JSON.stringify(draft())); } catch (e) {} }

function load(d) {
  HEADER_KEYS.forEach(k => { if (d[k] != null) $('#' + k).value = d[k]; });
  if (d.vatRate != null) $('#vatRate').value = d.vatRate;
  items = (d.items || []).map(it => ({
    name: it.name || '', spec: it.spec || '', unit: it.unit || '',
    qty: it.qty == null ? '' : it.qty, price: it.price == null ? '' : it.price
  }));
  renderItems();
}

/* ---------- download ---------- */
async function download() {
  const btn = $('#btn-download');
  const label = $('.dl-label', btn);
  const d = payload();
  const filled = d.items.length;
  if (!filled) { banner('err', '내용이 입력된 품목이 없습니다. 품목을 1행 이상 채워주세요.'); return; }

  btn.disabled = true; label.textContent = '생성 중…';
  try {
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(d)
    });
    if (!res.ok) {
      let msg = 'HTTP ' + res.status;
      try { msg = (await res.json()).error || msg; } catch (e) {}
      throw new Error(msg);
    }
    const blob = await res.blob();
    const cd = res.headers.get('Content-Disposition') || '';
    const m = /filename\*=UTF-8''([^;]+)/.exec(cd);
    const name = m ? decodeURIComponent(m[1]) : '견적서.xls';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    const wrote = +res.headers.get('X-Quote-Items');
    const dropped = items.length - filled;
    banner('ok', '다운로드 완료 — ' + name + '\n'
      + '품목 ' + wrote + '행 기재'
      + (dropped ? ' (빈 행 ' + dropped + '개 제외)' : '')
      + '  ·  공급가액 ' + won(+res.headers.get('X-Quote-Supply'))
      + '  ·  부가세 ' + won(+res.headers.get('X-Quote-Vat'))
      + '  ·  총 견적금액 ' + won(+res.headers.get('X-Quote-Total'))
      + '\n(서버의 out\\ 폴더에도 동일한 사본이 저장되었습니다)');
  } catch (e) {
    banner('err', '생성 실패\n' + e.message);
  } finally {
    btn.disabled = false; label.textContent = 'Excel 다운로드 (.xls)';
  }
}

/* ---------- sample ---------- */
const SAMPLE = {
  quoteNo: 'DSS-2026-0823-01',
  quoteDate: today(),
  customer: '㈜에스케이하이닉스 청주',
  contact: '김철수 책임',
  subject: 'RF Generator 부속 케이블 SET (60kW)',
  delivery: '발주 후 6주',
  payment: '납품 후 익월 말 현금 결제',
  remark: '설치 및 시운전 비용 별도 / 납품 장소: 청주 M15',
  vatRate: 0.1,
  items: [
    { name: 'INPUT CABLE (RFG)', spec: 'KR16840YA427(40M)', unit: 'ea', qty: 2, price: 1250000 },
    { name: 'INPUT CABLE (T/C)', spec: 'KR16840YA419(40M)', unit: 'ea', qty: 1, price: 1180000 },
    { name: 'INTERFACE', spec: 'KR16840YA620(20M)', unit: 'ea', qty: 1, price: 640000 },
    { name: 'DEVICE NET(RFG)', spec: 'KR16840YA599', unit: 'ea', qty: 2, price: 315000 },
    { name: 'MB-SIGNAL', spec: 'KR16840YA531(2.5M)', unit: 'ea', qty: 4, price: 187500 }
  ]
};

/* ---------- boot ---------- */
(async function boot() {
  try {
    CFG = await (await fetch('/api/config')).json();
  } catch (e) {
    banner('err', '백엔드에 연결하지 못했습니다. server.py 가 실행 중인지 확인하세요.');
  }
  $('#item-max').textContent = CFG.maxItems;
  Object.entries(CFG.headerCells || {}).forEach(([k, v]) => {
    const el = $('#c-' + k); if (el) el.textContent = v;
  });
  Object.entries(CFG.defaults || {}).forEach(([k, v]) => {
    const el = $('#' + k); if (el && !el.value) el.value = v;
  });
  $('#quoteDate').value = today();

  let draft = null;
  try { draft = JSON.parse(localStorage.getItem(STORE) || 'null'); } catch (e) {}
  if (draft && (draft.items || []).length) {
    load(draft);
    banner('info', '이전 작성 내용을 불러왔습니다. 새로 시작하려면 [초기화] 를 누르세요.');
  } else {
    renderItems();
  }

  $('#btn-add').addEventListener('click', addItem);
  $('#btn-download').addEventListener('click', download);
  $('#btn-sample').addEventListener('click', () => { load(SAMPLE); banner('', ''); save(); });
  $('#btn-reset').addEventListener('click', () => {
    localStorage.removeItem(STORE);
    HEADER_KEYS.forEach(k => { $('#' + k).value = ''; });
    Object.entries(CFG.defaults || {}).forEach(([k, v]) => {
      const el = $('#' + k); if (el) el.value = v;
    });
    $('#quoteDate').value = today();
    $('#vatRate').value = (CFG.defaults || {}).vatRate || 0.1;
    items = []; renderItems(); banner('', '');
  });
  $('#vatRate').addEventListener('input', () => { refreshTotals(); save(); });
  HEADER_KEYS.forEach(k => $('#' + k).addEventListener('input', save));
})();
