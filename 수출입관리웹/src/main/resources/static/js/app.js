/* 공용 동작: 행 클릭 이동 · 확인 대화 · 안건 화면 탭/단계 필드 · 신용장 자동계산 */
(function () {
  // 줄 클릭 → 이동
  document.querySelectorAll('.rowlink[data-href]').forEach(function (tr) {
    tr.addEventListener('click', function (e) {
      if (e.target.closest('a,button,input,label')) return;
      location.href = tr.dataset.href;
    });
  });
  // form[data-confirm]
  document.querySelectorAll('form[data-confirm]').forEach(function (f) {
    f.addEventListener('submit', function (e) { if (!confirm(f.dataset.confirm)) e.preventDefault(); });
  });

  var body = document.body;
  if (!body.dataset.tab) return;               // 안건 상세 화면이 아니면 끝

  // ---- 탭 전환 (같은 form 안에서 필드를 옮기므로 값이 유지됨)
  var fields = (body.dataset.fields || '').split(',').filter(Boolean);
  var stageBox = document.getElementById('stagefields');
  var homes = {};                                // key -> 원래 부모/다음 형제
  function moveToStage() {
    var n = 0;
    fields.forEach(function (k) {
      var row = document.getElementById('row-' + k);
      if (!row) return;
      if (!homes[k]) homes[k] = { parent: row.parentNode, next: row.nextSibling };
      stageBox.appendChild(row); n++;
    });
    document.getElementById('nofields').hidden = n > 0;
  }
  function moveBack() {
    fields.forEach(function (k) {
      var row = document.getElementById('row-' + k), h = homes[k];
      if (row && h) h.parent.insertBefore(row, h.next);
    });
  }
  function setTab(t) {
    if (t === 'flow') moveToStage(); else moveBack();
    body.dataset.tab = t;
    document.getElementById('tabField').value = t;
    document.querySelectorAll('.tab').forEach(function (a) { a.classList.toggle('on', a.dataset.tab === t); });
  }
  document.querySelectorAll('.tab').forEach(function (a) {
    a.addEventListener('click', function (e) { e.preventDefault(); setTab(a.dataset.tab); });
  });
  setTab(body.dataset.tab);

  // ---- 신용장 기한 자동계산
  var btn = document.getElementById('btnLc');
  if (btn) btn.addEventListener('click', function () {
    var fca = document.getElementById('fcaDate').value;
    if (!fca) { alert('먼저 FCA(선적예정일)를 입력하세요.'); return; }
    var cust = (document.getElementById('customerCode') || {}).value || '';
    fetch('/deals/lc-calc?fca=' + fca + '&customer=' + encodeURIComponent(cust))
      .then(function (r) { return r.json(); })
      .then(function (d) {
        document.getElementById('latestShipment').value = d.latestShipment;
        document.getElementById('expiryDate').value = d.expiryDate;
        alert((d.custom ? '고객사 전용 규칙' : '기본 규칙') + ' (+' + d.shipDays + '일 / +' + d.expiryDays + '일) 적용\n\n'
          + 'Latest shipment : ' + d.latestShipment + '\nExpiry date : ' + d.expiryDate
          + '\n\n개설요청서 수령 기한 : ' + d.requestDue);
      });
  });

  // 완료 체크 시 완료일 비어 있으면 오늘
  var done = document.querySelector('input[name=done]');
  if (done) done.addEventListener('change', function () {
    var dd = document.querySelector('input[name=doneDate]');
    if (done.checked && !dd.value) dd.value = new Date().toISOString().slice(0, 10);
    if (!done.checked) dd.value = '';
  });
})();
