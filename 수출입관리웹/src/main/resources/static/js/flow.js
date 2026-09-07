/* 업무 플로우차트 — 자유 배치 캔버스 편집기 (데스크톱 판을 웹으로 옮김) */
(function () {
  var NODE_W = 212, NODE_H = 86, SNAP = 5;
  var COLORS = { 0: '#0F62FE', 1: '#0F62FE', 2: '#5B4FCF', 3: '#5B4FCF', 4: '#D9820B', 5: '#1E8E4E', 6: '#1E8E4E', 7: '#1E8E4E', 8: '#0E7C86' };
  var cv = document.getElementById('flowCanvas'), ctx = cv.getContext('2d');
  var wrap = cv.parentNode;
  var nodes = [], edges = [], counts = {};
  var scale = 1, ox = 0, oy = 0;                  // 배율, 화면 오프셋(px)
  var sel = null;                                  // {kind:'node'|'edge', id}
  var drag = null, connect = false, connectFrom = null, panning = null;
  var detailOn = true;

  function api(method, url, body) {
    var o = { method: method, headers: {} };
    if (body !== undefined) { o.headers['Content-Type'] = 'application/json'; o.body = JSON.stringify(body); }
    return fetch(url, o).then(function (r) { return r.json(); });
  }
  function load(keepSel) {
    return api('GET', '/api/flow').then(function (d) {
      nodes = d.nodes; edges = d.edges; counts = d.counts || {};
      if (!keepSel) sel = null;
      draw(); showDetail();
    });
  }
  function nodeById(id) { for (var i = 0; i < nodes.length; i++) if (nodes[i].id === id) return nodes[i]; return null; }
  function color(n) { if (n.color) return n.color; if (n.stageNo == null || n.stageNo < 0) return '#5A6B7D'; return COLORS[n.stageNo] || '#5A6B7D'; }

  // ---- 좌표 --------------------------------------------------------------
  function resize() {
    var r = wrap.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
    cv.width = r.width * dpr; cv.height = r.height * dpr;
    cv.style.width = r.width + 'px'; cv.style.height = r.height + 'px';
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    draw();
  }
  function toModel(px, py) { return { x: (px - ox) / scale, y: (py - oy) / scale }; }
  function center(n) { return { x: n.x + NODE_W / 2, y: n.y + NODE_H / 2 }; }
  function anchor(a, b) {
    var ca = center(a), cb = center(b), dx = cb.x - ca.x, dy = cb.y - ca.y;
    if (!dx && !dy) return ca;
    var w2 = NODE_W / 2, h2 = NODE_H / 2, k;
    if (a.shape === 'diamond') k = 1 / (Math.abs(dx) / w2 + Math.abs(dy) / h2);
    else if (a.shape === 'oval') k = 1 / Math.sqrt((dx / w2) * (dx / w2) + (dy / h2) * (dy / h2));
    else k = Math.min(dx ? w2 / Math.abs(dx) : 1e9, dy ? h2 / Math.abs(dy) : 1e9);
    return { x: ca.x + dx * k, y: ca.y + dy * k };
  }
  function hit(px, py) {
    var m = toModel(px, py);
    for (var i = nodes.length - 1; i >= 0; i--) {
      var n = nodes[i];
      if (m.x >= n.x && m.x <= n.x + NODE_W && m.y >= n.y && m.y <= n.y + NODE_H) return { kind: 'node', id: n.id };
    }
    for (var j = 0; j < edges.length; j++) {
      var e = edges[j], a = nodeById(e.src), b = nodeById(e.dst); if (!a || !b) continue;
      var p1 = anchor(a, b), p2 = anchor(b, a);
      if (distSeg(m, p1, p2) <= 7 / scale) return { kind: 'edge', id: e.id };
    }
    return null;
  }
  function distSeg(p, a, b) {
    var l2 = (b.x - a.x) * (b.x - a.x) + (b.y - a.y) * (b.y - a.y); if (!l2) return Math.hypot(p.x - a.x, p.y - a.y);
    var t = Math.max(0, Math.min(1, ((p.x - a.x) * (b.x - a.x) + (p.y - a.y) * (b.y - a.y)) / l2));
    return Math.hypot(p.x - (a.x + t * (b.x - a.x)), p.y - (a.y + t * (b.y - a.y)));
  }

  // ---- 그리기 -------------------------------------------------------------
  function draw() {
    var W = cv.clientWidth, H = cv.clientHeight;
    ctx.clearRect(0, 0, W, H);
    ctx.save(); ctx.translate(ox, oy); ctx.scale(scale, scale);
    edges.forEach(drawEdge);
    nodes.forEach(drawNode);
    ctx.restore();
    document.getElementById('fZoom').textContent = Math.round(scale * 100) + '%';
  }
  function drawEdge(e) {
    var a = nodeById(e.src), b = nodeById(e.dst); if (!a || !b) return;
    var p1 = anchor(a, b), p2 = anchor(b, a), isSel = sel && sel.kind === 'edge' && sel.id === e.id;
    ctx.strokeStyle = isSel ? '#0F62FE' : '#93A2B4'; ctx.lineWidth = isSel ? 3 : 2; ctx.fillStyle = ctx.strokeStyle;
    ctx.beginPath(); ctx.moveTo(p1.x, p1.y); ctx.lineTo(p2.x, p2.y); ctx.stroke();
    var ang = Math.atan2(p2.y - p1.y, p2.x - p1.x), L = 13;
    ctx.beginPath(); ctx.moveTo(p2.x, p2.y);
    ctx.lineTo(p2.x - L * Math.cos(ang - 0.4), p2.y - L * Math.sin(ang - 0.4));
    ctx.lineTo(p2.x - L * Math.cos(ang + 0.4), p2.y - L * Math.sin(ang + 0.4)); ctx.closePath(); ctx.fill();
    if (e.label) {
      var mx = (p1.x + p2.x) / 2, my = (p1.y + p2.y) / 2;
      ctx.font = '12px "Malgun Gothic", sans-serif'; var tw = ctx.measureText(e.label).width;
      ctx.fillStyle = '#FAFBFD'; ctx.fillRect(mx - tw / 2 - 5, my - 9, tw + 10, 18);
      ctx.fillStyle = '#1B2430'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(e.label, mx, my);
    }
  }
  function roundRect(x, y, w, h, r) { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r); ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath(); }
  function drawNode(n) {
    var x = n.x, y = n.y, w = NODE_W, h = NODE_H, col = color(n);
    var isSel = sel && sel.kind === 'node' && sel.id === n.id, isFrom = connectFrom === n.id;
    var bg = isSel ? '#EAF1FD' : (isFrom ? '#E6F6EC' : '#FFFFFF');
    var outline = isSel ? '#0F62FE' : (isFrom ? '#1E8E4E' : '#D5DCE5');
    ctx.lineWidth = (isSel || isFrom) ? 3 : 1; ctx.strokeStyle = outline; ctx.fillStyle = bg;
    if (n.shape === 'oval') { ctx.beginPath(); ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2); ctx.fill(); ctx.stroke(); }
    else if (n.shape === 'diamond') { ctx.beginPath(); ctx.moveTo(x + w / 2, y); ctx.lineTo(x + w, y + h / 2); ctx.lineTo(x + w / 2, y + h); ctx.lineTo(x, y + h / 2); ctx.closePath(); ctx.fill(); ctx.stroke(); }
    else { roundRect(x, y, w, h, 3); ctx.fill(); ctx.stroke(); ctx.fillStyle = col; ctx.fillRect(x, y, 6, h); }
    var detail = scale >= 0.62;
    ctx.fillStyle = n.shape === 'box' ? '#1B2430' : col;
    ctx.font = 'bold 13px "Malgun Gothic", sans-serif'; ctx.textBaseline = 'top';
    if (n.shape === 'box') { ctx.textAlign = 'left'; ctx.fillText(clip(n.title || '(이름 없음)', w - 34), x + 16, detail ? y + 12 : y + h / 2 - 8); }
    else { ctx.textAlign = 'center'; ctx.fillText(clip(n.title || '(이름 없음)', w - 50), x + w / 2, detail ? y + 18 : y + h / 2 - 8); }
    if (n.shape === 'box' && detail) {
      ctx.fillStyle = '#6B7787'; ctx.font = '11px "Malgun Gothic", sans-serif'; ctx.textAlign = 'left';
      ctx.fillText(clip((n.descr || '').replace(/\n/g, ' '), w - 34), x + 16, y + 36);
      if (n.store) { ctx.fillStyle = '#8A96A4'; ctx.font = '10px "Malgun Gothic", sans-serif'; ctx.fillText(clip('📁 ' + n.store, w - 34), x + 16, y + h - 20); }
    }
    var cnt = (n.stageNo != null && n.stageNo >= 0) ? (counts[n.stageNo] || 0) : 0;
    if (cnt) {
      ctx.fillStyle = col; ctx.beginPath(); ctx.arc(x + w - 21, y + 19, 11, 0, Math.PI * 2); ctx.fill();
      ctx.fillStyle = '#fff'; ctx.font = 'bold 11px sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(String(cnt), x + w - 21, y + 19);
      ctx.textBaseline = 'top';
    }
  }
  function clip(s, w) { if (ctx.measureText(s).width <= w) return s; while (s.length && ctx.measureText(s + '…').width > w) s = s.slice(0, -1); return s + '…'; }

  // ---- 확대·축소 ------------------------------------------------------------
  function setScale(s, cx, cy) {
    s = Math.max(0.3, Math.min(2, s));
    if (cx == null) { cx = cv.clientWidth / 2; cy = cv.clientHeight / 2; }
    var m = toModel(cx, cy); scale = s; ox = cx - m.x * scale; oy = cy - m.y * scale; draw();
  }
  function fit() {
    if (!nodes.length) return;
    var maxX = 0, maxY = 0, minX = 1e9, minY = 1e9;
    nodes.forEach(function (n) { minX = Math.min(minX, n.x); minY = Math.min(minY, n.y); maxX = Math.max(maxX, n.x + NODE_W); maxY = Math.max(maxY, n.y + NODE_H); });
    var W = cv.clientWidth, H = cv.clientHeight;
    scale = Math.min(1.3, (W - 50) / (maxX - minX), (H - 50) / (maxY - minY));
    ox = (W - (maxX - minX) * scale) / 2 - minX * scale; oy = (H - (maxY - minY) * scale) / 2 - minY * scale; draw();
  }

  // ---- 마우스 -------------------------------------------------------------
  function pos(e) { var r = cv.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; }
  cv.addEventListener('mousedown', function (e) {
    var p = pos(e), h = hit(p.x, p.y);
    if (connect) {
      if (h && h.kind === 'node') {
        if (connectFrom == null) { connectFrom = h.id; hint('연결 그리기 : 도착할 항목을 클릭하세요 (Esc 취소)'); draw(); }
        else if (h.id !== connectFrom) api('POST', '/api/flow/edge?src=' + connectFrom + '&dst=' + h.id).then(function () { connectFrom = null; hint('연결 그리기 : 시작할 항목을 클릭하세요 (Esc 취소)'); load(true); });
      } else cancelConnect();
      return;
    }
    if (h && h.kind === 'node') {
      sel = h; var n = nodeById(h.id), m = toModel(p.x, p.y);
      drag = { id: h.id, dx: m.x - n.x, dy: m.y - n.y, moved: false }; draw(); showDetail();
    } else if (h) { sel = h; draw(); showDetail(); }
    else { sel = null; panning = { x: p.x - ox, y: p.y - oy }; draw(); showDetail(); }
  });
  cv.addEventListener('mousemove', function (e) {
    var p = pos(e);
    if (drag) { var n = nodeById(drag.id), m = toModel(p.x, p.y); n.x = Math.max(0, m.x - drag.dx); n.y = Math.max(0, m.y - drag.dy); drag.moved = true; draw(); }
    else if (panning) { ox = p.x - panning.x; oy = p.y - panning.y; draw(); }
  });
  window.addEventListener('mouseup', function () {
    if (drag && drag.moved) { var n = nodeById(drag.id); n.x = Math.round(n.x / SNAP) * SNAP; n.y = Math.round(n.y / SNAP) * SNAP; api('POST', '/api/flow/node/' + n.id + '/pos?x=' + n.x + '&y=' + n.y); draw(); }
    drag = null; panning = null;
  });
  cv.addEventListener('dblclick', function (e) {
    var p = pos(e), h = hit(p.x, p.y);
    if (h && h.kind === 'node') openEditor(nodeById(h.id));
    else if (h && h.kind === 'edge') { var ed = edges.filter(function (x) { return x.id === h.id; })[0]; var l = prompt('연결선에 표시할 글자 (예: 승인, 반려)', ed.label || ''); if (l !== null) api('POST', '/api/flow/edge/' + h.id + '/label?label=' + encodeURIComponent(l)).then(function () { load(true); }); }
  });
  cv.addEventListener('wheel', function (e) {
    e.preventDefault(); var p = pos(e);
    if (e.ctrlKey) setScale(scale * (e.deltaY < 0 ? 1.1 : 1 / 1.1), p.x, p.y);
    else if (e.shiftKey) { ox -= e.deltaY; draw(); } else { oy -= e.deltaY; draw(); }
  }, { passive: false });
  window.addEventListener('keydown', function (e) { if (e.key === 'Escape') cancelConnect(); });
  window.addEventListener('resize', resize);

  function hint(t) { document.getElementById('fHint').textContent = t; }
  function cancelConnect() { connect = false; connectFrom = null; hint(''); document.getElementById('fConn').classList.remove('on'); draw(); }

  // ---- 툴바 --------------------------------------------------------------
  document.getElementById('fConn').addEventListener('click', function () {
    connect = !connect; connectFrom = null; this.classList.toggle('on', connect);
    hint(connect ? '연결 그리기 : 시작할 항목을 클릭하세요 (Esc 취소)' : ''); draw();
  });
  document.getElementById('fAdd').addEventListener('click', function () { openEditor(null); });
  document.getElementById('fEdit').addEventListener('click', function () { if (sel && sel.kind === 'node') openEditor(nodeById(sel.id)); else alert('수정할 항목을 클릭해서 고르세요.'); });
  document.getElementById('fDel').addEventListener('click', function () {
    if (!sel) { alert('삭제할 항목이나 연결선을 클릭해서 고르세요.'); return; }
    if (sel.kind === 'edge') { if (confirm('선택한 연결선을 삭제할까요?')) api('DELETE', '/api/flow/edge/' + sel.id).then(function () { load(); }); }
    else { var n = nodeById(sel.id); if (confirm("'" + n.title + "' 항목과 연결선을 삭제할까요?")) api('DELETE', '/api/flow/node/' + sel.id).then(function () { load(); }); }
  });
  document.getElementById('fAuto').addEventListener('click', function () { if (confirm('모든 항목을 순서대로 격자 배치합니다. 직접 옮긴 위치는 사라집니다.')) api('POST', '/api/flow/auto-layout').then(function () { load(); fit(); }); });
  document.getElementById('fReset').addEventListener('click', function () { if (confirm('플로우를 엑셀 기준 0~8단계로 되돌립니다. 직접 추가·수정한 항목과 연결선은 사라집니다.')) api('POST', '/api/flow/reset').then(function () { load(); fit(); }); });
  document.getElementById('fFit').addEventListener('click', fit);
  document.getElementById('fZoomIn').addEventListener('click', function () { setScale(scale * 1.25); });
  document.getElementById('fZoomOut').addEventListener('click', function () { setScale(scale * 0.8); });
  document.getElementById('fDetail').addEventListener('click', function () {
    detailOn = !detailOn; document.querySelector('.flow-wrap').classList.toggle('wide', !detailOn);
    this.textContent = detailOn ? '상세 숨기기' : '상세 보이기'; setTimeout(function () { resize(); fit(); }, 30);
  });

  // ---- 편집 창 ------------------------------------------------------------
  var modal = document.getElementById('nodeModal');
  function openEditor(n) {
    document.getElementById('mTitle').textContent = n ? '플로우 항목 수정' : '플로우 항목 추가';
    document.getElementById('mId').value = n ? n.id : '';
    document.getElementById('mName').value = n ? n.title : '';
    document.getElementById('mStore').value = n ? (n.store || '') : '';
    document.getElementById('mDesc').value = n ? (n.descr || '') : '';
    document.getElementById('mTip').value = n ? (n.tip || '') : '';
    document.getElementById('mShape').value = n ? (n.shape || 'box') : 'box';
    document.getElementById('mColor').value = n ? (n.color || '') : '';
    document.getElementById('mFolder').value = n ? (n.folder || '') : '';
    document.getElementById('mStage').value = n && n.stageNo != null ? n.stageNo : -1;
    modal.hidden = false; document.getElementById('mName').focus();
  }
  document.getElementById('mCancel').addEventListener('click', function () { modal.hidden = true; });
  document.getElementById('mSave').addEventListener('click', function () {
    var title = document.getElementById('mName').value.trim(); if (!title) { alert('항목 이름을 입력하세요.'); return; }
    var id = document.getElementById('mId').value, cur = id ? nodeById(+id) : null;
    api('POST', '/api/flow/node', {
      id: id ? +id : null, title: title, store: document.getElementById('mStore').value, descr: document.getElementById('mDesc').value,
      tip: document.getElementById('mTip').value, shape: document.getElementById('mShape').value, color: document.getElementById('mColor').value,
      folder: document.getElementById('mFolder').value, stageNo: +document.getElementById('mStage').value,
      x: cur ? cur.x : null, y: cur ? cur.y : null
    }).then(function (n) { modal.hidden = true; sel = { kind: 'node', id: n.id }; load(true); });
  });

  // ---- 상세 패널 -----------------------------------------------------------
  function showDetail() {
    var t = document.getElementById('dTitle'), tx = document.getElementById('dText'), fo = document.getElementById('dFolder'), op = document.getElementById('dOpen');
    var fb = document.querySelector('#dFiles tbody'), db = document.querySelector('#dDeals tbody');
    fb.innerHTML = ''; db.innerHTML = ''; op.hidden = true; fo.textContent = '';
    if (!sel) { t.textContent = '항목을 클릭하세요'; tx.innerHTML = '· 박스를 끌어서 원하는 위치에 놓으세요<br>· [🔗 연결 그리기] 로 항목끼리 화살표를 잇습니다<br>· 박스를 더블클릭하면 내용을 고칠 수 있습니다'; return; }
    if (sel.kind === 'edge') {
      var e = edges.filter(function (x) { return x.id === sel.id; })[0], a = nodeById(e.src), b = nodeById(e.dst);
      t.textContent = '연결선'; tx.textContent = (a ? a.title : '?') + '\n    ↓  ' + (e.label || '') + '\n' + (b ? b.title : '?') + '\n\n· 더블클릭하면 글자를 넣을 수 있습니다\n· [삭제] 를 누르면 이 연결선이 지워집니다'; return;
    }
    var n = nodeById(sel.id); t.textContent = n.title;
    tx.textContent = (n.descr || '') + (n.tip ? '\n\n✔ ' + n.tip : '') + (n.store ? '\n\n저장소 : ' + n.store : '');
    api('GET', '/api/flow/node/' + n.id + '/detail').then(function (d) {
      fo.textContent = d.folder ? '📁 ' + d.folder : '';
      op.hidden = false; op.href = '/archive?path=' + encodeURIComponent(d.folder || '');
      d.files.forEach(function (f) {
        var tr = document.createElement('tr');
        tr.innerHTML = '<td>' + (f.dir ? '<a href="/archive?path=' + encodeURIComponent(f.rel) + '">📁 ' + esc(f.name) + '</a>' : '<a href="/archive/download?path=' + encodeURIComponent(f.rel) + '" target="_blank">' + esc(f.name) + '</a>') + '</td><td>' + esc(f.kind) + '</td><td class="r">' + esc(f.size) + '</td><td class="mono">' + esc(f.mtime) + '</td>';
        fb.appendChild(tr);
      });
      if (!d.files.length) fb.innerHTML = '<tr><td colspan="4" class="muted center">자료 없음</td></tr>';
      d.deals.forEach(function (x) {
        var tr = document.createElement('tr'); tr.className = 'rowlink ' + x.tag; tr.dataset.href = '/deals/' + x.id;
        tr.innerHTML = '<td class="mono">' + esc(x.docNo) + '</td><td>' + esc(x.customer) + '</td><td>' + esc(x.model) + '</td><td>' + esc(x.next) + '</td>';
        tr.addEventListener('click', function () { location.href = tr.dataset.href; }); db.appendChild(tr);
      });
      if (!d.deals.length) db.innerHTML = '<tr><td colspan="4" class="muted center">' + (n.stageNo >= 0 ? '이 단계에 진행중 안건 없음' : '연결 단계 없음') + '</td></tr>';
    });
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  resize(); load().then(fit);
})();
