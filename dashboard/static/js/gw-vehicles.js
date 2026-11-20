/* 그룹웨어 차량 운행 (2026-11-20) — U+커넥트 운행일지 올리기 + 차량별 한 달 + 날짜별 운행·추정 운전자.
   서버에서 온 글자는 모두 esc() 로만 넣는다. */
(function () {
  'use strict';

  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url, init) {
    init = init || {};
    init.credentials = 'same-origin';
    var method = init.method || 'GET';
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '올리지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }
  function km(n) {
    n = Number(n) || 0;
    return n.toLocaleString('ko-KR', { maximumFractionDigits: 1 }) + ' km';
  }
  function dayLabel(s) {
    var d = new Date(s + 'T00:00:00');
    return isNaN(d) ? s : (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ')';
  }
  function monthLabel(m) { return m.slice(0, 4) + '년 ' + Number(m.slice(5)) + '월'; }
  function shortAddr(a) {   // '경기도 김포시 걸포동' → '김포시 걸포동'
    return String(a || '').replace(/^(경기도|강원(특별자치)?도|충청[남북]도|전라[남북]도|전북특별자치도|경상[남북]도|제주특별자치도)\s+/, '')
      .replace(/^(서울|인천|부산|대구|광주|대전|울산)(특별시|광역시)\s+/, '$1 ');
  }
  function names(list) { return (list || []).length ? (list || []).map(esc).join(', ') : '<span class="text-muted">-</span>'; }
  // 운전자 출처(11-20) — 배정(차량 배정 화면)·고정(정해진 사람)·추정(방문 일정)
  var SRC = { assigned: '배정', fixed: '고정', guess: '추정' };
  function who(list, src) {
    if (!(list || []).length) return src === 'guess' || !src ? '추정 못 함' : '<span class="text-muted">-</span>';
    return '<span class="gw-vh-src ' + esc(src || 'guess') + '">' + esc(SRC[src] || '추정') + '</span> ' + names(list);
  }

  var state = { month: '', seq: 0 };

  function msg(text, kind) {
    var el = $('vhMsg');
    el.className = 'gw-vh-msg' + (kind ? ' ' + kind : '');
    el.textContent = text || '';
  }

  function fillMonths(list, cur) {
    var sel = $('vhMonth');
    var all = (list || []).slice();
    if (all.indexOf(cur) === -1) all.unshift(cur);
    all.sort().reverse();
    sel.innerHTML = all.map(function (m) {
      return '<option value="' + esc(m) + '"' + (m === cur ? ' selected' : '') + '>' + esc(monthLabel(m)) + '</option>';
    }).join('');
  }

  function render(j) {
    fillMonths(j.months, j.month);
    var up = j.last_upload;
    var cars = j.vehicles || [];
    $('vhStats').innerHTML =
      '<div class="gw-stat gw-stat-main"><div class="k">' + esc(monthLabel(j.month)) + ' 주행거리</div><div class="v">' + esc(km(j.total_km)) + '</div></div>' +
      '<div class="gw-stat"><div class="k">운행한 차량</div><div class="v">' + cars.length + '<small> 대</small></div></div>' +
      '<div class="gw-stat"><div class="k">마지막으로 올린 운행일지</div><div class="v gw-vh-up">' +
        (up ? esc(String(up.uploaded_at).slice(5, 16).replace('-', '/')) + '<small> · ' + esc(up.period || '') + '</small>' : '<small>아직 없어요</small>') +
      '</div></div>';
    if (j.resolving || j.pending_dongs) {
      msg('방문 주소의 동을 찾는 중이에요(1~2분). 끝나면 새로 고쳐 주세요 — 추정 운전자가 더 정확해져요.', 'info');
    }
    $('vhRows').innerHTML = cars.length ? cars.map(function (c) {
      return '<tr tabindex="0" role="button" data-plate="' + esc(c.plate) + '" aria-label="' + esc(c.plate + ' 운행 보기') + '">' +
        '<td><strong>' + esc(c.plate) + '</strong></td><td>' + esc(c.model) + '</td>' +
        '<td class="num">' + c.days + '일</td><td class="num">' + c.trips + '번</td><td class="num">' + esc(km(c.km)) + '</td>' +
        '<td>' + who(c.drivers, c.source) + '</td><td class="num">' + (c.visits ? c.visits + '건' : '-') + '</td></tr>';
    }).join('') : '<tr><td colspan="7" class="gw-empty">' + esc(monthLabel(j.month)) + ' 운행 기록이 없어요. 위 [운행일지 올리기]로 U+커넥트 운행일지를 올려 주세요.</td></tr>';
  }

  function load() {
    var my = ++state.seq;
    api('/api/vehicles' + (state.month ? '?month=' + encodeURIComponent(state.month) : '')).then(function (j) {
      if (my !== state.seq) return;
      state.month = j.month;
      render(j);
    }).catch(function (e) {
      if (my !== state.seq) return;
      $('vhRows').innerHTML = '<tr><td colspan="7" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  // ── 차량 한 대 ──
  var box = $('vhCar'), opener = null;
  function closeCar() { box.hidden = true; if (opener) opener.focus(); }
  function tripRow(t) {
    var v = (t.visits || []).map(function (x) {
      return '<div class="gw-vh-visit"><b>' + esc((x.who || []).join(', ')) + '</b> <span class="text-muted">' + esc(x.lead_no) + '</span><br>' +
        '<span class="gw-vh-addr">' + esc(x.addr) + '</span></div>';
    }).join('');
    return '<tr' + (t.home ? ' class="gw-vh-home"' : '') + '><td class="gw-vh-time">' + esc(t.start) + '<br><span class="text-muted">' + esc(t.end) + '</span></td>' +
      '<td>' + esc(shortAddr(t.to)) + (t.home ? ' <span class="gw-badge gw-badge-resigned">회사</span>' : '') +
        '<br><span class="text-muted small">' + esc(shortAddr(t.from)) + '에서</span></td>' +
      '<td class="num">' + esc(km(t.km)) + '</td><td>' + (v || '') + '</td></tr>';
  }
  function openCar(plate) {
    box.hidden = false;
    $('vhCarTitle').textContent = plate;
    $('vhCarBody').innerHTML = '<p class="text-muted">불러오는 중…</p>';
    $('vhCarTitle').focus();
    api('/api/vehicles/car?month=' + encodeURIComponent(state.month) + '&plate=' + encodeURIComponent(plate)).then(function (j) {
      $('vhCarTitle').textContent = j.plate + (j.model ? ' · ' + j.model : '');
      var days = j.days || [];
      $('vhCarBody').innerHTML = '<p class="gw-vh-sub">' + esc(monthLabel(j.month)) + ' · 최근 날짜부터. 같은 곳 방문은 그날 이 차가 도착한 동네의 방문 일정이에요.</p>' +
        (days.length ? days.map(function (d) {
          return '<section class="gw-vh-day"><div class="gw-vh-dayhead"><b>' + esc(dayLabel(d.day)) + '</b>' +
            '<span>' + esc(km(d.km)) + '</span><span class="gw-vh-who">' + who(d.drivers, d.source) + '</span></div>' +
            '<table class="gw-table gw-vh-trips"><thead><tr><th>시각</th><th>도착지</th><th class="num">거리</th><th>같은 곳 방문</th></tr></thead><tbody>' +
            d.trips.map(tripRow).join('') + '</tbody></table></section>';
        }).join('') : '<p class="text-muted">이 달 운행이 없어요.</p>');
    }).catch(function (e) { $('vhCarBody').innerHTML = '<p class="text-danger">' + esc(e.message) + '</p>'; });
  }
  $('vhRows').addEventListener('click', function (ev) {
    var tr = ev.target.closest('tr[data-plate]');
    if (tr) { opener = tr; openCar(tr.getAttribute('data-plate')); }
  });
  $('vhRows').addEventListener('keydown', function (ev) {
    var tr = ev.target.closest('tr[data-plate]');
    if (tr && (ev.key === 'Enter' || ev.key === ' ')) { ev.preventDefault(); opener = tr; openCar(tr.getAttribute('data-plate')); }
  });
  box.addEventListener('click', function (ev) { if (ev.target.closest('[data-close]')) closeCar(); });
  document.addEventListener('keydown', function (ev) { if (ev.key === 'Escape' && !box.hidden) closeCar(); });

  // ── 올리기 ──
  $('vhFile').addEventListener('change', function () {
    var f = this.files && this.files[0];
    if (!f) return;
    var fd = new FormData();
    fd.append('file', f);
    var btn = $('vhUploadBtn');
    btn.classList.add('disabled');
    msg('올리는 중이에요…', 'info');
    var input = this;
    api('/api/vehicles/upload', { method: 'POST', body: fd }).then(function (j) {
      msg('차량 ' + j.vehicles + '대 운행일지를 올렸어요. 새 운행 ' + j.added + '건' +
        (j.skipped ? ', 이미 있던 운행 ' + j.skipped + '건은 건너뛰었어요.' : '.') +
        ' 방문 주소의 동을 찾는 중이라 1~2분 뒤 새로 고치면 추정 운전자가 더 정확해져요.', 'ok');
      load();
    }).catch(function (e) { msg(e.message, 'err'); })
      .then(function () { btn.classList.remove('disabled'); input.value = ''; });
  });

  $('vhMonth').addEventListener('change', function () { state.month = this.value; msg(''); load(); });
  load();
})();
