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
  // '-' 같은 빈칸 표시는 사람이 아니다(서버에서도 거르지만 옛 응답 대비, 11-21 점검)
  function real(list) { return (list || []).filter(function (n) { return /[가-힣A-Za-z]/.test(String(n || '')); }); }
  function names(list) { list = real(list); return list.length ? list.map(esc).join(', ') : ''; }
  // 운전자 출처(11-20) — 배정(차량 배정 화면)·고정(정해진 사람)·추정(방문 일정)
  var SRC = { confirmed: '확인', assigned: '배정', fixed: '고정', guess: '추정' };   // 확인 = 저녁 슬랙 DM 으로 받은 실제 운전자(11-22)
  function who(list, src) {
    if (!real(list).length) return '<span class="gw-vh-none">' + (src === 'guess' || !src ? '알 수 없음' : '없음') + '</span>';
    return '<span class="gw-vh-src ' + esc(src || 'guess') + '">' + esc(SRC[src] || '추정') + '</span> ' + names(list);
  }
  function carName(c) { return c.label || c.plate; }

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
    // 큰 숫자는 검정(색은 상태에만 — 공통 모양 규칙). 기간은 아래 줄로 따로(가운데서 끊기던 것)
    $('vhStats').innerHTML =
      '<div class="gw-stat"><div class="k">' + esc(monthLabel(j.month)) + ' 주행거리</div><div class="v">' + esc(km(j.total_km)) + '</div></div>' +
      '<div class="gw-stat"><div class="k">운행한 차량</div><div class="v">' + cars.length + '<small> 대</small></div></div>' +
      '<div class="gw-stat"><div class="k">마지막으로 올린 운행일지</div><div class="v gw-vh-up">' +
        (up ? esc(String(up.uploaded_at).slice(5, 16).replace('-', '/')) +
          (up.period ? '<span class="gw-vh-period">' + esc(String(up.period).replace(/(\d{4})-(\d{2})-(\d{2})/g, function (m, y, mo, d) {
            return Number(mo) + '/' + Number(d); }).replace('~', ' ~ ')) + ' 운행</span>' : '')
          : '<small>아직 없어요</small>') +
      '</div></div>';
    renderMissing(j.missing_drivers || []);
    if (j.resolving || j.pending_dongs) {
      msg('방문 주소의 동을 찾는 중이에요(1~2분). 끝나면 저절로 다시 불러와요.', 'info');
      refreshLater();
    }
    $('vhRows').innerHTML = cars.length ? cars.map(function (c) {
      return '<tr tabindex="0" role="button" data-plate="' + esc(c.plate) + '" aria-label="' + esc(carName(c) + ' 운행 보기') + '">' +
        '<td><strong>' + esc(carName(c)) + '</strong>' + (c.label ? '<span class="gw-vh-plate">' + esc(c.plate) + '</span>' : '') + '</td>' +
        '<td class="num gw-vh-opt">' + c.days + '일</td><td class="num gw-vh-opt">' + c.trips + '번</td><td class="num">' + esc(km(c.km)) + '</td>' +
        '<td>' + who(c.drivers, c.source) + '</td><td class="num gw-vh-opt">' + (c.visits ? c.visits + '건' : '-') + '</td></tr>';
    }).join('') : '<tr><td colspan="6" class="gw-empty">' + esc(monthLabel(j.month)) + ' 운행 기록이 없어요. 위 [운행일지 올리기]로 U+커넥트 운행일지를 올려 주세요.</td></tr>';
  }
  // 운전자 미확인(11-24) — 확정한 날인데 운전자 답이 없는 차. 운행 종료 DM·다음 날 아침 다시 묻기 뒤에도 남은 것.
  //   상자는 JS 가 만든다(템플릿 재시작 없이). 없으면 숨김.
  function renderMissing(list) {
    var box = $('vhMissing');
    if (!box) {
      box = document.createElement('details');
      box.id = 'vhMissing';
      box.className = 'gw-panel gw-vh-missing';
      var after = $('vhStats');
      after.parentNode.insertBefore(box, after.nextSibling);
    }
    box.hidden = !list.length;
    if (!list.length) { box.innerHTML = ''; return; }
    var open = box.open;
    box.innerHTML = '<summary>운전자 미확인 <strong>' + list.length + '건</strong>' +
      '<span>배정을 확정한 날인데 탄 사람이 아직 운전자를 안 알려 준 차예요.</span></summary><ul>' +
      list.map(function (m) {
        var d = new Date(m.day + 'T00:00:00');
        var asked = m.asked === 'morning' ? '다음 날 아침에 다시 물음' : (m.asked === 'park' ? '운행 종료 때 물음' : '');
        return '<li><span class="d">' + (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ')</span> ' +
          '<strong>' + esc(m.car) + '</strong> <span class="p">' + esc(m.plate) + '</span> · ' + esc((m.riders || []).join(', ')) +
          (asked ? ' <span class="a">' + asked + '</span>' : '') + '</li>';
      }).join('') + '</ul>';
    box.open = open;
  }
  // 동 찾기가 끝나면 저절로 다시 불러온다(직접 새로 고치라던 것, 11-21 점검) — 30초마다, 최대 6번
  var refreshTimer = null, refreshLeft = 6;
  function refreshLater() {
    if (refreshTimer || refreshLeft <= 0) return;
    refreshLeft -= 1;
    refreshTimer = setTimeout(function () { refreshTimer = null; load(true); }, 30000);
  }

  function load(quiet) {
    var my = ++state.seq;
    if (!quiet) msg('');
    api('/api/vehicles' + (state.month ? '?month=' + encodeURIComponent(state.month) : '')).then(function (j) {
      if (my !== state.seq) return;
      state.month = j.month;
      if (quiet && !(j.resolving || j.pending_dongs)) msg('');
      render(j);
    }).catch(function (e) {
      if (my !== state.seq) return;
      $('vhStats').innerHTML = '';          // 앞 달 숫자가 남아 있지 않게
      $('vhRows').innerHTML = '<tr><td colspan="6" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  // 표 머리는 여기서 그린다 — 칸 구성이 바뀌어도 서버 재시작 전 옛 템플릿과 어긋나지 않게(11-21)
  var head = document.querySelector('.gw-vh-table thead tr');
  if (head) {
    head.innerHTML = '<th>차량</th><th class="num gw-vh-opt">운행일</th><th class="num gw-vh-opt">운행</th>' +
      '<th class="num">주행거리</th><th>운전자</th><th class="num gw-vh-opt">동일 지역 방문</th>';
  }

  // ── 차량 한 대 ──
  var box = $('vhCar'), opener = null;
  function closeCar() {
    box.hidden = true;
    document.body.classList.remove('gw-drawer-open');
    if (opener && document.contains(opener)) opener.focus();
  }
  // 창이 열린 동안 Tab 이 뒤쪽 표로 빠져나가지 않게(결재함 서랍과 같게, 11-21 점검)
  function trapTab(e) {
    if (e.key !== 'Tab' || box.hidden) return;
    var panel = box.querySelector('.gw-drawer-panel');
    var items = Array.prototype.filter.call(
      panel.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])'),
      function (el) { return el.offsetParent !== null; });
    if (!items.length) return;
    var first = items[0], last = items[items.length - 1], cur = document.activeElement;
    if (!panel.contains(cur)) { e.preventDefault(); first.focus(); return; }
    if (e.shiftKey && (cur === first || cur === $('vhCarTitle'))) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && cur === last) { e.preventDefault(); first.focus(); }
  }
  function tripRow(t) {
    var v = (t.visits || []).map(function (x) {
      return '<div class="gw-vh-visit"><b>' + names(x.who) + '</b> <span class="text-muted">' + esc(x.lead_no) + '</span><br>' +
        '<span class="gw-vh-addr">' + esc(x.addr) + '</span></div>';
    }).join('');
    return '<tr' + (t.home ? ' class="gw-vh-home"' : '') + '><td class="gw-vh-time">' + esc(t.start) + '<br><span class="text-muted">' + esc(t.end) + '</span>' +
      (t.driver ? '<br><span class="gw-vh-src assigned">' + esc(t.driver) + '</span>' : '') + '</td>' +
      '<td>' + esc(shortAddr(t.to)) + (t.home ? ' <span class="gw-badge gw-badge-resigned">회사</span>' : '') +
        '<br><span class="text-muted small">' + esc(shortAddr(t.from)) + '에서</span></td>' +
      '<td class="num">' + esc(km(t.km)) + '</td><td>' + (v || '') + '</td></tr>';
  }
  function openCar(plate) {
    box.hidden = false;
    document.body.classList.add('gw-drawer-open');
    $('vhCarTitle').textContent = plate;
    $('vhCarBody').innerHTML = '<p class="text-muted">불러오는 중…</p>';
    $('vhCarTitle').focus();
    api('/api/vehicles/car?month=' + encodeURIComponent(state.month) + '&plate=' + encodeURIComponent(plate)).then(function (j) {
      $('vhCarTitle').textContent = (j.label ? j.label + ' · ' + j.plate : j.plate) + (j.model ? ' · ' + j.model : '');
      var days = j.days || [];
      $('vhCarBody').innerHTML = '<p class="gw-vh-sub">' + esc(monthLabel(j.month)) + ' · 최근 날짜부터. 동일 지역 방문은 그날 이 차가 도착한 지역의 방문 일정이에요.</p>' +
        (days.length ? days.map(function (d) {
          return '<section class="gw-vh-day"><div class="gw-vh-dayhead"><b>' + esc(dayLabel(d.day)) + '</b>' +
            '<span>' + esc(km(d.km)) + '</span><span class="gw-vh-who">' + who(d.drivers, d.source) + '</span></div>' +
            '<table class="gw-table gw-vh-trips"><thead><tr><th>시각</th><th>도착지</th><th class="num">거리</th><th>동일 지역 방문</th></tr></thead><tbody>' +
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
  document.addEventListener('keydown', function (ev) {
    if (ev.key === 'Escape' && !box.hidden) closeCar();
    trapTab(ev);
  });

  // ── 올리기 ── (새 템플릿은 <button>, 재시작 전 옛 템플릿은 <label> 이 알아서 파일 창을 연다)
  if ($('vhUploadBtn').tagName === 'BUTTON') {
    $('vhUploadBtn').addEventListener('click', function () { $('vhFile').click(); });
  }
  $('vhFile').addEventListener('change', function () {
    var f = this.files && this.files[0];
    if (!f) return;
    var fd = new FormData();
    fd.append('file', f);
    var btn = $('vhUploadBtn');
    btn.classList.add('disabled');
    btn.setAttribute('aria-disabled', 'true');
    if (btn.tagName === 'BUTTON') btn.disabled = true;
    msg('올리는 중이에요…', 'info');
    var input = this;
    api('/api/vehicles/upload', { method: 'POST', body: fd }).then(function (j) {
      msg('차량 ' + j.vehicles + '대 운행일지를 올렸어요. 새 운행 ' + j.added + '건' +
        (j.skipped ? ', 이미 있던 운행 ' + j.skipped + '건은 건너뛰었어요.' : '.') +
        ' 방문 주소의 동을 찾는 1~2분 동안 추정 운전자가 조금씩 바뀔 수 있어요.', 'ok');
      refreshLeft = 6;
      load(true);
    }).catch(function (e) { msg(e.message, 'err'); })
      .then(function () {
        btn.classList.remove('disabled'); btn.removeAttribute('aria-disabled');
        if (btn.tagName === 'BUTTON') btn.disabled = false;
        input.value = '';
      });
  });

  $('vhMonth').addEventListener('change', function () { state.month = this.value; msg(''); load(); });
  load();
})();
