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
    renderAssignError(j.assign_error);
    if (j.resolving || j.pending_dongs) {
      msg('방문 주소의 동을 찾는 중이에요(1~2분). 끝나면 저절로 다시 불러와요.', 'info');
      refreshLater();
    }
    $('vhRows').innerHTML = cars.length ? cars.map(function (c) {
      return '<tr tabindex="0" role="button" data-plate="' + esc(c.plate) + '" aria-label="' + esc(carName(c) + ' 운행 보기') + '">' +
        '<td><strong>' + esc(carName(c)) + '</strong>' + (c.label ? '<span class="gw-vh-plate">' + esc(c.plate) + '</span>' : '') + '</td>' +
        '<td class="num gw-vh-opt">' + c.days + '일</td><td class="num gw-vh-opt">' + c.trips + '번</td><td class="num">' + esc(km(c.km)) + '</td>' +
        '<td>' + who(c.drivers, c.source) + '</td></tr>';
    }).join('') : '<tr><td colspan="5" class="gw-empty">' + esc(monthLabel(j.month)) + ' 운행 기록이 없어요. 위 [운행일지 올리기]로 U+커넥트 운행일지를 올려 주세요.</td></tr>';
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
        // DM 을 한 통도 못 보낸 차는 '물음'이 아니라 'DM 못 보냄'(슬랙 ID·연결 확인 필요, 11-24 조용한 실패 점검)
        var asked = m.dm_failed ? (m.asked === 'morning' ? '다음 날 아침 DM 못 보냄' : '운행 종료 DM 못 보냄')
          : m.asked === 'morning' ? '다음 날 아침에 다시 물음' : (m.asked === 'park' ? '운행 종료 때 물음' : '');
        return '<li><span class="d">' + (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ')</span> ' +
          '<strong>' + esc(m.car) + '</strong> <span class="p">' + esc(m.plate) + '</span> · ' + esc((m.riders || []).join(', ')) +
          (asked ? ' <span class="a' + (m.dm_failed ? ' text-danger' : '') + '">' + asked + '</span>' : '') + '</li>';
      }).join('') + '</ul>';
    box.open = open;
  }
  // 배정·운전자 기록을 못 읽은 달(11-24 조용한 실패 점검 23번) — 운전자 칸이 '추정'·'배정'만으로 보여도 정상처럼 믿지 않게 한 줄.
  //   'assign' = 차량 배정을 못 읽음(방문 지역 추정만) · 'driver' = 저녁 슬랙 운전자 답만 못 읽음(배정 기준). 상자는 JS 가 만든다
  function assignErrorText(kind) {
    if (kind === 'driver') return '저녁 슬랙 운전자 답을 못 읽어 운전자를 배정 기준으로 보여요. 잠시 뒤 새로 고쳐 보세요.';
    return kind ? '배정·운전자 기록을 못 읽어 운전자를 방문 지역으로 추정했어요. 잠시 뒤 새로 고쳐 보세요.' : '';
  }
  function renderAssignError(kind) {
    var el = $('vhAssignErr');
    if (!el) {
      el = document.createElement('div');
      el.id = 'vhAssignErr';
      el.className = 'gw-vh-msg err';
      el.setAttribute('role', 'alert');
      var after = $('vhStats');
      after.parentNode.insertBefore(el, after.nextSibling);
    }
    var t = assignErrorText(kind);
    el.textContent = t;
    el.hidden = !t;
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
      logbookLink();
      if (quiet && !(j.resolving || j.pending_dongs)) msg('');
      render(j);
    }).catch(function (e) {
      if (my !== state.seq) return;
      $('vhStats').innerHTML = '';          // 앞 달 숫자가 남아 있지 않게
      renderAssignError('');
      $('vhRows').innerHTML = '<tr><td colspan="5" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  // 표 머리는 여기서 그린다 — 칸 구성이 바뀌어도 서버 재시작 전 옛 템플릿과 어긋나지 않게(11-21)
  var head = document.querySelector('.gw-vh-table thead tr');
  if (head) {
    head.innerHTML = '<th>차량</th><th class="num gw-vh-opt">운행일</th><th class="num gw-vh-opt">운행</th>' +
      '<th class="num">주행거리</th><th>운전자</th>';   // 동일 지역 방문 칸은 뺌 — 차량 펼침창에만(11-24 부장님)
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
    var ev = t.events ? EVENT_NAMES.filter(function (e) { return t.events[e[0]]; })
      .map(function (e) { return e[1] + ' ' + t.events[e[0]]; }) : [];
    if (t.max_speed) ev.push('최고 ' + t.max_speed + 'km/h');
    var extra = (ev.length ? '<br><span class="gw-vh-ev">' + esc(ev.join(' · ')) + '</span>' : '') +
      (t.route ? '<br><button type="button" class="btn btn-sm btn-link p-0 gw-vh-routebtn" data-route="' + esc(t.start) + '" data-end="' + esc(t.end) +
        '" aria-label="' + esc(t.start + ' 출발 운행 경로 보기') + '"><i class="fas fa-route me-1" aria-hidden="true"></i>경로</button>' : '');
    return '<tr' + (t.home ? ' class="gw-vh-home"' : '') + '><td class="gw-vh-time">' + esc(t.start) + '<br><span class="text-muted">' + esc(t.end) + '</span>' +
      (t.driver ? '<br><span class="gw-vh-src assigned">' + esc(t.driver) + '</span>' : '') + '</td>' +
      '<td>' + esc(shortAddr(t.to)) + (t.home ? ' <span class="gw-badge gw-badge-resigned">회사</span>' : '') +
        (t.commute ? ' <span class="gw-badge gw-badge-field">출퇴근</span>' : '') +
        '<br><span class="text-muted small">' + esc(shortAddr(t.from)) + '에서</span>' + extra + '</td>' +
      '<td class="num">' + esc(km(t.km)) + '</td><td>' + (v || '') + '</td></tr>';
  }
  // U+ '운행 이력/경로' 아래 횟수 칸과 같은 이름(11-24)
  var EVENT_NAMES = [['rapid_acc', '급가속'], ['rapid_dec', '급감속'], ['sudden_stop', '급정지'], ['rapid_turn', '급회전'],
    ['rapid_u_turn', '급U턴'], ['rapid_lane', '급차선변경'], ['rapid_pass', '급추월']];
  var carPlate = '';
  function showRoute(btn) {
    var sec = btn.closest('.gw-vh-day');
    var day = sec && sec.getAttribute('data-day');
    if (!day) return;
    var boxEl = $('vhRoute');
    boxEl.hidden = false;
    $('vhRouteHead').textContent = dayLabel(day) + ' ' + btn.getAttribute('data-route') + '~' + btn.getAttribute('data-end') + ' 경로';
    $('vhRouteMap').innerHTML = '<p class="gw-vn-mapmsg">경로를 불러오는 중…</p>';
    boxEl.scrollIntoView({ block: 'nearest' });
    api('/api/vehicles/route?plate=' + encodeURIComponent(carPlate) + '&day=' + encodeURIComponent(day) +
        '&start=' + encodeURIComponent(btn.getAttribute('data-route')) + '&end=' + encodeURIComponent(btn.getAttribute('data-end'))).then(function (j) {
      if (!j.points || j.points.length < 2) { $('vhRouteMap').innerHTML = '<p class="gw-vn-mapmsg">이 운행의 경로 점이 없어요.</p>'; return; }
      if (!j.kakao_js) { $('vhRouteMap').innerHTML = '<p class="gw-vn-mapmsg">지도 키를 넣으면 경로가 지도에 그려져요.</p>'; return; }
      return window.GWKMap.route($('vhRouteMap'), j.kakao_js, j.points)
        .catch(function () { $('vhRouteMap').innerHTML = '<p class="gw-vn-mapmsg">지도를 불러오지 못했어요.</p>'; });
    }).catch(function (e) { $('vhRouteMap').innerHTML = '<p class="gw-vn-mapmsg">' + esc(e.message) + '</p>'; });
  }
  function openCar(plate) {
    carPlate = plate;
    box.hidden = false;
    document.body.classList.add('gw-drawer-open');
    $('vhCarTitle').textContent = plate;
    $('vhCarBody').innerHTML = '<p class="text-muted">불러오는 중…</p>';
    $('vhCarTitle').focus();
    api('/api/vehicles/car?month=' + encodeURIComponent(state.month) + '&plate=' + encodeURIComponent(plate)).then(function (j) {
      $('vhCarTitle').textContent = (j.label ? j.label + ' · ' + j.plate : j.plate) + (j.model ? ' · ' + j.model : '');
      var days = j.days || [];
      var aerr = assignErrorText(j.assign_error);
      $('vhCarBody').innerHTML = (aerr ? '<p class="gw-vh-msg err" role="alert">' + esc(aerr) + '</p>' : '') + '<p class="gw-vh-sub">' + esc(monthLabel(j.month)) + ' · 최근 날짜부터. 동일 지역 방문은 그날 이 차가 도착한 지역의 방문 일정이에요.' +
          ' <a href="/api/vehicles/logbook.xlsx?month=' + encodeURIComponent(j.month) + '&plate=' + encodeURIComponent(j.plate) + '" download>이 차 운행기록부 엑셀</a></p>' +
        '<section class="gw-vh-route" id="vhRoute" hidden><div class="gw-vh-routehead"><b id="vhRouteHead">경로</b>' +
          '<button type="button" class="btn-close" id="vhRouteClose" aria-label="경로 닫기"></button></div><div class="gw-vh-routemap" id="vhRouteMap"></div></section>' +
        (days.length ? days.map(function (d) {
          return '<section class="gw-vh-day" data-day="' + esc(d.day) + '"><div class="gw-vh-dayhead"><b>' + esc(dayLabel(d.day)) + '</b>' +
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
  box.addEventListener('click', function (ev) {
    if (ev.target.closest('[data-close]')) { closeCar(); return; }
    if (ev.target.closest('#vhRouteClose')) { $('vhRoute').hidden = true; $('vhRouteMap').innerHTML = ''; return; }
    var rb = ev.target.closest('[data-route]');
    if (rb) showRoute(rb);
  });
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

  // ── 탭(11-24) — 차량별 / 운전 습관(U+ 안전/경제운전 지수) / 가동률(U+ 가동률 리포트) ──
  var tab = 'cars', insight = { month: '', data: null };
  var PANES = { cars: 'vhPaneCars', habits: 'vhPaneHabits', util: 'vhPaneUtil' };
  function setTab(t) {
    tab = t;
    Array.prototype.forEach.call(document.querySelectorAll('#vhTabs button[data-tab]'), function (b) {
      var on = b.getAttribute('data-tab') === t;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
      b.tabIndex = on ? 0 : -1;
    });
    Object.keys(PANES).forEach(function (k) { if ($(PANES[k])) $(PANES[k]).hidden = k !== t; });
    if (t !== 'cars') loadInsights();
  }
  function logbookLink() {
    var a = $('vhLogbook');
    if (a) a.href = '/api/vehicles/logbook.xlsx' + (state.month ? '?month=' + encodeURIComponent(state.month) : '');
  }
  function score(v) {
    if (v == null) return '<span class="text-muted">-</span>';
    var band = v >= 80 ? 'good' : v >= 60 ? 'ok' : v >= 26 ? 'warn' : 'bad';
    return '<span class="gw-vh-score ' + band + '">' + esc(Math.round(v)) + '점</span>';
  }
  function bands(rows) {           // U+ 와 같은 구간 80-100 / 60-79 / 26-59 / 0-25
    var b = [[80, 100, 0], [60, 79, 0], [26, 59, 0], [0, 25, 0]];
    rows.forEach(function (r) { if (r.safe == null) return; b.forEach(function (x) { if (r.safe >= x[0] && r.safe <= x[1] + 0.99) x[2]++; }); });
    return b;
  }
  var habitBy = 'people';
  function cnt(v) { return v == null ? '-' : esc(v); }
  function renderHabits(h, month) {
    var rows = h[habitBy] || [];
    var avg = function (k) {
      var w = 0, t = 0;
      (habitBy === 'people' ? h.people : h.cars).forEach(function (r) { if (r[k] != null) { t += r[k] * r.km; w += r.km; } });
      return w ? Math.round(t / w) : null;
    };
    var bd = bands(rows);
    var head = '<div class="gw-stats"><div class="gw-stat"><div class="k">' + (habitBy === 'people' ? '운전한 사람' : '운행한 차') +
      '</div><div class="v">' + rows.length + '<small> ' + (habitBy === 'people' ? '명' : '대') + '</small></div></div>' +
      '<div class="gw-stat"><div class="k">평균 안전운전 점수</div><div class="v">' + (avg('safe') != null ? avg('safe') + '<small> 점</small>' : '-') + '</div>' +
        '<div class="gw-vm-kinds">' + bd.map(function (x) { return x[0] + '~' + x[1] + '점 ' + x[2]; }).join(' · ') + '</div></div>' +
      '<div class="gw-stat"><div class="k">평균 경제운전 점수</div><div class="v">' + (avg('eco') != null ? avg('eco') + '<small> 점</small>' : '-') + '</div></div></div>';
    var seg = '<div class="gw-seg gw-vh-seg" role="group" aria-label="묶는 기준">' +
      '<button type="button" class="gw-seg-btn" data-hby="people" aria-pressed="' + (habitBy === 'people') + '">사람별</button>' +
      '<button type="button" class="gw-seg-btn" data-hby="cars" aria-pressed="' + (habitBy === 'cars') + '">차량별</button></div>';
    var table = rows.length ? '<div class="gw-panel gw-ap-tablewrap"><table class="gw-table gw-vh-habit"><thead><tr>' +
      '<th>' + (habitBy === 'people' ? '운전자' : '차량') + '</th><th class="num">운행</th><th class="num">거리</th><th class="num">안전</th><th class="num">경제</th>' +
      '<th class="num gw-vh-opt">급가속</th><th class="num gw-vh-opt">급감속</th><th class="num gw-vh-opt">급정지</th>' +
      '<th class="num">위험 행동<br><span class="small">100km당</span></th><th class="num">최고 속도</th><th class="num">과속</th></tr></thead><tbody>' +
      rows.map(function (r) {
        return '<tr><td><strong>' + esc(r.name) + '</strong></td><td class="num">' + r.trips + '번</td><td class="num">' + esc(km(r.km)) + '</td>' +
          '<td class="num">' + score(r.safe) + '</td><td class="num">' + score(r.eco) + '</td>' +
          // 안전운전 기록이 없는 운행만이면 null — 0 회가 아니라 '-'(11-24 조용한 실패 점검)
          '<td class="num gw-vh-opt">' + cnt(r.acc) + '</td><td class="num gw-vh-opt">' + cnt(r.dec) + '</td><td class="num gw-vh-opt">' + cnt(r.stop) + '</td>' +
          '<td class="num">' + (r.rapid_per100 != null ? r.rapid_per100 : '-') + '</td><td class="num">' + (r.max_speed ? r.max_speed + 'km/h' : '-') + '</td>' +
          '<td class="num">' + (r.speeding ? '<span class="gw-badge gw-badge-reject">' + r.speeding + '번</span>' : '0') + '</td></tr>';
      }).join('') + '</tbody></table></div>' : '';
    $('vhPaneHabits').innerHTML = rows.length
      ? head + seg + table + '<p class="gw-vh-foot">점수는 U+ 가 운행마다 매긴 안전·경제운전 지수를 거리로 평균 낸 거예요(위험 행동·비경제 행동 횟수 기준). ' +
        '운전자는 운전자 답 › 배정 순서로 정해요. 과속은 최고 속도 ' + esc(h.speed_limit) + 'km/h 이상인 운행이에요.</p>'
      : '<div class="gw-panel gw-empty-box"><p class="gw-empty">' + esc(monthLabel(month)) + ' 운전 습관 기록이 아직 없어요. U+ 연동이 승인되면 매일 새벽 운행과 함께 자동으로 쌓여요.</p></div>';
  }
  function hm(min) {
    var h = Math.floor(min / 60), m = Math.round(min % 60);
    return (h ? h + '시간 ' : '') + m + '분';
  }
  function renderUtil(u, month) {
    var cars = u.cars || [];
    if (!u.has_data) {
      $('vhPaneUtil').innerHTML = '<div class="gw-panel gw-empty-box"><p class="gw-empty">' + esc(monthLabel(month)) + ' 가동률 기록이 아직 없어요. U+ 연동이 승인되면 운행 시간으로 계산돼요.</p></div>';
      return;
    }
    var tot = cars.reduce(function (a, c) { return a + c.hours * 60; }, 0);
    var dayAvg = cars.length && u.workdays ? tot / cars.length / u.workdays : 0;
    var rate = u.workdays ? Math.round(dayAvg / 540 * 1000) / 10 : 0;
    $('vhPaneUtil').innerHTML = '<div class="gw-stats">' +
      '<div class="gw-stat"><div class="k">근무일</div><div class="v">' + u.workdays + '<small> 일</small></div></div>' +
      '<div class="gw-stat"><div class="k">차 한 대 하루 평균 운행</div><div class="v gw-vh-up">' + esc(hm(dayAvg)) + '</div></div>' +
      '<div class="gw-stat"><div class="k">전체 가동률</div><div class="v">' + rate + '<small> %</small></div></div></div>' +
      '<div class="gw-panel gw-ap-tablewrap"><table class="gw-table gw-vh-util"><thead><tr><th>차량</th><th class="num">운행한 날</th>' +
      '<th class="num">하루 평균 운행</th><th class="num gw-vh-opt">하루 평균 횟수</th><th>가동률</th></tr></thead><tbody>' +
      cars.map(function (c) {
        var avgMin = u.workdays ? c.hours * 60 / u.workdays : 0;
        var r = c.rate || 0;
        return '<tr><td><strong>' + esc(c.label) + '</strong><span class="gw-vh-plate">' + esc(c.plate) + '</span></td>' +
          '<td class="num">' + c.days + '일</td><td class="num">' + esc(hm(avgMin)) + '</td>' +
          '<td class="num gw-vh-opt">' + (u.workdays ? Math.round(c.trips / u.workdays * 10) / 10 : 0) + '번</td>' +
          '<td><div class="gw-vh-bar" role="img" aria-label="가동률 ' + r + '%"><span style="width:' + Math.min(100, r) + '%"></span></div>' +
          '<span class="gw-vh-barnum">' + r + '%</span></td></tr>';
      }).join('') + '</tbody></table></div>' +
      '<p class="gw-vh-foot">가동률 = 운행 시간 ÷ 업무 시간(하루 9시간, 08~17시 — U+ 가동률 리포트와 같은 기준). 근무일은 월~토 중 공휴일을 뺀 날이에요. ' +
      '차를 줄이거나 늘릴 때 근거로 보세요.</p>';
  }
  function loadInsights() {
    if (insight.month === state.month && insight.data) { draw(); return; }
    var pane = $(PANES[tab]);
    pane.innerHTML = '<div class="gw-panel gw-empty-box"><p class="gw-empty">불러오는 중…</p></div>';
    api('/api/vehicles/insights' + (state.month ? '?month=' + encodeURIComponent(state.month) : '')).then(function (j) {
      insight = { month: j.month, data: j };
      draw();
    }).catch(function (e) { pane.innerHTML = '<div class="gw-panel gw-empty-box"><p class="gw-empty">' + esc(e.message) + '</p></div>'; });
    function draw() {
      if (tab === 'habits') renderHabits(insight.data.habits, insight.data.month);
      if (tab === 'util') renderUtil(insight.data.util, insight.data.month);
    }
  }
  if ($('vhTabs')) {
    $('vhTabs').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-tab]');
      if (b) setTab(b.getAttribute('data-tab'));
    });
    $('vhTabs').addEventListener('keydown', function (e) {
      if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].indexOf(e.key) === -1) return;
      var tabs = Array.prototype.slice.call($('vhTabs').querySelectorAll('button[data-tab]'));
      var i = tabs.indexOf(document.activeElement);
      if (i < 0) return;
      e.preventDefault();
      var n = e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : (i + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
      tabs[n].focus();
      tabs[n].click();
    });
    $('vhPaneHabits').addEventListener('click', function (e) {
      var b = e.target.closest('[data-hby]');
      if (b && insight.data) { habitBy = b.getAttribute('data-hby'); renderHabits(insight.data.habits, insight.data.month); }
    });
  }

  $('vhMonth').addEventListener('change', function () {
    state.month = this.value; msg(''); logbookLink(); load();
    if (tab !== 'cars') loadInsights();
  });
  load();
})();
