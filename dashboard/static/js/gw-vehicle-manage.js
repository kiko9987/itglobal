/* 그룹웨어 차량 관리 (2026-11-22) — 보험·검사 만기, 정비·비용(결재에서 자동), 계기판 km.
   서버에서 온 글자는 모두 esc() 로만 넣는다. 펼침창은 운행일지(gw-vehicles.js)와 같은 틀. */
(function () {
  'use strict';

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
            : '저장하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }
  function post(url, body) {
    return api(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  }
  function msg(text, kind) {
    var el = $('vmMsg');
    el.className = 'gw-vh-msg' + (text ? ' ' + (kind || 'info') : '');
    el.textContent = text || '';
  }
  function won(n) { return (Number(n) || 0).toLocaleString('ko-KR') + '원'; }
  function shortDay(d) { return d ? d.slice(2).replace(/-/g, '.') : ''; }
  function today() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  // 만기 — 지났거나 30일 안이면 표시(색만으로 알리지 않게 글자도 함께)
  function due(day, d) {
    if (!day) return '<span class="text-muted">-</span>';
    var tag = '';
    if (d < 0) tag = ' <span class="gw-badge gw-badge-reject">' + (-d) + '일 지남</span>';
    else if (d <= 30) tag = ' <span class="gw-badge gw-badge-wait">' + d + '일 남음</span>';
    return esc(shortDay(day)) + tag;
  }

  var S = { kinds: [], cars: [] };

  // ── 첫 화면 ──
  function load() {
    return api('/api/vehicles/manage').then(function (j) {
      S.kinds = j.kinds; S.cars = j.cars;
      var top = (j.kinds_year || []).slice(0, 3).map(function (k) { return esc(k.kind) + ' ' + esc(won(k.amount)); }).join(' · ');
      $('vmStats').innerHTML =
        '<div class="gw-stat"><div class="k">만기 지남·30일 안</div><div class="v">' + j.alerts + '<small> 대</small></div></div>' +
        '<div class="gw-stat"><div class="k">' + Number(j.month.slice(5)) + '월 차량비</div><div class="v">' + esc(won(j.month_total)) + '</div></div>' +
        '<div class="gw-stat"><div class="k">' + esc(j.year) + '년 차량비</div><div class="v">' + esc(won(j.year_total)) + '</div>' +
          (top ? '<div class="gw-vm-kinds">' + top + '</div>' : '') + '</div>';
      $('vmRows').innerHTML = j.cars.map(function (c) {
        var tag = c.usage === 'off' ? ' <span class="gw-badge gw-badge-resigned">배정 안 함</span>' : '';
        var own = c.ownership && c.ownership !== '법인 소유' ? '<span class="gw-vm-own">' + esc(c.ownership) + '</span>' : '';
        var kmCell = c.km != null ? esc(Number(c.km).toLocaleString('ko-KR')) + '<br><span class="text-muted small">' + esc(shortDay(c.km_day)) + '</span>'
          : (c.uplus ? '<span class="text-muted small">U+ 연동 후</span>' : '<span class="gw-badge gw-badge-wait">입력 필요</span>');
        return '<tr tabindex="0" role="button" data-plate="' + esc(c.plate) + '" aria-label="' + esc(c.label + ' 자세히') + '"' +
          (c.warn ? ' class="gw-vm-warn"' : '') + '>' +
          '<td><strong>' + esc(c.label) + '</strong>' + tag + '<br><span class="text-muted small">' + esc(c.plate) + ' · ' + esc(c.model) + '</span> ' + own + '</td>' +
          '<td>' + due(c.insurance_end, c.insurance_d) + (c.insurer ? '<br><span class="text-muted small">' + esc(c.insurer) + '</span>' : '') + '</td>' +
          '<td>' + due(c.inspection_end, c.inspection_d) + '</td>' +
          '<td class="num">' + (c.cost_year ? esc(won(c.cost_year)) : '<span class="text-muted">-</span>') + '</td>' +
          '<td>' + (c.last_fix_day ? esc(shortDay(c.last_fix_day)) + '<br><span class="text-muted small">' + esc(c.last_fix_item) + '</span>' : '<span class="text-muted">-</span>') + '</td>' +
          '<td class="num gw-vm-opt">' + kmCell + '</td></tr>';
      }).join('');
    }).catch(function (e) { msg(e.message, 'err'); });
  }

  // ── 차량 한 대 ──
  var box = $('vmCar'), opener = null, cur = null;
  function closeCar() {
    box.hidden = true;
    document.body.classList.remove('gw-drawer-open');
    if (opener && document.contains(opener)) opener.focus();
  }
  function trapTab(e) {
    if (e.key !== 'Tab' || box.hidden) return;
    var panel = box.querySelector('.gw-drawer-panel');
    var items = Array.prototype.filter.call(
      panel.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])'),
      function (el) { return el.offsetParent !== null; });
    if (!items.length) return;
    var first = items[0], last = items[items.length - 1], c = document.activeElement;
    if (!panel.contains(c)) { e.preventDefault(); first.focus(); return; }
    if (e.shiftKey && (c === first || c === $('vmCarTitle'))) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && c === last) { e.preventDefault(); first.focus(); }
  }
  function infoForm(car, info) {
    return '<section class="gw-vm-sec"><h3>보험·검사</h3><form class="gw-vm-form" id="vmInfo">' +
      '<label>소유<input class="form-control form-control-sm" autocomplete="off" name="ownership" maxlength="60" value="' + esc(info.ownership || '') + '"></label>' +
      '<label>보험사<input class="form-control form-control-sm" autocomplete="off" name="insurer" maxlength="60" value="' + esc(info.insurer || '') + '"></label>' +
      '<label>보험 만기<input type="date" class="form-control form-control-sm" autocomplete="off" name="insurance_end" value="' + esc(info.insurance_end || '') + '"></label>' +
      '<label>검사 만기<input type="date" class="form-control form-control-sm" autocomplete="off" name="inspection_end" value="' + esc(info.inspection_end || '') + '"></label>' +
      '<label class="wide">메모<input class="form-control form-control-sm" autocomplete="off" name="note" maxlength="60" value="' + esc(info.note || '') + '"></label>' +
      '<div class="gw-vm-actions"><button type="submit" class="btn btn-sm btn-outline-secondary">저장</button></div></form></section>';
  }
  function odoSection(car, rows) {
    var list = rows.length ? '<table class="gw-table gw-vm-list"><thead><tr><th>날짜</th><th class="num">계기판 km</th><th class="num">그 사이 주행</th></tr></thead><tbody>' +
      rows.map(function (r) {
        return '<tr><td>' + esc(shortDay(r.day)) + '</td><td class="num">' + esc(Number(r.km).toLocaleString('ko-KR')) + '</td>' +
          '<td class="num">' + (r.driven != null ? esc(Number(r.driven).toLocaleString('ko-KR')) + ' km' : '-') + '</td></tr>';
      }).join('') + '</tbody></table>' : '<p class="text-muted small">아직 넣은 기록이 없어요. 한 달에 한 번쯤 계기판 숫자를 넣으면 그 사이 주행거리가 운행일지로 이어져요.</p>';
    return '<section class="gw-vm-sec"><h3>계기판 km <span class="gw-vm-sub">U+ 단말이 없는 차</span></h3>' +
      '<form class="gw-vm-form" id="vmOdo"><label>날짜<input type="date" class="form-control form-control-sm" autocomplete="off" name="day" value="' + today() + '" required></label>' +
      '<label>계기판 km<input type="number" inputmode="decimal" step="0.1" min="0" class="form-control form-control-sm" autocomplete="off" name="km" required></label>' +
      '<div class="gw-vm-actions"><button type="submit" class="btn btn-sm btn-primary">넣기</button></div></form>' + list + '</section>';
  }
  function costSection(car, rows) {
    var kinds = S.kinds.map(function (k) { return '<option>' + esc(k) + '</option>'; }).join('');
    // 세 칸(날짜·내용·금액) — 종류는 내용 앞 꼬리표, 결재/삭제는 금액 아래(휴대폰에서 내용 칸이 좁아지지 않게, 11-22)
    var list = rows.length ? '<table class="gw-table gw-vm-list"><thead><tr><th>날짜</th><th>내용</th><th class="num">금액</th></tr></thead><tbody>' +
      rows.map(function (r) {
        var src = r.source === 'manual'
          ? '<button type="button" class="btn btn-sm btn-link text-danger p-0" data-del="' + r.id + '">삭제</button>'
          : '<span class="gw-badge gw-badge-resigned" title="결재에서 자동으로 들어온 줄">결재</span>';
        var shared = r.shared > 1 ? '<br><span class="text-muted small">' + r.shared + '대 나눔</span>' : '';
        return '<tr><td class="gw-vm-day">' + esc(shortDay(r.day)) + '</td>' +
          '<td><span class="gw-vm-kind">' + esc(r.kind) + '</span> ' + esc(r.item) +
            (r.vendor ? '<br><span class="text-muted small">' + esc(r.vendor) + '</span>' : '') + '</td>' +
          '<td class="num">' + esc(won(r.amount)) + shared + '<br>' + src + '</td></tr>';
      }).join('') + '</tbody></table>' : '<p class="text-muted small">결재에 이 차 번호가 적힌 비용이 아직 없어요.</p>';
    return '<section class="gw-vm-sec"><h3>정비·비용</h3>' +
      '<details class="gw-vm-add"><summary>결재 없이 쓴 비용 직접 넣기</summary><form class="gw-vm-form" id="vmCost">' +
      '<label>날짜<input type="date" class="form-control form-control-sm" autocomplete="off" name="day" value="' + today() + '" required></label>' +
      '<label>종류<select class="form-select form-select-sm" name="kind" autocomplete="off">' + kinds + '</select></label>' +
      '<label class="wide">내용<input class="form-control form-control-sm" autocomplete="off" name="item" maxlength="120" required></label>' +
      '<label>금액(원)<input type="number" inputmode="numeric" min="0" step="1" class="form-control form-control-sm" autocomplete="off" name="amount" required></label>' +
      '<label>업체<input class="form-control form-control-sm" autocomplete="off" name="vendor" maxlength="40"></label>' +
      '<div class="gw-vm-actions"><button type="submit" class="btn btn-sm btn-primary">넣기</button></div></form></details>' + list + '</section>';
  }
  function openCar(plate) {
    cur = S.cars.filter(function (c) { return c.plate === plate; })[0];
    box.hidden = false;
    document.body.classList.add('gw-drawer-open');
    $('vmCarTitle').textContent = cur ? cur.label + ' · ' + cur.plate : plate;
    $('vmCarBody').innerHTML = '<p class="text-muted">불러오는 중…</p>';
    $('vmCarTitle').focus();
    return api('/api/vehicles/manage/car?plate=' + encodeURIComponent(plate)).then(function (j) {
      $('vmCarBody').innerHTML = '<div class="gw-vh-msg" id="vmCarMsg" role="status" aria-live="polite"></div>' +
        infoForm(cur, j.info) + costSection(cur, j.costs) + ((cur && !cur.uplus) || j.odometer.length ? odoSection(cur, j.odometer) : '');
    }).catch(function (e) { $('vmCarBody').innerHTML = '<p class="text-danger">' + esc(e.message) + '</p>'; });
  }
  function carMsg(text, kind) {
    var el = $('vmCarMsg');
    if (!el) return;
    el.className = 'gw-vh-msg' + (text ? ' ' + (kind || 'info') : '');
    el.textContent = text || '';
    // 아래쪽 칸에서 저장해도 결과 글이 보이게(펼침창 맨 위에 있음 — 점검 11-22)
    if (text && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' });
  }
  function formData(form) {
    var o = {};
    Array.prototype.forEach.call(form.elements, function (el) { if (el.name) o[el.name] = el.value; });
    return o;
  }
  function after(text) {
    var plate = cur.plate;
    return Promise.all([load(), openCar(plate)]).then(function () { carMsg(text, 'ok'); });
  }
  box.addEventListener('submit', function (e) {
    e.preventDefault();
    var f = e.target, btn = f.querySelector('button[type=submit]');
    var data = formData(f), url, done;
    if (f.id === 'vmInfo') { url = '/api/vehicles/manage/info/' + encodeURIComponent(cur.plate); done = '보험·검사 정보를 저장했어요.'; }
    else if (f.id === 'vmCost') { data.plate = cur.plate; url = '/api/vehicles/manage/costs'; done = '비용을 넣었어요.'; }
    else if (f.id === 'vmOdo') { data.plate = cur.plate; url = '/api/vehicles/manage/odometer'; done = '계기판 km 를 넣었어요.'; }
    else return;
    btn.disabled = true;
    post(url, data).then(function () { return after(done); })
      .catch(function (err) { carMsg(err.message, 'err'); btn.disabled = false; });
  });
  box.addEventListener('click', function (ev) {
    if (ev.target.closest('[data-close]')) { closeCar(); return; }
    var del = ev.target.closest('[data-del]');
    if (del && window.confirm('직접 넣은 이 비용을 지울까요?')) {
      post('/api/vehicles/manage/costs/' + del.getAttribute('data-del') + '/delete', {})
        .then(function () { return after('지웠어요.'); }).catch(function (err) { carMsg(err.message, 'err'); });
    }
  });
  $('vmRows').addEventListener('click', function (ev) {
    var tr = ev.target.closest('tr[data-plate]');
    if (tr) { opener = tr; openCar(tr.getAttribute('data-plate')); }
  });
  $('vmRows').addEventListener('keydown', function (ev) {
    var tr = ev.target.closest('tr[data-plate]');
    if (tr && (ev.key === 'Enter' || ev.key === ' ')) { ev.preventDefault(); opener = tr; openCar(tr.getAttribute('data-plate')); }
  });
  document.addEventListener('keydown', function (ev) {
    if (ev.key === 'Escape' && !box.hidden) closeCar();
    trapTab(ev);
  });

  load();
})();
