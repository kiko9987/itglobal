/* 그룹웨어 쉬는 날(11-25) — 연도별 목록, 경영지원은 임시공휴일·회사 지정 휴일을 넣고 고치고 지운다.
   저장 전에 그 날짜에 걸린 휴가(승인·대기)를 미리 보여 주고, 저장하면 서버가 다시 세고 본인에게 DM 한다.
   권한·규칙은 서버(services/company_holidays.py)가 다시 확인한다. */
(function () {
  'use strict';

  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function api(url, opts) {
    opts = opts || {};
    var method = opts.method || 'GET';
    var init = { method: method, credentials: 'same-origin' };
    if (opts.body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(opts.body); }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '처리하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function toDate(iso) { return new Date(Number(iso.slice(0, 4)), Number(iso.slice(5, 7)) - 1, Number(iso.slice(8, 10))); }
  function md(iso, withMonth) {
    var d = toDate(iso);
    return (withMonth === false ? '' : (d.getMonth() + 1) + '월 ') + d.getDate() + '일 (' + WD[d.getDay()] + ')';
  }
  function span(s, e) {
    if (s === e) return md(s);
    var same = s.slice(0, 7) === e.slice(0, 7);
    return md(s) + ' – ' + md(e, !same);
  }
  // 받침 있으면 a(이·을), 없으면 b(가·를) — 이름 뒤 조사
  function josa(word, a, b) {
    var c = String(word || ' ').slice(-1).charCodeAt(0);
    return (c >= 0xAC00 && c <= 0xD7A3 && (c - 0xAC00) % 28) ? a : b;
  }
  function dayCount(x) { return (Math.round(x * 10) / 10) + '일'; }
  function weekendOnly(s, e) {
    for (var d = toDate(s); d <= toDate(e); d.setDate(d.getDate() + 1)) {
      if (d.getDay() !== 0 && d.getDay() !== 6) return false;
    }
    return true;
  }

  var page = document.getElementById('hdPage');
  if (!page) return;
  var canEdit = page.getAttribute('data-can-edit') === '1';
  var listEl = document.getElementById('hdList');
  var titleEl = document.getElementById('hdTitle');
  var sumEl = document.getElementById('hdSum');
  var nextEl = document.getElementById('hdNextOff');
  var flashEl = document.getElementById('hdFlash');

  // 한 해 달력(11-25 다듬기) — 왼쪽에 12달, 오른쪽에 넣기 칸·목록. 자리는 JS 가 만든다(재시작 전 템플릿에서도 바로)
  page.classList.remove('gw-page-narrow');
  page.classList.add('gw-hd-wide');
  var bodyEl = document.getElementById('hdBody');
  var yearCalEl = document.getElementById('hdYearCal');
  if (!bodyEl) {
    bodyEl = document.createElement('div');
    bodyEl.className = 'gw-hd-body';
    bodyEl.id = 'hdBody';
    listEl.parentNode.insertBefore(bodyEl, listEl);
    yearCalEl = document.createElement('section');
    yearCalEl.className = 'gw-hd-yearcal';
    yearCalEl.id = 'hdYearCal';
    yearCalEl.setAttribute('aria-label', '한 해 달력');
    var mainEl = document.createElement('div');
    mainEl.className = 'gw-hd-main';
    var formBox = document.getElementById('hdForm');
    if (formBox) mainEl.appendChild(formBox);
    mainEl.appendChild(listEl);
    bodyEl.appendChild(yearCalEl);
    bodyEl.appendChild(mainEl);
  }

  var now = new Date();
  var todayIso = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
  var year = now.getFullYear();
  var data = null;
  var seq = 0;

  function flash(msg) {
    flashEl.textContent = msg;
    flashEl.className = 'gw-lv-flash ok';
    flashEl.hidden = !msg;
  }

  // ── 목록 ──────────────────────────────────────────────────────────────
  var KIND_CLASS = { legal: 'legal', temp: 'temp', company: 'company' };

  function rowHtml(e) {
    var past = e.end < todayIso;
    var cls = 'gw-hd-row' + (past ? ' past' : '') + (e.id === nextId ? ' next' : '');
    var h = '<li class="' + cls + '" id="hd-r-' + e.start + '" tabindex="-1">';
    h += '<div class="gw-hd-when">' + esc(span(e.start, e.end)) +
      (e.days > 1 ? ' <span class="gw-hd-days">' + e.days + '일</span>' : '') +
      (weekendOnly(e.start, e.end) ? ' <span class="gw-hd-wkend">주말</span>' : '') + '</div>';
    h += '<div class="gw-hd-what"><span class="gw-hd-dot ' + (KIND_CLASS[e.kind] || 'legal') + '" aria-hidden="true"></span>' +
      '<span class="gw-hd-name">' + esc(e.name) + '</span>' +
      (e.kind !== 'legal' ? '<span class="gw-badge gw-hd-kind-' + (KIND_CLASS[e.kind] || 'legal') + '">' + esc(e.kind_label) + '</span>'
        : '<span class="visually-hidden">(법정 공휴일)</span>') +
      (e.yearly && e.kind !== 'legal' ? '<span class="gw-hd-tag">매년</span>' : '') +
      (e.id === nextId ? '<span class="gw-hd-tag next">다음 쉬는 날</span>' : '') +
      (e.note ? '<span class="gw-hd-memo">' + esc(e.note) + '</span>' : '') + '</div>';
    if (canEdit && e.editable) {
      h += '<div class="gw-hd-row-act">' +
        '<button type="button" class="btn btn-sm btn-outline-secondary" data-act="edit" data-id="' + e.id + '" ' +
        'aria-label="' + esc(e.name + ' 고치기') + '">고치기</button>' +
        '<button type="button" class="btn btn-sm btn-outline-danger" data-act="delete" data-id="' + e.id + '" ' +
        'aria-label="' + esc(e.name + ' 지우기') + '">지우기</button></div>';
    }
    return h + '</li>';
  }

  var nextId = null;

  function render(j) {
    data = j;
    var ents = j.entries || [];
    titleEl.textContent = j.year + '년';
    var bk = j.by_kind || {};
    var parts = ['법정 공휴일 ' + (bk.legal || 0) + '일'];
    if (bk.temp) parts.push('임시공휴일 ' + bk.temp + '일');
    parts.push('회사 지정 ' + (bk.company || 0) + '일');
    sumEl.innerHTML = '<strong>쉬는 날 ' + j.total + '일</strong><span class="gw-hd-sep" aria-hidden="true">·</span>평일 ' +
      j.weekdays + '일<span class="gw-hd-sep" aria-hidden="true">·</span>' + esc(parts.join(' · '));

    // 다음 쉬는 날 — 올해 목록에서 오늘 뒤 첫 줄(법정 연휴는 id 가 없어 시작 날짜로 표시)
    var nx = null;
    ents.forEach(function (e) { if (!nx && e.end >= todayIso) nx = e; });
    nextId = nx ? (nx.id != null ? nx.id : 'L' + nx.start) : null;
    ents.forEach(function (e) { if (e.id == null) e.id = 'L' + e.start; });
    if (nx && String(j.year) === todayIso.slice(0, 4)) {
      var left = Math.round((toDate(nx.start) - toDate(todayIso)) / 86400000);
      nextEl.innerHTML = '<i class="fas fa-umbrella-beach" aria-hidden="true"></i> 다음 쉬는 날 <strong>' +
        esc(span(nx.start, nx.end)) + ' ' + esc(nx.name) + '</strong>' +
        (left > 0 ? ' · ' + left + '일 남았어요' : ' · 오늘이에요');
      nextEl.hidden = false;
    } else {
      nextEl.hidden = true;
    }

    if (!ents.length) {
      listEl.innerHTML = '<div class="gw-empty gw-empty-box">' + j.year + '년 쉬는 날이 없어요.</div>';
      renderYearCal(j);
      return;
    }
    var thisYear = String(j.year) === todayIso.slice(0, 4);
    var past = thisYear ? ents.filter(function (e) { return e.end < todayIso; }) : [];
    var rest = thisYear ? ents.filter(function (e) { return e.end >= todayIso; }) : ents;
    var h = '';
    if (past.length) {
      h += '<details class="gw-hd-pastbox" id="hdPastBox"><summary><i class="fas fa-chevron-right" aria-hidden="true"></i> 지난 쉬는 날 ' +
        past.length + '개</summary>' + monthsHtml(past) + '</details>';
    }
    h += rest.length ? monthsHtml(rest) : '<div class="gw-empty gw-empty-box">올해 남은 쉬는 날이 없어요.</div>';
    listEl.innerHTML = h;
    renderYearCal(j);
  }

  function monthsHtml(list) {
    var h = '', month = 0;
    list.forEach(function (e) {
      var m = Number(e.start.slice(5, 7));
      if (m !== month) {
        if (month) h += '</ul></section>';
        month = m;
        h += '<section class="gw-hd-month"><h3 class="gw-hd-mh">' + m + '월</h3><ul class="gw-hd-rows">';
      }
      h += rowHtml(e);
    });
    return h + '</ul></section>';
  }

  // ── 한 해 달력 — 12달, 쉬는 날은 구분 색, 누르면 목록의 그 줄로 ──────────────
  var KIND_ORDER = { legal: 0, temp: 1, company: 2 };
  function renderYearCal(j) {
    if (!yearCalEl) return;
    var map = {};
    (j.entries || []).forEach(function (e) {
      for (var d = toDate(e.start); d <= toDate(e.end); d.setDate(d.getDate() + 1)) {
        var k = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
        if (!map[k] || KIND_ORDER[e.kind] < KIND_ORDER[map[k].kind]) map[k] = { kind: e.kind, name: e.name, start: e.start };
      }
    });
    // 휴대폰은 접어 둠 — 칸이 20px 라 누르기 어렵고 목록까지 스크롤이 길다. 펼친 상태는 기억(다시 그릴 때)
    var wide = !(window.matchMedia && window.matchMedia('(max-width: 767.98px)').matches);
    var wasOpen = yearCalEl.querySelector('details') ? yearCalEl.querySelector('details').open : wide;
    var y = j.year, h = '<details class="gw-hd-calfold"' + (wide || wasOpen ? ' open' : '') + '><summary>' +
      '<i class="fas fa-calendar-days" aria-hidden="true"></i> ' + y + '년 한눈에 보기</summary><div class="gw-hd-months">';
    for (var m = 1; m <= 12; m++) {
      var first = new Date(y, m - 1, 1), last = new Date(y, m, 0).getDate();
      var lead = (first.getDay() + 6) % 7;
      var n = 0;
      var cells = '';
      for (var i = 0; i < lead; i++) cells += '<span class="gw-hd-c out"></span>';
      for (var dd = 1; dd <= last; dd++) {
        var iso = y + '-' + pad(m) + '-' + pad(dd);
        var wd = (lead + dd - 1) % 7;
        var hol = map[iso];
        var cls = 'gw-hd-c' + (wd >= 5 ? ' we' : '') + (iso < todayIso ? ' past' : '') + (iso === todayIso ? ' today' : '');
        if (hol) {
          n++;
          cells += '<button type="button" class="' + cls + ' hol ' + hol.kind + '" data-goto="' + hol.start + '" title="' + esc(hol.name) +
            '" aria-label="' + m + '월 ' + dd + '일 ' + esc(hol.name) + '">' + dd + '</button>';
        } else {
          cells += '<span class="' + cls + '"' + (iso === todayIso ? ' aria-label="오늘 ' + m + '월 ' + dd + '일"' : '') + '>' + dd + '</span>';
        }
      }
      var monthPast = thisMonthKey(y, m) < todayIso.slice(0, 7);
      h += '<div class="gw-hd-mcal' + (monthPast ? ' past' : '') + '"><div class="gw-hd-mcal-h"><span>' + m + '월</span>' +
        (n ? '<span class="gw-hd-mcal-n">' + n + '일</span>' : '') + '</div>' +
        '<div class="gw-hd-mgrid">' +
        '<span class="gw-hd-wd">월</span><span class="gw-hd-wd">화</span><span class="gw-hd-wd">수</span><span class="gw-hd-wd">목</span>' +
        '<span class="gw-hd-wd">금</span><span class="gw-hd-wd we">토</span><span class="gw-hd-wd we">일</span>' + cells + '</div></div>';
    }
    h += '</div><ul class="gw-hd-legend" aria-label="색 설명"><li><span class="gw-hd-dot legal"></span>법정 공휴일</li>' +
      '<li><span class="gw-hd-dot temp"></span>임시공휴일</li><li><span class="gw-hd-dot company"></span>회사 지정</li>' +
      '<li><span class="gw-hd-todaymark"></span>오늘</li></ul></details>';
    yearCalEl.innerHTML = h;
  }
  function thisMonthKey(y, m) { return y + '-' + pad(m); }

  if (yearCalEl) {
    yearCalEl.addEventListener('click', function (ev) {
      var b = ev.target.closest('[data-goto]');
      if (!b) return;
      var row = document.getElementById('hd-r-' + b.getAttribute('data-goto'));
      if (!row) return;
      var box = row.closest('details');
      if (box && !box.open) box.open = true;
      row.scrollIntoView({ behavior: 'smooth', block: 'center' });
      row.classList.remove('is-hit');
      void row.offsetWidth;
      row.classList.add('is-hit');
      try { row.focus({ preventScroll: true }); } catch (e) { row.focus(); }
    });
  }

  function load(focusTitle) {
    titleEl.textContent = year + '년';
    var my = ++seq;
    listEl.setAttribute('aria-busy', 'true');
    api('/api/leave/holidays?year=' + year).then(function (j) {
      if (my !== seq) return;
      listEl.removeAttribute('aria-busy');
      render(j);
      if (focusTitle) titleEl.focus();
    }).catch(function (err) {
      if (my !== seq) return;
      listEl.removeAttribute('aria-busy');
      listEl.innerHTML = '<div class="gw-empty gw-empty-box">' + esc(err.message) + '</div>';
    });
  }

  document.getElementById('hdPrev').addEventListener('click', function () { year -= 1; load(); });
  document.getElementById('hdNext').addEventListener('click', function () { year += 1; load(); });
  document.getElementById('hdThis').addEventListener('click', function () { year = now.getFullYear(); load(); });
  document.addEventListener('keydown', function (ev) {
    if (ev.target.closest('input, textarea, select, [contenteditable]')) return;
    if (ev.key === 'ArrowLeft' && ev.altKey) { year -= 1; load(); }
    if (ev.key === 'ArrowRight' && ev.altKey) { year += 1; load(); }
  });

  // ── 넣기·고치기·지우기(경영지원) ─────────────────────────────────────────
  var form = document.getElementById('hdForm');
  if (canEdit && form) {
    var addBtn = document.getElementById('hdAdd');
    var formTitle = document.getElementById('hdFormTitle');
    var fields = document.getElementById('hdFields');
    var delWhat = document.getElementById('hdDelWhat');
    var impactEl = document.getElementById('hdImpact');
    var errEl = document.getElementById('hdErr');
    var saveBtn = document.getElementById('hdSave');
    var yearlyBox = document.getElementById('hdYearlyBox');
    var f = {
      start: document.getElementById('hdStart'), end: document.getElementById('hdEnd'),
      name: document.getElementById('hdName'), yearly: document.getElementById('hdYearly'),
      note: document.getElementById('hdNote')
    };
    var mode = 'add', editId = null, expect = null, busy = false, opener = null;

    function kind() {
      var k = form.querySelector('input[name="kind"]:checked');
      return k ? k.value : 'company';
    }
    function syncYearly() {
      var company = kind() === 'company';
      yearlyBox.hidden = !company;
      if (!company) f.yearly.checked = false;
    }
    function resetImpact() {
      expect = null;
      impactEl.hidden = true;
      impactEl.innerHTML = '';
      saveBtn.textContent = mode === 'delete' ? '지우기' : '저장';
    }
    function body() {
      return { kind: kind(), start: f.start.value, end: f.end.value, name: f.name.value.trim(),
        yearly: f.yearly.checked, note: f.note.value.trim() };
    }
    function url() {
      if (mode === 'add') return '/api/leave/holidays/add';
      return '/api/leave/holidays/' + encodeURIComponent(editId) + '/' + mode;
    }

    function open(m, ent) {
      mode = m;
      editId = ent ? ent.id : null;
      opener = document.activeElement;
      errEl.textContent = '';
      resetImpact();
      flash('');
      form.classList.toggle('is-delete', m === 'delete');
      saveBtn.className = 'btn ' + (m === 'delete' ? 'btn-danger' : 'btn-primary');
      fields.hidden = m === 'delete';
      delWhat.hidden = m !== 'delete';
      if (m === 'add') {
        formTitle.textContent = '쉬는 날 추가';
        form.reset();
        document.getElementById('hdKindCompany').checked = true;
      } else {
        // 매년 반복 줄은 처음 넣은 날짜(base)로 고친다 — 다른 해 화면에서 열어도 같은 줄
        var s = ent.base_start || ent.start, e = ent.base_end || ent.end;
        formTitle.textContent = m === 'edit' ? '쉬는 날 고치기' : '쉬는 날 지우기';
        document.getElementById(ent.kind === 'temp' ? 'hdKindTemp' : 'hdKindCompany').checked = true;
        f.start.value = s;
        f.end.value = e !== s ? e : '';
        f.name.value = ent.name;
        f.yearly.checked = !!ent.yearly;
        f.note.value = ent.note || '';
        delWhat.innerHTML = '지울 쉬는 날: <strong>' + esc(span(ent.start, ent.end)) + ' ' + esc(ent.name) + '</strong>' +
          (ent.yearly ? ' (매년 반복 — 모든 해에서 빠져요)' : '') + '<br>지우면 이 날은 다시 근무일로 셉니다.';
      }
      syncYearly();
      form.hidden = false;
      if (addBtn) addBtn.setAttribute('aria-expanded', 'true');
      form.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      (m === 'delete' ? saveBtn : (m === 'add' ? f.start : f.name)).focus({ preventScroll: true });
      if (m === 'delete') check();         // 지우기는 바로 걸린 휴가부터 보여 줌
    }

    function close() {
      form.hidden = true;
      if (addBtn) addBtn.setAttribute('aria-expanded', 'false');
      if (opener && document.contains(opener)) opener.focus();
      else if (addBtn) addBtn.focus();
    }

    function showImpact(j) {
      var a = j.affected || [];
      expect = a.length;
      if (!a.length) {
        impactEl.hidden = true;
        return;
      }
      var h = '<p class="gw-hd-impact-head"><i class="fas fa-circle-info" aria-hidden="true"></i> ' +
        '이 날짜에 걸린 휴가 <strong>' + a.length + '건</strong>의 일수를 다시 세어요. ' +
        (mode === 'delete' ? '지우면' : '저장하면') + ' 아래처럼 바뀌고 본인에게 슬랙 DM이 가요.</p>';
      h += '<ul class="gw-hd-impact-list">';
      a.forEach(function (x) {
        var diff = x.delta < 0 ? '연차 ' + dayCount(-x.delta) + ' 돌려줌' : '연차 ' + dayCount(x.delta) + ' 더 씀';
        h += '<li><span class="gw-hd-who">' + esc(x.name) + '</span>' +
          '<span>' + esc(x.period) + ' ' + esc(x.type_label) + (x.status === 'approved' ? '' : ' · ' + esc(x.status_label)) + '</span>' +
          '<span class="gw-hd-delta">' + dayCount(x.old_days) + ' → <strong>' + dayCount(x.new_days) + '</strong>' +
          (x.deducts ? ' <small>' + diff + '</small>' : '') + '</span></li>';
      });
      h += '</ul>';
      impactEl.innerHTML = h;
      impactEl.hidden = false;
      saveBtn.textContent = (mode === 'delete' ? '휴가 ' + a.length + '건 다시 세고 지우기' : '휴가 ' + a.length + '건 다시 세고 저장');
    }

    function check() {
      busy = true;
      saveBtn.disabled = true;
      return api(url(), { method: 'POST', body: Object.assign({ preview: true }, mode === 'delete' ? {} : body()) })
        .then(function (j) { showImpact(j); return j; })
        .catch(function (err) { errEl.textContent = err.message; throw err; })
        .finally(function () { busy = false; saveBtn.disabled = false; });
    }

    function save() {
      busy = true;
      saveBtn.disabled = true;
      var b = Object.assign({ expect: expect }, mode === 'delete' ? {} : body());
      api(url(), { method: 'POST', body: b }).then(function (j) {
        var nm = mode === 'delete' ? '' : body().name;
        var obj = '「' + nm + '」' + josa(nm, '을', '를');
        var msg = mode === 'add' ? obj + ' 쉬는 날로 넣었어요.' :
          (mode === 'edit' ? obj + ' 고쳤어요.' : '쉬는 날에서 지웠어요.');
        if (j.count) msg += ' 걸린 휴가 ' + j.count + '건을 다시 셌고, DM ' + (j.dm_sent || 0) + '건을 보냈어요.';
        if (mode !== 'delete' && f.start.value) year = Number(f.start.value.slice(0, 4));
        form.hidden = true;
        if (addBtn) addBtn.setAttribute('aria-expanded', 'false');
        flash(msg);
        flashEl.focus();
        load();
      }).catch(function (err) {
        errEl.textContent = err.message;
        if (/다시 확인/.test(err.message)) check().catch(function () {});   // 그사이 바뀌었으면 새로 보여 줌
      }).finally(function () { busy = false; saveBtn.disabled = false; });
    }

    if (addBtn) addBtn.addEventListener('click', function () { open('add'); });
    document.getElementById('hdCancel').addEventListener('click', close);
    form.addEventListener('keydown', function (ev) { if (ev.key === 'Escape') { ev.preventDefault(); close(); } });
    form.querySelectorAll('input[name="kind"]').forEach(function (r) { r.addEventListener('change', function () { syncYearly(); resetImpact(); }); });
    [f.start, f.end, f.yearly].forEach(function (el) { el.addEventListener('change', resetImpact); });
    f.start.addEventListener('change', function () {
      if (f.end.value && f.end.value < f.start.value) f.end.value = '';
      f.end.min = f.start.value;
    });

    form.addEventListener('submit', function (ev) {
      ev.preventDefault();
      if (busy) return;
      errEl.textContent = '';
      if (mode !== 'delete') {
        if (!f.start.value) { errEl.textContent = '날짜를 골라 주세요.'; f.start.focus(); return; }
        if (!f.name.value.trim()) { errEl.textContent = '이름을 적어 주세요.'; f.name.focus(); return; }
      }
      if (expect === null) {                // 처음 누르면 걸린 휴가부터 — 없으면 바로 저장
        check().then(function (j) { if (!(j.affected || []).length) save(); }).catch(function () {});
      } else {
        save();
      }
    });

    listEl.addEventListener('click', function (ev) {
      var b = ev.target.closest('button[data-act]');
      if (!b || !data) return;
      var id = Number(b.getAttribute('data-id'));
      var ent = (data.entries || []).filter(function (e) { return e.id === id; })[0];
      if (ent) open(b.getAttribute('data-act'), ent);
    });
  }

  load();
})();
