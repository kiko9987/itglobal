/* 그룹웨어 휴가 화면 — 내 휴가 / 전사 현황(달력·직원별). 보기 전용(신청은 슬랙 휴가 앱).
   window.GW_LEAVE = { mode: 'me' | 'team' } */
(function () {
  'use strict';

  var CFG = window.GW_LEAVE || {};
  var WD = ['월', '화', '수', '목', '금', '토', '일'];

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function api(url) {
    return fetch(url, { credentials: 'same-origin' }).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok || j.success === false) throw new Error(j.message || ('불러오지 못했습니다 (' + r.status + ')'));
        return j;
      });
    });
  }

  function days(x) {
    if (x == null) return '-';
    var n = Math.round(Number(x) * 10) / 10;
    return (n % 1 === 0 ? n.toFixed(0) : n.toFixed(1)) + '일';
  }

  function md(s) {
    if (!s) return '';
    var p = s.split('-');
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return Number(p[1]) + '월 ' + Number(p[2]) + '일 (' + WD[(d.getDay() + 6) % 7] + ')';
  }

  function period(e, withYear) {
    var y = withYear ? e.start.slice(0, 4) + '년 ' : '';
    if (e.half) return y + md(e.start) + ' ' + e.half;
    if (!e.end || e.end === e.start) return y + md(e.start);
    return y + md(e.start) + ' ~ ' + md(e.end);
  }

  function statusBadge(s) {
    var cls = { '승인': 'gw-badge-active', '대기': 'gw-badge-leave', '반려': 'gw-badge-reject', '취소': 'gw-badge-resigned' }[s] || 'gw-badge-resigned';
    return '<span class="gw-badge ' + cls + '">' + esc(s === '대기' ? '승인 대기' : s) + '</span>';
  }

  // ── 한 사람 휴가(내 휴가·직원별 상세 공용) ───────────────────────
  function renderSummary(L, opts) {
    opts = opts || {};
    var h = '';
    if (!L.linked) {
      h += '<div class="gw-note mb-2">휴가 장부에 아직 연결되지 않아 남은 연차를 계산하지 못했습니다. 지난 기록만 보입니다.</div>';
    }
    h += '<div class="gw-stats">' +
      '<div class="gw-stat gw-stat-main"><div class="k">남은 연차</div><div class="v">' + (L.balance == null ? '-' : days(L.balance)) + '</div></div>' +
      '<div class="gw-stat"><div class="k">' + L.year + '년 쓴 연차</div><div class="v">' + days(L.annual_used) + '</div></div>' +
      '<div class="gw-stat"><div class="k">' + L.year + '년 병가</div><div class="v">' + days(L.sick_used) + ' <small>/ ' + L.sick_limit + '일</small></div></div>' +
      '<div class="gw-stat"><div class="k">승인 대기</div><div class="v">' + L.pending + '건</div></div>' +
      '</div>';

    if (L.grants && L.grants.length) {
      h += '<div class="gw-sec-title">언제까지 써야 하나요</div><ul class="gw-plain-list">';
      L.grants.forEach(function (g) {
        h += '<li><strong>' + days(g.days) + '</strong> — ' + esc(g.valid_to.slice(0, 4)) + '년 ' + esc(md(g.valid_to)) + '까지</li>';
      });
      h += '</ul>';
    }

    if (L.upcoming && L.upcoming.length) {
      h += '<div class="gw-sec-title">다가오는 휴가</div><ul class="gw-plain-list">';
      L.upcoming.forEach(function (e) {
        h += '<li>' + esc(period(e)) + ' · ' + esc(e.type) + ' · ' + days(e.days) + ' ' + statusBadge(e.status) + '</li>';
      });
      h += '</ul>';
    }

    if (!opts.hideApplyNote) {
      h += '<div class="gw-callout"><i class="fab fa-slack me-1" aria-hidden="true"></i>휴가 신청·취소는 슬랙 <strong>휴가</strong> 앱에서 해요.</div>';
    }

    var years = [];
    (L.entries || []).forEach(function (e) { var y = e.start.slice(0, 4); if (years.indexOf(y) === -1) years.push(y); });
    h += '<div class="gw-sec-head"><div class="gw-sec-title m-0">휴가 기록</div>' +
      (years.length > 1 ? '<select class="form-select form-select-sm gw-year-select" aria-label="연도">' +
        '<option value="">전체</option>' + years.map(function (y) { return '<option value="' + y + '">' + y + '년</option>'; }).join('') +
        '</select>' : '') + '</div>';
    h += '<div class="gw-table-wrap"><table class="gw-table gw-leave-table"><thead><tr><th>기간</th><th>종류</th><th class="n">일수</th><th>상태</th><th>사유</th></tr></thead><tbody>';
    if (!L.entries || !L.entries.length) {
      h += '<tr><td colspan="5" class="gw-empty">기록이 없습니다.</td></tr>';
    } else {
      L.entries.forEach(function (e) {
        h += '<tr data-year="' + esc(e.start.slice(0, 4)) + '"' + (e.status === '취소' || e.status === '반려' ? ' class="muted"' : '') + '>' +
          '<td>' + esc(period(e, true)) + '</td><td>' + esc(e.type) + '</td><td class="n">' + days(e.days) + '</td>' +
          '<td>' + statusBadge(e.status) + '</td><td class="gw-reason">' + esc(e.reason) + '</td></tr>';
      });
    }
    h += '</tbody></table></div>';
    h += '<div class="gw-note">2026년 11월 이전 기록은 플렉스에서 옮겨 온 것입니다.</div>';
    return h;
  }

  function bindYearFilter(root) {
    var sel = root.querySelector('.gw-year-select');
    if (!sel) return;
    sel.addEventListener('change', function () {
      var y = sel.value;
      root.querySelectorAll('.gw-leave-table tbody tr[data-year]').forEach(function (tr) {
        tr.hidden = !!y && tr.getAttribute('data-year') !== y;
      });
    });
  }

  // ── 내 휴가 ────────────────────────────────────────────────
  if (CFG.mode === 'me') {
    var mine = document.getElementById('lvMine');
    api('/api/leave/me').then(function (j) {
      mine.innerHTML = renderSummary(j.leave);
      bindYearFilter(mine);
    }).catch(function (e) {
      mine.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
    });
    return;
  }

  if (CFG.mode !== 'team') return;

  // ── 전사 현황: 탭 ───────────────────────────────────────────
  var tabs = { cal: document.getElementById('lvTabCal'), list: document.getElementById('lvTabList') };
  var panes = { cal: document.getElementById('lvCalPane'), list: document.getElementById('lvListPane') };
  var listLoaded = false;

  function showTab(key) {
    Object.keys(tabs).forEach(function (k) {
      tabs[k].setAttribute('aria-selected', String(k === key));
      panes[k].hidden = k !== key;
    });
    try { localStorage.setItem('gw.leave.tab', key); } catch (e) { /* 무시 */ }
    if (key === 'list' && !listLoaded) loadTeam();
  }
  tabs.cal.addEventListener('click', function () { showTab('cal'); });
  tabs.list.addEventListener('click', function () { showTab('list'); });

  // ── 달력 ───────────────────────────────────────────────────
  var calEl = document.getElementById('lvCal');
  var titleEl = document.getElementById('lvMonthTitle');
  var now = new Date();
  var cur = { y: now.getFullYear(), m: now.getMonth() + 1 };

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }

  function renderCal(entries) {
    var first = new Date(cur.y, cur.m - 1, 1);
    var lastDay = new Date(cur.y, cur.m, 0).getDate();
    var lead = (first.getDay() + 6) % 7;     // 월요일 시작
    var todayIso = iso(now.getFullYear(), now.getMonth() + 1, now.getDate());
    var h = '<div class="gw-cal-grid">';
    WD.forEach(function (w, i) { h += '<div class="gw-cal-wd' + (i >= 5 ? ' we' : '') + '">' + w + '</div>'; });
    for (var i = 0; i < lead; i++) h += '<div class="gw-cal-day out"></div>';
    for (var d = 1; d <= lastDay; d++) {
      var di = iso(cur.y, cur.m, d);
      var wd = (lead + d - 1) % 7;
      var todays = entries.filter(function (e) { return e.start <= di && di <= e.end; });
      h += '<div class="gw-cal-day' + (wd >= 5 ? ' we' : '') + (di === todayIso ? ' today' : '') + '">' +
        '<div class="gw-cal-num">' + d + '</div>';
      todays.forEach(function (e) {
        var label = e.half ? e.half + ' 반차' : (e.type === '연차' ? '' : e.type);
        h += '<div class="gw-cal-chip ' + (e.status === '대기' ? 'wait' : 'ok') + '" title="' +
          esc(e.name + ' · ' + e.type + (e.half ? ' ' + e.half : '') + (e.status === '대기' ? ' · 승인 대기' : '')) + '">' +
          esc(e.name) + (label ? ' <small>' + esc(label) + '</small>' : '') + '</div>';
      });
      h += '</div>';
    }
    var tail = (7 - ((lead + lastDay) % 7)) % 7;
    for (var t = 0; t < tail; t++) h += '<div class="gw-cal-day out"></div>';
    h += '</div>';
    calEl.innerHTML = h;
  }

  function loadCal() {
    titleEl.textContent = cur.y + '년 ' + cur.m + '월';
    calEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    api('/api/leave/calendar?month=' + cur.y + '-' + pad(cur.m)).then(function (j) {
      renderCal(j.entries || []);
    }).catch(function (e) {
      calEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
    });
  }

  document.getElementById('lvPrev').addEventListener('click', function () {
    cur.m -= 1; if (cur.m < 1) { cur.m = 12; cur.y -= 1; } loadCal();
  });
  document.getElementById('lvNext').addEventListener('click', function () {
    cur.m += 1; if (cur.m > 12) { cur.m = 1; cur.y += 1; } loadCal();
  });
  document.getElementById('lvToday').addEventListener('click', function () {
    cur = { y: now.getFullYear(), m: now.getMonth() + 1 }; loadCal();
  });

  // ── 직원별 ─────────────────────────────────────────────────
  var tbody = document.querySelector('#lvTeamTable tbody');

  function loadTeam() {
    api('/api/leave/team').then(function (j) {
      listLoaded = true;
      var rows = j.people || [];
      tbody.innerHTML = rows.map(function (p) {
        var nxt = p.next ? esc(period(p.next)) + (p.next.status === '대기' ? ' <span class="gw-badge gw-badge-leave">대기</span>' : '') : '<span class="text-muted">-</span>';
        return '<tr tabindex="0" data-id="' + p.person_id + '" data-name="' + esc(p.name) + '">' +
          '<td><strong>' + esc(p.name) + '</strong>' + (p.access === 'field' ? ' <span class="gw-badge gw-badge-field">' + esc(p.job_role || '현장') + '</span>' : '') + '</td>' +
          '<td>' + esc(p.department) + '</td><td>' + esc(p.hire_date) + '</td>' +
          '<td class="n">' + (p.linked ? days(p.balance) : '<span class="text-muted" title="휴가 장부에 없음">-</span>') + '</td>' +
          '<td class="n">' + days(p.annual_used) + '</td><td class="n">' + days(p.sick_used) + '</td>' +
          '<td class="n">' + (p.pending ? '<strong>' + p.pending + '</strong>' : '0') + '</td><td>' + nxt + '</td></tr>';
      }).join('') || '<tr><td colspan="8" class="gw-empty">재직자가 없습니다.</td></tr>';
    }).catch(function (e) {
      tbody.innerHTML = '<tr><td colspan="8" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  var modalEl = document.getElementById('lvPersonModal');
  var bodyEl = document.getElementById('lvPersonBody');

  function openPerson(id, name) {
    document.getElementById('lvPersonTitle').textContent = name + ' 휴가';
    bodyEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    if (window.bootstrap) window.bootstrap.Modal.getOrCreateInstance(modalEl).show();
    api('/api/leave/person/' + id).then(function (j) {
      bodyEl.innerHTML = renderSummary(j.leave, { hideApplyNote: true });
      bindYearFilter(bodyEl);
    }).catch(function (e) {
      bodyEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
    });
  }

  tbody.addEventListener('click', function (e) {
    var tr = e.target.closest('tr[data-id]');
    if (tr) openPerson(tr.getAttribute('data-id'), tr.getAttribute('data-name'));
  });
  tbody.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var tr = e.target.closest('tr[data-id]');
    if (tr) { e.preventDefault(); openPerson(tr.getAttribute('data-id'), tr.getAttribute('data-name')); }
  });

  var saved = 'cal';
  try { saved = localStorage.getItem('gw.leave.tab') || 'cal'; } catch (e) { /* 무시 */ }
  loadCal();
  showTab(saved === 'list' ? 'list' : 'cal');
})();
