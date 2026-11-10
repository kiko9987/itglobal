/* 그룹웨어 홈 — 결재·내 휴가·휴가자·메일. 각 칸은 따로 불러오고, 하나가 실패해도 나머지는 뜬다. */
(function () {
  'use strict';

  var WD = ['일', '월', '화', '수', '목', '금', '토'];

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

  function el(id) { return document.getElementById(id); }
  function fail(id, msg) { var e = el(id); if (e) e.innerHTML = '<div class="gw-empty">' + esc(msg) + '</div>'; }

  function days(x) {
    if (x == null) return '-';
    var n = Math.round(Number(x) * 10) / 10;
    return (n % 1 === 0 ? n.toFixed(0) : n.toFixed(1)) + '일';
  }

  function md(s) {
    var p = s.split('-');
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return Number(p[1]) + '/' + Number(p[2]) + '(' + WD[d.getDay()] + ')';
  }

  function period(e) {
    if (e.half) return md(e.start) + ' ' + e.half;
    if (!e.end || e.end === e.start) return md(e.start);
    return md(e.start) + '~' + md(e.end);
  }

  var now = new Date();
  el('homeDate').textContent = now.getFullYear() + '년 ' + (now.getMonth() + 1) + '월 ' + now.getDate() + '일 ' + WD[now.getDay()] + '요일';

  // ── 결재 ─────────────────────────────────────────
  function renderApprovals(A) {
    var box = el('homeApprovals');
    if (!box) return;
    if (!A) { fail('homeApprovals', '결재를 불러오지 못했습니다.'); return; }
    if (!A.linked) {
      box.innerHTML = '<div class="gw-empty">인사카드에 슬랙 계정이 연결되지 않아 결재를 찾지 못했습니다.</div>';
      return;
    }
    function list(rows, empty) {
      if (!rows.length) return '<div class="gw-home-none">' + esc(empty) + '</div>';
      return '<ul class="gw-home-list">' + rows.map(function (r) {
        var inner = '<span class="gw-home-li-main">' + esc(r.title) + '</span>' +
          '<span class="gw-home-li-sub">' + esc([r.requester, r.what].filter(Boolean).join(' · ')) +
          (r.amount ? ' · ' + r.amount.toLocaleString('ko-KR') + '원' : '') + '</span>' +
          '<span class="gw-badge gw-badge-leave">' + esc(r.status) + '</span>';
        return '<li>' + (r.link ? '<a href="' + esc(r.link) + '" target="_blank" rel="noopener">' + inner + '</a>' : '<div>' + inner + '</div>') + '</li>';
      }).join('') + '</ul>';
    }
    box.innerHTML =
      '<div class="gw-home-split">' +
      '<div><div class="gw-home-label">내가 처리할 것 <strong>' + A.todo_count + '</strong></div>' + list(A.todo, '처리할 결재가 없어요.') + '</div>' +
      '<div><div class="gw-home-label">내가 올린 진행 중 <strong>' + A.mine_count + '</strong></div>' + list(A.mine, '진행 중인 요청이 없어요.') + '</div>' +
      '</div>';
  }

  // ── 내 휴가 ──────────────────────────────────────
  function renderLeave(L) {
    var box = el('homeLeave');
    if (!L) { box.innerHTML = '<div class="gw-empty">인사카드가 연결되지 않았습니다.</div>'; return; }
    var h = '<div class="gw-home-big"><span class="k">남은 연차</span><span class="v">' + (L.balance == null ? '-' : days(L.balance)) + '</span></div>';
    if (L.grants && L.grants.length) {
      var g = L.grants[0];
      h += '<div class="gw-home-li-sub">' + esc(days(g.days)) + ' — ' + esc(g.valid_to.slice(0, 4)) + '년 ' + esc(md(g.valid_to)) + '까지</div>';
    }
    if (L.upcoming && L.upcoming.length) {
      h += '<div class="gw-home-label mt-2">다가오는 휴가</div><ul class="gw-home-list">' + L.upcoming.slice(0, 3).map(function (e) {
        return '<li><div><span class="gw-home-li-main">' + esc(period(e)) + '</span><span class="gw-home-li-sub">' + esc(e.type) + '</span>' +
          (e.status === '대기' ? '<span class="gw-badge gw-badge-leave">승인 대기</span>' : '<span class="gw-badge gw-badge-active">승인</span>') + '</div></li>';
      }).join('') + '</ul>';
    } else {
      h += '<div class="gw-home-none mt-2">잡힌 휴가가 없어요.</div>';
    }
    h += '<div class="gw-callout gw-callout-sm"><i class="fab fa-slack me-1" aria-hidden="true"></i>신청은 슬랙 <strong>휴가</strong> 앱에서 해요.</div>';
    box.innerHTML = h;
  }

  // ── 휴가자 ───────────────────────────────────────
  function renderOff(O) {
    var box = el('homeOff');
    if (!box) return;
    if (!O) { fail('homeOff', '휴가자를 불러오지 못했습니다.'); return; }
    function chips(rows, withDate) {
      return rows.map(function (x) {
        return '<span class="gw-home-chip">' + esc(x.name) +
          (withDate ? ' <small>' + esc(period(x)) + '</small>' : (x.half ? ' <small>' + esc(x.half) + ' 반차</small>' : '')) + '</span>';
      }).join('');
    }
    box.innerHTML =
      '<div class="gw-home-label">오늘 <strong>' + O.today.length + '</strong></div>' +
      (O.today.length ? '<div class="gw-home-chips">' + chips(O.today, false) + '</div>' : '<div class="gw-home-none">오늘 휴가자가 없어요.</div>') +
      '<div class="gw-home-label mt-3">앞으로 7일</div>' +
      (O.soon.length ? '<div class="gw-home-chips">' + chips(O.soon, true) + '</div>' : '<div class="gw-home-none">잡힌 휴가가 없어요.</div>');
  }

  // ── 메일 ─────────────────────────────────────────
  function renderMail(M) {
    var box = el('homeMail');
    if (!box) return;
    if (M && M.inbox) el('homeMailInbox').href = M.inbox;
    if (!M || !M.ok) {
      box.innerHTML = '<div class="gw-empty">' + (M && M.reason === 'no_company_mail'
        ? '회사 메일 계정이 아니라 메일 칸이 없어요.' : '메일을 불러오지 못했어요. 「Gmail 열기」로 확인해 주세요.') + '</div>';
      return;
    }
    var h = '<div class="gw-home-label">안 읽은 메일 <strong>' + M.unread + '</strong>통</div>';
    if (!M.items.length) {
      h += '<div class="gw-home-none">받은편지함이 비어 있어요.</div>';
    } else {
      h += '<ul class="gw-home-list gw-home-mail">' + M.items.map(function (m) {
        return '<li' + (m.unread ? ' class="unread"' : '') + '><a href="' + esc(m.link) + '" target="_blank" rel="noopener">' +
          '<span class="gw-home-mail-top"><span class="gw-home-li-main">' + esc(m.from) + '</span><span class="gw-home-when">' + esc(m.when) + '</span></span>' +
          '<span class="gw-home-mail-subj">' + esc(m.subject) + '</span>' +
          '<span class="gw-home-li-sub">' + esc(m.snippet) + '</span></a></li>';
      }).join('') + '</ul>';
    }
    box.innerHTML = h;
  }

  // ── 지금 할 일 ────────────────────────────────────
  // [닫기] = 오늘 하루 숨김. 내용(sig)이 바뀌면(건수 변화 등) 다시 뜬다.
  var todayKey = now.getFullYear() + '-' + (now.getMonth() + 1) + '-' + now.getDate();
  function hiddenTodo(t) {
    try { return localStorage.getItem('gw.todo.hide.' + t.id) === t.sig + '|' + todayKey; } catch (e) { return false; }
  }
  function hideTodo(t) {
    try { localStorage.setItem('gw.todo.hide.' + t.id, t.sig + '|' + todayKey); } catch (e) { /* noop */ }
  }

  function renderTodos(list) {
    var box = el('homeTodos');
    if (!box) return;
    box.textContent = '';
    (list || []).filter(function (t) { return !hiddenTodo(t); }).forEach(function (t) {
      var card = document.createElement('div');
      card.className = 'gw-todo ' + (t.tone || '');
      card.innerHTML =
        '<span class="gw-todo-ic"><i class="' + esc(t.icon) + '" aria-hidden="true"></i></span>' +
        '<div class="gw-todo-main"><div class="gw-todo-title">' + esc(t.title) + '</div>' +
        (t.desc ? '<div class="gw-todo-desc">' + esc(t.desc) + '</div>' : '') + '</div>' +
        (t.href ? '<a class="gw-todo-act" href="' + esc(t.href) + '"' +
          (t.external ? ' target="_blank" rel="noopener"' : '') + '>' + esc(t.action) + '</a>' : '') +
        '<button type="button" class="gw-todo-close">닫기</button>';
      card.querySelector('.gw-todo-close').addEventListener('click', function () {
        hideTodo(t);
        card.remove();
      });
      box.appendChild(card);
    });
  }

  // ── 이번 달 영업: 숫자 칸 → 날짜별 막대 ──────────────────
  var METRICS = [
    { key: 'inquiry', label: '문의', goal: true },
    { key: 'visit', label: '방문 잡음', goal: true },
    { key: 'contract', label: '계약', goal: false }
  ];
  var salesData = null;
  var salesMetric = 'inquiry';

  function won(n) {
    if (!n) return '0원';
    if (n >= 1e8) return (Math.round(n / 1e7) / 10).toLocaleString('ko-KR') + '억 원';
    return Math.round(n / 1e4).toLocaleString('ko-KR') + '만 원';
  }

  function diff(a, b) {
    var d = a - b;
    if (!d) return '<span>같음</span>';
    return '<span class="' + (d > 0 ? 'up' : 'down') + '">' + (d > 0 ? '+' : '−') + Math.abs(d).toLocaleString('ko-KR') + '</span>';
  }

  function renderSalesTiles() {
    var S = salesData, box = el('homeSalesTiles');
    box.innerHTML = METRICS.map(function (m) {
      var v = S.this[m.key], p = S.prev[m.key];
      var sub = '지난달 같은 기간(' + esc(S.prev_label) + ') ' + p.toLocaleString('ko-KR') + '건 ' + diff(v, p);
      var extra = m.goal
        ? '평일 하루 ' + (S.workday_avg[m.key] == null ? '-' : S.workday_avg[m.key]) + '건 · 목표 ' + S.goals[m.key] + (m.key === 'inquiry' ? '~25건' : '건 이상')
        : '계약 금액 ' + won(S.this.amount);
      return '<button type="button" class="gw-sales-tile' + (m.key === salesMetric ? ' active' : '') + '" data-k="' + m.key + '" aria-pressed="' + (m.key === salesMetric) + '">' +
        '<span class="k">' + m.label + '</span><span class="v">' + v.toLocaleString('ko-KR') + '건</span>' +
        '<span class="s">' + extra + '</span><span class="s">' + sub + '</span></button>';
    }).join('');
  }

  function niceMax(v) {
    if (v <= 4) return 4;
    var unit = v <= 10 ? 2 : v <= 50 ? 5 : v <= 200 ? 10 : 50;
    return Math.ceil(v / unit) * unit;
  }

  function renderSalesChart() {
    var S = salesData, box = el('homeSalesChart');
    var m = METRICS.filter(function (x) { return x.key === salesMetric; })[0];
    var days = S.days, n = days.length;
    var goal = m.goal ? S.goals[m.key] : null;
    // 목표선이 맨 위에 붙지 않게 15% 여유. 그림은 칸의 실제 폭으로 그리고 높이는 고정(넓은 화면에서 커지지 않게)
    var maxV = niceMax(1.15 * Math.max(goal || 0, Math.max.apply(null, days.map(function (d) { return d[m.key]; }))));
    var W = Math.max(320, Math.round(box.clientWidth - 24)), H = 200, L = 30, R = 8, T = 12, B = 24;
    var pw = W - L - R, ph = H - T - B, slot = pw / n, bw = Math.max(2, slot - 2);
    function y(v) { return T + ph - (v / maxV) * ph; }
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="최근 ' + n + '일 날짜별 ' + m.label + ' 막대 그래프">';
    [0, Math.round(maxV / 2), maxV].forEach(function (t) {
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '" stroke="#eef0f3" stroke-width="1"/>' +
        '<text x="' + (L - 6) + '" y="' + (y(t) + 4) + '" text-anchor="end" font-size="11" fill="#868b94">' + t + '</text>';
    });
    days.forEach(function (d, i) {
      var v = d[m.key], x = L + i * slot + 1, top = y(v), h = T + ph - top;
      var fill = d.workday ? '#2f6fde' : '#b9cdf4';
      if (h > 0) {
        var r = Math.min(4, h, bw / 2);
        svg += '<path d="M' + x + ',' + (T + ph) + 'V' + (top + r) + 'Q' + x + ',' + top + ' ' + (x + r) + ',' + top +
          'H' + (x + bw - r) + 'Q' + (x + bw) + ',' + top + ' ' + (x + bw) + ',' + (top + r) + 'V' + (T + ph) + 'Z" fill="' + fill + '"/>';
      }
      var p = d.date.split('-'), dd = Number(p[2]);
      if (i === n - 1 || (i % 5 === 0 && i < n - 3)) {
        svg += '<text x="' + (x + bw / 2) + '" y="' + (H - 6) + '" text-anchor="middle" font-size="11" fill="#868b94">' + Number(p[1]) + '/' + dd + '</text>';
      }
      svg += '<rect class="hit" data-i="' + i + '" x="' + (L + i * slot) + '" y="' + T + '" width="' + slot + '" height="' + ph + '" fill="transparent"/>';
    });
    svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + (T + ph) + '" y2="' + (T + ph) + '" stroke="#d6d9de" stroke-width="1"/>';
    if (goal) {
      // 목표 글자는 막대에 겹치지 않게 그래프 밖(제목 줄)에 둔다
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(goal) + '" y2="' + y(goal) + '" stroke="#555d6d" stroke-width="1.5" stroke-dasharray="5 4"/>';
    }
    svg += '</svg>';
    var total = days.reduce(function (a, d) { return a + d[m.key]; }, 0);
    box.innerHTML = '<div class="gw-sales-chart-head"><span><strong>최근 ' + n + '일 날짜별 ' + m.label + '</strong>' +
      (goal ? ' <span class="gw-sales-goal">점선 = 평일 목표 ' + goal + '건</span>' : '') +
      '</span><span>합계 ' + total.toLocaleString('ko-KR') + '건</span></div>' +
      svg + '<div class="gw-sales-tip" id="homeSalesTip"></div>';
    var tip = el('homeSalesTip'), svgEl = box.querySelector('svg');
    box.querySelectorAll('rect.hit').forEach(function (r) {
      r.addEventListener('mousemove', function () {
        var d = days[Number(r.getAttribute('data-i'))];
        var line = md(d.date) + ' · ' + m.label + ' ' + d[m.key] + '건';
        if (m.key === 'contract' && d.amount) line += ' · ' + won(d.amount);
        if (!d.workday) line += ' (휴일)';
        tip.textContent = line;
        var br = svgEl.getBoundingClientRect(), bb = box.getBoundingClientRect(), rb = r.getBoundingClientRect();
        tip.style.left = (rb.left - bb.left + rb.width / 2) + 'px';
        tip.style.top = (br.top - bb.top + (y(d[m.key]) / H) * br.height - 6) + 'px';
        tip.style.display = 'block';
      });
      r.addEventListener('mouseleave', function () { tip.style.display = 'none'; });
    });
  }

  function renderSales() {
    renderSalesTiles();
    renderSalesChart();
    el('homeSalesNote').textContent = '문의 = 리드 시트 상담 시간 날짜(모든 경로) · 방문 잡음 = 그 문의 중 방문 예약 이후 상태 · ' +
      '계약 = 공사현황 공사 확정 날짜(취소 제외) · 흐린 막대 = 주말·공휴일 · 5분마다 새로 셈';
  }

  if (el('homeSalesTiles')) {
    el('homeSalesTiles').addEventListener('click', function (e) {
      var b = e.target.closest('.gw-sales-tile');
      if (!b || !salesData) return;
      salesMetric = b.getAttribute('data-k');
      renderSalesTiles();
      renderSalesChart();
    });
    api('/api/home/sales').then(function (j) { salesData = j.sales; renderSales(); })
      .catch(function (e) { fail('homeSalesTiles', e.message); });
    var resizeTimer = null;
    window.addEventListener('resize', function () {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(function () { if (salesData) renderSalesChart(); }, 150);
    });
  }

  api('/api/home').then(function (j) {
    renderTodos(j.todos);
    renderLeave(j.leave);
    if (!j.field) {
      renderApprovals(j.approvals);
      renderOff(j.off);
    }
  }).catch(function (e) {
    ['homeApprovals', 'homeLeave', 'homeOff'].forEach(function (id) { fail(id, e.message); });
  });

  if (el('homeMail')) {
    api('/api/home/mail').then(function (j) { renderMail(j.mail); })
      .catch(function () { renderMail(null); });
  }
})();
