/* 그룹웨어 홈 — 결재·내 휴가·휴가자·메일. 각 칸은 따로 불러오고, 하나가 실패해도 나머지는 뜬다. */
(function () {
  'use strict';

  // 로그인이 끝나 로그인 화면(200 HTML)으로 돌려보내졌거나 401 이면 멈춘다 — 예전엔 그 화면을 {} 로 읽어
  // 처리가 '된' 것처럼 보였음(11-24 조용한 실패 점검 4번, 결재함 gw-approval.js 와 같은 판정)
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function jsonOrEnd(r) {
    var ct = r.headers.get('content-type') || '';
    if (r.status === 401 || r.redirected || (r.ok && ct.indexOf('application/json') === -1)) {
      return Promise.reject(new Error(SESSION_ENDED));
    }
    return r.json().catch(function () { return {}; });
  }

  var WD = ['일', '월', '화', '수', '목', '금', '토'];

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // 구글 메일 미리보기(snippet)는 &#39; 처럼 이미 HTML 기호로 와서 esc 하면 기호가 그대로 보였다(11-15 점검) → 먼저 글자로 되돌림
  function unent(v) {
    var t = document.createElement('textarea');
    t.innerHTML = String(v == null ? '' : v);
    return t.value;
  }

  function api(url) {
    return fetch(url, { credentials: 'same-origin' }).then(function (r) {
      return jsonOrEnd(r).then(function (j) {
        if (!r.ok || j.success === false) throw new Error(j.message || ('불러오지 못했습니다 (' + r.status + ')'));
        return j;
      });
    });
  }

  // 홈 칸 먼저 보여 주기(11-28 부장님 점검 B3) — 결재·메일은 슬랙·지메일을 그때그때 읽어 3~8초 '불러오는 중'이었음.
  // 이 탭에서 10분 안에 본 결과가 있으면 바로 그리고, 새 결과가 오면 바꿔 그림. 같은 PC 를 다른 사람이 쓰더라도
  // 남지 않게 탭을 닫으면 지워지는 sessionStorage + 사용자 이메일로 나눠 둠. 새로 받기 실패 시엔 먼저 그린 것 유지
  var HOME_CACHE_MS = 10 * 60 * 1000;
  // 누구 자료인지 템플릿이 적어 준 이메일로 나눔. 재시작 전 옛 템플릿처럼 이메일이 없으면 먼저 보여 주기는 쓰지 않음
  var ME = (document.querySelector('.gw-home[data-me]') || { getAttribute: function () { return ''; } }).getAttribute('data-me') || '';
  function cacheKey(url) { return 'gw.home.' + ME + url; }
  function cacheGet(url) {
    if (!ME) return null;
    try {
      var v = JSON.parse(sessionStorage.getItem(cacheKey(url)) || 'null');
      return v && Date.now() - v.t < HOME_CACHE_MS ? v.j : null;
    } catch (e) { return null; }
  }
  function cacheSet(url, j) {
    if (!ME) return;
    try { sessionStorage.setItem(cacheKey(url), JSON.stringify({ t: Date.now(), j: j })); } catch (e) { /* 저장소가 차면 그냥 넘어감 */ }
  }
  function apiCached(url, draw) {
    var c = cacheGet(url);
    if (c) { try { draw(c); } catch (e) { /* 예전 모양 자료면 새 결과로 다시 그림 */ } }
    return api(url).then(function (j) { cacheSet(url, j); draw(j); return j; })
      .catch(function (e) { if (!c) throw e; });
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
      box.innerHTML = '<div class="gw-empty">' + (A.unavailable
        ? '슬랙 확인이 잠시 안 돼요. 잠시 뒤 새로 고쳐 주세요.'
        : '인사카드에 슬랙 계정이 연결되지 않아 결재를 볼 수 없어요. 경영지원에 연결을 요청해 주세요.') + '</div>';
      return;
    }
    function list(rows, empty) {
      if (!rows.length) return '<div class="gw-home-none">' + esc(empty) + '</div>';
      return '<ul class="gw-home-list">' + rows.map(function (r) {
        var inner = '<span class="gw-home-li-main">' + esc(r.title) + '</span>' +
          '<span class="gw-home-li-sub">' + esc([r.requester, r.what].filter(Boolean).join(' · ')) +
          (r.amount ? ' · ' + r.amount.toLocaleString('ko-KR') + '원' : '') + '</span>' +
          '<span class="gw-badge gw-badge-leave">' + esc(r.status) + '</span>';
        return '<li>' + (r.link ? '<a href="' + esc(r.link) + '">' + inner + '</a>' : '<div>' + inner + '</div>') + '</li>';
      }).join('') + '</ul>';
    }
    box.innerHTML =
      '<div class="gw-home-split">' +
      '<div><div class="gw-home-label">내가 처리할 것 <strong>' + A.todo_count + '</strong></div>' + list(A.todo, '처리할 결재가 없어요.') + '</div>' +
      '<div><div class="gw-home-label">내가 올린 진행 중 <strong>' + A.mine_count + '</strong></div>' + list(A.mine, '진행 중인 요청이 없어요.') + '</div>' +
      '</div>';
  }

  // ── 내 휴가 ──────────────────────────────────────
  function renderLeave(L, loadError) {
    var box = el('homeLeave');
    // 불러오기 실패를 '인사카드 없음'으로 안내하던 것(11-24 조용한 실패 점검)
    if (loadError) { fail('homeLeave', '휴가 정보를 불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'); return; }
    if (!L) { box.innerHTML = '<div class="gw-empty">내 계정에 연결된 인사카드가 없어 휴가를 볼 수 없어요. 경영지원에 연결을 요청해 주세요.</div>'; return; }
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
    h += '<div class="gw-callout gw-callout-sm"><i class="far fa-calendar-check me-1" aria-hidden="true"></i>휴가 신청·취소는 <a href="/leave">내 휴가</a> 화면에서 해요.</div>';
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
          '<span class="gw-home-li-sub">' + esc(unent(m.snippet)) + '</span></a></li>';
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
        '<button type="button" class="gw-todo-close" aria-label="' + esc(t.title) + ' 닫기">닫기</button>';
      card.querySelector('.gw-todo-close').addEventListener('click', function () {
        hideTodo(t);
        // 누른 단추가 사라지며 포커스가 페이지 맨 앞으로 튀던 것(11-16 점검) → 다음 카드 닫기, 없으면 이전, 없으면 화면 제목
        var near = card.nextElementSibling || card.previousElementSibling;
        card.remove();
        var to = near && near.querySelector('.gw-todo-close');
        if (!to) {
          to = document.querySelector('.gw-page-title');
          if (to && !to.hasAttribute('tabindex')) to.setAttribute('tabindex', '-1');
        }
        if (to) to.focus();
      });
      box.appendChild(card);
    });
  }

  // ── 이번 달 영업: 숫자 칸 → 날짜별 막대 ──────────────────
  var METRICS = [
    { key: 'inquiry', label: '문의', goal: true },
    { key: 'visit', label: '방문 예약', goal: true },
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

  // 비율(11-18 부장님): 방문 예약 = 문의 가운데 몇 %, 계약 = 문의·방문 예약 대비 몇 %. 같은 기간 건수끼리 나눈 값
  function pct(a, b) {
    return b ? Math.round((a / b) * 100) + '%' : '-';
  }

  function rateText(key, T) {
    // '문의의 78%'는 잘 안 읽힌다(11-18 부장님) → '문의량 대비'
    if (key === 'visit') return '문의량 대비 ' + pct(T.visit, T.inquiry);
    if (key === 'contract') return '문의량 대비 ' + pct(T.contract, T.inquiry) + ' · 방문 예약 대비 ' + pct(T.contract, T.visit);
    return '';
  }

  function renderSalesTiles() {
    var S = salesData, box = el('homeSalesTiles');
    box.innerHTML = METRICS.map(function (m) {
      var v = S.this[m.key], p = S.prev[m.key];
      var rate = rateText(m.key, S.this);
      var sub = '지난달 같은 기간(' + esc(S.prev_label) + ') ' + p.toLocaleString('ko-KR') + '건 ' + diff(v, p);
      var extra = m.goal
        ? '평일 하루 ' + (S.workday_avg[m.key] == null ? '-' : S.workday_avg[m.key]) + '건 · 목표 ' + S.goals[m.key] + (m.key === 'inquiry' ? '~25건' : '건 이상')
        : '계약 금액 ' + won(S.this.amount);
      return '<button type="button" class="gw-sales-tile' + (m.key === salesMetric ? ' active' : '') + '" data-k="' + m.key + '" aria-pressed="' + (m.key === salesMetric) + '">' +
        '<span class="k">' + m.label + '</span><span class="v">' + v.toLocaleString('ko-KR') + '건' +
        (rate ? '<span class="r">' + rate + '</span>' : '') + '</span>' +
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
    // tabindex: 키보드로 그래프에 와서 ←→ 로 날짜별 값을 봄(마우스를 올려야만 보이던 것)
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" tabindex="0" aria-label="최근 ' + n + '일 날짜별 ' + m.label + ' 막대 그래프. 좌우 화살표로 날짜별 값을 봅니다.">';
    [0, Math.round(maxV / 2), maxV].forEach(function (t) {
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '" stroke="#eef0f3" stroke-width="1"/>' +
        '<text x="' + (L - 6) + '" y="' + (y(t) + 4) + '" text-anchor="end" font-size="11" fill="#868b94" class="ax">' + t + '</text>';
    });
    days.forEach(function (d, i) {
      var v = d[m.key], x = L + i * slot + 1, top = y(v), h = T + ph - top;
      var fill = d.workday ? chartColor('--s-chart', '#2f6fde') : chartColor('--s-chart-light', '#b9cdf4');
      if (h > 0) {
        var r = Math.min(4, h, bw / 2);
        svg += '<path d="M' + x + ',' + (T + ph) + 'V' + (top + r) + 'Q' + x + ',' + top + ' ' + (x + r) + ',' + top +
          'H' + (x + bw - r) + 'Q' + (x + bw) + ',' + top + ' ' + (x + bw) + ',' + (top + r) + 'V' + (T + ph) + 'Z" fill="' + fill + '"/>';
      }
      var p = d.date.split('-'), dd = Number(p[2]);
      // 마지막 날짜와 겹치지 않게, 글자 폭(약 40) 안에 드는 앞 눈금은 건너뜀(휴대폰에서 '11/1111/15'로 붙던 것)
      if (i === n - 1 || (i % 5 === 0 && (n - 1 - i) * slot >= 40)) {
        // 마지막 날짜는 오른쪽 끝에 맞춰 — 가운데 정렬이면 휴대폰에서 '11/15'가 잘렸다(11-15 점검)
        var last = i === n - 1;
        svg += '<text x="' + (last ? x + bw : x + bw / 2) + '" y="' + (H - 6) + '" text-anchor="' + (last ? 'end' : 'middle') + '" font-size="11" fill="#6b7079">' + Number(p[1]) + '/' + dd + '</text>';
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
      svg + '<div class="gw-sales-tip" id="homeSalesTip" aria-live="polite"></div>';
    var tip = el('homeSalesTip'), svgEl = box.querySelector('svg');
    var hits = box.querySelectorAll('rect.hit'), cur = -1;
    function show(i) {
      var r = hits[i], d = days[i];
      if (!r || !d) return;
      cur = i;
      var line = md(d.date) + ' · ' + m.label + ' ' + d[m.key] + '건';
      if (m.key === 'contract' && d.amount) line += ' · ' + won(d.amount);
      if (!d.workday) line += ' (휴일)';
      tip.textContent = line;
      tip.style.display = 'block';   // 폭을 재려면 먼저 보이게
      var br = svgEl.getBoundingClientRect(), bb = box.getBoundingClientRect(), rb = r.getBoundingClientRect();
      // 말풍선은 가운데 기준(translate -50%)이라 첫날·마지막 날에선 반쯤 칸 밖으로 잘렸다(11-16 점검) → 칸 안으로 제한
      var half = tip.offsetWidth / 2, cx = rb.left - bb.left + rb.width / 2;
      var lo = half + 4, hi = box.clientWidth - half - 4;
      tip.style.left = (hi < lo ? box.clientWidth / 2 : Math.max(lo, Math.min(hi, cx))) + 'px';
      tip.style.top = (br.top - bb.top + (y(d[m.key]) / H) * br.height - 6) + 'px';
    }
    hits.forEach(function (r) {
      var i = Number(r.getAttribute('data-i'));
      r.addEventListener('mousemove', function () { show(i); });
      r.addEventListener('click', function () { show(i); });   // 휴대폰: 눌러서 보기
      r.addEventListener('mouseleave', function () { tip.style.display = 'none'; });
    });
    svgEl.addEventListener('keydown', function (e) {
      var k = e.key;
      if (k === 'ArrowRight' || k === 'ArrowLeft' || k === 'Home' || k === 'End') {
        e.preventDefault();
        var i = k === 'Home' ? 0 : k === 'End' ? n - 1 : cur < 0 ? n - 1 : cur + (k === 'ArrowRight' ? 1 : -1);
        show(Math.max(0, Math.min(n - 1, i)));
      } else if (k === 'Escape') { tip.style.display = 'none'; }
    });
    svgEl.addEventListener('focus', function () { if (cur < 0) show(n - 1); });
    svgEl.addEventListener('blur', function () { tip.style.display = 'none'; });
  }

  // 그래프 색 — 테마 색 값(--s-chart, 1안 주황)이 있으면 그걸, 없으면(옛 남색 ?theme=b 등) 예전 파랑
  function chartColor(name, fallback) {
    try { return getComputedStyle(document.body).getPropertyValue(name).trim() || fallback; } catch (e) { return fallback; }
  }
  function renderSales() {
    renderSalesTiles();
    renderSalesChart();
    // 집계 규칙은 services/gw_home.py _sales_counts 와 맞춤: 방문은 문의 행의 상태로 세서 문의 날짜에 들어간다
    // 설명 다섯 줄이 늘 깔려 있던 것 → 접힌 도움말(11-28 부장님, 토스 원칙 '한 화면 하나'). 문구는 그대로
    var note = el('homeSalesNote');
    if (note && !note.querySelector('details')) {
      note.innerHTML = '<details class="gw-help sm"><summary><i class="fas fa-circle-question" aria-hidden="true"></i> 숫자는 어떻게 세나요?' +
        '<i class="fas fa-chevron-down gw-help-fold" aria-hidden="true"></i></summary><div class="gw-help-body"><ul>' +
        '<li>문의는 리드 시트의 상담 날짜 기준이에요(모든 경로).</li>' +
        '<li>방문 예약은 그 문의 가운데 방문 예약까지 간 건이고, 문의한 날에 들어가요.</li>' +
        '<li>계약은 공사현황의 공사 확정 날짜 기준이며 취소된 공사는 빼요.</li>' +
        '<li>비율은 이번 달 건수끼리 나눈 값이라, 계약 비율은 이번 달 문의가 실제로 계약까지 간 비율과는 달라요.</li>' +
        '<li>흐린 막대는 주말·공휴일이고, 숫자는 5분마다 새로 계산해요.</li></ul></div></details>';
    }
  }

  if (el('homeSalesTiles')) {
    el('homeSalesTiles').addEventListener('click', function (e) {
      var b = e.target.closest('.gw-sales-tile');
      if (!b || !salesData) return;
      salesMetric = b.getAttribute('data-k');
      renderSalesTiles();
      renderSalesChart();
      // 칸을 다시 그려 누른 단추가 사라지며 포커스를 잃던 것 → 같은 칸으로
      var nb = el('homeSalesTiles').querySelector('[data-k="' + salesMetric + '"]');
      if (nb) nb.focus({ preventScroll: true });
    });
    api('/api/home/sales').then(function (j) { salesData = j.sales; renderSales(); })
      .catch(function (e) { fail('homeSalesTiles', e.message); });
    var resizeTimer = null;
    window.addEventListener('resize', function () {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(function () { if (salesData) renderSalesChart(); }, 150);
    });
  }

  apiCached('/api/home', function (j) {
    renderTodos(j.todos);
    renderLeave(j.leave, j.leave_error);
    if (!j.field) {
      renderApprovals(j.approvals);
      renderOff(j.off);
    }
  }).catch(function (e) {
    ['homeApprovals', 'homeLeave', 'homeOff'].forEach(function (id) { fail(id, e.message); });
  });

  // 공지(11-25) — 서버 재시작 전 화면엔 칸이 없을 수 있어 칸이 있을 때만
  if (el('homeNotice')) {
    api('/api/notices/unread').then(function (j) {
      var box = el('homeNotice');
      var n = el('homeNoticeN');
      if (n) { n.hidden = !j.notice; n.textContent = '안 읽음 ' + j.notice; }
      if (!j.recent || !j.recent.length) { box.innerHTML = '<div class="gw-empty">아직 공지가 없어요.</div>'; return; }
      box.innerHTML = '<ul class="gw-home-nt">' + j.recent.map(function (x) {
        var d = new Date(String(x.created_at).replace(' ', 'T'));
        return '<li class="' + (x.read ? '' : 'unread') + '"><a href="/notices/' + x.id + '">' +
          (x.pinned ? '<i class="fas fa-thumbtack" aria-label="고정"></i>' : '') +
          '<span class="gw-home-nt-t">' + esc(x.title) + '</span>' +
          '<span class="gw-home-nt-d">' + (d.getMonth() + 1) + '.' + d.getDate() + '</span></a></li>';
      }).join('') + '</ul>';
    }).catch(function (e) { fail('homeNotice', e.message); });
  }

  if (el('homeMail')) {
    apiCached('/api/home/mail', function (j) { renderMail(j.mail); })
      .catch(function () { renderMail(null); });
  }
})();
