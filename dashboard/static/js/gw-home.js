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

  api('/api/home').then(function (j) {
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
