/* 그룹웨어 입찰 공고 (2026-11-18) — 나라장터 공고 목록 + 검토 표시.
   서버에서 온 글자는 모두 esc() 로만 넣는다. 링크는 나라장터(g2b.go.kr) 주소만 그대로 쓴다. */
(function () {
  'use strict';

  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  var MARKS = [['', '표시 없음'], ['review', '검토 중'], ['join', '참여'], ['skip', '안 함']];
  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function safeUrl(u) {   // 나라장터 주소만 링크로
    return /^https:\/\/(www\.)?g2b\.go\.kr\//.test(String(u || '')) ? u : '';
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url, opts) {
    opts = opts || {};
    var method = opts.method || 'GET';
    var init = { method: method, credentials: 'same-origin' };
    if (opts.body) {
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify(opts.body);
    }
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
  function won(n) {
    n = Number(n) || 0;
    if (!n) return '-';
    if (n >= 100000000) {
      var eok = Math.floor(n / 100000000), man = Math.floor((n % 100000000) / 10000);
      return eok + '억' + (man ? ' ' + man.toLocaleString('ko-KR') + '만' : '') + ' 원';
    }
    return Math.floor(n / 10000).toLocaleString('ko-KR') + '만 원';
  }
  function when(s) {
    if (!s) return '-';
    var d = new Date(s.replace(' ', 'T'));
    if (isNaN(d)) return s;
    return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ') ' + s.slice(11, 16);
  }
  function soon(s) {   // 48시간 안에 마감
    var d = new Date(String(s || '').replace(' ', 'T'));
    return !isNaN(d) && d - new Date() < 48 * 3600 * 1000 && d > new Date();
  }


  var state = { tab: 'eligible', q: '', closed: false, seq: 0, items: {}, open: '', anim: false };

  function row(it) {
    var sel = '<select class="gw-chip-select gw-bd-mark" data-no="' + esc(it.bid_no) + '" aria-label="' + esc(it.title) + ' 표시">' +
      MARKS.map(function (m) {
        return '<option value="' + m[0] + '"' + (m[0] === it.mark ? ' selected' : '') + '>' + m[1] + '</option>';
      }).join('') + '</select>' +
      (it.mark ? '<div class="gw-bd-by">' + esc(it.mark_by) + ' · ' + esc(when(it.mark_at)) + '</div>' : '');
    var close = '<span class="' + (it.closed ? 'gw-bd-closed-t' : (soon(it.close_at) ? 'gw-bd-soon' : '')) + '">' +
      esc(when(it.close_at)) + '</span>' + (it.closed ? '<div class="gw-bd-sub">마감</div>' : '');
    var title = '<div class="gw-bd-title">' + (it.is_new ? '<span class="gw-badge gw-badge-new">New</span> ' : '') +
      esc(it.title) + '</div><div class="gw-bd-sub">' + esc(it.demand_org || it.org) +
      (it.contact ? ' · 담당 ' + esc(it.contact) : '') + '</div>';
    var bAt = (it.detail || {}).briefing_at;   // 현장설명회 — 놓치면 못 들어가는 공고가 있어 줄에서 바로 보이게
    if (bAt) {
      title += '<div class="gw-bd-brief ' + (past(bAt) ? 'gone' : (soon(bAt) ? 'soon' : '')) + '">' +
        '<i class="fas fa-person-chalkboard me-1" aria-hidden="true"></i>현장설명회 ' + esc(when(bAt)) +
        (past(bAt) ? ' (지남)' : '') + '</div>';
    }
    var cond = '<div>' + esc((it.regions || []).join(', ')) + '</div><div class="gw-bd-sub">' +
      esc((it.licenses || []).join(' 또는 ') || '면허 제한 없음') + '</div>' +
      (it.category !== 'eligible' && it.reason ? '<div class="gw-bd-miss">' + esc(it.reason) + '</div>' : '');
    var open = state.open === it.bid_no;
    return '<tr class="gw-bd-row' + (it.mark === 'skip' ? ' gw-bd-skip' : '') + (open ? ' gw-bd-open' : '') +
      '" data-no="' + esc(it.bid_no) + '" tabindex="0" aria-expanded="' + (open ? 'true' : 'false') +
      '" title="눌러서 자세히 보기">' +
      '<td class="gw-bd-close">' + close + '</td>' +
      '<td>' + title + '</td>' +
      '<td class="gw-bd-nw">' + esc(it.award_method || it.contract_method) + '</td>' +
      '<td class="num gw-bd-nw">' + esc(won(it.price)) + '</td>' +
      '<td>' + cond + '</td>' +
      '<td>' + sel + '</td></tr>';
  }

  // ── 펼친 상세 — 나라장터 공고 상세처럼 표로, 맨 위엔 우리가 판단할 것만 먼저 ──
  function full(s) { return s ? esc(when(s)) : ''; }
  function past(s) { var d = new Date(String(s || '').replace(' ', 'T')); return !isNaN(d) && d < new Date(); }
  function chk(ok, label, text) {   // ok: true=됨, false=안 됨, null=주의
    var cls = ok === true ? 'ok' : ok === false ? 'no' : 'warn';
    var ic = ok === true ? 'fa-circle-check' : ok === false ? 'fa-circle-xmark' : 'fa-circle-exclamation';
    return '<li class="gw-bd-chk ' + cls + '"><i class="fas ' + ic + '" aria-hidden="true"></i>' +
      '<span class="gw-bd-chk-l">' + esc(label) + '</span><span class="gw-bd-chk-t">' + text + '</span></li>';
  }
  function tr2(l1, v1, l2, v2) {   // 나라장터처럼 [항목|값|항목|값], 두 번째가 없으면 값이 끝까지
    return '<tr><th scope="row">' + esc(l1) + '</th><td' + (l2 ? '' : ' colspan="3"') + '>' + (v1 || '-') + '</td>' +
      (l2 ? '<th scope="row">' + esc(l2) + '</th><td>' + (v2 || '-') + '</td>' : '') + '</tr>';
  }
  function sect(title, rows) {
    return '<h3 class="gw-bd-sec">' + esc(title) + '</h3><table class="gw-bd-kv">' +
      '<colgroup><col class="k"><col><col class="k"><col></colgroup>' + rows + '</table>';
  }
  function detail(it) {
    var d = it.detail || {};
    var has = Object.keys(d).length > 0;
    var lic = (it.licenses || []).join(' 또는 ');
    var held = (state.held || []).join(', ');
    // 1) 확인할 것
    var rg = esc((it.regions || []).join(', '));
    var checks = it.region_ok ? chk(true, '참가 지역', rg + ' — 김포 업체 참가 가능')
      : it.nationwide ? chk(null, '참가 지역', '<b>지역 제한 없음(전국 공고)</b> — 경기도 한정 공고가 아니에요')
      : chk(false, '참가 지역', '<b>' + rg + '</b> — 김포 업체는 참가할 수 없어요');
    checks += it.license_ok
      ? chk(true, '면허', esc(lic || '제한 없음') + (lic ? ' — 보유' : ''))
      : chk(false, '면허', '<b>' + esc(it.missing || '우리 면허로 안 됨') + (it.missing ? ' 없음' : '') + '</b> — ' + esc(lic) +
          (lic.indexOf('+') > -1 ? ' 모두 있어야 해요' : ''));
    checks += it.hard_reason ? chk(false, '참가 조건', esc(it.hard_reason))
      : chk(true, '참가 조건', '실적제한·PQ 없음 · ' + esc(it.award_method || it.contract_method));
    if (d.briefing_at) {
      checks += chk(past(d.briefing_at) ? false : null, '현장설명회',
        '<b>' + full(d.briefing_at) + '</b>' + (d.briefing_place ? ' · ' + esc(d.briefing_place) : '') +
        (past(d.briefing_at) ? ' — 지났어요.' : '') + ' <span class="gw-bd-sub">참석이 필수인지 공고문에서 확인하세요</span>');
    }
    if (d.qual_at) {
      checks += chk(past(d.qual_at) ? false : (soon(d.qual_at) ? null : true), '참가 등록 마감',
        '<b>' + full(d.qual_at) + '</b>' + (past(d.qual_at) ? ' — 지났어요' : '') + ' <span class="gw-bd-sub">(나라장터 입찰참가자격 등록)</span>');
    }
    checks += chk(it.closed ? false : (soon(it.close_at) ? null : true), '입찰 마감',
      '<b>' + full(it.close_at) + '</b>' + (it.closed ? ' — 지났어요' : ''));
    // 2) 표
    var tel = d.contact_tel ? '<a href="tel:' + esc(d.contact_tel.replace(/[^0-9+]/g, '')) + '">' + esc(d.contact_tel) + '</a>' : '';
    var mail = d.contact_email && /^[^\s@<>"']+@[^\s@<>"']+$/.test(d.contact_email)
      ? '<a href="mailto:' + esc(d.contact_email) + '">' + esc(d.contact_email) + '</a>' : '';
    var files = (it.files || []).map(function (f) {
      var u = safeUrl(f.url);
      return u ? '<li><a href="' + esc(u) + '" target="_blank" rel="noopener">' + esc(f.name) + '</a></li>' : '';
    }).join('');
    var du = safeUrl(it.detail_url);
    var body =
      sect('일정',
        tr2('입찰 시작', full(d.begin_at), '입찰 마감', '<b>' + full(it.close_at) + '</b>') +
        tr2('참가 등록 마감', full(d.qual_at), '개찰', full(it.open_at)) +
        tr2('현장설명회', full(d.briefing_at), '설명회 장소', esc(d.briefing_place)) +
        tr2('공고 게시', full(it.posted_at), '개찰 장소', esc(d.open_place))) +
      sect('금액·낙찰',
        tr2('추정가격', '<b>' + esc(won(it.price)) + '</b>', '배정예산', d.budget ? esc(won(d.budget)) : '') +
        tr2('부가세', d.vat ? esc(won(d.vat)) : '', '낙찰하한율', esc(d.lower_rate)) +
        tr2('계약 방법', esc([it.contract_method, d.bid_method].filter(Boolean).join(' · ')),
            '예정가격', esc([d.price_method, d.price_count].filter(Boolean).join(' · '))) +
        tr2('낙찰 방법', esc(d.award_full || it.award_method))) +
      sect('참가 조건',
        tr2('참가가능지역', esc((it.regions || []).join(', '))) +
        tr2('면허', esc(lic || '제한 없음') + (held ? ' <span class="gw-bd-sub">(우리 면허: ' + esc(held) + ')</span>' : '')) +
        tr2('주공종', esc(d.main_license), '공사현장', esc(d.site_region || it.site_region)) +
        tr2('공동수급', esc(d.joint), '지역 공동도급', esc(d.joint_region))) +
      sect('공고·문의',
        tr2('공고번호', esc(it.bid_no + '-' + it.bid_ord) + (d.re_notice === '예' ? ' (재공고)' : ''), '참조번호', esc(d.ref_no)) +
        tr2('발주기관', esc(it.org), '수요기관', esc(it.demand_org)) +
        tr2('담당자', esc(d.contact_name || it.contact), '연락처', [tel, mail].filter(Boolean).join(' · '))) +
      '<h3 class="gw-bd-sec">첨부파일</h3>' +
      (files ? '<ul class="gw-bd-files">' + files + '</ul>' : '<p class="gw-bd-sub">첨부파일이 없어요.</p>');
    var foot = '<div class="gw-bd-acc-foot">처음 본 때 ' + esc(when(it.first_seen)) +
      (it.mark ? ' · ' + esc(it.mark_label) + ' 표시: ' + esc(it.mark_by) + ' ' + esc(when(it.mark_at)) : '') +
      (has ? '' : ' · 일정·낙찰하한율 같은 세부 항목은 다음 모으기 때 채워져요.') + '</div>';
    return '<tr class="gw-bd-detail"><td colspan="6"><div class="gw-bd-acc' + (state.anim ? ' gw-bd-slide' : '') + '">' +
      '<div class="gw-bd-acc-top"><ul class="gw-bd-chks" aria-label="확인할 것">' + checks + '</ul>' +
      (du ? '<a class="btn btn-sm btn-primary gw-bd-acc-go" href="' + esc(du) + '" target="_blank" rel="noopener">' +
        '<i class="fas fa-up-right-from-square me-1" aria-hidden="true"></i>나라장터에서 보기</a>' : '') + '</div>' +
      body + foot + '</div></td></tr>';
  }
  function closeDetail() {
    var rows = $('bdRows');
    Array.prototype.forEach.call(rows.querySelectorAll('tr.gw-bd-detail'), function (x) { x.remove(); });
    Array.prototype.forEach.call(rows.querySelectorAll('tr.gw-bd-open'), function (x) {
      x.classList.remove('gw-bd-open');
      x.setAttribute('aria-expanded', 'false');
    });
  }
  function openDetail(tr, animate) {
    var it = state.items[tr.getAttribute('data-no')];
    if (!it) return;
    closeDetail();
    state.open = it.bid_no;
    state.anim = animate;
    tr.classList.add('gw-bd-open');
    tr.setAttribute('aria-expanded', 'true');
    tr.insertAdjacentHTML('afterend', detail(it));
  }

  function load() {
    var my = ++state.seq;
    var qs = '?tab=' + encodeURIComponent(state.tab) + (state.closed ? '&closed=1' : '') +
      (state.q ? '&q=' + encodeURIComponent(state.q) : '');
    api('/api/bids' + qs).then(function (j) {
      if (my !== state.seq) return;
      $('bdCntEligible').textContent = j.counts.eligible;
      $('bdCntBlocked').textContent = j.counts.blocked;
      $('bdRule').textContent = '나라장터 냉난방 관련 공사 공고를 모았어요. 입찰 가능 기준: 참가가능지역에 경기도·김포시, ' +
        '지금 면허(' + (j.held || []).join(', ') + ')로 참가 가능, 실적제한·PQ·지명경쟁·4억 이상 적격심사 없음.' +
        (j.planned && j.planned.length ? ' 준비 중인 면허: ' + j.planned.join(', ') + '.' : '');
      state.held = j.held || [];
      var rows = j.items || [];
      state.items = {};
      rows.forEach(function (it) { state.items[it.bid_no] = it; });
      if (state.open && !state.items[state.open]) state.open = '';
      $('bdCount').textContent = rows.length + '건';
      $('bdRows').innerHTML = rows.length ? rows.map(row).join('') :
        '<tr><td colspan="6" class="gw-empty">' + (j.enabled ? (state.tab === 'eligible'
          ? '지금 들어갈 수 있는 마감 전 공고가 없어요.' : '입찰 불가능한 마감 전 공고가 없어요.')
          : '나라장터 연결 키가 서버에 없어요.') + '</td></tr>';
      if (state.open) {   // 표시를 바꿔 다시 그려도 펼친 줄은 그대로(애니메이션 없이)
        var tr = $('bdRows').querySelector('tr.gw-bd-row[data-no="' + CSS.escape(state.open) + '"]');
        if (tr) openDetail(tr, false);
      }
      $('bdFoot').textContent = j.last_collect ? '마지막으로 모은 때: ' + j.last_collect + ' (하루 세 번 07:20·12:20·16:20)' : '';
    }).catch(function (e) {
      if (my !== state.seq) return;
      $('bdRows').innerHTML = '<tr><td colspan="6" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  $('bdTabs').addEventListener('click', function (ev) {
    var b = ev.target.closest('button[data-tab]');
    if (!b) return;
    state.tab = b.getAttribute('data-tab');
    Array.prototype.forEach.call(this.querySelectorAll('button'), function (x) {
      var on = x === b;
      x.classList.toggle('active', on);
      x.setAttribute('aria-selected', on ? 'true' : 'false');
      x.tabIndex = on ? 0 : -1;
    });
    load();
  });
  var qt;
  $('bdQ').addEventListener('input', function () {
    clearTimeout(qt);
    var v = this.value.trim();
    qt = setTimeout(function () { state.q = v; load(); }, 250);
  });
  $('bdClosed').addEventListener('change', function () { state.closed = this.checked; load(); });
  $('bdRows').addEventListener('click', function (ev) {
    if (ev.target.closest('a, button, select, input, label, .gw-bd-detail')) return;   // 링크·표시 칸은 펼치기 제외
    var tr = ev.target.closest('tr.gw-bd-row');
    if (!tr) return;
    if (state.open === tr.getAttribute('data-no')) {
      state.open = '';
      closeDetail();
    } else {
      openDetail(tr, true);
    }
  });
  $('bdRows').addEventListener('keydown', function (ev) {   // 줄에 커서가 있을 때 Enter·Space 로 펼치기
    if ((ev.key !== 'Enter' && ev.key !== ' ') || !ev.target.classList || !ev.target.classList.contains('gw-bd-row')) return;
    ev.preventDefault();
    ev.target.click();
  });
  document.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape' || !state.open) return;
    var tr = $('bdRows').querySelector('tr.gw-bd-open');
    state.open = '';
    closeDetail();
    if (tr) tr.focus();
  });
  $('bdRows').addEventListener('change', function (ev) {
    var s = ev.target.closest('select.gw-bd-mark');
    if (!s) return;
    s.disabled = true;
    api('/api/bids/' + encodeURIComponent(s.getAttribute('data-no')) + '/mark', { method: 'POST', body: { mark: s.value } })
      .then(function () { load(); })
      .catch(function (e) { s.disabled = false; alert(e.message); load(); });
  });

  load();
})();
