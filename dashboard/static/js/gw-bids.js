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
  function fileLabel(name) {   // 공고문이 hwp·pdf 두 개로 오는 일이 많아 형식을 붙인다
    var ext = (String(name).match(/\.([A-Za-z0-9]{2,5})$/) || [])[1];
    var kind = /공고/.test(name) ? '공고문' : /내역/.test(name) ? '내역서' : /시방/.test(name) ? '시방서'
      : /도면/.test(name) ? '도면' : '';
    if (!kind) return name.length > 12 ? name.slice(0, 12) + '…' : name;
    return kind + (ext ? ' ' + ext.toUpperCase() : '');
  }

  var state = { tab: 'eligible', q: '', closed: false, seq: 0, items: {}, open: '', anim: false };

  function row(it) {
    var links = [];
    var du = safeUrl(it.detail_url);
    if (du) links.push('<a href="' + esc(du) + '" target="_blank" rel="noopener">나라장터에서 보기</a>');
    (it.files || []).slice(0, 4).forEach(function (f) {
      var u = safeUrl(f.url);
      if (u) links.push('<a href="' + esc(u) + '" target="_blank" rel="noopener" title="' + esc(f.name) + '">' + esc(fileLabel(f.name)) + '</a>');
    });
    var sel = '<select class="gw-chip-select gw-bd-mark" data-no="' + esc(it.bid_no) + '" aria-label="' + esc(it.title) + ' 표시">' +
      MARKS.map(function (m) {
        return '<option value="' + m[0] + '"' + (m[0] === it.mark ? ' selected' : '') + '>' + m[1] + '</option>';
      }).join('') + '</select>' +
      (it.mark ? '<div class="gw-bd-by">' + esc(it.mark_by) + ' · ' + esc(when(it.mark_at)) + '</div>' : '');
    var close = '<span class="' + (it.closed ? 'gw-bd-closed-t' : (soon(it.close_at) ? 'gw-bd-soon' : '')) + '">' +
      esc(when(it.close_at)) + '</span>' + (it.closed ? '<div class="gw-bd-sub">마감</div>' : '');
    var title = '<div class="gw-bd-title">' + (it.is_new ? '<span class="gw-badge gw-badge-new">새</span> ' : '') +
      esc(it.title) + '</div><div class="gw-bd-sub">' + esc(it.demand_org || it.org) +
      (it.contact ? ' · 담당 ' + esc(it.contact) : '') + '</div>';
    var cond = '<div>' + esc((it.regions || []).join(', ')) + '</div><div class="gw-bd-sub">' +
      esc((it.licenses || []).join(' 또는 ') || '면허 제한 없음') + '</div>' +
      (it.missing ? '<div class="gw-bd-miss">' + esc(it.missing) + ' 필요</div>' : '');
    var open = state.open === it.bid_no;
    return '<tr class="gw-bd-row' + (it.mark === 'skip' ? ' gw-bd-skip' : '') + (open ? ' gw-bd-open' : '') +
      '" data-no="' + esc(it.bid_no) + '" tabindex="0" aria-expanded="' + (open ? 'true' : 'false') +
      '" title="눌러서 자세히 보기">' +
      '<td class="gw-bd-close">' + close + '</td>' +
      '<td>' + title + '</td>' +
      '<td class="gw-bd-nw">' + esc(it.award_method || it.contract_method) + '</td>' +
      '<td class="num gw-bd-nw">' + esc(won(it.price)) + '</td>' +
      '<td>' + cond + '</td>' +
      '<td class="gw-bd-links">' + (links.join('<br>') || '-') + '</td>' +
      '<td>' + sel + '</td></tr>';
  }

  // ── 펼친 상세(PM 프로젝트 줄 아코디언과 같은 방식: 줄을 누르면 바로 아래에 한 개만) ──
  function item(label, html) {
    return '<div class="gw-bd-item"><small>' + esc(label) + '</small><div>' + (html || '-') + '</div></div>';
  }
  function card(icon, tone, title, body) {
    return '<section class="gw-bd-card"><h3 class="gw-bd-card-t ' + tone + '"><i class="' + icon + ' me-1" aria-hidden="true"></i>' +
      esc(title) + '</h3><div class="gw-bd-grid">' + body + '</div></section>';
  }
  function full(s) { return s ? esc(when(s)) : ''; }
  function detail(it) {
    var d = it.detail || {};
    var has = Object.keys(d).length > 0;
    var verdict = it.category === 'eligible'
      ? '<span class="gw-badge gw-badge-active">지금 면허로 참가 가능</span>'
      : '<span class="gw-badge gw-badge-leave">' + esc(it.missing || '면허') + ' 있어야 참가 가능</span>';
    var files = (it.files || []).map(function (f) {
      var u = safeUrl(f.url);
      return u ? '<li><a href="' + esc(u) + '" target="_blank" rel="noopener">' + esc(f.name) + '</a></li>' : '';
    }).join('');
    var tel = d.contact_tel ? '<a href="tel:' + esc(d.contact_tel.replace(/[^0-9+]/g, '')) + '">' + esc(d.contact_tel) + '</a>' : '';
    var mail = d.contact_email && /^[^\s@<>"']+@[^\s@<>"']+$/.test(d.contact_email)
      ? '<a href="mailto:' + esc(d.contact_email) + '">' + esc(d.contact_email) + '</a>' : '';
    var du = safeUrl(it.detail_url);
    var head =
      '<div class="gw-bd-acc-head">' +
        '<div class="gw-bd-acc-no"><i class="fas fa-tag me-1" aria-hidden="true"></i>' + esc(it.bid_no + '-' + it.bid_ord) +
          (d.ref_no ? ' · ' + esc(d.ref_no) : '') + (d.kind ? ' · ' + esc(d.kind) : '') +
          (d.re_notice === '예' ? ' · 재공고' : '') + '</div>' +
        '<div class="gw-bd-acc-title">' + esc(it.title) + '</div>' +
        '<div class="gw-bd-sub"><i class="fas fa-building me-1" aria-hidden="true"></i>' + esc(it.org) +
          (it.demand_org && it.demand_org !== it.org ? ' · 수요기관 ' + esc(it.demand_org) : '') +
          ((d.site_region || it.site_region) ? ' · <i class="fas fa-map-marker-alt mx-1" aria-hidden="true"></i>' + esc(d.site_region || it.site_region) : '') +
        '</div>' +
        (du ? '<a class="btn btn-sm btn-outline-primary gw-bd-acc-go" href="' + esc(du) + '" target="_blank" rel="noopener">' +
          '<i class="fas fa-up-right-from-square me-1" aria-hidden="true"></i>나라장터에서 보기</a>' : '') +
      '</div>';
    var cards =
      card('fas fa-calendar-alt', 't-blue', '일정',
        item('공고 게시', full(it.posted_at)) +
        item('입찰 시작', full(d.begin_at)) +
        item('참가 등록 마감', full(d.qual_at)) +
        item('입찰 마감', '<b class="' + (soon(it.close_at) ? 'gw-bd-soon' : '') + '">' + full(it.close_at) + '</b>') +
        item('개찰', full(it.open_at) + (d.open_place ? '<div class="gw-bd-sub">' + esc(d.open_place) + '</div>' : ''))) +
      card('fas fa-won-sign', 't-green', '금액·방식',
        item('추정가격', esc(won(it.price))) +
        item('배정예산', d.budget ? esc(won(d.budget)) : '') +
        item('부가세', d.vat ? esc(won(d.vat)) : '') +
        item('계약 방법', esc(it.contract_method) + (d.bid_method ? ' <span class="gw-bd-sub">' + esc(d.bid_method) + '</span>' : '')) +
        item('낙찰 방법', esc(d.award_full || it.award_method)) +
        item('낙찰하한율', esc(d.lower_rate)) +
        item('예정가격', esc([d.price_method, d.price_count].filter(Boolean).join(' · ')))) +
      card('fas fa-clipboard-check', 't-orange', '참가 조건',
        item('우리 판정', verdict) +
        item('참가가능지역', esc((it.regions || []).join(', '))) +
        item('면허', esc((it.licenses || []).join(' 또는 ') || '제한 없음')) +
        item('주공종', esc(d.main_license) + (d.main_price ? ' <span class="gw-bd-sub">' + esc(won(d.main_price)) + '</span>' : '')) +
        item('공동수급', esc(d.joint)) +
        item('지역 공동도급', esc(d.joint_region))) +
      card('fas fa-user', 't-purple', '담당자·첨부',
        item('담당자', esc(d.contact_name || it.contact)) +
        item('전화', tel) +
        item('메일', mail) +
        '<div class="gw-bd-item gw-bd-item-wide"><small>첨부파일</small>' +
          (files ? '<ul class="gw-bd-files">' + files + '</ul>' : '<div>-</div>') + '</div>');
    var foot = '<div class="gw-bd-acc-foot">처음 본 때 ' + esc(when(it.first_seen)) +
      (it.mark ? ' · ' + esc(it.mark_label) + ' 표시: ' + esc(it.mark_by) + ' ' + esc(when(it.mark_at)) : '') +
      (has ? '' : ' · 일정·낙찰하한율 같은 세부 항목은 다음 모으기 때 채워져요.') + '</div>';
    return '<tr class="gw-bd-detail"><td colspan="7"><div class="gw-bd-acc' + (state.anim ? ' gw-bd-slide' : '') + '">' +
      head + '<div class="gw-bd-cards">' + cards + '</div>' + foot + '</div></td></tr>';
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
      $('bdCntLicense').textContent = j.counts.license_only;
      $('bdRule').textContent = '나라장터 냉난방 관련 공사 공고 중 경기도·김포시 업체가 참가할 수 있는 것만 모았어요. ' +
        '지금 면허: ' + (j.held || []).join(', ') + (j.planned && j.planned.length ? ' · 준비 중: ' + j.planned.join(', ') : '') +
        '. 실적제한·PQ·지명경쟁·4억 이상 적격심사는 뺐어요.';
      var rows = j.items || [];
      state.items = {};
      rows.forEach(function (it) { state.items[it.bid_no] = it; });
      if (state.open && !state.items[state.open]) state.open = '';
      $('bdCount').textContent = rows.length + '건';
      $('bdRows').innerHTML = rows.length ? rows.map(row).join('') :
        '<tr><td colspan="7" class="gw-empty">' + (j.enabled ? (state.tab === 'eligible'
          ? '지금 들어갈 수 있는 마감 전 공고가 없어요.' : '면허만 없어서 못 들어가는 마감 전 공고가 없어요.')
          : '나라장터 연결 키가 서버에 없어요.') + '</td></tr>';
      if (state.open) {   // 표시를 바꿔 다시 그려도 펼친 줄은 그대로(애니메이션 없이)
        var tr = $('bdRows').querySelector('tr.gw-bd-row[data-no="' + CSS.escape(state.open) + '"]');
        if (tr) openDetail(tr, false);
      }
      $('bdFoot').textContent = j.last_collect ? '마지막으로 모은 때: ' + j.last_collect + ' (하루 세 번 07:20·12:20·16:20)' : '';
    }).catch(function (e) {
      if (my !== state.seq) return;
      $('bdRows').innerHTML = '<tr><td colspan="7" class="gw-empty">' + esc(e.message) + '</td></tr>';
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
