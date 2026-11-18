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

  var state = { tab: 'eligible', q: '', closed: false, seq: 0 };

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
    return '<tr' + (it.mark === 'skip' ? ' class="gw-bd-skip"' : '') + '>' +
      '<td class="gw-bd-close">' + close + '</td>' +
      '<td>' + title + '</td>' +
      '<td class="gw-bd-nw">' + esc(it.award_method || it.contract_method) + '</td>' +
      '<td class="num gw-bd-nw">' + esc(won(it.price)) + '</td>' +
      '<td>' + cond + '</td>' +
      '<td class="gw-bd-links">' + (links.join('<br>') || '-') + '</td>' +
      '<td>' + sel + '</td></tr>';
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
      $('bdCount').textContent = rows.length + '건';
      $('bdRows').innerHTML = rows.length ? rows.map(row).join('') :
        '<tr><td colspan="7" class="gw-empty">' + (j.enabled ? (state.tab === 'eligible'
          ? '지금 들어갈 수 있는 마감 전 공고가 없어요.' : '면허만 없어서 못 들어가는 마감 전 공고가 없어요.')
          : '나라장터 연결 키가 서버에 없어요.') + '</td></tr>';
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
