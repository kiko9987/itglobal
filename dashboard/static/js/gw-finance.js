/* 그룹웨어 재무 > 입출금(11-26). 은행 문자로 쌓이는 장부(잔액 포함)를 보고, 문자 없는 계좌는 직접 넣는다.
   화면 틀은 결재함·운행일지·차량 관리와 같다: 요약 타일 → 탭 → 칩 거르기 → 표(줄 누르면 오른쪽 창).
   권한·입력 검사는 서버(services/bank_ledger.py)가 다시 한다 — 여기서 잠그는 칸은 편의일 뿐. */
(function () {
  'use strict';

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  var RANGE_LABEL = { '0': '오늘', '1': '어제', '7': '최근 7일', month: '이번 달', '': '고른 기간' };

  var page = document.getElementById('fbPage');
  if (!page || !document.getElementById('fbRows')) return;
  var $ = function (id) { return document.getElementById(id); };
  var S = { accounts: [], rows: [], tab: 'tx', end: '', noOut: [] };
  // 거르기 칩 — 결재 칩(fbAppr)은 11-26 추가라 재시작 전 옛 템플릿엔 없을 수 있음
  var CHIPS = ['fbAcct', 'fbDir', 'fbUn', 'fbAppr'].filter(function (id) { return document.getElementById(id); });

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function won(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('ko-KR'); }
  function iso(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function dayLabel(s) {
    var d = new Date(s.slice(0, 10) + 'T00:00:00');
    return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ')';
  }
  function when(s) { return s ? dayLabel(s) + ' ' + s.slice(11, 16) : ''; }
  function acctName(a) { return a.code + ' ' + a.bank; }
  // 줄 출처 꼬리표(일시 칸 옆) · 계좌 기록 방식
  var SRC_TAG = { manual: '직접', report: '자금일보' };
  function srcBadge(a) {
    if (a.sms) return '<span class="gw-badge gw-fb-sms">문자 자동</span>';
    if (a.in_report) return '<span class="gw-badge gw-fb-report" title="SB 자금일보 계좌 표에서 아침마다 옮겨 와요">자금일보 잔액</span>';
    return '<span class="gw-badge gw-fb-hand">직접 입력</span>';
  }

  function api(url, body) {
    var init = { method: body ? 'POST' : 'GET', credentials: 'same-origin' };
    if (body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(body); }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || (body
            ? '처리하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'
            : '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }
  function msg(text, kind) { var m = $('fbMsg'); m.className = 'gw-vh-msg ' + (kind || 'ok'); m.textContent = text || ''; }

  // ── 기간 ──────────────────────────────────────────────────────────
  function applyRange(key) {
    if (key === '') return;                         // 기간 직접 — 날짜 칸 그대로
    var t = new Date(), s = new Date(t), e = new Date(t);
    if (key === '1') { s.setDate(t.getDate() - 1); e = new Date(s); }
    else if (key === '7') s.setDate(t.getDate() - 6);
    else if (key === 'month') s = new Date(t.getFullYear(), t.getMonth(), 1);
    $('fbStart').value = iso(s);
    $('fbEnd').value = iso(e);
  }
  function markChips() {
    CHIPS.forEach(function (id) { $(id).classList.toggle('on', !!$(id).value); });
    $('fbRange').classList.toggle('on', $('fbRange').value !== '7');   // 기본(최근 7일)에서 바꿨을 때만 진하게
    var active = $('fbAcct').value || $('fbDir').value || $('fbUn').value || $('fbQ').value.trim() || $('fbRange').value !== '7';
    $('fbFilters').classList.toggle('has-active-filters', !!active);
  }

  // 탭·거르기를 주소에 남김(웹 가이드라인 'URL reflects state', 11-26 점검)
  function saveUrl() {
    var q = new URLSearchParams();
    if (S.tab !== 'tx') q.set('tab', S.tab);
    var r = $('fbRange').value;
    if (r !== '7') q.set('range', r);
    if (r === '') { q.set('start', $('fbStart').value); q.set('end', $('fbEnd').value); }
    if ($('fbAcct').value) q.set('acct', $('fbAcct').value);
    if ($('fbDir').value) q.set('dir', $('fbDir').value);
    if ($('fbUn').value) q.set('un', $('fbUn').value);
    if ($('fbAppr') && $('fbAppr').value) q.set('appr', $('fbAppr').value);
    if ($('fbQ').value.trim()) q.set('q', $('fbQ').value.trim());
    var qs = q.toString();
    try { history.replaceState(null, '', location.pathname + (qs ? '?' + qs : '')); } catch (e) { /* 미리보기 등 */ }
  }

  // ── 불러오기 ──────────────────────────────────────────────────────
  var seq = 0;
  function load() {
    var my = ++seq;
    markChips();
    saveUrl();
    var p = new URLSearchParams({
      start: $('fbStart').value, end: $('fbEnd').value, account: $('fbAcct').value, direction: $('fbDir').value,
      unassigned: $('fbUn').value, appr: $('fbAppr') ? $('fbAppr').value : '', q: $('fbQ').value.trim()
    });
    return api('/api/finance/bank?' + p.toString()).then(function (j) {
      if (my !== seq) return;
      S.accounts = j.accounts || [];
      S.rows = j.rows || [];
      S.noOut = j.no_out || [];
      S.end = j.end;
      renderStats(j.total || {});
      renderAccounts();
      renderRows();
    }).catch(function (e) {
      if (my !== seq) return;
      $('fbRows').innerHTML = '<tr><td colspan="8" class="gw-empty">' + esc(e.message) + '</td></tr>';
      msg(e.message, 'err');
    });
  }

  function renderStats(tot) {
    var rep = S.accounts.filter(function (a) { return a.in_report; });
    var known = rep.filter(function (a) { return a.balance != null; });
    var total = known.reduce(function (s, a) { return s + Number(a.balance); }, 0);
    var label = RANGE_LABEL[$('fbRange').value] || '고른 기간';
    $('fbStats').innerHTML =
      '<div class="gw-stat gw-stat-main"><div class="k">자금일보 잔액 · ' + esc(dayLabel(S.end)) + '</div><div class="v">' + won(total) + '<small> 원</small></div>' +
        (known.length < rep.length ? '<div class="gw-fb-stat-note">잔액 넣은 계좌 ' + known.length + '/' + rep.length + '개만 더했어요</div>' : '') + '</div>' +
      '<div class="gw-stat"><div class="k">' + esc(label) + ' 입금</div><div class="v">' + won(tot['in'] || 0) + '<small> 원</small></div></div>' +
      '<div class="gw-stat"><div class="k">' + esc(label) + ' 출금</div><div class="v">' + won(tot.out || 0) + '<small> 원</small></div></div>' +
      '<div class="gw-stat"><div class="k">적요 없는 거래</div><div class="v">' + (tot.unassigned || 0) + '<small> 건</small></div></div>' +
      '<div class="gw-stat' + (tot.no_out ? ' gw-fb-stat-warn' : '') + '"><div class="k">출금 문자 없는 이체 완료</div><div class="v">' + (tot.no_out || 0) + '<small> 건</small></div>' +
        '<div class="gw-fb-stat-sub">결재 없는 출금 ' + (tot.no_approval || 0) + '건</div></div>';
  }

  function renderAccounts() {
    var keep = $('fbAcct').value;
    $('fbAcct').innerHTML = '<option value="">계좌</option>' + S.accounts.map(function (a) {
      return '<option value="' + a.id + '">' + esc(acctName(a)) + '</option>';
    }).join('');
    $('fbAcct').value = keep;
    markChips();
    // 입금·출금 칸은 기간 끝 날 하루치 — 머리 칸에 그 날을 적음('그날'은 어느 날인지 몰라서, 11-26 부장님)
    var heads = document.querySelectorAll('#fbPaneAcct thead th');
    if (heads.length >= 6 && S.end) {
      var dl = S.end === iso(new Date()) ? '오늘' : dayLabel(S.end);
      heads[4].textContent = dl + ' 입금';
      heads[5].textContent = dl + ' 출금';
    }
    var row = function (a) {
      var bal = a.balance == null ? '<span class="gw-badge gw-badge-wait">잔액 넣기</span>' : won(a.balance);
      return '<tr tabindex="0" role="button" data-acct="' + a.id + '" aria-label="' + esc(acctName(a) + ' 잔액 맞추기') + '"' +
        (a.in_report ? '' : ' class="gw-fb-side"') + '>' +
        '<td><strong>' + esc(a.code) + '</strong> <span class="text-muted">' + esc(a.bank) + '</span></td>' +
        '<td class="c">' + esc(a.entity) + '</td>' +
        '<td class="num"><strong>' + bal + '</strong></td>' +
        '<td class="text-muted">' + esc(when(a.balance_at)) + '</td>' +
        '<td class="num gw-fb-in">' + (a.day_in ? won(a.day_in) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="num gw-fb-out">' + (a.day_out ? won(a.day_out) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="c">' + srcBadge(a) + '</td></tr>';
    };
    var rep = S.accounts.filter(function (a) { return a.in_report; });
    var side = S.accounts.filter(function (a) { return !a.in_report; });
    $('fbAccts').innerHTML = rep.map(row).join('') +
      (side.length ? '<tr class="gw-fb-group"><th colspan="7" scope="rowgroup">자금일보 표에 없는 계좌</th></tr>' + side.map(row).join('') : '');
  }

  function renderNoOut() {
    var box = $('fbNoOut');
    if (!box) return;
    box.hidden = S.tab !== 'tx' || !S.noOut.length;
    if (box.hidden) { box.innerHTML = ''; return; }
    box.innerHTML = '<p class="gw-fb-noout-h"><i class="fas fa-triangle-exclamation me-1" aria-hidden="true"></i>' +
      '전자결재에서 이체 완료를 눌렀는데 같은 금액의 출금 문자가 없는 결재 ' + S.noOut.length + '건이에요. ' +
      '문자가 안 오는 통장(1012 등)에서 보냈다면 그대로 두셔도 돼요.</p><ul class="gw-fb-noout-list">' +
      S.noOut.map(function (a) {
        return '<li><a href="/approval?id=' + a.id + '" target="_blank" rel="noopener">결재 #' + a.id + '</a> ' +
          '<span class="text-muted">' + esc(when(a.done_at)) + '</span> ' + esc([a.form, a.partner, a.content].filter(Boolean).join(' · ')) +
          ' <strong>' + won(a.amount) + '원</strong>' + (a.grn ? ' <span class="text-muted small">(' + esc(a.grn) + ' 통장)</span>' : '') + '</li>';
      }).join('') + '</ul>';
  }

  function renderRows() {
    renderNoOut();
    $('fbCntTx').textContent = S.rows.length;
    $('fbCount').textContent = S.tab === 'tx' ? '' : (S.end ? dayLabel(S.end) + ' 끝 기준 잔액' : '');   // 거래 건수는 탭에 있음
    if (!S.rows.length) {
      var filtered = $('fbUn').value || $('fbQ').value.trim() || $('fbAcct').value || $('fbDir').value || ($('fbAppr') && $('fbAppr').value);
      $('fbRows').innerHTML = '<tr><td colspan="8" class="gw-empty">' +
        ($('fbUn').value ? '이 기간에 적요 없는 거래가 없어요.' : filtered ? '조건에 맞는 거래가 없어요.' : '이 기간에 거래가 없어요.') + '</td></tr>';
      return;
    }
    $('fbRows').innerHTML = S.rows.map(function (r) {
      var memo = r.memo || '';
      var need = (r.direction === 'in' || r.direction === 'out') && !memo;
      var memoCell = need ? '<span class="gw-badge gw-badge-wait">적요 없음</span>' : esc(memo);
      if (r.project_code && memo.indexOf(r.project_code) === -1) memoCell += ' <span class="text-muted small">' + esc(r.project_code) + '</span>';
      if (r.approval) memoCell += ' <a class="gw-fb-appr" href="/approval?id=' + r.approval.id + '" target="_blank" rel="noopener" title="결재함에서 열기">결재 #' + r.approval.id + '</a>';
      else if (r.no_approval) memoCell += ' <span class="gw-badge gw-fb-hand" title="같은 금액·날짜의 이체 완료 결재가 없어요">결재 없음</span>';
      return '<tr tabindex="0" role="button" data-id="' + r.id + '" aria-label="' + esc(when(r.at) + ' ' + (r.counterparty || '') + ' 고치기') + '">' +
        '<td class="text-nowrap">' + esc(when(r.at)) + (SRC_TAG[r.source] ? ' <span class="text-muted small">' + SRC_TAG[r.source] + '</span>' : '') + '</td>' +
        '<td class="text-nowrap c"><strong>' + esc(r.acct_code || '') + '</strong></td>' +
        '<td class="gw-fb-wrap">' + esc(r.counterparty || (r.direction === 'bal' ? '잔액 맞추기' : '')) + '</td>' +
        '<td class="num gw-fb-in">' + (r.direction === 'in' ? won(r.amount) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="num gw-fb-out">' + (r.direction === 'out' ? won(r.amount) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="num">' + won(r.balance) + '</td>' +
        '<td class="text-nowrap c">' + esc(r.title || '') + '</td>' +
        '<td class="gw-fb-wrap">' + memoCell + '</td></tr>';
    }).join('');
  }

  // ── 탭 ────────────────────────────────────────────────────────────
  function setTab(tab, focus) {
    S.tab = tab;
    $('fbTabs').querySelectorAll('[role=tab]').forEach(function (b) {
      var on = b.getAttribute('data-tab') === tab;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
      b.tabIndex = on ? 0 : -1;
      if (on && focus) b.focus();
    });
    $('fbPaneTx').hidden = tab !== 'tx';
    $('fbPaneAcct').hidden = tab !== 'acct';
    page.querySelector('.gw-fb-txonly').hidden = tab !== 'tx';
    page.querySelector('.gw-fb-rangeonly').hidden = tab !== 'tx';
    $('fbEnd').setAttribute('aria-label', tab === 'tx' ? '끝 날짜' : '기준 날짜');
    saveUrl();
    renderRows();
  }

  // ── 오른쪽 창(넣기·고치기) ─────────────────────────────────────────
  var box = $('fbDrawer'), opener = null, cur = null;
  function closeDrawer() {
    box.hidden = true;
    document.body.classList.remove('gw-drawer-open');
    cur = null;
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
    if (e.shiftKey && (c === first || c === $('fbDrawerTitle'))) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && c === last) { e.preventDefault(); first.focus(); }
  }
  function nowLocal() {
    var d = new Date();
    return iso(d) + 'T' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
  }
  function inp(name, label, val, attrs, wide) {
    return '<label' + (wide ? ' class="wide"' : '') + '>' + label +
      '<input class="form-control form-control-sm" autocomplete="off" name="' + name + '" value="' + esc(val == null ? '' : val) + '"' + (attrs || '') + '></label>';
  }

  // mode: 'add' | 'bal' | 'edit'
  function openDrawer(mode, row, acctId, from) {
    opener = from || null;
    cur = { mode: mode, row: row || null };
    var sms = row && row.source !== 'manual';
    var dir = row ? row.direction : (mode === 'bal' ? 'bal' : 'in');
    var isBal = dir === 'bal';
    var pick = row ? row.account_id : (acctId || Number($('fbAcct').value) || 0);
    var dis = sms ? ' disabled' : '';
    var fromReport = row && row.source === 'report';
    $('fbDrawerTitle').textContent = row ? (fromReport ? '자금일보 잔액' : sms ? '계정·적요 고치기' : '직접 넣은 거래 고치기') : (isBal ? '잔액 맞추기' : '직접 입력');
    var acctOpts = S.accounts.map(function (a) {
      return '<option value="' + a.id + '"' + (a.id === pick ? ' selected' : '') + '>' + esc(acctName(a)) + '</option>';
    }).join('');
    var h = '<div class="gw-vh-msg" id="fbDMsg" role="status" aria-live="polite"></div>';
    if (fromReport) {
      h += '<dl class="gw-fb-facts"><dt>날짜</dt><dd>' + esc(dayLabel(row.at)) + ' 끝</dd><dt>계좌</dt><dd>' + esc(row.acct_code) + '</dd>' +
        '<dt>잔액</dt><dd>' + won(row.balance) + '원</dd></dl>' +
        '<p class="gw-fb-help">SB 자금일보 계좌 표에서 아침마다 옮겨 오는 잔액이에요. 잔액이 틀리면 자금일보를 고쳐 주세요. 다음 날 아침에 다시 맞춰져요.</p>';
    } else if (sms) {
      h += '<dl class="gw-fb-facts"><dt>일시</dt><dd>' + esc(when(row.at)) + '</dd><dt>계좌</dt><dd>' + esc(row.acct_code) + '</dd>' +
        '<dt>상대방</dt><dd>' + esc(row.counterparty || '') + '</dd><dt>' + (row.direction === 'in' ? '입금' : '출금') + '</dt><dd>' + won(row.amount) + '원</dd>' +
        '<dt>거래 후 잔액</dt><dd>' + (row.balance == null ? '-' : won(row.balance) + '원') + '</dd>' +
        (row.approval ? '<dt>전자결재</dt><dd><a href="/approval?id=' + row.approval.id + '" target="_blank" rel="noopener">#' + row.approval.id + '</a> ' +
          esc([row.approval.form, row.approval.partner, row.approval.content].filter(Boolean).join(' · ')) +
          (row.approval.done_at ? ' <span class="text-muted small">이체 ' + esc(when(row.approval.done_at)) + '</span>' : '') + '</dd>'
          : (row.no_approval ? '<dt>전자결재</dt><dd class="text-muted">같은 금액·날짜의 이체 완료 결재가 없어요</dd>' : '')) +
        '</dl>' +
        '<p class="gw-fb-help">은행 문자로 들어온 거래라 위 내용은 고칠 수 없어요.</p>';
    } else if (isBal) {
      h += '<p class="gw-fb-help">지금 통장에 찍힌 잔액을 넣으면, 이 시각부터 이 금액을 기준으로 잔액을 이어 가요.</p>';
    }
    h += '<section class="gw-vm-sec"><form class="gw-vm-form" id="fbForm" novalidate>';
    if (!sms) {
      h += '<label>계좌<select class="form-select form-select-sm" name="account_id" required><option value="">고르기</option>' + acctOpts + '</select></label>' +
        '<label>날짜·시각<input type="datetime-local" class="form-control form-control-sm" name="at" required value="' +
          esc(row ? row.at.replace(' ', 'T').slice(0, 16) : nowLocal()) + '"></label>';
      if (!isBal) {
        h += '<label>구분<select class="form-select form-select-sm" name="direction">' +
            '<option value="in"' + (dir === 'in' ? ' selected' : '') + '>입금</option><option value="out"' + (dir === 'out' ? ' selected' : '') + '>출금</option></select></label>' +
          inp('amount', '금액(원)', row ? won(row.amount) : '', ' inputmode="numeric" data-num') +
          inp('balance', '거래 후 잔액(원) 모르면 비워 두세요', row ? won(row.balance) : '', ' inputmode="numeric" data-num') +
          inp('counterparty', '상대방', row ? row.counterparty : '', ' maxlength="60"');
      } else {
        h += '<input type="hidden" name="direction" value="bal">' +
          inp('balance', '지금 잔액(원)', row ? won(row.balance) : '', ' inputmode="numeric" data-num required');
      }
    }
    if (!isBal) {
      h += inp('title', '계정', row ? row.title : '', ' maxlength="30" placeholder="매출수금, 외상매입"') +
        inp('project_code', '공사 코드', row ? row.project_code : '', ' maxlength="60"');
    }
    h += inp('memo', isBal ? '메모' : '적요', row ? row.memo : '', ' maxlength="200"', true) +
      '<div class="gw-vm-actions">' +
      (row && !sms ? '<button type="button" class="btn btn-sm btn-outline-danger me-auto" data-del>지우기</button>' : '') +
      '<button type="submit" class="btn btn-sm btn-primary">저장</button></div></form></section>';
    $('fbDrawerBody').innerHTML = h;
    box.hidden = false;
    document.body.classList.add('gw-drawer-open');
    var first = sms ? box.querySelector('[name=title]') : (acctId ? box.querySelector('[name=balance]') : box.querySelector('[name=account_id]'));
    (first || $('fbDrawerTitle')).focus();
  }

  function formData(form) {
    var d = {};
    Array.prototype.forEach.call(form.elements, function (el) {
      if (!el.name || el.disabled) return;
      d[el.name] = el.value.trim();
    });
    if (d.at) d.at = d.at.replace('T', ' ');
    ['amount', 'balance'].forEach(function (k) { if (k in d) d[k] = d[k].replace(/[^\d-]/g, ''); });
    return d;
  }

  box.addEventListener('submit', function (ev) {
    ev.preventDefault();
    if (!cur) return;
    var form = ev.target, btn = form.querySelector('[type=submit]');
    btn.disabled = true;
    var url = cur.row ? '/api/finance/bank/' + cur.row.id : '/api/finance/bank';
    api(url, formData(form)).then(function () {
      closeDrawer();
      msg('저장했어요.', 'ok');
      return load();
    }).catch(function (e) {
      var m = $('fbDMsg'); if (m) { m.className = 'gw-vh-msg err'; m.textContent = e.message; }
      btn.disabled = false;
    });
  });
  box.addEventListener('click', function (ev) {
    if (ev.target.closest('[data-close]')) { closeDrawer(); return; }
    if (ev.target.closest('[data-del]') && cur && cur.row && window.confirm('직접 넣은 이 거래를 지울까요? 지우면 되돌릴 수 없어요.')) {
      api('/api/finance/bank/' + cur.row.id + '/delete', {}).then(function () {
        closeDrawer();
        msg('지웠어요.', 'ok');
        return load();
      }).catch(function (e) { var m = $('fbDMsg'); if (m) { m.className = 'gw-vh-msg err'; m.textContent = e.message; } });
    }
  });
  // 금액 칸: 쓰는 동안 천 단위 쉼표
  box.addEventListener('input', function (ev) {
    var el = ev.target;
    if (!el.hasAttribute('data-num')) return;
    var digits = el.value.replace(/\D/g, '');
    el.value = digits ? Number(digits).toLocaleString('ko-KR') : '';
  });

  // ── 이벤트 ────────────────────────────────────────────────────────
  $('fbAdd').addEventListener('click', function () { openDrawer('add', null, 0, this); });
  $('fbBal').addEventListener('click', function () { openDrawer('bal', null, 0, this); });
  function rowOpen(ev) {
    if (ev.target.closest('a')) return;            // 결재 #번호 링크는 그대로 열림
    var tr = ev.target.closest('tr[data-id], tr[data-acct]');
    if (!tr || (ev.type === 'keydown' && ev.key !== 'Enter' && ev.key !== ' ')) return;
    ev.preventDefault();
    if (tr.hasAttribute('data-acct')) { openDrawer('bal', null, Number(tr.getAttribute('data-acct')), tr); return; }
    var id = Number(tr.getAttribute('data-id'));
    var row = S.rows.filter(function (r) { return r.id === id; })[0];
    if (row) openDrawer('edit', row, 0, tr);
  }
  ['fbRows', 'fbAccts'].forEach(function (id) {
    $(id).addEventListener('click', rowOpen);
    $(id).addEventListener('keydown', rowOpen);
  });
  $('fbTabs').addEventListener('click', function (ev) {
    var b = ev.target.closest('[role=tab]');
    if (b) setTab(b.getAttribute('data-tab'));
  });
  $('fbTabs').addEventListener('keydown', function (ev) {
    var k = ev.key;
    if (k === 'Home') { ev.preventDefault(); setTab('tx', true); return; }
    if (k === 'End') { ev.preventDefault(); setTab('acct', true); return; }
    if (k !== 'ArrowRight' && k !== 'ArrowLeft') return;
    setTab(S.tab === 'tx' ? 'acct' : 'tx', true);
  });
  $('fbRange').addEventListener('change', function () { applyRange(this.value); load(); });
  ['fbStart', 'fbEnd'].forEach(function (id) {
    $(id).addEventListener('change', function () { $('fbRange').value = ''; load(); });
  });
  CHIPS.forEach(function (id) { $(id).addEventListener('change', load); });
  var qTimer;
  $('fbQ').addEventListener('input', function () { clearTimeout(qTimer); qTimer = setTimeout(load, 300); });
  $('fbReset').addEventListener('click', function () {
    $('fbRange').value = '7'; applyRange('7');
    CHIPS.forEach(function (id) { $(id).value = ''; }); $('fbQ').value = '';
    load();
  });
  document.addEventListener('keydown', function (ev) {
    if (ev.key === 'Escape' && !box.hidden) closeDrawer();
    trapTab(ev);
  });

  // 휴대폰: 결재함처럼 찾기 + [필터 n] 만 보이고 누르면 칩이 펼쳐짐(넓은 화면에선 CSS 가 단추를 숨김)
  (function () {
    var box = $('fbFilters');
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'gw-chip-more';
    btn.setAttribute('aria-expanded', 'false');
    btn.setAttribute('aria-controls', 'fbFilters');
    btn.innerHTML = '<i class="fas fa-sliders-h me-1" aria-hidden="true"></i>필터<b class="gw-chip-more-n" hidden></b>';
    box.insertBefore(btn, box.firstChild);
    var nEl = btn.querySelector('b');
    function sync() {
      var n = CHIPS.filter(function (id) { return !!$(id).value; }).length + ($('fbRange').value !== '7' ? 1 : 0);
      nEl.textContent = n;
      nEl.hidden = !n;
      btn.classList.toggle('on', n > 0);
    }
    btn.addEventListener('click', function () {
      var open = !box.classList.contains('open');
      box.classList.toggle('open', open);
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    box.addEventListener('change', sync);
    $('fbReset').addEventListener('click', function () { setTimeout(sync, 0); });
    sync();
  })();

  // 주소에 남은 탭·거르기로 열기(새로 고침·즐겨찾기 해도 그대로)
  (function () {
    var q = new URLSearchParams(location.search);
    var r = q.get('range');
    $('fbRange').value = r != null && $('fbRange').querySelector('option[value="' + r + '"]') ? r : '7';
    applyRange($('fbRange').value);
    if ($('fbRange').value === '') {
      if (/^\d{4}-\d{2}-\d{2}$/.test(q.get('start') || '')) $('fbStart').value = q.get('start');
      if (/^\d{4}-\d{2}-\d{2}$/.test(q.get('end') || '')) $('fbEnd').value = q.get('end');
      if (!$('fbStart').value) applyRange('7');
    }
    ['acct', 'dir', 'un', 'appr'].forEach(function (k) {
      var id = { acct: 'fbAcct', dir: 'fbDir', un: 'fbUn', appr: 'fbAppr' }[k], v = q.get(k) || '';
      if (!$(id)) return;
      if (k === 'acct' && v) { $(id).innerHTML += '<option value="' + esc(v) + '"></option>'; }   // 계좌 목록은 불러온 뒤 채움
      $(id).value = v;
    });
    $('fbQ').value = (q.get('q') || '').slice(0, 40);
    if (q.get('tab') === 'acct') setTab('acct');
  })();
  load();
})();
