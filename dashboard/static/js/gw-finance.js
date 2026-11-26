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
  var S = { accounts: [], rows: [], tab: 'tx', end: '' };

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
    ['fbAcct', 'fbDir', 'fbUn'].forEach(function (id) { $(id).classList.toggle('on', !!$(id).value); });
    $('fbRange').classList.toggle('on', $('fbRange').value !== '7');   // 기본(최근 7일)에서 바꿨을 때만 진하게
    var active = $('fbAcct').value || $('fbDir').value || $('fbUn').value || $('fbQ').value.trim() || $('fbRange').value !== '7';
    $('fbFilters').classList.toggle('has-active-filters', !!active);
  }

  // ── 불러오기 ──────────────────────────────────────────────────────
  var seq = 0;
  function load() {
    var my = ++seq;
    markChips();
    var p = new URLSearchParams({
      start: $('fbStart').value, end: $('fbEnd').value, account: $('fbAcct').value, direction: $('fbDir').value,
      unassigned: $('fbUn').value, q: $('fbQ').value.trim()
    });
    return api('/api/finance/bank?' + p.toString()).then(function (j) {
      if (my !== seq) return;
      S.accounts = j.accounts || [];
      S.rows = j.rows || [];
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
      '<div class="gw-stat"><div class="k">적요 없는 거래</div><div class="v">' + (tot.unassigned || 0) + '<small> 건</small></div></div>';
  }

  function renderAccounts() {
    var keep = $('fbAcct').value;
    $('fbAcct').innerHTML = '<option value="">계좌</option>' + S.accounts.map(function (a) {
      return '<option value="' + a.id + '">' + esc(acctName(a)) + '</option>';
    }).join('');
    $('fbAcct').value = keep;
    markChips();
    var row = function (a) {
      var bal = a.balance == null ? '<span class="gw-badge gw-badge-wait">잔액 넣기</span>' : won(a.balance);
      return '<tr tabindex="0" role="button" data-acct="' + a.id + '" aria-label="' + esc(acctName(a) + ' 잔액 맞추기') + '"' +
        (a.in_report ? '' : ' class="gw-fb-side"') + '>' +
        '<td><strong>' + esc(a.code) + '</strong> <span class="text-muted">' + esc(a.bank) + '</span></td>' +
        '<td>' + esc(a.entity) + '</td>' +
        '<td class="num"><strong>' + bal + '</strong></td>' +
        '<td class="text-muted">' + esc(when(a.balance_at)) + '</td>' +
        '<td class="num gw-fb-in">' + (a.day_in ? won(a.day_in) : '') + '</td>' +
        '<td class="num gw-fb-out">' + (a.day_out ? won(a.day_out) : '') + '</td>' +
        '<td><span class="gw-badge ' + (a.sms ? 'gw-fb-sms' : 'gw-fb-hand') + '">' + (a.sms ? '문자 자동' : '직접 입력') + '</span></td></tr>';
    };
    var rep = S.accounts.filter(function (a) { return a.in_report; });
    var side = S.accounts.filter(function (a) { return !a.in_report; });
    $('fbAccts').innerHTML = rep.map(row).join('') +
      (side.length ? '<tr class="gw-fb-group"><th colspan="7" scope="rowgroup">자금일보 표에 없는 계좌</th></tr>' + side.map(row).join('') : '');
  }

  function renderRows() {
    $('fbCntTx').textContent = S.rows.length;
    $('fbCount').textContent = S.tab === 'tx' ? S.rows.length + '건' : (S.end ? dayLabel(S.end) + ' 끝 기준' : '');
    if (!S.rows.length) {
      var filtered = $('fbUn').value || $('fbQ').value.trim() || $('fbAcct').value || $('fbDir').value;
      $('fbRows').innerHTML = '<tr><td colspan="8" class="gw-empty">' +
        ($('fbUn').value ? '이 기간에 적요 없는 거래가 없어요.' : filtered ? '조건에 맞는 거래가 없어요.' : '이 기간에 거래가 없어요.') + '</td></tr>';
      return;
    }
    $('fbRows').innerHTML = S.rows.map(function (r) {
      var memo = r.memo || '';
      var need = (r.direction === 'in' || r.direction === 'out') && !memo;
      var memoCell = need ? '<span class="gw-badge gw-badge-wait">적요 없음</span>' : esc(memo);
      if (r.project_code && memo.indexOf(r.project_code) === -1) memoCell += ' <span class="text-muted small">' + esc(r.project_code) + '</span>';
      return '<tr tabindex="0" role="button" data-id="' + r.id + '" aria-label="' + esc(when(r.at) + ' ' + (r.counterparty || '') + ' 고치기') + '">' +
        '<td class="text-nowrap">' + esc(when(r.at)) + (r.source === 'manual' ? ' <span class="text-muted small">직접</span>' : '') + '</td>' +
        '<td class="text-nowrap"><strong>' + esc(r.acct_code || '') + '</strong></td>' +
        '<td class="gw-fb-wrap">' + esc(r.counterparty || (r.direction === 'bal' ? '잔액 맞추기' : '')) + '</td>' +
        '<td class="num gw-fb-in">' + (r.direction === 'in' ? won(r.amount) : '') + '</td>' +
        '<td class="num gw-fb-out">' + (r.direction === 'out' ? won(r.amount) : '') + '</td>' +
        '<td class="num">' + won(r.balance) + '</td>' +
        '<td class="text-nowrap">' + esc(r.title || '') + '</td>' +
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
    $('fbDrawerTitle').textContent = row ? (sms ? '계정·적요 고치기' : '직접 넣은 거래 고치기') : (isBal ? '잔액 맞추기' : '직접 입력');
    var acctOpts = S.accounts.map(function (a) {
      return '<option value="' + a.id + '"' + (a.id === pick ? ' selected' : '') + '>' + esc(acctName(a)) + '</option>';
    }).join('');
    var h = '<div class="gw-vh-msg" id="fbDMsg" role="status" aria-live="polite"></div>';
    if (sms) {
      h += '<dl class="gw-fb-facts"><dt>일시</dt><dd>' + esc(when(row.at)) + '</dd><dt>계좌</dt><dd>' + esc(row.acct_code) + '</dd>' +
        '<dt>상대방</dt><dd>' + esc(row.counterparty || '') + '</dd><dt>' + (row.direction === 'in' ? '입금' : '출금') + '</dt><dd>' + won(row.amount) + '원</dd>' +
        '<dt>거래 후 잔액</dt><dd>' + (row.balance == null ? '-' : won(row.balance) + '원') + '</dd></dl>' +
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
          inp('balance', '거래 후 잔액(원) · 모르면 비움', row ? won(row.balance) : '', ' inputmode="numeric" data-num') +
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
    if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
    setTab(S.tab === 'tx' ? 'acct' : 'tx', true);
  });
  $('fbRange').addEventListener('change', function () { applyRange(this.value); load(); });
  ['fbStart', 'fbEnd'].forEach(function (id) {
    $(id).addEventListener('change', function () { $('fbRange').value = ''; load(); });
  });
  ['fbAcct', 'fbDir', 'fbUn'].forEach(function (id) { $(id).addEventListener('change', load); });
  var qTimer;
  $('fbQ').addEventListener('input', function () { clearTimeout(qTimer); qTimer = setTimeout(load, 300); });
  $('fbReset').addEventListener('click', function () {
    $('fbRange').value = '7'; applyRange('7');
    $('fbAcct').value = ''; $('fbDir').value = ''; $('fbUn').value = ''; $('fbQ').value = '';
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
      var n = ['fbAcct', 'fbDir', 'fbUn'].filter(function (id) { return !!$(id).value; }).length + ($('fbRange').value !== '7' ? 1 : 0);
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

  applyRange('7');
  load();
})();
