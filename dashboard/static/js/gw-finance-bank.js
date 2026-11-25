/* 그룹웨어 재무 > 입출금(11-26). 은행 문자로 쌓이는 장부(잔액 포함)를 보고, 문자 없는 계좌는 직접 넣는다.
   권한·입력 검사는 서버(services/bank_ledger.py)가 다시 한다 — 여기서 숨기는 칸은 편의일 뿐. */
(function () {
  'use strict';

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  var WD = ['일', '월', '화', '수', '목', '금', '토'];

  var page = document.getElementById('fbPage');
  if (!page) return;
  var $ = function (id) { return document.getElementById(id); };
  var flashBox = $('fbFlash'), form = $('fbForm'), listBody = $('fbList').tBodies[0];
  var acctTable = $('fbAccts'), acctSel = $('fbAcct'), unBtn = $('fbUn');
  var state = { accounts: [], rows: [], editing: null };

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function won(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('ko-KR'); }
  function iso(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function dayLabel(s) {   // '2026-11-26' → '11/26(목)'
    var d = new Date(s.slice(0, 10) + 'T00:00:00');
    return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ')';
  }
  function when(s) { return s ? dayLabel(s) + ' ' + s.slice(11, 16) : ''; }

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

  function flash(msg, kind) {
    flashBox.className = 'gw-lv-flash ' + (kind || 'ok');
    flashBox.textContent = msg;
    flashBox.hidden = !msg;
    if (msg) flashBox.focus();
  }

  // ── 기간 ──────────────────────────────────────────────────────────
  function setRange(key) {
    var t = new Date(), s = new Date(t), e = new Date(t);
    if (key === '1') { s.setDate(t.getDate() - 1); e = new Date(s); }
    else if (key === '7') s.setDate(t.getDate() - 6);
    else if (key === 'month') s = new Date(t.getFullYear(), t.getMonth(), 1);
    $('fbStart').value = iso(s);
    $('fbEnd').value = iso(e);
    pressRange(key);
  }
  function pressRange(key) {
    page.querySelectorAll('[data-range]').forEach(function (b) {
      b.setAttribute('aria-pressed', b.getAttribute('data-range') === key ? 'true' : 'false');
    });
  }

  // ── 불러오기 ──────────────────────────────────────────────────────
  var seq = 0;
  function load() {
    var my = ++seq;
    var p = new URLSearchParams({
      start: $('fbStart').value, end: $('fbEnd').value, account: acctSel.value, direction: $('fbDir').value,
      unassigned: unBtn.getAttribute('aria-pressed') === 'true' ? '1' : '', q: $('fbQ').value.trim()
    });
    return api('/api/finance/bank?' + p.toString()).then(function (j) {
      if (my !== seq) return;
      state.accounts = j.accounts || [];
      state.rows = j.rows || [];
      renderAccounts(j.end);
      renderList(j.total || {});
    }).catch(function (e) {
      if (my !== seq) return;
      listBody.innerHTML = '<tr><td colspan="9" class="gw-empty">' + esc(e.message) + '</td></tr>';
      flash(e.message, 'err');
    });
  }

  function acctName(a) { return a.code + ' ' + a.bank; }

  function acctRow(a) {
    var bal = a.balance == null
      ? '<button type="button" class="btn btn-link btn-sm gw-fb-setbal" data-bal-acct="' + a.id + '" aria-label="' + esc(acctName(a)) + ' 잔액 넣기">잔액 넣기</button>'
      : won(a.balance);
    return '<tr' + (a.in_report ? '' : ' class="gw-fb-side"') + '>' +
      '<th scope="row"><span class="gw-fb-code">' + esc(a.code) + '</span> <span class="gw-fb-bank">' + esc(a.bank) + '</span>' +
      ' <span class="gw-fb-ent" title="사업자">' + esc(a.entity) + '</span></th>' +
      '<td class="num gw-fb-bal">' + bal + '</td>' +
      '<td class="gw-fb-at">' + esc(when(a.balance_at)) + '</td>' +
      '<td class="num in">' + (a.day_in ? won(a.day_in) : '') + '</td>' +
      '<td class="num out">' + (a.day_out ? won(a.day_out) : '') + '</td>' +
      '<td><span class="gw-fb-src' + (a.sms ? ' sms' : '') + '">' + (a.sms ? '문자 자동' : '직접 입력') + '</span></td></tr>';
  }

  function renderAccounts(day) {
    var keep = acctSel.value;
    acctSel.innerHTML = '<option value="">모든 계좌</option>' + state.accounts.map(function (a) {
      return '<option value="' + a.id + '">' + esc(acctName(a)) + '</option>';
    }).join('');
    acctSel.value = keep;
    $('fbAcctDay').textContent = day ? dayLabel(day) + ' 기준' : '';
    var rep = state.accounts.filter(function (a) { return a.in_report; });
    var side = state.accounts.filter(function (a) { return !a.in_report; });
    var known = rep.filter(function (a) { return a.balance != null; });
    var total = known.reduce(function (s, a) { return s + Number(a.balance); }, 0);
    var head = '<thead><tr><th scope="col">계좌·사업자</th><th scope="col" class="num">잔액</th>' +
      '<th scope="col">마지막 잔액 시각</th><th scope="col" class="num">그날 입금</th><th scope="col" class="num">그날 출금</th>' +
      '<th scope="col">기록</th></tr></thead>';
    var body = '<tbody>' + rep.map(acctRow).join('') +
      '<tr class="gw-fb-total"><th scope="row">자금일보 계좌 합계</th><td class="num">' + won(total) + '</td>' +
      '<td colspan="4" class="gw-fb-at">' + (known.length < rep.length ? '잔액 넣은 계좌 ' + known.length + '/' + rep.length + '개만 더함' : '') + '</td></tr>' +
      '</tbody>';
    if (side.length) {
      body += '<tbody class="gw-fb-side-group"><tr class="gw-fb-group"><th scope="rowgroup" colspan="6">자금일보 표에 없는 계좌</th></tr>' +
        side.map(acctRow).join('') + '</tbody>';
    }
    acctTable.innerHTML = head + body;
  }

  function renderList(tot) {
    var n = tot.unassigned || 0;
    var unN = $('fbUnN');   // 재시작 전 옛 템플릿엔 없음(JS 는 바로 나감)
    if (unN) unN.textContent = n ? ' ' + n : '';
    $('fbSum').innerHTML = state.rows.length
      ? '<strong>' + state.rows.length + '건</strong><span class="gw-fb-sep" aria-hidden="true">·</span>입금 <strong class="in">' + won(tot['in'] || 0) +
        '원</strong><span class="gw-fb-sep" aria-hidden="true">·</span>출금 <strong class="out">' + won(tot.out || 0) + '원</strong>'
      : '';
    if (!state.rows.length) {
      var un = unBtn.getAttribute('aria-pressed') === 'true';
      var filtered = un || $('fbQ').value.trim() || acctSel.value || $('fbDir').value;
      listBody.innerHTML = '<tr><td colspan="9" class="gw-empty">' +
        (un ? '이 기간에 적요 없는 거래가 없어요.' : filtered ? '조건에 맞는 거래가 없어요.' : '이 기간에 거래가 없어요.') + '</td></tr>';
      return;
    }
    listBody.innerHTML = state.rows.map(function (r) {
      var memo = r.memo || '';
      var need = (r.direction === 'in' || r.direction === 'out') && !memo;
      var memoCell = need ? '<span class="gw-fb-need">적요 없음</span>' : esc(memo);
      if (r.project_code && memo.indexOf(r.project_code) === -1) memoCell += ' <span class="gw-fb-tag">' + esc(r.project_code) + '</span>';
      return '<tr data-id="' + r.id + '"' + (r.direction === 'bal' ? ' class="gw-fb-balrow"' : '') + '>' +
        '<td class="gw-fb-at">' + esc(when(r.at)) + (r.source === 'manual' ? ' <span class="gw-fb-tag">직접</span>' : '') + '</td>' +
        '<td class="gw-fb-code">' + esc(r.acct_code || '') + '</td>' +
        '<td class="gw-fb-who">' + esc(r.counterparty || (r.direction === 'bal' ? '잔액 맞추기' : '')) + '</td>' +
        '<td class="num in">' + (r.direction === 'in' ? won(r.amount) : '') + '</td>' +
        '<td class="num out">' + (r.direction === 'out' ? won(r.amount) : '') + '</td>' +
        '<td class="num">' + won(r.balance) + '</td>' +
        '<td>' + esc(r.title || '') + '</td>' +
        '<td class="gw-fb-memo">' + memoCell + '</td>' +
        '<td class="gw-fb-act"><button type="button" class="btn btn-sm btn-outline-secondary" data-edit="' + r.id + '"' +
        ' aria-label="' + esc(when(r.at) + ' ' + (r.counterparty || '') + ' 고치기') + '">고치기</button></td></tr>';
    }).join('');
  }

  // ── 입력 창(쉬는 날 추가 창과 같은 모양) ─────────────────────────
  function nowLocal() {
    var d = new Date();
    return iso(d) + 'T' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
  }
  function field(id, label, input, opt, wide) {
    return '<div class="gw-fb-field' + (wide ? ' wide' : '') + '"><label class="form-label" for="' + id + '">' + label +
      (opt ? ' <span class="gw-fb-opt">' + opt + '</span>' : '') + '</label>' + input + '</div>';
  }
  function text(id, name, val, attrs) {
    return '<input type="text" class="form-control" id="' + id + '" name="' + name + '" autocomplete="off" value="' + esc(val) + '"' + (attrs || '') + '>';
  }

  function openForm(mode, row, opener, acctId) {
    if (state.editing && state.editing.opener) state.editing.opener.setAttribute('aria-expanded', 'false');
    state.editing = { mode: mode, row: row || null, opener: opener || null };
    if (opener && opener.hasAttribute('aria-expanded')) opener.setAttribute('aria-expanded', 'true');
    var sms = row && row.source !== 'manual';
    var dir = row ? row.direction : (mode === 'bal' ? 'bal' : 'in');
    var isBal = dir === 'bal';
    var title = row ? (sms ? '계정·적요 고치기' : '직접 넣은 거래 고치기') : (isBal ? '잔액 맞추기' : '직접 입력');
    var pick = row ? row.account_id : (acctId || Number(acctSel.value) || 0);
    var dis = sms ? ' disabled' : '';
    var opts = state.accounts.map(function (a) {
      return '<option value="' + a.id + '"' + (a.id === pick ? ' selected' : '') + '>' + esc(acctName(a)) + '</option>';
    }).join('');
    var h = '<h3 class="gw-fb-form-title" id="fbFormH">' + title + '</h3>';
    if (sms) h += '<p class="gw-fb-form-desc">은행 문자로 들어온 거래예요. 날짜·금액·잔액은 문자 그대로 두고 계정·적요·공사 코드만 고칠 수 있어요.</p>';
    else if (isBal) h += '<p class="gw-fb-form-desc">지금 통장에 찍힌 잔액을 넣으면, 이 시각부터 이 금액을 기준으로 잔액을 이어 가요.</p>';
    h += '<div class="gw-fb-fields">' +
      field('fbFAcct', '계좌', '<select class="form-select" id="fbFAcct" name="account_id" required' + dis + '><option value="">고르기</option>' + opts + '</select>') +
      field('fbFAt', '날짜·시각', '<input type="datetime-local" class="form-control" id="fbFAt" name="at" required value="' +
        esc(row ? row.at.replace(' ', 'T').slice(0, 16) : nowLocal()) + '"' + dis + '>');
    if (!isBal) {
      h += '<fieldset class="gw-fb-field gw-fb-dir"' + dis + '><legend class="form-label">구분</legend><div class="btn-group" role="radiogroup">' +
        '<input type="radio" class="btn-check" name="direction" id="fbFIn" value="in"' + (dir === 'in' ? ' checked' : '') + '>' +
        '<label class="btn btn-outline-secondary" for="fbFIn">입금</label>' +
        '<input type="radio" class="btn-check" name="direction" id="fbFOut" value="out"' + (dir === 'out' ? ' checked' : '') + '>' +
        '<label class="btn btn-outline-secondary" for="fbFOut">출금</label></div></fieldset>' +
        field('fbFAmt', '금액(원)', text('fbFAmt', 'amount', row ? won(row.amount) : '', ' inputmode="numeric"' + dis)) +
        field('fbFBal', '거래 후 잔액(원)', text('fbFBal', 'balance', row ? won(row.balance) : '', ' inputmode="numeric"' + dis), '모르면 비워 두세요') +
        field('fbFWho', '상대방', text('fbFWho', 'counterparty', row ? row.counterparty : '', ' maxlength="60"' + dis));
    } else {
      h += '<input type="hidden" name="direction" value="bal">' +
        field('fbFBal', '지금 잔액(원)', text('fbFBal', 'balance', row ? won(row.balance) : '', ' inputmode="numeric" required' + dis));
    }
    if (!isBal) {   // 잔액 맞추기엔 계정·공사 코드가 없음
      h += field('fbFTitle', '계정', text('fbFTitle', 'title', row ? row.title : '', ' maxlength="30" placeholder="매출수금, 외상매입"'), '선택') +
        field('fbFCode', '공사 코드', text('fbFCode', 'project_code', row ? row.project_code : '', ' maxlength="60"'), '선택');
    }
    h += field('fbFMemo', isBal ? '메모' : '적요', text('fbFMemo', 'memo', row ? row.memo : '', ' maxlength="200"'), '선택', true) +
      '</div><div class="gw-lv-err" id="fbErr" role="alert"></div><div class="gw-fb-form-acts">' +
      (row && !sms ? '<button type="button" class="btn btn-outline-danger me-auto" id="fbDel">지우기</button>' : '') +
      '<button type="button" class="btn btn-light" id="fbCancel">닫기</button>' +
      '<button type="submit" class="btn btn-primary" id="fbSave">저장</button></div>';
    form.innerHTML = h;
    // 금액·잔액 칸은 오른쪽 맞춤(text() 의 class 를 덮어씀)
    form.querySelectorAll('#fbFAmt, #fbFBal').forEach(function (el) { el.classList.add('num'); });
    form.hidden = false;
    form.scrollIntoView({ block: 'nearest' });
    var first = form.querySelector('select:enabled, input:enabled:not([type=hidden]):not(.btn-check)');
    if (sms) first = $('fbFTitle');
    else if (acctId) first = $('fbFBal');   // 계좌를 골라 연 잔액 넣기는 바로 금액부터
    if (first) first.focus();
  }

  function closeForm() {
    var opener = state.editing && state.editing.opener;
    form.hidden = true;
    form.innerHTML = '';
    state.editing = null;
    if (opener && opener.hasAttribute('aria-expanded')) opener.setAttribute('aria-expanded', 'false');
    if (opener && document.body.contains(opener)) opener.focus();
  }

  function formData() {
    var d = {};
    Array.prototype.forEach.call(form.elements, function (el) {
      if (!el.name || el.matches(':disabled')) return;   // 잠긴 칸(묶음 fieldset 포함)은 안 보냄
      if (el.type === 'radio') { if (el.checked) d[el.name] = el.value; return; }
      d[el.name] = el.value.trim();
    });
    if (d.at) d.at = d.at.replace('T', ' ');
    ['amount', 'balance'].forEach(function (k) { if (k in d) d[k] = d[k].replace(/[^\d-]/g, ''); });
    return d;
  }

  var busy = false;
  form.addEventListener('submit', function (ev) {
    ev.preventDefault();
    if (busy || !state.editing) return;
    var row = state.editing.row;
    var data = formData();
    var err = $('fbErr'), btn = $('fbSave');
    busy = true;
    btn.disabled = true;
    err.textContent = '';
    (row ? api('/api/finance/bank/' + row.id, { method: 'POST', body: data })
      : api('/api/finance/bank', { method: 'POST', body: data })).then(function () {
      closeForm();
      flash('저장했어요.', 'ok');
      return load();
    }).catch(function (e) {
      if (document.body.contains(err)) err.textContent = e.message; else flash(e.message, 'err');
    }).then(function () {
      busy = false;
      if (document.body.contains(btn)) btn.disabled = false;
    });
  });

  form.addEventListener('click', function (ev) {
    if (ev.target.id === 'fbCancel') closeForm();
    if (ev.target.id === 'fbDel' && state.editing && state.editing.row) {
      if (!window.confirm('직접 넣은 이 거래를 지울까요?')) return;
      api('/api/finance/bank/' + state.editing.row.id + '/delete', { method: 'POST', body: {} }).then(function () {
        closeForm();
        flash('지웠어요.', 'ok');
        return load();
      }).catch(function (e) { $('fbErr').textContent = e.message; });
    }
  });
  form.addEventListener('keydown', function (ev) { if (ev.key === 'Escape') closeForm(); });

  // 금액 칸: 쓰는 동안 천 단위 쉼표
  form.addEventListener('input', function (ev) {
    var el = ev.target;
    if (el.tagName !== 'INPUT' || !el.classList.contains('num')) return;
    var neg = /^-/.test(el.value), digits = el.value.replace(/\D/g, '');
    el.value = (neg ? '-' : '') + (digits ? Number(digits).toLocaleString('ko-KR') : '');
  });

  // ── 이벤트 ────────────────────────────────────────────────────────
  $('fbAdd').addEventListener('click', function () { openForm('add', null, this); });
  $('fbBal').addEventListener('click', function () { openForm('bal', null, this); });
  acctTable.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-bal-acct]');
    if (b) openForm('bal', null, b, Number(b.getAttribute('data-bal-acct')));
  });
  listBody.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-edit]');
    if (!b) return;
    var id = Number(b.getAttribute('data-edit'));
    var row = state.rows.filter(function (r) { return r.id === id; })[0];
    if (row) openForm('edit', row, b);
  });
  page.querySelectorAll('[data-range]').forEach(function (b) {
    b.addEventListener('click', function () { setRange(b.getAttribute('data-range')); load(); });
  });
  ['fbStart', 'fbEnd'].forEach(function (id) {
    $(id).addEventListener('change', function () { pressRange(''); load(); });
  });
  ['fbAcct', 'fbDir'].forEach(function (id) { $(id).addEventListener('change', load); });
  unBtn.addEventListener('click', function () {
    unBtn.setAttribute('aria-pressed', unBtn.getAttribute('aria-pressed') === 'true' ? 'false' : 'true');
    load();
  });
  var qTimer;
  $('fbQ').addEventListener('input', function () { clearTimeout(qTimer); qTimer = setTimeout(load, 300); });

  setRange('7');
  load();
})();
