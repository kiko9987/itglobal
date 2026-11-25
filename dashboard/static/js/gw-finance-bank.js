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
  var acctTable = $('fbAccts'), acctSel = $('fbAcct');
  var state = { accounts: [], rows: [], editing: null };

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function won(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('ko-KR'); }
  function iso(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function when(s) {   // '2026-11-26 09:12' → '11/26(목) 09:12'
    if (!s) return '';
    var d = new Date(s.slice(0, 10) + 'T00:00:00');
    return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ') ' + s.slice(11, 16);
  }

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
      start: $('fbStart').value, end: $('fbEnd').value, account: acctSel.value,
      direction: $('fbDir').value, unassigned: $('fbUn').checked ? '1' : '', q: $('fbQ').value.trim()
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

  function renderAccounts(day) {
    var keep = acctSel.value;
    acctSel.innerHTML = '<option value="">모든 계좌</option>' + state.accounts.map(function (a) {
      return '<option value="' + a.id + '">' + esc(acctName(a)) + '</option>';
    }).join('');
    acctSel.value = keep;
    $('fbAcctDay').textContent = day ? '(' + when(day + ' 00:00').slice(0, -6) + ' 끝 기준)' : '';
    var head = '<thead><tr><th scope="col">계좌</th><th scope="col">사업자</th><th scope="col" class="num">잔액</th>' +
      '<th scope="col">잔액 시각</th><th scope="col" class="num">그날 입금</th><th scope="col" class="num">그날 출금</th>' +
      '<th scope="col">들어오는 길</th></tr></thead>';
    var total = 0, known = 0;
    var body = state.accounts.map(function (a) {
      if (a.balance != null && a.in_report) { total += Number(a.balance); known++; }
      return '<tr' + (a.in_report ? '' : ' class="gw-fb-side"') + '>' +
        '<th scope="row">' + esc(a.code) + ' <span class="gw-fb-bank">' + esc(a.bank) + '</span></th>' +
        '<td>' + esc(a.entity) + '</td>' +
        '<td class="num">' + (a.balance == null ? '<span class="gw-fb-unknown">모름</span>' : won(a.balance)) + '</td>' +
        '<td class="gw-fb-at">' + esc(when(a.balance_at)) + '</td>' +
        '<td class="num">' + (a.day_in ? won(a.day_in) : '') + '</td>' +
        '<td class="num">' + (a.day_out ? won(a.day_out) : '') + '</td>' +
        '<td class="gw-fb-src">' + (a.sms ? '은행 문자' : '직접 입력') + (a.in_report ? '' : ' · 자금일보 표 밖') + '</td></tr>';
    }).join('');
    var reportCnt = state.accounts.filter(function (a) { return a.in_report; }).length;
    var foot = '<tfoot><tr><th scope="row" colspan="2">자금일보 계좌 합계</th><td class="num">' + won(total) + '</td>' +
      '<td colspan="4" class="gw-fb-at">' + (known < reportCnt ? '잔액을 아는 ' + known + '/' + reportCnt + '개 계좌만 더함' : '') + '</td></tr></tfoot>';
    acctTable.innerHTML = head + '<tbody>' + body + '</tbody>' + foot;
  }

  function renderList(tot) {
    $('fbSum').textContent = state.rows.length
      ? state.rows.length + '건 · 입금 ' + won(tot['in'] || 0) + '원 · 출금 ' + won(tot.out || 0) + '원' +
        (tot.unassigned ? ' · 적요 없는 것 ' + tot.unassigned + '건' : '')
      : '';
    if (!state.rows.length) {
      listBody.innerHTML = '<tr><td colspan="9" class="gw-empty">이 기간에 거래가 없어요.</td></tr>';
      return;
    }
    listBody.innerHTML = state.rows.map(function (r) {
      var memo = r.memo || '';
      var need = (r.direction === 'in' || r.direction === 'out') && !memo;
      return '<tr data-id="' + r.id + '"' + (r.direction === 'bal' ? ' class="gw-fb-balrow"' : '') + '>' +
        '<td class="gw-fb-at">' + esc(when(r.at)) + (r.source === 'manual' ? ' <span class="gw-fb-tag">직접</span>' : '') + '</td>' +
        '<td>' + esc(r.acct_code || '') + '</td>' +
        '<td class="gw-fb-who">' + esc(r.counterparty || (r.direction === 'bal' ? '잔액 맞추기' : '')) + '</td>' +
        '<td class="num in">' + (r.direction === 'in' ? won(r.amount) : '') + '</td>' +
        '<td class="num out">' + (r.direction === 'out' ? won(r.amount) : '') + '</td>' +
        '<td class="num">' + won(r.balance) + '</td>' +
        '<td>' + esc(r.title || '') + '</td>' +
        '<td class="gw-fb-memo">' + (need ? '<span class="gw-fb-need">적요 없음</span>' : esc(memo)) + '</td>' +
        '<td class="gw-fb-act"><button type="button" class="btn btn-sm btn-outline-secondary" data-edit="' + r.id + '"' +
        ' aria-label="' + esc(when(r.at) + ' ' + (r.counterparty || '') + ' 고치기') + '">고치기</button></td></tr>';
    }).join('');
  }

  // ── 입력 창 ───────────────────────────────────────────────────────
  function nowLocal() { var d = new Date(); return iso(d) + 'T' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }

  function openForm(mode, row, opener) {
    state.editing = { mode: mode, row: row || null, opener: opener || null };
    var sms = row && row.source !== 'manual';
    var dir = row ? row.direction : (mode === 'bal' ? 'bal' : 'in');
    var title = row ? (sms ? '계정·적요 고치기' : '직접 넣은 거래 고치기') : (mode === 'bal' ? '잔액 맞추기' : '직접 입력');
    var opts = state.accounts.map(function (a) {
      var sel = row ? row.account_id === a.id : (Number(acctSel.value) === a.id);
      return '<option value="' + a.id + '"' + (sel ? ' selected' : '') + '>' + esc(acctName(a)) + '</option>';
    }).join('');
    var dis = sms ? ' disabled' : '';
    var h = '<h2 class="gw-panel-title" id="fbFormH">' + title + '</h2>';
    if (sms) h += '<p class="gw-note">은행 문자로 들어온 거래라 날짜·금액·잔액은 그대로 두고 계정·적요·공사 코드만 고칠 수 있어요.</p>';
    else if (mode === 'bal' || dir === 'bal') h += '<p class="gw-note">지금 통장 잔액을 넣으면 그 시각부터 이 금액을 기준으로 잔액을 보여 줘요.</p>';
    h += '<div class="gw-fb-grid">' +
      '<label>계좌<select class="form-select" name="account_id" required' + dis + '><option value="">고르기</option>' + opts + '</select></label>' +
      '<label>날짜·시각<input type="datetime-local" class="form-control" name="at" required value="' +
        esc(row ? row.at.replace(' ', 'T').slice(0, 16) : nowLocal()) + '"' + dis + '></label>';
    if (mode !== 'bal' && dir !== 'bal') {
      h += '<fieldset class="gw-fb-dir"' + dis + '><legend>구분</legend>' +
        '<label class="form-check form-check-inline"><input class="form-check-input" type="radio" name="direction" value="in"' + (dir === 'in' ? ' checked' : '') + '><span class="form-check-label">입금</span></label>' +
        '<label class="form-check form-check-inline"><input class="form-check-input" type="radio" name="direction" value="out"' + (dir === 'out' ? ' checked' : '') + '><span class="form-check-label">출금</span></label></fieldset>' +
        '<label>금액(원)<input type="text" inputmode="numeric" class="form-control num" name="amount" autocomplete="off" value="' + esc(row ? won(row.amount) : '') + '"' + dis + '></label>' +
        '<label>거래 후 잔액(원, 모르면 비움)<input type="text" inputmode="numeric" class="form-control num" name="balance" autocomplete="off" value="' + esc(row ? won(row.balance) : '') + '"' + dis + '></label>' +
        '<label>상대방<input type="text" class="form-control" name="counterparty" maxlength="60" autocomplete="off" value="' + esc(row ? row.counterparty : '') + '"' + dis + '></label>';
    } else {
      h += '<input type="hidden" name="direction" value="bal">' +
        '<label>지금 잔액(원)<input type="text" inputmode="numeric" class="form-control num" name="balance" required autocomplete="off" value="' + esc(row ? won(row.balance) : '') + '"' + dis + '></label>';
    }
    h += '<label>계정<input type="text" class="form-control" name="title" maxlength="30" autocomplete="off" placeholder="예: 매출수금, 외상매입" value="' + esc(row ? row.title : '') + '"></label>' +
      '<label class="gw-fb-wide">적요<input type="text" class="form-control" name="memo" maxlength="200" autocomplete="off" value="' + esc(row ? row.memo : '') + '"></label>' +
      '<label>공사 코드<input type="text" class="form-control" name="project_code" maxlength="60" autocomplete="off" value="' + esc(row ? row.project_code : '') + '"></label>' +
      '</div><div class="gw-fb-form-acts">' +
      (row && !sms ? '<button type="button" class="btn btn-outline-danger me-auto" id="fbDel">지우기</button>' : '') +
      '<button type="button" class="btn btn-outline-secondary" id="fbCancel">닫기</button>' +
      '<button type="submit" class="btn btn-primary">저장</button></div>';
    form.innerHTML = h;
    form.hidden = false;
    form.scrollIntoView({ block: 'nearest' });
    var first = form.querySelector('select:enabled, input:enabled:not([type=hidden])');
    if (first) first.focus();
  }

  function closeForm() {
    var opener = state.editing && state.editing.opener;
    form.hidden = true;
    form.innerHTML = '';
    state.editing = null;
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
    busy = true;
    var btn = form.querySelector('[type=submit]');
    btn.disabled = true;
    (row ? api('/api/finance/bank/' + row.id, { method: 'POST', body: data })
      : api('/api/finance/bank', { method: 'POST', body: data })).then(function () {
      closeForm();
      flash('저장했어요.', 'ok');
      return load();
    }).catch(function (e) { flash(e.message, 'err'); }).then(function () {
      busy = false;
      if (btn) btn.disabled = false;
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
      }).catch(function (e) { flash(e.message, 'err'); });
    }
  });
  form.addEventListener('keydown', function (ev) { if (ev.key === 'Escape') closeForm(); });

  // 금액 칸: 쓰는 동안 천 단위 쉼표
  form.addEventListener('input', function (ev) {
    var el = ev.target;
    if (!el.classList.contains('num') || el.tagName !== 'INPUT') return;
    var neg = /^-/.test(el.value), digits = el.value.replace(/\D/g, '');
    el.value = (neg ? '-' : '') + (digits ? Number(digits).toLocaleString('ko-KR') : '');
  });

  // ── 이벤트 ────────────────────────────────────────────────────────
  $('fbAdd').addEventListener('click', function () { openForm('add', null, this); });
  $('fbBal').addEventListener('click', function () { openForm('bal', null, this); });
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
  ['fbAcct', 'fbDir', 'fbUn'].forEach(function (id) { $(id).addEventListener('change', load); });
  var qTimer;
  $('fbQ').addEventListener('input', function () { clearTimeout(qTimer); qTimer = setTimeout(load, 300); });

  setRange('7');
  load();
})();
