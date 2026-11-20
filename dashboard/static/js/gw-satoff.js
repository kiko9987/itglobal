/* 그룹웨어 토요일 휴무(기술영업 격주휴무, 11-20) — 달력의 토요일을 눌러 이름을 넣고 빼고, 결재를 올리고 승인한다.
   규칙·권한은 서버(services/sat_off.py)가 다시 확인한다. */
(function () {
  'use strict';

  var WD = ['월', '화', '수', '목', '금', '토', '일'];
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
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

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function md(iso) { return Number(iso.slice(5, 7)) + '/' + Number(iso.slice(8, 10)); }
  function showModal(el) {
    if (!window.bootstrap) return null;
    var opener = document.activeElement;
    el.addEventListener('hidden.bs.modal', function h() {
      el.removeEventListener('hidden.bs.modal', h);
      if (opener && opener.focus && document.contains(opener)) opener.focus();
    });
    var m = window.bootstrap.Modal.getOrCreateInstance(el);
    m.show();
    return m;
  }

  var now = new Date();
  var todayIso = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
  var qm = (new URLSearchParams(location.search).get('month') || '').match(/^(\d{4})-(\d{2})$/);
  var cur = qm ? { y: Number(qm[1]), m: Number(qm[2]) } : { y: now.getFullYear(), m: now.getMonth() + 1 };
  var data = null;          // 지금 보이는 달
  var draft = { add: [], remove: [] };   // 확정된 달 변경 신청 고르는 중(본인)
  var seq = 0;

  var calEl = document.getElementById('soCal');
  var barEl = document.getElementById('soBar');
  var flashEl = document.getElementById('soFlash');

  function monthKey() { return cur.y + '-' + pad(cur.m); }

  function flash(text, kind) {
    flashEl.textContent = text;
    flashEl.className = 'gw-lv-flash ' + (kind || 'ok');
    flashEl.hidden = !text;
  }

  function satOf(day) { return (data.saturdays || []).filter(function (s) { return s.day === day; })[0]; }

  // 토요일 칸을 누를 수 있나 — 'self'(내 이름 넣고 빼기) · 'pick'(사람 고르기 창) · 'draft'(변경 신청 고르기) · ''
  function clickMode(s) {
    if (!data || s.holiday) return '';
    var me = data.me || {};
    if (s.past && !me.approver) return '';
    if (data.status === 'confirmed') {
      if (me.approver) return 'pick';
      return data.can_request_change && me.member ? 'draft' : '';
    }
    if (!data.can_toggle) return '';
    if (me.approver || data.pick_others) return 'pick';
    return me.member ? 'self' : '';
  }

  function inDraft(list, day) { return list.indexOf(day) !== -1; }

  function render() {
    var d = data;
    document.getElementById('soTitle').textContent = cur.y + '년 ' + cur.m + '월';
    var st = document.getElementById('soStatus');
    st.className = 'gw-so-status ' + d.status;
    st.textContent = d.status_label;

    // 위쪽 안내·단추
    var me = d.me || {};
    var h = '';
    if (d.why) h += '<div class="gw-so-note warn">' + esc(d.why) + '</div>';
    if (d.status === 'open') {
      if (d.note && d.decided_by) h += '<div class="gw-so-note warn">↩️ ' + esc(d.decided_by) + ' 님이 돌려보냈어요: ' + esc(d.note) + '</div>';
      if (d.can_toggle) {
        h += '<p class="gw-so-how">' + (me.member && !d.pick_others
          ? '쉬고 싶은 토요일 칸을 누르면 내 이름이 들어가고, 한 번 더 누르면 빠져요. 한 달 ' + d.limit + '번이 기본이에요.'
          : '토요일 칸을 눌러 쉬는 사람을 고르세요. 다 모이면 [결재 올리기]를 눌러 주세요.') + '</p>';
      } else {
        h += '<p class="gw-so-how">아직 고를 수 없는 달이에요.</p>';
      }
      if (d.can_submit) h += '<button type="button" class="btn btn-primary btn-sm" data-act="submit">결재 올리기</button>';
    } else if (d.status === 'wait') {
      h += '<p class="gw-so-how">' + esc(d.submitted_by) + ' 님이 결재를 올렸어요(' + esc((d.submitted_at || '').slice(5, 16)) +
        '). 박정우 실장님이나 대표님 중 한 분이 승인하면 확정돼요.</p>';
      if (d.can_decide) {
        h += '<button type="button" class="btn btn-primary btn-sm" data-act="approve">승인</button> ' +
          '<button type="button" class="btn btn-outline-secondary btn-sm" data-act="return">돌려보내기</button>';
      }
    } else {
      h += '<p class="gw-so-how">' + esc(d.decided_by) + ' 님이 승인해 확정됐어요(' + esc((d.decided_at || '').slice(5, 16)) + ').' +
        (me.approver ? ' 토요일 칸을 눌러 바로 고칠 수 있어요(기록이 남아요).'
          : d.can_request_change ? ' 바꾸려면 토요일 칸을 눌러 고른 뒤 [변경 신청]을 눌러 주세요.' : '') + '</p>';
      if (draft.add.length || draft.remove.length) {
        var parts = draft.remove.map(function (x) { return md(x) + ' 빼기'; }).concat(draft.add.map(function (x) { return md(x) + ' 넣기'; }));
        h += '<div class="gw-so-draft"><span>바꿀 내용: ' + esc(parts.join(', ')) + '</span> ' +
          '<button type="button" class="btn btn-primary btn-sm" data-act="change">변경 신청</button> ' +
          '<button type="button" class="btn btn-light btn-sm" data-act="undo">되돌리기</button></div>';
      }
    }
    barEl.innerHTML = h;

    // 달력
    var first = new Date(cur.y, cur.m - 1, 1);
    var lastDay = new Date(cur.y, cur.m, 0).getDate();
    var lead = (first.getDay() + 6) % 7;
    var g = '<div class="gw-cal-grid">';
    WD.forEach(function (w, i) { g += '<div class="gw-cal-wd' + (i >= 5 ? ' we' : '') + '">' + w + '</div>'; });
    for (var i = 0; i < lead; i++) g += '<div class="gw-cal-day out"></div>';
    for (var dd = 1; dd <= lastDay; dd++) {
      var di = cur.y + '-' + pad(cur.m) + '-' + pad(dd);
      var wd = (lead + dd - 1) % 7;
      var s = wd === 5 ? satOf(di) : null;
      var cls = 'gw-cal-day' + (wd >= 5 ? ' we' : '') + (di === todayIso ? ' today' : '');
      if (!s) {
        g += '<div class="' + cls + (wd < 5 ? ' gw-so-dim' : '') + '"><div class="gw-cal-num">' + dd + '</div></div>';
        continue;
      }
      var mode = clickMode(s);
      var names = s.names.slice();
      var label = md(di) + ' 토요일, 쉬는 사람 ' + (names.length ? names.join(', ') : '없음');
      var attrs = mode ? ' role="button" tabindex="0" data-day="' + di + '" data-mode="' + mode + '" aria-label="' +
        esc(label + (mode === 'self' ? (s.mine ? ' — 눌러서 내 이름 빼기' : ' — 눌러서 내 이름 넣기') : ' — 눌러서 고치기')) + '"' : '';
      g += '<div class="' + cls + ' gw-so-sat' + (mode ? ' can' : '') + (s.past ? ' past' : '') + '"' + attrs + '>' +
        '<div class="gw-cal-num">' + dd + (s.holiday ? ' <small class="gw-so-hol">' + esc(s.holiday) + '</small>' : '') + '</div>';
      names.forEach(function (n) {
        var gone = me.name === n && inDraft(draft.remove, di);
        g += '<div class="gw-cal-chip gw-so-chip' + (n === me.name ? ' mine' : '') + (gone ? ' gone' : '') + '">' + esc(n) +
          (gone ? '<span class="visually-hidden"> (빼기로 고름)</span>' : '') + '</div>';
      });
      if (inDraft(draft.add, di)) g += '<div class="gw-cal-chip gw-so-chip mine add">' + esc(me.name) + '<span class="visually-hidden"> (넣기로 고름)</span></div>';
      if (mode === 'self' && !s.mine) g += '<div class="gw-so-plus" aria-hidden="true">+ 내 이름</div>';
      if (mode === 'pick') g += '<div class="gw-so-plus" aria-hidden="true">+ 고르기</div>';
      if (mode === 'draft' && !s.mine && !inDraft(draft.add, di)) g += '<div class="gw-so-plus" aria-hidden="true">+ 바꿔 쉬기</div>';
      g += '</div>';
    }
    var tail = (7 - ((lead + lastDay) % 7)) % 7;
    for (var t = 0; t < tail; t++) g += '<div class="gw-cal-day out"></div>';
    calEl.innerHTML = g + '</div>';

    // 사람별 횟수
    document.getElementById('soLimit').textContent = '(한 달 ' + d.limit + '번 기본)';
    document.getElementById('soCounts').innerHTML = (d.members || []).map(function (m) {
      var over = m.count > d.limit;
      return '<li class="' + (over ? 'over' : m.count ? '' : 'zero') + (m.name === me.name ? ' mine' : '') + '"><span>' + esc(m.name) + '</span>' +
        '<b>' + m.count + '번' + (over ? ' <span class="gw-so-warn">기본보다 많음</span>' : '') + '</b></li>';
    }).join('') || '<li class="zero">대상자가 없어요.</li>';

    // 변경 신청
    var ch = d.changes || [];
    document.getElementById('soChangesBox').hidden = !ch.length;
    document.getElementById('soChanges').innerHTML = ch.map(function (c) {
      var btns = '';
      if (c.status === 'wait' && me.approver) {
        btns = '<button type="button" class="btn btn-primary btn-sm" data-cact="approve" data-id="' + c.id + '">승인</button> ' +
          '<button type="button" class="btn btn-outline-secondary btn-sm" data-cact="reject" data-id="' + c.id + '">반려</button>';
      } else if (c.status === 'wait' && (c.name === me.name || c.requested_by === me.name)) {
        btns = '<button type="button" class="btn btn-light btn-sm" data-cact="cancel" data-id="' + c.id + '">신청 취소</button>';
      }
      return '<li><div><strong>' + esc(c.name) + '</strong> ' + esc(c.what) + ' <span class="gw-so-cst ' + c.status + '">' + esc(c.status_label) + '</span></div>' +
        '<div class="gw-so-sub">사유: ' + esc(c.reason || '') + (c.decide_note ? ' · 처리 메모: ' + esc(c.decide_note) : '') +
        (c.decided_by && c.status !== 'wait' ? ' · ' + esc(c.decided_by) : '') + '</div>' + (btns ? '<div class="gw-so-btns">' + btns + '</div>' : '') + '</li>';
    }).join('');

    // 기록
    var log = d.log || [];
    document.getElementById('soLogBox').hidden = !log.length;
    document.getElementById('soLog').innerHTML = log.map(function (l) {
      return '<li><span>' + esc((l.at || '').slice(5, 16)) + '</span> ' + esc(l.actor || '') + ' · ' + esc(l.detail || '') + '</li>';
    }).join('');
  }

  function load(keepDraft) {
    var k = monthKey();
    document.getElementById('soTitle').textContent = cur.y + '년 ' + cur.m + '월';
    if (!keepDraft) draft = { add: [], remove: [] };
    var my = ++seq;
    if (!data) calEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    return api('/api/satoff?month=' + k).then(function (j) {
      if (my !== seq) return;
      data = j;
      render();
      try { history.replaceState(null, '', '?month=' + k); } catch (e) { /* 무시 */ }
    }).catch(function (e) {
      if (my !== seq) return;
      calEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
      barEl.innerHTML = '';
    });
  }

  function move(n) {
    cur.m += n;
    if (cur.m < 1) { cur.m = 12; cur.y -= 1; }
    if (cur.m > 12) { cur.m = 1; cur.y += 1; }
    flash('');
    load();
  }

  document.getElementById('soPrev').addEventListener('click', function () { move(-1); });
  document.getElementById('soNext').addEventListener('click', function () { move(1); });
  document.getElementById('soToday').addEventListener('click', function () {
    cur = { y: now.getFullYear(), m: now.getMonth() + 1 }; flash(''); load();
  });

  // ── 토요일 칸 ──
  var busy = false;
  function onDay(cell) {
    if (busy || !data) return;
    var day = cell.getAttribute('data-day');
    var mode = cell.getAttribute('data-mode');
    var s = satOf(day);
    if (!s) return;
    if (mode === 'self') {
      busy = true;
      api('/api/satoff/toggle', { method: 'POST', body: { day: day, on: !s.mine } }).then(function (j) {
        flash(j.warn || (s.mine ? md(day) + ' 토요일에서 이름을 뺐어요.' : md(day) + ' 토요일에 이름을 넣었어요.'), j.warn ? 'err' : 'ok');
        return load();
      }).catch(function (e) { flash(e.message, 'err'); }).then(function () { busy = false; });
    } else if (mode === 'draft') {
      var name = data.me.name;
      if (s.mine) {
        if (inDraft(draft.remove, day)) draft.remove.splice(draft.remove.indexOf(day), 1); else draft.remove.push(day);
      } else if (s.names.indexOf(name) === -1) {
        if (inDraft(draft.add, day)) draft.add.splice(draft.add.indexOf(day), 1); else draft.add.push(day);
      }
      draft.add.sort(); draft.remove.sort();
      render();
      var again = calEl.querySelector('[data-day="' + day + '"]');
      if (again) again.focus();
    } else if (mode === 'pick') {
      openPick(s);
    }
  }
  calEl.addEventListener('click', function (e) {
    var cell = e.target.closest('.gw-so-sat[data-day]');
    if (cell) onDay(cell);
  });
  calEl.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var cell = e.target.closest('.gw-so-sat[data-day]');
    if (cell) { e.preventDefault(); onDay(cell); }
  });

  // ── 사람 고르기 창(취합 담당·승인권자) ──
  var pickEl = document.getElementById('soPickModal');
  var pickDay = null;
  var pickModal = null;
  function openPick(s) {
    pickDay = s.day;
    document.getElementById('soPickTitle').textContent = md(s.day) + ' 토요일 쉬는 사람';
    document.getElementById('soPickHint').textContent = data.status === 'confirmed' || data.status === 'wait'
      ? (data.status === 'confirmed' ? '확정된 달이에요. 저장하면 바로 바뀌고 기록이 남아요.' : '결재 대기 중인 달이에요. 승인권자만 고칠 수 있어요.')
      : '쉬는 사람을 고르세요. 이름 옆 숫자는 이번 달 고른 횟수예요.';
    var counts = {};
    (data.members || []).forEach(function (m) { counts[m.name] = m.count; });
    document.getElementById('soPickList').innerHTML = (data.members || []).map(function (m, i) {
      var on = s.names.indexOf(m.name) !== -1;
      return '<div class="form-check"><input class="form-check-input" type="checkbox" id="soPk' + i + '" value="' + esc(m.name) + '"' + (on ? ' checked' : '') + '>' +
        '<label class="form-check-label" for="soPk' + i + '">' + esc(m.name) + ' <small class="text-muted">' + (counts[m.name] || 0) + '번</small></label></div>';
    }).join('');
    document.getElementById('soPickErr').textContent = '';
    pickModal = showModal(pickEl);
  }
  document.getElementById('soPickForm').addEventListener('submit', function (ev) {
    ev.preventDefault();
    var s = satOf(pickDay);
    if (!s) return;
    var jobs = [];
    pickEl.querySelectorAll('#soPickList input').forEach(function (i) {
      var was = s.names.indexOf(i.value) !== -1;
      if (i.checked !== was) jobs.push({ name: i.value, on: i.checked });
    });
    var btn = document.getElementById('soPickSave');
    btn.disabled = true;
    var warns = [];
    var chain = Promise.resolve();
    jobs.forEach(function (j) {
      chain = chain.then(function () {
        return api('/api/satoff/toggle', { method: 'POST', body: { day: pickDay, name: j.name, on: j.on } })
          .then(function (r) { if (r.warn) warns.push(r.warn); });
      });
    });
    chain.then(function () {
      if (pickModal) pickModal.hide();
      flash(warns.length ? warns.join(' ') : (jobs.length ? md(pickDay) + ' 토요일을 고쳤어요.' : ''), warns.length ? 'err' : 'ok');
      return load();
    }).catch(function (e) {
      document.getElementById('soPickErr').textContent = e.message;
      load();
    }).then(function () { btn.disabled = false; });
  });

  // ── 사유 창 ──
  var noteEl = document.getElementById('soNoteModal');
  var noteJob = null;
  var noteModal = null;
  function openNote(title, what, label, okText, job) {
    noteJob = job;
    document.getElementById('soNoteTitle').textContent = title;
    document.getElementById('soNoteWhat').textContent = what;
    document.getElementById('soNoteLabel').textContent = label;
    document.getElementById('soNoteOk').textContent = okText;
    document.getElementById('soNoteText').value = '';
    document.getElementById('soNoteErr').textContent = '';
    noteModal = showModal(noteEl);
    setTimeout(function () { document.getElementById('soNoteText').focus(); }, 300);
  }
  document.getElementById('soNoteForm').addEventListener('submit', function (ev) {
    ev.preventDefault();
    var text = document.getElementById('soNoteText').value.trim();
    if (!text) { document.getElementById('soNoteErr').textContent = '사유를 적어 주세요.'; return; }
    var btn = document.getElementById('soNoteOk');
    btn.disabled = true;
    noteJob(text).then(function (msg) {
      if (noteModal) noteModal.hide();
      flash(msg, 'ok');
      return load();
    }).catch(function (e) { document.getElementById('soNoteErr').textContent = e.message; })
      .then(function () { btn.disabled = false; });
  });

  // ── 위쪽 단추 ──
  barEl.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-act]');
    if (!b || !data) return;
    var act = b.getAttribute('data-act');
    var k = data.month;
    var title = cur.y + '년 ' + cur.m + '월';
    if (act === 'submit') {
      if (!window.confirm(title + ' 토요일 휴무표를 결재에 올릴까요? 올린 뒤에는 승인권자만 고칠 수 있어요.')) return;
      b.disabled = true;
      api('/api/satoff/submit', { method: 'POST', body: { month: k } }).then(function () {
        flash('결재를 올렸어요. 박정우 실장님과 대표님께 슬랙으로 알려 드렸어요.', 'ok');
        return load();
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    } else if (act === 'approve') {
      b.disabled = true;
      api('/api/satoff/decide', { method: 'POST', body: { month: k, approve: true } }).then(function (j) {
        var extra = (j.affected || []).length ? ' 휴가 일수가 달라지는 신청 ' + j.affected.length + '건은 경영지원에 알렸어요.' : '';
        flash('승인했어요. #휴가_일정에 공지했어요.' + extra, 'ok');
        return load();
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    } else if (act === 'return') {
      openNote('돌려보내기', title + ' 토요일 휴무표를 다시 작성 중으로 돌려요.', '돌려보내는 이유', '돌려보내기', function (note) {
        return api('/api/satoff/decide', { method: 'POST', body: { month: k, approve: false, note: note } })
          .then(function () { return '돌려보냈어요. 올린 분께 슬랙으로 알려 드렸어요.'; });
      });
    } else if (act === 'undo') {
      draft = { add: [], remove: [] };
      render();
    } else if (act === 'change') {
      var add = draft.add.slice(), remove = draft.remove.slice();
      var what = remove.map(function (x) { return md(x) + '(토) 빼기'; }).concat(add.map(function (x) { return md(x) + '(토) 넣기'; })).join(', ');
      openNote('변경 신청', what, '바꾸는 이유', '신청', function (reason) {
        return api('/api/satoff/change', { method: 'POST', body: { month: k, add: add, remove: remove, reason: reason } })
          .then(function () { draft = { add: [], remove: [] }; return '변경 신청을 했어요. 박정우 실장님이나 대표님이 승인하면 바뀌어요.'; });
      });
    }
  });

  // ── 변경 신청 처리 ──
  document.getElementById('soChanges').addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-cact]');
    if (!b) return;
    var id = b.getAttribute('data-id');
    var act = b.getAttribute('data-cact');
    var c = (data.changes || []).filter(function (x) { return String(x.id) === id; })[0];
    if (!c) return;
    if (act === 'approve') {
      b.disabled = true;
      api('/api/satoff/change/' + id + '/decide', { method: 'POST', body: { approve: true } }).then(function () {
        flash(c.name + ' 님 변경을 승인했어요.', 'ok'); return load(true);
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    } else if (act === 'reject') {
      openNote('변경 반려', c.name + ' 님 ' + c.what, '반려하는 이유', '반려', function (note) {
        return api('/api/satoff/change/' + id + '/decide', { method: 'POST', body: { approve: false, note: note } })
          .then(function () { return c.name + ' 님 변경 신청을 반려했어요.'; });
      });
    } else if (act === 'cancel') {
      if (!window.confirm('변경 신청을 취소할까요?')) return;
      b.disabled = true;
      api('/api/satoff/change/' + id + '/cancel', { method: 'POST' }).then(function () {
        flash('변경 신청을 취소했어요.', 'ok'); return load(true);
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    }
  });

  load();
})();
