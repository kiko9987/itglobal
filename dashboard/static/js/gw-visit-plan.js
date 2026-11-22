/* ① 방문 일정 (2026-11-22 부장님) — 슬랙 캔버스2 + /일정확정 대신 그룹웨어에서 다음 날 방문을 팀에 나눠 넣고 확정한다.
   왼쪽 '배정 안 된 방문' → 오른쪽 팀 칸으로 끌어 놓기. 휴대폰·키보드는 방문을 누르고(고르기) → 팀의 [여기로 옮기기].
   [임시 저장]은 DM 없이 저장, [일정 확정]은 /일정확정과 같은 처리(시트·슬랙 List·담당자/당번 DM·캔버스1)를 뒤에서 돌린다.
   날짜는 gw-vehicle-assign.js 가 고르고 'va:day' 로 알려 준다. 서버에서 온 글자는 모두 esc() 로만 넣는다. */
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
  var root = $('vpPlan');
  if (!root) return;

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url, init) {
    init = init || {};
    init.credentials = 'same-origin';
    var method = init.method || 'GET';
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          var e = new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '저장하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
          e.status = r.status;
          throw e;
        }
        return j;
      });
    });
  }
  function post(url, body) {
    return api(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  }
  function cssq(v) { return String(v).replace(/["\\]/g, '\\$&'); }

  var P = { day: '', label: '', data: null, plan: null, canEdit: false, dirty: false, picked: null,
            step: '', seq: 0, poll: null, names: {}, byNo: {}, focus: null };

  function say(text, kind) {
    var el = $('vpBarMsg');
    var useBar = el && !$('vpBar').hidden;
    var top = $('vpInfo');
    if (useBar) { el.className = 'gw-vh-msg gw-va-barmsg' + (text ? ' ' + (kind || 'info') : ''); el.textContent = text || ''; }
    else if (top) { top.className = 'gw-vp-info' + (text ? ' gw-vh-msg ' + (kind || 'info') : ''); top.textContent = text || ''; }
  }
  function setDirty(v) {
    P.dirty = v;
    if ($('vpSave')) $('vpSave').disabled = !v;
  }
  window.vpDirty = function () { return P.dirty && P.canEdit; };

  // ── 단계(① 방문 일정 / ② 차량 배정) ──
  function showStep(step, push) {
    P.step = step;
    var plan = step === 'plan';
    root.hidden = !plan;
    $('vaStepCars').hidden = plan;
    $('vpTabPlan').setAttribute('aria-selected', plan ? 'true' : 'false');
    $('vpTabCars').setAttribute('aria-selected', plan ? 'false' : 'true');
    $('vpTabPlan').tabIndex = plan ? 0 : -1;
    $('vpTabCars').tabIndex = plan ? -1 : 0;
    if (push !== false) setUrl();
  }
  function setUrl() {
    if (!P.day) return;
    try { history.replaceState(null, '', '/vehicles/assign?day=' + P.day + (P.step === 'plan' ? '&step=plan' : '')); } catch (e) { /* 무시 */ }
  }
  $('vpTabPlan').addEventListener('click', function () { showStep('plan'); });
  $('vpTabCars').addEventListener('click', function () { showStep('cars'); });
  $('vpSteps').addEventListener('keydown', function (e) {      // 탭 사이 ←/→
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    e.preventDefault();
    showStep(P.step === 'plan' ? 'cars' : 'plan');
    (P.step === 'plan' ? $('vpTabPlan') : $('vpTabCars')).focus();
  });

  // ── 자료 ──
  function team(i) { return P.plan.teams[i]; }
  function teamOf(no) {
    for (var i = 0; i < P.plan.teams.length; i++) if (P.plan.teams[i].leads.indexOf(no) !== -1) return i;
    return -1;
  }
  function leftVisits() { return P.data.visits.filter(function (v) { return teamOf(v.lead_no) === -1; }); }
  function offWhy(code) {
    var o = (P.data.off || []).filter(function (x) { return x.code === code; })[0];
    return o ? o.why : '';
  }
  function halfWhy(code) {
    var o = (P.data.half || []).filter(function (x) { return x.code === code; })[0];
    return o ? o.why : '';
  }
  function name(code) { return P.names[code] || code; }
  function teamName(t) { return t.members.length ? t.members.map(name).join(' + ') : '사람을 넣어 주세요'; }
  function inTeams(code) { return P.plan.teams.filter(function (t) { return t.members.indexOf(code) !== -1; }).length; }
  function isDuty(code) { return P.plan.online.some(function (o) { return o.code === code; }); }

  // 확정 전 점검 — 서버(visit_plan.warnings)와 같은 규칙. 확정할 때 서버가 다시 본다
  function warnings() {
    var out = [], offs = {};
    (P.data.off || []).forEach(function (o) { offs[o.code] = o.why; });
    var left = leftVisits();
    if (left.length) out.push({ level: 'warn', text: '배정 안 된 방문 ' + left.length + '건 — 담당 DM 이 가지 않아요.' });
    var count = {};
    P.plan.teams.forEach(function (t, i) {
      if (t.leads.length && !t.members.length) out.push({ level: 'error', team: i, text: '사람을 넣지 않은 팀에 방문 ' + t.leads.length + '건이 있어요.' });
      t.members.forEach(function (c) {
        count[c] = (count[c] || 0) + 1;
        if (offs[c]) out.push({ level: 'warn', team: i, text: name(c) + '님은 이날 ' + offs[c] + '예요.' });
      });
      t.leads.forEach(function (n) {
        var v = P.byNo[n] || {};
        if (v.owner && t.members.indexOf(v.owner) === -1) {
          out.push({ level: 'warn', team: i, text: '본인 방문 필수 건(' + String(v.addr || n).slice(0, 30) + ')인데 상담한 ' + name(v.owner) + '님이 팀에 없어요.' });
        }
      });
    });
    Object.keys(count).forEach(function (c) { if (count[c] > 1) out.push({ level: 'warn', text: name(c) + '님이 팀 ' + count[c] + '곳에 들어가 있어요.' }); });
    P.plan.online.forEach(function (o) { if (offs[o.code]) out.push({ level: 'warn', text: '온라인 당번 ' + name(o.code) + '님은 이날 ' + offs[o.code] + '예요.' }); });
    return out;
  }

  // ── 그리기 ──
  function marks(v) {
    return (v.platform ? '<span class="gw-va-smark plat">' + esc(v.platform) + '</span>' : '') +
      v.marks.map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('');
  }
  function visitBody(v) {
    return '<span class="gw-vp-vtop">' + marks(v) + '<strong>' + esc(v.date) + '</strong>' +
      '<span class="gw-va-sno">' + esc(v.lead_no) + '</span></span>' +
      (v.addr ? '<span class="gw-vp-vaddr">' + esc(v.addr) + '</span>' : '') +
      (v.text ? '<span class="gw-vp-vtext">' + esc(v.text) + '</span>' : '');
  }
  function leftCard(v) {
    var on = P.picked === v.lead_no;
    if (!P.canEdit) return '<div class="gw-vp-visit">' + visitBody(v) + '</div>';
    return '<button type="button" class="gw-vp-visit' + (on ? ' picked' : '') + '" draggable="true" data-lead="' + esc(v.lead_no) +
      '" data-act="pick" aria-pressed="' + (on ? 'true' : 'false') + '">' + visitBody(v) + '</button>';
  }
  function teamVisit(v, i, k, n) {
    var on = P.picked === v.lead_no;
    var acts = P.canEdit ? '<span class="gw-vp-vacts">' +
      '<button type="button" class="gw-vp-ic" data-act="up" data-lead="' + esc(v.lead_no) + '"' + (k === 0 ? ' disabled' : '') + ' aria-label="' + esc(v.lead_no) + ' 위로"><i class="fas fa-arrow-up" aria-hidden="true"></i></button>' +
      '<button type="button" class="gw-vp-ic" data-act="down" data-lead="' + esc(v.lead_no) + '"' + (k === n - 1 ? ' disabled' : '') + ' aria-label="' + esc(v.lead_no) + ' 아래로"><i class="fas fa-arrow-down" aria-hidden="true"></i></button>' +
      '<button type="button" class="gw-vp-tx' + (on ? ' on' : '') + '" data-act="pick" data-lead="' + esc(v.lead_no) + '" aria-pressed="' + (on ? 'true' : 'false') + '">옮기기</button>' +
      '<button type="button" class="gw-vp-tx" data-act="out" data-lead="' + esc(v.lead_no) + '">빼기</button></span>' : '';
    return '<li class="gw-vp-visit in' + (on ? ' picked' : '') + '"' + (P.canEdit ? ' draggable="true"' : '') + ' data-lead="' + esc(v.lead_no) + '">' +
      '<span class="gw-vp-no" aria-hidden="true">' + (k + 1) + '</span><span class="gw-vp-vmain">' + visitBody(v) + '</span>' + acts + '</li>';
  }
  function personTag(c) {
    var why = offWhy(c) || halfWhy(c);
    return why ? ' <small class="gw-va-leave">' + esc(why) + '</small>' : '';
  }
  function addPersonSelect(i, t) {
    var opts = P.data.people.filter(function (p) { return t.members.indexOf(p.code) === -1; }).map(function (p) {
      var bits = [];
      if (offWhy(p.code)) bits.push(offWhy(p.code));
      else if (halfWhy(p.code)) bits.push(halfWhy(p.code));
      if (inTeams(p.code)) bits.push('다른 팀');
      if (isDuty(p.code)) bits.push('온라인 당번');
      return { free: !inTeams(p.code) && !offWhy(p.code), html: '<option value="' + esc(p.code) + '">' + esc(p.name) + (bits.length ? ' (' + esc(bits.join(', ')) + ')' : '') + '</option>' };
    });
    var free = opts.filter(function (o) { return o.free; }).map(function (o) { return o.html; }).join('');
    var rest = opts.filter(function (o) { return !o.free; }).map(function (o) { return o.html; }).join('');
    return '<select class="form-select form-select-sm gw-vp-addp" data-team="' + i + '" aria-label="' + esc(teamName(t)) + ' 팀에 사람 넣기">' +
      '<option value="">+ 사람 넣기</option>' + (free ? '<optgroup label="팀 없는 사람">' + free + '</optgroup>' : '') +
      (rest ? '<optgroup label="그 밖의 사람">' + rest + '</optgroup>' : '') + '</select>';
  }
  function teamCard(t, i, ws) {
    var bad = ws.filter(function (w) { return w.team === i; });
    var vs = t.leads.map(function (n) { return P.byNo[n]; }).filter(Boolean);
    var chips = t.members.map(function (c) {
      return '<span class="gw-va-chip"><span class="gw-va-cname">' + esc(name(c)) + '</span>' + personTag(c) +
        (P.canEdit ? '<button type="button" class="gw-va-x" data-act="rmp" data-team="' + i + '" data-code="' + esc(c) + '" aria-label="' + esc(name(c)) + ' 빼기">×</button>' : '') + '</span>';
    }).join('');
    var put = P.canEdit && P.picked && teamOf(P.picked) !== i
      ? '<button type="button" class="btn btn-sm btn-outline-secondary gw-va-put" data-act="put" data-team="' + i + '">여기로 옮기기</button>' : '';
    var memo = P.canEdit
      ? '<label class="gw-vp-memo"><span>메모 <small>담당자 DM 에 함께 가요</small></span><input type="text" class="form-control form-control-sm" maxlength="200" data-act="memo" data-team="' + i + '" value="' + esc(t.memo) + '"></label>'
      : (t.memo ? '<p class="gw-vp-memotext"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(t.memo) + '</p>' : '');
    return '<section class="gw-vp-team' + (bad.some(function (w) { return w.level === 'error'; }) ? ' bad' : '') + (P.picked && P.canEdit ? ' can-drop' : '') +
      '" data-drop="team" data-team="' + i + '" aria-label="' + esc(teamName(t)) + ' 팀">' +
      '<div class="gw-vp-thead"><div class="gw-va-people">' + (chips || '<span class="gw-va-none">사람을 넣어 주세요</span>') + '</div>' +
      '<span class="gw-vp-tcount">' + vs.length + '건</span></div>' +
      (P.canEdit ? '<div class="gw-vp-ttools">' + addPersonSelect(i, t) +
        '<button type="button" class="btn btn-sm btn-link gw-vp-tdel" data-act="deltm" data-team="' + i + '">팀 지우기</button></div>' : '') +
      put + '<ol class="gw-vp-tlist">' + vs.map(function (v, k) { return teamVisit(v, i, k, vs.length); }).join('') + '</ol>' +
      (!vs.length ? '<p class="gw-vp-tempty">' + (P.canEdit ? '방문을 끌어 놓거나 [여기로 옮기기]를 눌러 주세요.' : '방문 없음') + '</p>' : '') +
      memo + '</section>';
  }
  function dutyPanel() {
    var rows = P.plan.online.map(function (o, k) {
      if (!P.canEdit) return '<span class="gw-va-chip">' + esc(name(o.code)) + (o.shift ? ' (' + esc(o.shift) + ')' : '') + '</span>';
      return '<span class="gw-va-chip gw-vp-dchip"><span class="gw-va-cname">' + esc(name(o.code)) + '</span>' + personTag(o.code) +
        '<select class="form-select form-select-sm gw-vp-shift" data-act="shift" data-k="' + k + '" aria-label="' + esc(name(o.code)) + ' 당번 시간">' +
        ['', '오전', '오후'].map(function (s) { return '<option value="' + s + '"' + (o.shift === s ? ' selected' : '') + '>' + (s || '종일') + '</option>'; }).join('') +
        '</select><button type="button" class="gw-va-x" data-act="rmduty" data-k="' + k + '" aria-label="' + esc(name(o.code)) + ' 당번 빼기">×</button></span>';
    }).join('');
    var add = '';
    if (P.canEdit) {
      add = '<select class="form-select form-select-sm gw-vp-addduty" data-act="addduty" aria-label="온라인 당번 넣기"><option value="">+ 당번 넣기</option>' +
        P.data.people.filter(function (p) { return !isDuty(p.code); }).map(function (p) {
          var w = offWhy(p.code) || halfWhy(p.code);
          return '<option value="' + esc(p.code) + '">' + esc(p.name) + (w ? ' (' + esc(w) + ')' : '') + '</option>';
        }).join('') + '</select>' +
        '<label class="gw-vp-shared"><input type="checkbox" data-act="shared"' + (P.plan.online_shared ? ' checked' : '') +
        (P.plan.online.length ? ' disabled' : '') + '> 당번 없이 전원이 나눠서 받기</label>';
    }
    var none = !P.plan.online.length && !P.canEdit
      ? '<span class="gw-va-none">' + (P.plan.online_shared ? '전원이 나눠서 받기' : '정하지 않음') + '</span>' : '';
    var offs = (P.data.off || []).map(function (o) { return o.name + '(' + o.why + ')'; });
    var halves = (P.data.half || []).map(function (o) { return o.name + '(' + o.why + ')'; });
    var leave = '<p class="gw-vp-leave"><strong>휴무</strong> ' + esc(offs.join(', ') || '없음') +
      (halves.length ? ' · <strong>반차</strong> ' + esc(halves.join(', ')) : '') + ' <small>그룹웨어 휴가·토요 휴무표에서 자동</small></p>';
    return '<h2 class="gw-vp-h">온라인 당번</h2><div class="gw-vp-drow">' + rows + none + add + '</div>' + leave;
  }
  function freeLine() {
    var free = P.data.people.filter(function (p) { return !inTeams(p.code) && !isDuty(p.code) && !offWhy(p.code); });
    return free.length ? '팀 없는 사람: ' + free.map(function (p) { return p.name + (halfWhy(p.code) ? '(' + halfWhy(p.code) + ')' : ''); }).join(', ') : '';
  }
  function status() {
    var d = P.data, parts = ['<strong>' + esc(d.day_label) + ' 방문 ' + d.visits.length + '건</strong>'];
    if (d.confirmed_at) {
      parts.push('<span class="gw-badge gw-badge-ok">확정</span> <span class="gw-va-when">' + esc(d.confirmed_at.slice(5, 16)) +
        (d.confirmed_name ? ' · ' + esc(d.confirmed_name) : '') + '</span>');
    } else {
      parts.push('<span class="gw-badge gw-badge-wait">확정 전</span>');
    }
    if (d.saved_at && !d.confirmed_at) parts.push('<span class="gw-va-when">임시 저장 ' + esc(d.saved_at.slice(11, 16)) + (d.saved_name ? ' · ' + esc(d.saved_name) : '') + ' · DM 아직 안 보냄</span>');
    var notes = [];
    if (P.canEdit && d.source === 'sheet') notes.push('시트에 이미 적힌 담당자대로 팀을 묶어 두었어요. 바뀐 곳만 고쳐 주세요.');
    if (P.canEdit && d.source === 'empty' && d.visits.length) notes.push('왼쪽 방문을 팀에 나눠 넣어 주세요. [새 팀]으로 팀을 만들어요.');
    if (d.dropped) notes.push('저장한 뒤 날짜가 바뀌거나 취소된 방문 ' + d.dropped + '건은 팀에서 뺐어요.');
    if (d.changed_since_confirm || (d.confirmed_at && P.dirty)) notes.push('확정한 뒤 고친 내용이 있어요. 다시 확정해야 담당자 DM 이 바뀌어요.');
    if (!d.editable) notes.push('오늘과 지난 날 일정은 고칠 수 없어요.');
    else if (!P.canEdit) notes.push('방문 일정은 일정 담당자만 고칠 수 있어요.');
    $('vpStatus').innerHTML = parts.join(' ') + notes.map(function (n) { return '<span class="gw-va-draft">' + esc(n) + '</span>'; }).join('');
    $('vpTabPlanSt').textContent = d.confirmed_at ? (d.changed_since_confirm || P.dirty ? '· 고친 내용 있음' : '· 확정') : '· 확정 전';
  }
  function render() {
    var ws = warnings();
    var left = leftVisits();
    $('vpLeftCount').textContent = left.length + '건';
    $('vpLeft').innerHTML = left.map(leftCard).join('') ||
      '<p class="gw-va-sempty">' + (P.data.visits.length ? '모든 방문을 팀에 넣었어요.' : '이날 방문 일정이 없어요.') + '</p>';
    $('vpLeftBox').classList.toggle('can-drop', !!(P.picked && P.canEdit && teamOf(P.picked) !== -1));
    $('vpHow').innerHTML = P.canEdit
      ? (P.picked ? '고른 방문을 넣을 팀에서 [여기로 옮기기]를 눌러 주세요. <button type="button" class="btn btn-sm btn-link" data-act="unpick">고르기 취소</button>' +
        (teamOf(P.picked) !== -1 ? ' <button type="button" class="btn btn-sm btn-outline-secondary gw-va-put" data-act="put" data-team="-1">배정 안 됨으로</button>' : '')
        : '방문 카드를 오른쪽 팀으로 끌어 놓거나, 눌러서 고른 뒤 팀에서 [여기로 옮기기]를 눌러 주세요.')
      : '';
    $('vpTeams').innerHTML = P.plan.teams.map(function (t, i) { return teamCard(t, i, ws); }).join('') ||
      '<p class="gw-va-sempty">아직 팀이 없어요.' + (P.canEdit ? ' [새 팀]을 눌러 만들어 주세요.' : '') + '</p>';
    $('vpAddTeam').hidden = !P.canEdit;
    $('vpFree').textContent = freeLine();
    $('vpDuty').innerHTML = dutyPanel();
    $('vpWarns').innerHTML = ws.map(function (w) {
      return '<li class="' + (w.level === 'error' ? 'err' : 'warn') + '"><i class="fas ' + (w.level === 'error' ? 'fa-circle-xmark' : 'fa-triangle-exclamation') +
        '" aria-hidden="true"></i> ' + esc(w.text) + '</li>';
    }).join('');
    $('vpBar').hidden = !P.canEdit;
    $('vpConfirm').textContent = P.data.confirmed_at ? '다시 확정 · 바뀐 사람에게 DM' : '일정 확정 · DM 보내기';
    setDirty(P.dirty);
    status();
    if (P.focus) {
      var el = root.querySelector(P.focus);
      P.focus = null;
      if (el) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
    }
  }

  // ── 고치기 ──
  function changed(focusSel) { setDirty(true); P.focus = focusSel || null; render(); }
  function moveTo(no, ti) {                  // ti = -1 이면 배정 안 됨으로
    var from = teamOf(no);
    if (from === ti) return;
    if (from !== -1) team(from).leads.splice(team(from).leads.indexOf(no), 1);
    if (ti !== -1) team(ti).leads.push(no);
    P.picked = null;
    changed(ti === -1 ? '#vpLeft [data-lead="' + cssq(no) + '"]' : '.gw-vp-team[data-team="' + ti + '"] [data-act="pick"][data-lead="' + cssq(no) + '"]');
    say(ti === -1 ? '배정 안 됨으로 옮겼어요.' : teamName(team(ti)) + ' 팀으로 옮겼어요.', 'info');
  }
  root.addEventListener('click', function (e) {
    var b = e.target.closest('[data-act]');
    if (!b || b.disabled || !P.canEdit) return;
    var act = b.getAttribute('data-act');
    var no = b.getAttribute('data-lead');
    var ti = parseInt(b.getAttribute('data-team'), 10);
    if (act === 'pick') {
      P.picked = P.picked === no ? null : no;
      P.focus = (teamOf(no) === -1 ? '#vpLeft' : '.gw-vp-team') + ' [data-act="pick"][data-lead="' + cssq(no) + '"]';
      render();
    } else if (act === 'unpick') {
      P.picked = null; render();
    } else if (act === 'put') {
      if (P.picked) moveTo(P.picked, ti);
    } else if (act === 'out') {
      moveTo(no, -1);
    } else if (act === 'up' || act === 'down') {
      var t = team(teamOf(no)), k = t.leads.indexOf(no), j = act === 'up' ? k - 1 : k + 1;
      if (j < 0 || j >= t.leads.length) return;
      t.leads.splice(k, 1); t.leads.splice(j, 0, no);
      changed('.gw-vp-team [data-act="' + act + '"][data-lead="' + cssq(no) + '"]');
    } else if (act === 'rmp') {
      var tm = team(ti), c = b.getAttribute('data-code');
      tm.members.splice(tm.members.indexOf(c), 1);
      changed('.gw-vp-addp[data-team="' + ti + '"]');
    } else if (act === 'deltm') {
      var dt = team(ti);
      if (dt.leads.length && !window.confirm(teamName(dt) + ' 팀을 지울까요?\n방문 ' + dt.leads.length + '건은 배정 안 됨으로 돌아가요.')) return;
      P.plan.teams.splice(ti, 1);
      P.picked = null;
      changed('#vpAddTeam');
    } else if (act === 'rmduty') {
      P.plan.online.splice(parseInt(b.getAttribute('data-k'), 10), 1);
      changed('.gw-vp-addduty');
    }
  });
  root.addEventListener('change', function (e) {
    var el = e.target;
    if (!P.canEdit) return;
    var act = el.getAttribute('data-act');
    if (el.classList.contains('gw-vp-addp') && el.value) {
      var ti = parseInt(el.getAttribute('data-team'), 10);
      team(ti).members.push(el.value);
      changed('.gw-vp-addp[data-team="' + ti + '"]');
    } else if (act === 'addduty' && el.value) {
      P.plan.online.push({ code: el.value, shift: '' });
      P.plan.online_shared = false;
      changed('.gw-vp-addduty');
    } else if (act === 'shift') {
      P.plan.online[parseInt(el.getAttribute('data-k'), 10)].shift = el.value;
      setDirty(true);
    } else if (act === 'shared') {
      P.plan.online_shared = el.checked;
      setDirty(true);
    }
  });
  root.addEventListener('input', function (e) {
    var el = e.target;
    if (P.canEdit && el.getAttribute('data-act') === 'memo') {
      team(parseInt(el.getAttribute('data-team'), 10)).memo = el.value;
      setDirty(true);
      $('vpTabPlanSt').textContent = P.data.confirmed_at ? '· 고친 내용 있음' : '· 확정 전';
    }
  });
  root.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && P.picked) { P.picked = null; render(); }
  });
  $('vpAddTeam').addEventListener('click', function () {
    P.plan.teams.push({ members: [], leads: [], memo: '' });
    changed('.gw-vp-addp[data-team="' + (P.plan.teams.length - 1) + '"]');
  });

  // 끌어 놓기(마우스) — 놓을 곳: 팀 칸·왼쪽 '배정 안 된 방문'
  var dragNo = null;
  root.addEventListener('dragstart', function (e) {
    var el = e.target.closest && e.target.closest('[data-lead][draggable="true"]');
    if (!el || !P.canEdit) return;
    dragNo = el.getAttribute('data-lead');
    try { e.dataTransfer.setData('text/plain', dragNo); e.dataTransfer.effectAllowed = 'move'; } catch (er) { /* 무시 */ }
  });
  function dropZone(e) { return e.target.closest && e.target.closest('[data-drop]'); }
  root.addEventListener('dragover', function (e) {
    if (!dragNo) return;
    var z = dropZone(e);
    if (!z) return;
    e.preventDefault();
    root.querySelectorAll('.over').forEach(function (x) { if (x !== z) x.classList.remove('over'); });
    z.classList.add('over');
  });
  root.addEventListener('dragleave', function (e) {
    var z = dropZone(e);
    if (z && !z.contains(e.relatedTarget)) z.classList.remove('over');
  });
  root.addEventListener('drop', function (e) {
    var z = dropZone(e);
    if (!dragNo || !z) return;
    e.preventDefault();
    var no = dragNo;
    dragNo = null;
    z.classList.remove('over');
    moveTo(no, z.getAttribute('data-drop') === 'left' ? -1 : parseInt(z.getAttribute('data-team'), 10));
  });
  root.addEventListener('dragend', function () {
    dragNo = null;
    root.querySelectorAll('.over').forEach(function (x) { x.classList.remove('over'); });
  });

  // ── 불러오기·저장·확정 ──
  function take(j) {
    P.data = j;
    P.plan = JSON.parse(JSON.stringify(j.plan));
    P.canEdit = !!j.can_edit;
    P.names = {}; P.byNo = {};
    j.people.forEach(function (p) { P.names[p.code] = p.name; });
    j.visits.forEach(function (v) { P.byNo[v.lead_no] = v; });
    P.picked = null;
  }
  function load(day, keepMsg) {
    var my = ++P.seq;
    clearTimeout(P.poll);
    return api('/api/vehicles/plan?day=' + encodeURIComponent(day)).then(function (j) {
      if (my !== P.seq) return;
      take(j);
      setDirty(false);
      $('vpSteps').hidden = false;
      if (!P.step) {
        var q = new URLSearchParams(location.search).get('step');
        showStep(q === 'plan' || (P.canEdit && !j.confirmed_at) ? 'plan' : 'cars', true);
      } else setUrl();
      if (!keepMsg) say('');
      render();
      runState(j.run, false);
    }).catch(function (e) {
      if (my !== P.seq) return;
      if (e.status === 403 || e.message === SESSION_ENDED) {       // 현장 세션 등 — 단계 없이 차량 배정만
        $('vpSteps').hidden = true; root.hidden = true; $('vaStepCars').hidden = false;
        $('vaStepCars').removeAttribute('role');
        return;
      }
      $('vpSteps').hidden = false;
      say(e.message, 'err');
    });
  }
  function onDay(d) {
    d = d || {};
    if (!d.day) return;
    if (d.day === P.day && P.data) { setUrl(); return; }     // 차량 배정만 다시 불러온 것 — 주소의 단계만 되살림
    P.day = d.day; P.label = d.label || '';
    load(d.day);
  }
  document.addEventListener('va:day', function (e) { onDay(e.detail); });

  function body() { return { day: P.day, plan: P.plan }; }
  function lock(on) { ['vpSave', 'vpConfirm', 'vpAddTeam'].forEach(function (id) { if ($(id)) $(id).disabled = on; }); }

  $('vpSave').addEventListener('click', function () {
    lock(true);
    post('/api/vehicles/plan', body()).then(function (j) {
      take(j); setDirty(false); render();
      say('임시 저장했어요. 담당자 DM 은 아직 안 보냈어요.', 'ok');
    }).catch(function (e) { say(e.message, 'err'); }).then(function () { lock(false); setDirty(P.dirty); });
  });

  $('vpConfirm').addEventListener('click', function () {
    var ws = warnings();
    var err = ws.filter(function (w) { return w.level === 'error'; })[0];
    if (err) { say(err.text, 'err'); return; }
    var placed = P.plan.teams.filter(function (t) { return t.members.length && t.leads.length; });
    if (!placed.length) { say('팀에 넣은 방문이 없어요.', 'err'); return; }
    var people = {};
    placed.forEach(function (t) { t.members.forEach(function (c) { people[c] = 1; }); });
    var lines = [P.data.day_label + ' 방문 일정을 ' + (P.data.confirmed_at ? '다시 ' : '') + '확정할까요?',
      '팀 ' + placed.length + '곳 · 담당자 ' + Object.keys(people).length + '명' +
      (P.data.confirmed_at ? ' — 바뀐 사람에게만 DM 이 가요.' : '에게 DM 이 가요.')];
    if (ws.length) {
      lines.push('', '확인할 것:');
      ws.forEach(function (w) { lines.push('· ' + w.text); });
    }
    if (!window.confirm(lines.join('\n'))) return;
    lock(true);
    say('확정하는 중이에요. 30초쯤 걸려요…', 'info');
    post('/api/vehicles/plan/confirm', body()).then(function (j) {
      setDirty(false);
      runState(j.run, true);
    }).catch(function (e) { lock(false); setDirty(P.dirty); say(e.message, 'err'); });
  });

  function resultText(r) {
    var bits = ['시트 담당자 ' + (r.sheet || 0) + '건', '담당자 DM ' + (r.dm || 0) + '명'];
    if (r.duty) bits.push('당번 DM ' + r.duty + '명');
    if (r.deassign) bits.push('재배정 안내 ' + r.deassign + '명');
    var bad = [];
    if ((r.sheet_failed || []).length) bad.push('시트 저장 실패 ' + r.sheet_failed.length + '건');
    if (r.dm_failed) bad.push('DM 실패 ' + r.dm_failed + '명');
    if (r.list_failed) bad.push('슬랙 List 실패 ' + r.list_failed + '건');
    if ((r.list_missing || []).length) bad.push('슬랙 List 담당자 칸에 선택지가 없는 사람이 있어요(' +
      r.list_missing.map(function (x) { return (x.names || []).join('·'); }).join(', ') + ')');
    if ((r.cards || []).length) bad.push('캔버스1에 없는 방문 카드 ' + r.cards.length + '건(사장님께 DM 으로 알렸어요)');
    return { ok: bits.join(' · '), bad: bad };
  }
  // 확정 진행 상태 — running 이면 2초마다 다시 묻고, 끝나면 결과를 단추 위에
  function runState(run, fresh) {
    clearTimeout(P.poll);
    run = run || {};
    if (run.state === 'running') {
      lock(true);
      say((run.by ? run.by + '님이 ' : '') + '확정하는 중이에요. 30초쯤 걸려요…', 'info');
      var my = P.seq, day = P.day;
      P.poll = setTimeout(function () {
        api('/api/vehicles/plan/state?day=' + encodeURIComponent(day)).then(function (j) {
          if (my === P.seq) runState(j.run, true);
        }).catch(function () { if (my === P.seq) runState(run, true); });
      }, 2000);
      return;
    }
    lock(false); setDirty(P.dirty);
    if (!fresh) {
      if (run.state === 'stale') say('지난 확정이 중간에 끊겼어요. 다시 확정해 주세요.', 'err');
      else if (run.state === 'failed' && !P.data.confirmed_at) say('지난 확정이 안 됐어요: ' + ((run.result || {}).message || ''), 'err');
      return;
    }
    if (run.state === 'done') {
      var t = resultText(run.result || {});
      load(P.day, true).then(function () {
        say('확정했어요. ' + t.ok + (t.bad.length ? ' — 확인 필요: ' + t.bad.join(' / ') : ''), t.bad.length ? 'warn' : 'ok');
        var go = document.createElement('button');
        go.type = 'button'; go.className = 'btn btn-sm btn-outline-primary gw-vp-next'; go.textContent = '② 차량 배정으로';
        go.addEventListener('click', function () { showStep('cars'); window.scrollTo(0, 0); });
        var el = $('vpBarMsg');
        if (el && !$('vpBar').hidden) { el.appendChild(document.createTextNode(' ')); el.appendChild(go); }
      });
      if (window.vaReload) window.vaReload();
    } else if (run.state === 'failed' || run.state === 'stale') {
      say(run.state === 'stale' ? '확정이 중간에 끊겼어요. 다시 확정해 주세요.' : ((run.result || {}).message || '확정하지 못했어요.'), 'err');
    }
  }

  window.addEventListener('beforeunload', function (e) {
    if (P.dirty && P.canEdit) { e.preventDefault(); e.returnValue = ''; }
  });
  if (window.vaCurrent) onDay(window.vaCurrent);     // 차량 배정이 이미 날짜를 불러왔으면 바로
})();
