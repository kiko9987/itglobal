/* ① 방문 일정 (2026-11-22 부장님) — 슬랙 캔버스2 + /일정확정 대신 그룹웨어에서 다음 날 방문을 팀에 나눠 넣고 확정한다.
   차량 배정과 같은 모양(11-22 부장님): 왼쪽 사람·[함께] 조합 → 오른쪽 캔버스1 현장으로 끌어 놓기.
   휴대폰·키보드는 사람을 누르고(고르기) → 현장의 [여기에 넣기]. 사람 없는 현장은 [위 현장과 같게]로 빠르게.
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
  // P.plan.sites = {리드 No: {members: [이니셜], memo}} · P.picked = 'YM' 또는 조합 'YM,LSH'(차량 배정과 같은 고르기)
  function siteOf(no) { return P.plan.sites[no] || null; }
  function membersOf(no) { var s = siteOf(no); return s ? s.members : []; }
  function ensure(no) { if (!P.plan.sites[no]) P.plan.sites[no] = { members: [], memo: '' }; return P.plan.sites[no]; }
  function tidy(no) { var s = siteOf(no); if (s && !s.members.length && !s.memo) delete P.plan.sites[no]; }
  function keyCodes(key) { return String(key || '').split(',').filter(Boolean); }
  function offWhy(code) {
    var o = (P.data.off || []).filter(function (x) { return x.code === code; })[0];
    return o ? o.why : '';
  }
  function halfWhy(code) {
    var o = (P.data.half || []).filter(function (x) { return x.code === code; })[0];
    return o ? o.why : '';
  }
  function name(code) { return P.names[code] || code; }
  function keyName(key) { return keyCodes(key).map(name).join(' + '); }
  function siteCount(code) {
    return P.data.visits.filter(function (v) { return membersOf(v.lead_no).indexOf(code) !== -1; }).length;
  }
  function isDuty(code) { return P.plan.online.some(function (o) { return o.code === code; }); }
  function emptySites() { return P.data.visits.filter(function (v) { return !membersOf(v.lead_no).length; }); }
  // 같이 가는 조합(두 명 이상 현장) — 왼쪽 [함께] 단추로 한 번에 다른 현장에 넣는다
  function combos() {
    var seen = {}, out = [];
    P.data.visits.forEach(function (v) {
      var m = membersOf(v.lead_no);
      if (m.length < 2) return;
      var k = m.join(',');
      if (!seen[k]) { seen[k] = { key: k, n: 0 }; out.push(seen[k]); }
      seen[k].n++;
    });
    return out;
  }

  // 확정 전 점검 — 서버(visit_plan.warnings)와 같은 규칙. 확정할 때 서버가 다시 본다
  function warnings() {
    var out = [], offs = {}, warned = {};
    (P.data.off || []).forEach(function (o) { offs[o.code] = o.why; });
    var left = emptySites();
    if (left.length) out.push({ level: 'warn', text: '사람을 넣지 않은 현장 ' + left.length + '건 — 담당 DM 이 가지 않아요.' });
    P.data.visits.forEach(function (v) {
      var m = membersOf(v.lead_no);
      m.forEach(function (c) {
        if (offs[c] && !warned[c]) { warned[c] = 1; out.push({ level: 'warn', lead: v.lead_no, text: name(c) + '님은 이날 ' + offs[c] + '예요.' }); }
      });
      if (m.length && v.owner && m.indexOf(v.owner) === -1) {
        out.push({ level: 'warn', lead: v.lead_no, text: '본인 방문 필수 현장(' + String(v.addr || v.lead_no).slice(0, 30) + ')인데 상담한 ' + name(v.owner) + '님이 없어요.' });
      }
    });
    P.plan.online.forEach(function (o) { if (offs[o.code]) out.push({ level: 'warn', text: '온라인 당번 ' + name(o.code) + '님은 이날 ' + offs[o.code] + '예요.' }); });
    return out;
  }

  // ── 그리기 ──
  function marks(v) {
    return (v.platform ? '<span class="gw-va-smark plat">' + esc(v.platform) + '</span>' : '') +
      v.marks.map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('');
  }
  function personTag(c) {
    var why = offWhy(c) || halfWhy(c);
    return why ? ' <small class="gw-va-leave">' + esc(why) + '</small>' : '';
  }
  function personBtn(p) {
    var n = siteCount(p.code), on = P.picked === p.code;
    var tags = (n ? '<small class="gw-va-visit">현장 ' + n + '</small>' : '') + personTag(p.code) +
      (isDuty(p.code) ? ' <small class="gw-va-leave">온라인 당번</small>' : '');
    if (!P.canEdit) return '<div class="gw-va-person">' + '<span class="gw-va-pname">' + esc(p.name) + '</span>' + tags + '</div>';
    return '<button type="button" class="gw-va-person' + (on ? ' picked' : '') + (n ? ' placed' : '') + '" draggable="true" data-pick="' + esc(p.code) +
      '" aria-pressed="' + (on ? 'true' : 'false') + '"><span class="gw-va-pname">' + esc(p.name) + '</span>' + tags + '</button>';
  }
  function peoplePanel() {
    var ppl = P.data.people;
    var off = ppl.filter(function (p) { return offWhy(p.code); });
    var free = ppl.filter(function (p) { return !offWhy(p.code) && !siteCount(p.code); });
    var busy = ppl.filter(function (p) { return !offWhy(p.code) && siteCount(p.code); });
    var html = '';
    var cs = combos();
    if (cs.length && P.canEdit) {
      html += '<div class="gw-va-group"><div class="gw-va-gtitle">함께 가는 조합 <span>' + cs.length + '</span></div><div class="gw-va-plist">' +
        cs.map(function (c) {
          var on = P.picked === c.key;
          return '<button type="button" class="gw-va-teambtn' + (on ? ' picked' : '') + '" draggable="true" data-pick="' + esc(c.key) +
            '" aria-pressed="' + (on ? 'true' : 'false') + '"><i class="fas fa-people-group" aria-hidden="true"></i> ' + esc(keyName(c.key)) +
            ' <small>현장 ' + c.n + '</small></button>';
        }).join('') + '</div></div>';
    }
    [['현장 없음', free], ['현장 있음', busy], ['휴무', off]].forEach(function (g) {
      if (!g[1].length) return;
      html += '<div class="gw-va-group"><div class="gw-va-gtitle">' + esc(g[0]) + ' <span>' + g[1].length + '</span></div>' +
        '<div class="gw-va-plist">' + g[1].map(personBtn).join('') + '</div></div>';
    });
    return html;
  }
  function siteCard(v, prev) {
    var no = v.lead_no, s = siteOf(no) || { members: [], memo: '' }, m = s.members;
    var chips = m.map(function (c) {
      return '<span class="gw-va-chip"><span class="gw-va-cname">' + esc(name(c)) + '</span>' + personTag(c) +
        (P.canEdit ? '<button type="button" class="gw-va-x" data-act="rmp" data-lead="' + esc(no) + '" data-code="' + esc(c) + '" aria-label="' + esc(name(c)) + ' 빼기">×</button>' : '') + '</span>';
    }).join('');
    var acts = '';
    if (P.canEdit) {
      var add = P.picked ? keyCodes(P.picked).filter(function (c) { return m.indexOf(c) === -1; }) : [];
      if (add.length) acts += '<button type="button" class="btn btn-sm btn-outline-secondary gw-va-put" data-act="put" data-lead="' + esc(no) + '">여기에 ' + esc(add.map(name).join(' + ')) + ' 넣기</button>';
      var pm = prev ? membersOf(prev.lead_no) : [];
      if (!m.length && pm.length && !P.picked) acts += '<button type="button" class="btn btn-sm btn-outline-secondary gw-vp-same" data-act="same" data-lead="' + esc(no) + '" data-from="' + esc(prev.lead_no) + '">위 현장과 같게(' + esc(pm.map(name).join(' + ')) + ')</button>';
      acts += P.memoOpen[no] || s.memo
        ? '<label class="gw-vp-memo"><span>현장 메모 <small>담당자 DM 에 함께 가요</small></span><input type="text" class="form-control form-control-sm" maxlength="200" data-act="memo" data-lead="' + esc(no) + '" value="' + esc(s.memo) + '"></label>'
        : '<button type="button" class="btn btn-sm btn-link gw-vp-memobtn" data-act="memoopen" data-lead="' + esc(no) + '">메모 쓰기</button>';
    } else if (s.memo) {
      acts = '<p class="gw-vp-memotext"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(s.memo) + '</p>';
    }
    return '<section class="gw-va-card gw-vp-site' + (m.length ? '' : ' need') + (P.picked && P.canEdit ? ' can-drop' : '') + '" data-drop="site" data-lead="' + esc(no) + '" aria-label="' + esc(v.addr || no) + '">' +
      '<div class="gw-vp-vtop">' + marks(v) + '<strong>' + esc(v.date) + '</strong><span class="gw-va-sno">' + esc(no) + '</span></div>' +
      (v.addr ? '<div class="gw-vp-vaddr">' + esc(v.addr) + '</div>' : '') +
      (v.text ? '<div class="gw-vp-vtext">' + esc(v.text) + '</div>' : '') +
      '<div class="gw-va-people">' + (chips || '<span class="gw-va-none">' + (P.canEdit ? '여기에 사람을 놓으세요' : '사람 없음') + '</span>') + '</div>' +
      acts + '</section>';
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
  function status() {
    var d = P.data, parts = ['<strong>' + esc(d.day_label) + ' 현장 ' + d.visits.length + '건</strong>'];
    if (d.confirmed_at) {
      parts.push('<span class="gw-badge gw-badge-ok">확정</span> <span class="gw-va-when">' + esc(d.confirmed_at.slice(5, 16)) +
        (d.confirmed_name ? ' · ' + esc(d.confirmed_name) : '') + '</span>');
    } else {
      parts.push('<span class="gw-badge gw-badge-wait">확정 전</span>');
    }
    if (d.saved_at && !d.confirmed_at) parts.push('<span class="gw-va-when">임시 저장 ' + esc(d.saved_at.slice(11, 16)) + (d.saved_name ? ' · ' + esc(d.saved_name) : '') + ' · DM 아직 안 보냄</span>');
    var notes = [];
    if (d.dropped) notes.push('저장한 뒤 날짜가 바뀌거나 취소된 현장 ' + d.dropped + '건은 뺐어요.');
    if (d.changed_since_confirm || (d.confirmed_at && P.dirty)) notes.push('확정한 뒤 고친 내용이 있어요. 다시 확정해야 담당자 DM 이 바뀌어요.');
    if (!d.editable) notes.push('오늘과 지난 날 일정은 고칠 수 없어요.');
    else if (!P.canEdit) notes.push('방문 일정은 일정 담당자만 고칠 수 있어요.');
    $('vpStatus').innerHTML = parts.join(' ') + notes.map(function (n) { return '<span class="gw-va-draft">' + esc(n) + '</span>'; }).join('');
    $('vpTabPlanSt').textContent = d.confirmed_at ? (d.changed_since_confirm || P.dirty ? '· 고친 내용 있음' : '· 확정') : '· 확정 전';
  }
  function render() {
    var ws = warnings();
    var empty = emptySites().length;
    $('vpPeopleCount').textContent = P.data.people.length + '명';
    $('vpPeople').innerHTML = peoplePanel();
    $('vpHow').innerHTML = P.canEdit
      ? (P.picked ? esc(keyName(P.picked)) + ' 고름 — 넣을 현장에서 [여기에 넣기]를 눌러 주세요. <button type="button" class="btn btn-sm btn-link" data-act="unpick">고르기 취소</button>'
        : '이름이나 [함께] 조합을 오른쪽 현장으로 끌어 놓으세요. 누른 뒤 현장의 [여기에 넣기]를 눌러도 돼요.')
      : '';
    $('vpSiteCount').textContent = P.data.visits.length + '건' + (empty ? ' · 사람 없는 현장 ' + empty + '건' : '');
    var only = $('vpOnlyLeft').checked;
    // 캔버스1과 같은 묶음(기타·거래처·온라인 방문) 제목 아래에, 캔버스1 순서대로
    var groups = [];
    P.data.visits.forEach(function (v, i) {
      var g = groups.filter(function (x) { return x.title === v.title; })[0];
      if (!g) { g = { title: v.title, items: [] }; groups.push(g); }
      g.items.push({ v: v, prev: i ? P.data.visits[i - 1] : null });
    });
    $('vpSites').innerHTML = groups.map(function (g) {
      var items = g.items.filter(function (x) { return !only || !membersOf(x.v.lead_no).length; });
      if (!items.length) return '';
      return '<h3 class="gw-vp-gtitle">' + esc(g.title) + ' <span>' + g.items.length + '</span></h3><div class="gw-va-list gw-vp-sitelist">' +
        items.map(function (x) { return siteCard(x.v, x.prev && x.prev.title === g.title ? x.prev : null); }).join('') + '</div>';
    }).join('') || '<p class="gw-va-sempty">' + (P.data.visits.length ? '사람 없는 현장이 없어요.' : '이날 방문 일정이 없어요.') + '</p>';
    $('vpDuty').innerHTML = dutyPanel();
    $('vpWarns').innerHTML = ws.map(function (w) {
      return '<li class="' + (w.level === 'error' ? 'err' : 'warn') + '"><i class="fas ' + (w.level === 'error' ? 'fa-circle-xmark' : 'fa-triangle-exclamation') +
        '" aria-hidden="true"></i> ' + esc(w.text) + '</li>';
    }).join('');
    $('vpBoard').classList.toggle('editing', P.canEdit);
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
  function putOn(no, key) {
    var s = ensure(no), add = keyCodes(key).filter(function (c) { return s.members.indexOf(c) === -1; });
    add.forEach(function (c) { s.members.push(c); });
    P.picked = null;
    var v = P.byNo[no] || {};
    changed('[data-drop="site"][data-lead="' + cssq(no) + '"] .gw-va-x');
    say(add.length ? keyName(add.join(',')) + ' → ' + String(v.addr || no).slice(0, 24) + ' 넣었어요.' : '이미 들어가 있어요.', 'info');
  }
  P.memoOpen = {};
  $('vpOnlyLeft').addEventListener('change', function () { render(); });
  root.addEventListener('click', function (e) {
    var b = e.target.closest('[data-act], [data-pick]');
    if (!b || b.disabled || !P.canEdit) return;
    var pick = b.getAttribute('data-pick');
    if (pick) {
      P.picked = P.picked === pick ? null : pick;
      P.focus = '[data-pick="' + cssq(pick) + '"]';
      render();
      return;
    }
    var act = b.getAttribute('data-act'), no = b.getAttribute('data-lead');
    if (act === 'unpick') {
      P.picked = null; render();
    } else if (act === 'put') {
      if (P.picked) putOn(no, P.picked);
    } else if (act === 'same') {
      putOn(no, membersOf(b.getAttribute('data-from')).join(','));
    } else if (act === 'rmp') {
      var s = ensure(no), c = b.getAttribute('data-code');
      s.members.splice(s.members.indexOf(c), 1);
      tidy(no);
      changed('[data-drop="site"][data-lead="' + cssq(no) + '"] .gw-va-x, [data-drop="site"][data-lead="' + cssq(no) + '"] button');
    } else if (act === 'memoopen') {
      P.memoOpen[no] = true;
      P.focus = 'input[data-act="memo"][data-lead="' + cssq(no) + '"]';
      render();
    } else if (act === 'rmduty') {
      P.plan.online.splice(parseInt(b.getAttribute('data-k'), 10), 1);
      changed('.gw-vp-addduty');
    }
  });
  root.addEventListener('change', function (e) {
    var el = e.target;
    if (!P.canEdit) return;
    var act = el.getAttribute('data-act');
    if (act === 'addduty' && el.value) {
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
      var no = el.getAttribute('data-lead');
      ensure(no).memo = el.value;
      setDirty(true);
      $('vpTabPlanSt').textContent = P.data.confirmed_at ? '· 고친 내용 있음' : '· 확정 전';
    }
  });
  root.addEventListener('focusout', function (e) {       // 비운 메모 칸은 닫고, 사람·메모 없는 현장은 지운다
    var el = e.target;
    if (el.getAttribute && el.getAttribute('data-act') === 'memo') {
      var no = el.getAttribute('data-lead');
      if (!el.value.trim()) { delete P.memoOpen[no]; tidy(no); }
    }
  });
  root.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && P.picked) { P.picked = null; render(); }
  });

  // 끌어 놓기(마우스) — 사람·[함께] 조합 → 현장
  var dragKey = null;
  root.addEventListener('dragstart', function (e) {
    var el = e.target.closest && e.target.closest('[data-pick][draggable="true"]');
    if (!el || !P.canEdit) return;
    dragKey = el.getAttribute('data-pick');
    try { e.dataTransfer.setData('text/plain', dragKey); e.dataTransfer.effectAllowed = 'copy'; } catch (er) { /* 무시 */ }
  });
  function dropZone(e) { return e.target.closest && e.target.closest('[data-drop="site"]'); }
  root.addEventListener('dragover', function (e) {
    if (!dragKey) return;
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
    if (!dragKey || !z) return;
    e.preventDefault();
    var key = dragKey;
    dragKey = null;
    z.classList.remove('over');
    putOn(z.getAttribute('data-lead'), key);
  });
  root.addEventListener('dragend', function () {
    dragKey = null;
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
    P.memoOpen = {};
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
  function lock(on) { ['vpSave', 'vpConfirm'].forEach(function (id) { if ($(id)) $(id).disabled = on; }); }

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
    var placed = P.data.visits.filter(function (v) { return membersOf(v.lead_no).length; });
    if (!placed.length) { say('사람을 넣은 현장이 없어요.', 'err'); return; }
    var people = {};
    placed.forEach(function (v) { membersOf(v.lead_no).forEach(function (c) { people[c] = 1; }); });
    var lines = [P.data.day_label + ' 방문 일정을 ' + (P.data.confirmed_at ? '다시 ' : '') + '확정할까요?',
      '현장 ' + placed.length + '곳 · 담당자 ' + Object.keys(people).length + '명' +
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
