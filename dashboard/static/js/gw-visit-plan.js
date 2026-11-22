/* ① 방문 일정 (2026-11-22 부장님) — 슬랙 캔버스2 + /일정확정 대신 그룹웨어에서 다음 날 방문을 팀에 나눠 넣고 확정한다.
   차량 배정과 같은 모양(11-22 부장님): 위 코스 카드(= 차 카드), 왼쪽 사람 → 코스에 넣기.
   JW 가 메모장에서 하던 순서 — 아래 현장 목록(캔버스1, 지역끼리 모아 보임)을 하나씩 [새 코스로]·끌어 올리기 또는
   체크해서 [코스 N으로] 한꺼번에 올려 코스를 먼저 묶고 → 사람 넣기 → 넣었다 뺐다. 휴대폰·키보드는 끌기 대신 단추.
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

  var P = { day: '', label: '', data: null, plan: null, canEdit: false, dirty: false,
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
  // P.plan.courses = [{name, members: [이니셜], leads: [리드 No(가는 순서)], memo}]
  // P.pp = 체크한 사람(→ 코스로 끌기·[코스 고르기]) · P.checked = 체크한 현장(→ [코스 N으로 올리기]) — 둘 다 같은 방식(11-22 부장님)
  function courseOf(no) {
    for (var i = 0; i < P.plan.courses.length; i++) if (P.plan.courses[i].leads.indexOf(no) !== -1) return i;
    return -1;
  }
  function label(i) { var c = P.plan.courses[i]; return (c && c.name) || ('코스 ' + (i + 1)); }
  function offWhy(code) {
    var o = (P.data.off || []).filter(function (x) { return x.code === code; })[0];
    return o ? o.why : '';
  }
  function halfWhy(code) {
    var o = (P.data.half || []).filter(function (x) { return x.code === code; })[0];
    return o ? o.why : '';
  }
  function name(code) { return P.names[code] || code; }
  function coursesWith(code) {
    var out = [];
    P.plan.courses.forEach(function (c, i) { if (c.members.indexOf(code) !== -1) out.push(i); });
    return out;
  }
  function isDuty(code) { return P.plan.online.some(function (o) { return o.code === code; }); }
  function inbox() { return P.data.visits.filter(function (v) { return courseOf(v.lead_no) === -1; }); }
  function checkedList() { return inbox().filter(function (v) { return P.checked[v.lead_no]; }).map(function (v) { return v.lead_no; }); }

  // 확정 전 점검 — 서버(visit_plan.warnings)와 같은 규칙. 확정할 때 서버가 다시 본다
  function warnings() {
    var out = [], offs = {}, count = {};
    (P.data.off || []).forEach(function (o) { offs[o.code] = o.why; });
    var done = {};
    P.plan.courses.forEach(function (c) { if (c.members.length) c.leads.forEach(function (n) { done[n] = 1; }); });
    var left = P.data.visits.filter(function (v) { return !done[v.lead_no]; }).length;
    if (left) out.push({ level: 'warn', text: '코스·사람이 안 정해진 현장 ' + left + '건 — 담당 DM 이 가지 않아요.' });
    P.plan.courses.forEach(function (c, i) {
      if (c.leads.length && !c.members.length) out.push({ level: 'warn', course: i, text: label(i) + '에 사람을 넣지 않았어요(현장 ' + c.leads.length + '건).' });
      c.members.forEach(function (m) {
        count[m] = (count[m] || 0) + 1;
        if (offs[m]) out.push({ level: 'warn', course: i, text: label(i) + '의 ' + name(m) + '님은 이날 ' + offs[m] + '예요.' });
      });
      if (c.members.length) c.leads.forEach(function (n) {
        var v = P.byNo[n] || {};
        if (v.owner && c.members.indexOf(v.owner) === -1) {
          out.push({ level: 'warn', course: i, text: label(i) + '의 본인 방문 필수 현장(' + String(v.addr || n).slice(0, 30) + ')에 상담한 ' + name(v.owner) + '님이 없어요.' });
        }
      });
    });
    Object.keys(count).forEach(function (m) { if (count[m] > 1) out.push({ level: 'warn', text: name(m) + '님이 코스 ' + count[m] + '곳에 들어가 있어요.' }); });
    P.plan.online.forEach(function (o) { if (offs[o.code]) out.push({ level: 'warn', text: '온라인 당번 ' + name(o.code) + '님은 이날 ' + offs[o.code] + '예요.' }); });
    return out;
  }

  // ── 그리기 ──
  function marks(v) {
    return (v.platform ? '<span class="gw-va-smark plat">' + esc(v.platform) + '</span>' : '') +
      v.marks.map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('');
  }
  function placeTag(v) { return v.place ? '<span class="gw-vp-place">' + esc(v.place) + '</span>' : ''; }
  function personTag(c) {
    var why = offWhy(c) || halfWhy(c);
    return why ? ' <small class="gw-va-leave">' + esc(why) + '</small>' : '';
  }
  function checkedPeople() { return P.data.people.filter(function (p) { return P.pp[p.code]; }).map(function (p) { return p.code; }); }
  function personBtn(p) {
    var cs = coursesWith(p.code), on = !!P.pp[p.code];
    var tags = (cs.length ? '<small class="gw-va-visit">' + esc(cs.map(label).join(', ')) + '</small>' : '') + personTag(p.code) +
      (isDuty(p.code) ? ' <small class="gw-va-leave">온라인 당번</small>' : '');
    if (!P.canEdit) return '<div class="gw-va-person"><span class="gw-va-pname">' + esc(p.name) + '</span>' + tags + '</div>';
    // 현장 목록과 같은 방식 — 줄 아무 곳이나 눌러 체크, 체크한 사람을 끌면 모두 함께
    return '<div class="gw-va-person gw-vp-pperson' + (on ? ' on' : '') + (cs.length ? ' placed' : '') + '" draggable="true" data-person="' + esc(p.code) + '">' +
      '<input type="checkbox" class="form-check-input gw-vp-check" data-act="pcheck" data-code="' + esc(p.code) + '"' + (on ? ' checked' : '') +
      ' aria-label="' + esc(p.name) + ' 고르기"><span class="gw-va-pname">' + esc(p.name) + '</span>' + tags + '</div>';
  }
  // 사람 칸 위 한 줄 — 안내 글 ↔ '체크한 n명 → [코스 고르기]' 가 같은 자리(높이 고정)
  function peopleBar() {
    var n = checkedPeople().length;
    if (!P.canEdit) return '';
    if (!n) return '<span class="gw-vp-how">체크해서 코스로 끌어 놓으세요.</span>';
    return '<strong>' + n + '명</strong> → <select class="form-select form-select-sm gw-vp-liftsel" data-act="psel" aria-label="체크한 사람을 넣을 코스">' +
      '<option value="">코스에 넣기</option>' + P.plan.courses.map(function (c, i) { return '<option value="' + i + '">' + esc(label(i)) + '</option>'; }).join('') +
      '</select><button type="button" class="btn btn-sm btn-link" data-act="punchk">풀기</button>';
  }
  function peoplePanel() {
    var ppl = P.data.people;
    var off = ppl.filter(function (p) { return offWhy(p.code); });
    var free = ppl.filter(function (p) { return !offWhy(p.code) && !coursesWith(p.code).length; });
    var busy = ppl.filter(function (p) { return !offWhy(p.code) && coursesWith(p.code).length; });
    return [['코스 없음', free], ['코스 있음', busy], ['휴무', off]].map(function (g) {
      if (!g[1].length) return '';
      return '<div class="gw-va-group"><div class="gw-va-gtitle">' + esc(g[0]) + ' <span>' + g[1].length + '</span></div>' +
        '<div class="gw-va-plist">' + g[1].map(personBtn).join('') + '</div></div>';
    }).join('');
  }
  function courseRow(v, i, k, n) {
    var acts = P.canEdit ? '<span class="gw-vp-racts">' +
      '<button type="button" class="gw-vp-ic" data-act="up" data-lead="' + esc(v.lead_no) + '"' + (k === 0 ? ' disabled' : '') + ' aria-label="' + esc(v.place || v.lead_no) + ' 앞으로"><i class="fas fa-arrow-up" aria-hidden="true"></i></button>' +
      '<button type="button" class="gw-vp-ic" data-act="down" data-lead="' + esc(v.lead_no) + '"' + (k === n - 1 ? ' disabled' : '') + ' aria-label="' + esc(v.place || v.lead_no) + ' 뒤로"><i class="fas fa-arrow-down" aria-hidden="true"></i></button>' +
      '<button type="button" class="gw-vp-ic" data-act="out" data-lead="' + esc(v.lead_no) + '" aria-label="' + esc(v.place || v.lead_no) + ' 현장 목록으로 내리기" title="현장 목록으로 내리기"><i class="fas fa-arrow-turn-down" aria-hidden="true"></i></button></span>' : '';
    return '<li class="gw-vp-row"' + (P.canEdit ? ' draggable="true" data-drag-lead="' + esc(v.lead_no) + '"' : '') + '>' +
      '<span class="gw-vp-no" aria-hidden="true">' + (k + 1) + '</span><span class="gw-vp-rmain">' +
      '<span class="gw-vp-vtop">' + placeTag(v) + marks(v) + '<span class="gw-va-sno">' + esc(v.date) + ' · ' + esc(v.lead_no) + '</span></span>' +
      (v.addr ? '<span class="gw-vp-vaddr">' + esc(v.addr) + '</span>' : '') +
      (v.text ? '<span class="gw-vp-vtext">' + esc(v.text) + '</span>' : '') + '</span>' + acts + '</li>';
  }
  function courseCard(c, i, ws) {
    var bad = ws.some(function (w) { return w.course === i; });
    var vs = c.leads.map(function (n) { return P.byNo[n]; }).filter(Boolean);
    var chips = c.members.map(function (m) {
      return '<span class="gw-va-chip"><span class="gw-va-cname">' + esc(name(m)) + '</span>' + personTag(m) +
        (P.canEdit ? '<button type="button" class="gw-va-x" data-act="rmp" data-ci="' + i + '" data-code="' + esc(m) + '" aria-label="' + esc(label(i)) + '에서 ' + esc(name(m)) + ' 빼기">×</button>' : '') + '</span>';
    }).join('');
    var acts = '';
    if (P.canEdit) {
      // 체크한 현장은 코스 카드에 단추를 늘리지 않고(화면이 출렁임, 11-22 부장님) 위 '체크한 현장 → [코스 N로]' 줄·끌기로 올린다
    }
    // 캔버스2 섹션처럼 한 줄: 코스 N · 사람 칩 · 현장 수 · [지우기]
    var head = '<div class="gw-vp-chead"><strong>' + esc(label(i)) + '</strong>' +
      '<div class="gw-va-people">' + (chips || '<span class="gw-va-none">' + (P.canEdit ? '여기에 사람을 놓으세요' : '사람 없음') + '</span>') + '</div>' +
      '<span class="gw-vp-ccount">현장 ' + vs.length + '</span>' +
      (P.canEdit ? '<button type="button" class="btn btn-sm btn-link gw-vp-cdel" data-act="delc" data-ci="' + i + '">지우기</button>' : '') + '</div>';
    var memo = P.canEdit
      ? '<label class="gw-vp-memo"><span>메모</span><input type="text" class="form-control form-control-sm" maxlength="200" data-act="memo" data-ci="' + i + '" value="' + esc(c.memo) + '" aria-describedby="vpMemoHelp"></label>'
      : (c.memo ? '<p class="gw-vp-memotext"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(c.memo) + '</p>' : '');
    return '<section class="gw-va-card gw-vp-course' + (bad ? ' warn' : '') + (c.members.length ? '' : ' need') + (checkedPeople().length && P.canEdit ? ' can-drop' : '') +
      '" data-drop="course" data-ci="' + i + '" aria-label="' + esc(label(i)) + '">' + head +
      '<ol class="gw-vp-rows">' + vs.map(function (v, k) { return courseRow(v, i, k, vs.length); }).join('') + '</ol>' +
      (!vs.length ? '<p class="gw-vp-tempty">' + (P.canEdit ? '아래 현장 목록에서 끌어 올리거나, 체크해서 올려 주세요.' : '현장 없음') + '</p>' : '') +
      (acts ? '<div class="gw-vp-cacts">' + acts + '</div>' : '') + memo + '</section>';
  }
  // 현장 목록 — 지역(region)끼리 모아, 지역은 캔버스1에서 처음 나온 순서대로
  function inboxRows() {
    var groups = [];
    inbox().forEach(function (v) {
      var key = v.region || '지역 모름';
      var g = groups.filter(function (x) { return x.key === key; })[0];
      if (!g) { g = { key: key, items: [] }; groups.push(g); }
      g.items.push(v);
    });
    return groups.map(function (g) {
      var allOn = g.items.every(function (v) { return P.checked[v.lead_no]; });
      var head = '<h3 class="gw-vp-gtitle">' + esc(g.key) + ' <span>' + g.items.length + '</span>' +
        (P.canEdit && g.items.length > 1 ? ' <button type="button" class="btn btn-sm btn-link gw-vp-gall" data-act="checkgroup" data-region="' + esc(g.key) + '">' + (allOn ? '체크 풀기' : '모두 체크') + '</button>' : '') + '</h3>';
      // 체크한 현장은 제목 줄의 '체크한 현장 → [새 코스로]·[코스 N로]' 로 올린다(지역 단추는 두지 않음, 11-22 부장님)
      return '<div class="gw-vp-group">' + head + '<ul class="gw-vp-inrows">' + g.items.map(function (v) {
        var on = !!P.checked[v.lead_no];
        var check = P.canEdit ? '<input type="checkbox" class="form-check-input gw-vp-check" data-act="check" data-lead="' + esc(v.lead_no) + '"' + (on ? ' checked' : '') +
          ' aria-label="' + esc((v.place || '') + ' ' + v.lead_no) + ' 고르기">' : '';
        return '<li class="gw-vp-inrow' + (on ? ' on' : '') + '"' + (P.canEdit ? ' draggable="true" data-drag-lead="' + esc(v.lead_no) + '"' : '') + '>' + check +
          '<span class="gw-vp-rmain"><span class="gw-vp-vtop">' + placeTag(v) + marks(v) + '<strong>' + esc(v.date) + '</strong><span class="gw-va-sno">' + esc(v.lead_no) + '</span></span>' +
          (v.addr ? '<span class="gw-vp-vaddr">' + esc(v.addr) + '</span>' : '') +
          (v.text ? '<span class="gw-vp-vtext">' + esc(v.text) + '</span>' : '') + '</span>' +
          // 체크한 줄은 [새 코스로]를 숨긴다 — 체크한 것은 위 [새 코스로]·지역 단추로 함께 올림
          (P.canEdit && !on ? '<button type="button" class="btn btn-sm btn-outline-secondary gw-vp-newc" data-act="newc" data-lead="' + esc(v.lead_no) + '">새 코스로</button>' : '') + '</li>';
      }).join('') + '</ul></div>';
    }).join('');
  }
  function selBar() {
    var n = checkedList().length;
    var bar = $('vpSelBar');
    if (!P.canEdit) { bar.innerHTML = ''; return; }
    if (!n) {
      bar.innerHTML = '<span class="gw-vp-how">같은 지역끼리 모아 보여 줘요. 위 코스로 끌어 올리거나(코스 안 줄 위에 놓으면 그 앞에), 체크해서 한꺼번에 올리세요.</span>';
      return;
    }
    // 코스가 많으면 단추가 두 줄로 넘어가 목록이 밀림 → 4개부터는 목록 상자 하나로
    var to = P.plan.courses.length > 3
      ? '<select class="form-select form-select-sm gw-vp-liftsel" data-act="liftsel" aria-label="체크한 현장을 올릴 코스"><option value="">코스 고르기</option>' +
        P.plan.courses.map(function (c, i) { return '<option value="' + i + '">' + esc(label(i)) + '</option>'; }).join('') + '</select>'
      : P.plan.courses.map(function (c, i) {
        return '<button type="button" class="btn btn-sm btn-outline-secondary" data-act="lift" data-ci="' + i + '">' + esc(label(i)) + '로</button>';
      }).join('');
    bar.innerHTML = '<strong>체크한 현장 ' + n + '건</strong> → ' + to + '<button type="button" class="btn btn-sm btn-primary" data-act="liftnew">새 코스로</button>' +
      '<button type="button" class="btn btn-sm btn-link" data-act="uncheck">체크 풀기</button>';
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
    if (d.dropped) notes.push('저장한 뒤 날짜가 바뀌거나 취소된 현장 ' + d.dropped + '건은 코스에서 뺐어요.');
    if (d.changed_since_confirm || (d.confirmed_at && P.dirty)) notes.push('확정한 뒤 고친 내용이 있어요. 다시 확정해야 담당자 DM 이 바뀌어요.');
    if (!d.editable) notes.push('오늘과 지난 날 일정은 고칠 수 없어요.');
    else if (!P.canEdit) notes.push('방문 일정은 일정 담당자만 고칠 수 있어요.');
    $('vpStatus').innerHTML = parts.join(' ') + notes.map(function (n) { return '<span class="gw-va-draft">' + esc(n) + '</span>'; }).join('');
    $('vpTabPlanSt').textContent = d.confirmed_at ? (d.changed_since_confirm || P.dirty ? '· 고친 내용 있음' : '· 확정') : '· 확정 전';
  }
  function render() {
    var ws = warnings(), left = inbox();
    Object.keys(P.checked).forEach(function (n) { if (courseOf(n) !== -1) delete P.checked[n]; });
    $('vpPeopleCount').textContent = P.data.people.length + '명';
    $('vpPeople').innerHTML = peoplePanel();
    $('vpPeopleHow').innerHTML = peopleBar();
    $('vpCourseCount').textContent = P.plan.courses.length ? P.plan.courses.length + '개' : '';
    $('vpCourses').innerHTML = P.plan.courses.map(function (c, i) { return courseCard(c, i, ws); }).join('') ||
      '<p class="gw-va-sempty">' + (P.canEdit ? '아래 현장의 [새 코스로]를 누르거나, 현장을 체크한 뒤 [새 코스로]를 눌러 코스를 만들어 주세요.' : '아직 코스가 없어요.') + '</p>';
    $('vpAddCourse').hidden = !P.canEdit;
    $('vpInboxCount').textContent = left.length + '건' + (left.length !== P.data.visits.length ? ' (전체 ' + P.data.visits.length + '건)' : '');
    selBar();
    $('vpInList').innerHTML = inboxRows() ||
      '<p class="gw-va-sempty">' + (P.data.visits.length ? '모든 현장을 코스에 넣었어요.' : '이날 방문 일정이 없어요.') + '</p>';
    $('vpDuty').innerHTML = dutyPanel();
    $('vpWarns').innerHTML = ws.map(function (w) {
      return '<li class="warn"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> ' + esc(w.text) + '</li>';
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
  function newCourse() { P.plan.courses.push({ name: '', members: [], leads: [], memo: '' }); return P.plan.courses.length - 1; }
  function moveTo(nos, ci, before) {        // ci = -1 이면 현장 목록으로 내림, before = 이 현장 앞에 끼움
    nos.forEach(function (no) {
      var from = courseOf(no);
      if (from !== -1) P.plan.courses[from].leads.splice(P.plan.courses[from].leads.indexOf(no), 1);
      delete P.checked[no];
    });
    if (ci !== -1) {
      var leads = P.plan.courses[ci].leads;
      var at = before && nos.indexOf(before) === -1 ? leads.indexOf(before) : -1;
      if (at === -1) at = leads.length;
      Array.prototype.splice.apply(leads, [at, 0].concat(nos));
    }
    changed(ci === -1 ? '#vpInList [data-act="check"]' : '.gw-vp-course[data-ci="' + ci + '"] [data-act="memo"]');
    say(ci === -1 ? '현장 목록으로 내렸어요.' : label(ci) + '에 현장 ' + nos.length + '건을 올렸어요.', 'info');
  }
  function putPeople(codes, ci) {
    var c = P.plan.courses[ci];
    codes.forEach(function (code) { if (c.members.indexOf(code) === -1) c.members.push(code); delete P.pp[code]; });
    changed('.gw-vp-course[data-ci="' + ci + '"] .gw-va-x');
    say(label(ci) + '에 ' + codes.map(name).join(', ') + '님을 넣었어요.', 'info');
  }
  P.checked = {};
  P.pp = {};
  $('vpAddCourse').addEventListener('click', function () {
    var ci = newCourse();
    changed('.gw-vp-course[data-ci="' + ci + '"] [data-act="memo"]');
  });
  root.addEventListener('click', function (e) {
    // 현장 목록 줄은 아무 곳이나 눌러도 체크(11-22 부장님) — [새 코스로] 단추·체크 칸 자체는 그대로
    var inrow = P.canEdit && e.target.closest('.gw-vp-inrow');
    if (inrow && !e.target.closest('button, input, a, select')) {
      var cb = inrow.querySelector('[data-act="check"]');
      if (cb) { cb.checked = !cb.checked; cb.dispatchEvent(new Event('change', { bubbles: true })); }
      return;
    }
    var prow = P.canEdit && e.target.closest('.gw-vp-pperson');      // 사람 줄도 아무 곳이나 눌러 체크
    if (prow && !e.target.closest('button, input, a, select')) {
      var pcb = prow.querySelector('[data-act="pcheck"]');
      if (pcb) { pcb.checked = !pcb.checked; pcb.dispatchEvent(new Event('change', { bubbles: true })); }
      return;
    }
    var b = e.target.closest('[data-act]');
    if (!b || b.disabled || !P.canEdit) return;
    if (b.matches('input')) return;                          // 체크·메모 칸은 change/input 에서
    var act = b.getAttribute('data-act'), no = b.getAttribute('data-lead'), ci = parseInt(b.getAttribute('data-ci'), 10);
    if (act === 'punchk') { P.pp = {}; P.focus = '#vpPeople [data-act="pcheck"]'; render(); }
    else if (act === 'newc') { moveTo([no], newCourse()); }
    else if (act === 'lift') { moveTo(checkedList(), ci); }
    else if (act === 'liftnew') { var list = checkedList(); if (list.length) moveTo(list, newCourse()); }
    else if (act === 'uncheck') { P.checked = {}; render(); }
    else if (act === 'checkgroup') {
      var region = b.getAttribute('data-region');
      var items = inbox().filter(function (v) { return (v.region || '지역 모름') === region; });
      var allOn = items.every(function (v) { return P.checked[v.lead_no]; });
      items.forEach(function (v) { if (allOn) delete P.checked[v.lead_no]; else P.checked[v.lead_no] = true; });
      P.focus = '[data-act="checkgroup"][data-region="' + cssq(region) + '"]';
      render();
    } else if (act === 'out') { moveTo([no], -1); }
    else if (act === 'up' || act === 'down') {
      var c = P.plan.courses[courseOf(no)], k = c.leads.indexOf(no), j = act === 'up' ? k - 1 : k + 1;
      if (j < 0 || j >= c.leads.length) return;
      c.leads.splice(k, 1); c.leads.splice(j, 0, no);
      changed('.gw-vp-course [data-act="' + act + '"][data-lead="' + cssq(no) + '"]');
    } else if (act === 'rmp') {
      var cm = P.plan.courses[ci], code = b.getAttribute('data-code');
      cm.members.splice(cm.members.indexOf(code), 1);
      changed('.gw-vp-course[data-ci="' + ci + '"] [data-act="memo"]');
    } else if (act === 'delc') {
      var dc = P.plan.courses[ci];
      if ((dc.leads.length || dc.members.length) && !window.confirm(label(ci) + '를 지울까요?\n현장 ' + dc.leads.length + '건은 현장 목록으로 돌아가요.')) return;
      P.plan.courses.splice(ci, 1);
      changed('#vpAddCourse');
    } else if (act === 'rmduty') {
      P.plan.online.splice(parseInt(b.getAttribute('data-k'), 10), 1);
      changed('.gw-vp-addduty');
    }
  });
  root.addEventListener('change', function (e) {
    var el = e.target;
    if (!P.canEdit) return;
    var act = el.getAttribute('data-act');
    if (act === 'liftsel' && el.value !== '') {
      moveTo(checkedList(), parseInt(el.value, 10));
    } else if (act === 'psel' && el.value !== '') {
      putPeople(checkedPeople(), parseInt(el.value, 10));
    } else if (act === 'pcheck') {
      var code = el.getAttribute('data-code');
      if (el.checked) P.pp[code] = true; else delete P.pp[code];
      P.focus = '[data-act="pcheck"][data-code="' + cssq(code) + '"]';
      render();
    } else if (act === 'check') {
      if (el.checked) P.checked[el.getAttribute('data-lead')] = true; else delete P.checked[el.getAttribute('data-lead')];
      P.focus = '[data-act="check"][data-lead="' + cssq(el.getAttribute('data-lead')) + '"]';
      render();
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
    var el = e.target, act = el.getAttribute('data-act');
    if (!P.canEdit || act !== 'memo') return;
    P.plan.courses[parseInt(el.getAttribute('data-ci'), 10)].memo = el.value;
    setDirty(true);
    $('vpTabPlanSt').textContent = P.data.confirmed_at ? '· 고친 내용 있음' : '· 확정 전';
  });
  root.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && checkedPeople().length) { P.pp = {}; render(); }
  });

  // 끌어 놓기(마우스) — 사람 → 코스, 현장(목록·코스 안) → 코스 또는 현장 목록
  var drag = null;
  root.addEventListener('dragstart', function (e) {
    if (!P.canEdit || !e.target.closest) return;
    var p = e.target.closest('[data-person][draggable="true"]'), l = e.target.closest('[data-drag-lead]');
    if (p) {
      var code = p.getAttribute('data-person');
      drag = { people: P.pp[code] ? checkedPeople() : [code] };     // 체크한 사람을 끌면 체크한 사람 모두
    }
    else if (l) {
      var no = l.getAttribute('data-drag-lead');
      drag = { leads: P.checked[no] ? checkedList() : [no] };     // 체크한 현장을 끌면 체크한 것 모두
    } else return;
    try { e.dataTransfer.setData('text/plain', 'x'); e.dataTransfer.effectAllowed = 'move'; } catch (er) { /* 무시 */ }
  });
  function dropZone(e) {
    var z = e.target.closest && e.target.closest('[data-drop]');
    if (!z || !drag) return null;
    if (drag.people && z.getAttribute('data-drop') !== 'course') return null;
    return z;
  }
  root.addEventListener('dragover', function (e) {
    var z = dropZone(e);
    if (!z) return;
    e.preventDefault();
    root.querySelectorAll('.over, .before').forEach(function (x) { if (x !== z) x.classList.remove('over', 'before'); });
    z.classList.add('over');
    var row = drag.leads && z.getAttribute('data-drop') === 'course' && e.target.closest('.gw-vp-row');
    if (row) row.classList.add('before');                    // 이 줄 앞에 들어간다는 표시
  });
  root.addEventListener('dragleave', function (e) {
    var r = e.target.closest && e.target.closest('.gw-vp-row');
    if (r && !r.contains(e.relatedTarget)) r.classList.remove('before');
    var z = e.target.closest && e.target.closest('[data-drop]');
    if (z && !z.contains(e.relatedTarget)) z.classList.remove('over');
  });
  root.addEventListener('drop', function (e) {
    var z = dropZone(e);
    if (!z) return;
    e.preventDefault();
    var d = drag;
    drag = null;
    z.classList.remove('over');
    var ci = z.getAttribute('data-drop') === 'course' ? parseInt(z.getAttribute('data-ci'), 10) : -1;
    var row = ci !== -1 && e.target.closest('.gw-vp-row');
    root.querySelectorAll('.before').forEach(function (x) { x.classList.remove('before'); });
    if (d.people) putPeople(d.people, ci);
    else moveTo(d.leads, ci, row ? row.getAttribute('data-drag-lead') : null);
  });
  root.addEventListener('dragend', function () {
    drag = null;
    root.querySelectorAll('.over, .before').forEach(function (x) { x.classList.remove('over', 'before'); });
  });

  // ── 불러오기·저장·확정 ──
  function take(j) {
    P.data = j;
    P.plan = JSON.parse(JSON.stringify(j.plan));
    P.canEdit = !!j.can_edit;
    P.names = {}; P.byNo = {};
    j.people.forEach(function (p) { P.names[p.code] = p.name; });
    j.visits.forEach(function (v) { P.byNo[v.lead_no] = v; });
    P.checked = {};
    P.pp = {};
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
    var placed = P.plan.courses.filter(function (c) { return c.members.length && c.leads.length; });
    if (!placed.length) { say('사람과 현장을 넣은 코스가 없어요.', 'err'); return; }
    var people = {}, nSites = 0;
    placed.forEach(function (c) { nSites += c.leads.length; c.members.forEach(function (m) { people[m] = 1; }); });
    var lines = [P.data.day_label + ' 방문 일정을 ' + (P.data.confirmed_at ? '다시 ' : '') + '확정할까요?',
      '코스 ' + placed.length + '개 · 현장 ' + nSites + '건 · 담당자 ' + Object.keys(people).length + '명' +
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
