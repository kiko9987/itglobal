/* 그룹웨어 차량 배정 (2026-11-20) — 다음 날 차마다 탈 사람. 고치기는 배정 담당자(SJ)·인사 관리자만.
   카톡 배정을 대신한다(카톡 글 붙여넣기는 넣지 않음 — 부장님 11-20).
   보드 방식(부장님 11-20): 왼쪽 사람 명단 → 오른쪽 차로 마우스로 끌어 놓기.
   휴대폰·키보드는 사람을 누르고(고르기) → 차를 누른다. 차 안 사람 위에 놓으면 그 앞자리로(맨 앞 = 운전).
   서버에서 온 글자는 모두 esc() 로만 넣는다. */
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
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
          throw new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '저장하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }
  function post(url, body) {
    return api(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  }
  // near = true 면 아래 단추 줄 바로 위에(저장 결과 — 휴대폰에서 위쪽 안내는 화면 밖, 11-21 점검)
  function msg(text, kind, near) {
    var bar = $('vaBarMsg');
    var useBar = near && bar && !$('vaBar').hidden;
    [$('vaMsg'), bar].forEach(function (el) {
      if (!el) return;
      var on = useBar ? el === bar : el.id === 'vaMsg';
      el.className = 'gw-vh-msg' + (el === bar ? ' gw-va-barmsg' : '') + (on && text ? ' ' + (kind || 'info') : '');
      el.textContent = on ? (text || '') : '';
    });
  }

  var TOUCH = !!(window.matchMedia && window.matchMedia('(pointer: coarse)').matches);   // 휴대폰: 끌기 대신 누르기
  var S = { day: '', cars: [], staff: [], names: {}, canEdit: false, confirmed: false, dirty: false,
            visits: {}, places: {}, teams: [], picked: null, seq: 0, focus: null,
            placeTries: 0, placeTimer: null,
            off: {}, half: {}, online: {}, schedule: [], courses: null, plan: null, dayLabel: '',
            noteOpen: {}, doneOpen: true, carWhen: '', draftNote: '' };   // 배정됨은 펼친 채로 시작(11-23 부장님)
  // 화면 정리(11-23 부장님 "너무 어지럽다"): 쓰는 차만 큰 카드·빈 차는 아래 작은 칸, 메모·나누기는 카드 머리 작은 단추,
  //   명단 '배정됨'은 접어 두기, 코스는 한 줄씩(눌러서 현장 펼침), 위 안내는 한 줄로
  // courses = 그룹웨어에서 확정한 방문 일정의 코스(위 칸, 11-23) — 슬랙으로 확정한 날은 null(캔버스1 모양 그대로)
  // online = {이니셜: '온라인 당번(오전)'} 그날 사무실 문의 당번(/일정확정 때 캔버스2 '온라인' 섹션, 11-21)
  // picked = 'YM' 또는 팀 'YM,LSH' · places = {이니셜: [{label, addr}]} 그날 갈 곳(차 카드)
  // off = {이니셜: 까닭} 그날 종일 쉬는 사람, half = 반차(그룹웨어 휴가·토요일 휴무표, 11-21)

  function name(code) { return S.names[code] || code; }
  function carOf(plate) { return S.cars.filter(function (c) { return c.plate === plate; })[0]; }
  function keyCodes(key) { return String(key || '').split(',').filter(Boolean); }
  function keyName(key) { return keyCodes(key).map(name).join(' + '); }
  function anyPlaced(key) { return keyCodes(key).some(function (c) { return whereIs(c); }); }
  // 오전·오후(11-22 부장님): c.people = 오전(나누지 않았으면 종일), c.pm = 오후(null 이면 오전과 같음)
  //   같은 사람은 반나절에 한 차 — 오전 레이·오후 포터처럼 나눠 탈 수 있다
  function isSplit(c) { return !!c.pm; }
  function pmOf(c) { return c.pm ? c.pm : c.people; }
  function ridersOf(c) {
    var out = c.people.slice();
    pmOf(c).forEach(function (p) { if (out.indexOf(p) === -1) out.push(p); });
    return out;
  }
  function whereAt(code, half) {
    for (var i = 0; i < S.cars.length; i++) {
      if ((half === 'pm' ? pmOf(S.cars[i]) : S.cars[i].people).indexOf(code) !== -1) return S.cars[i];
    }
    return null;
  }
  function whereIs(code) { return whereAt(code, 'am') || whereAt(code, 'pm'); }
  function sameList(a, b) { return a.length === b.length && a.every(function (x, i) { return x === b[i]; }); }
  function tidy(c) { if (c.pm && sameList(c.pm, c.people)) c.pm = null; }
  // 한 사람을 그 반나절의 차에서 뺀다. 종일 탄 차에서 반나절만 빼면 그 차는 오전·오후로 나뉜다
  function removeHalf(code, half) {
    S.cars.forEach(function (c) {
      if (half === 'am' && c.people.indexOf(code) !== -1) {
        if (!isSplit(c)) c.pm = c.people.slice();
        c.people.splice(c.people.indexOf(code), 1);
      } else if (half === 'pm' && pmOf(c).indexOf(code) !== -1) {
        if (!isSplit(c)) c.pm = c.people.slice();
        c.pm.splice(c.pm.indexOf(code), 1);
      }
      tidy(c);
    });
  }
  function addDays(day, n) {
    var d = new Date(day + 'T00:00:00');
    d.setDate(d.getDate() + n);
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function setDirty(v) {
    S.dirty = v;
    var b = $('vaSave');
    if (b) b.disabled = !v;                              // 바뀐 게 없으면 [저장만]은 꺼 둔다
    var c = $('vaConfirm');
    if (c && S.confirmed) c.disabled = !v;               // 확정 뒤에는 고친 게 있을 때만
  }
  // 다시 그린 뒤 포커스를 같은 자리로(innerHTML 로 통째 바꿔 body 로 빠지던 것, 11-21 점검)
  function focusAfter(sel) { S.focus = sel; }
  function restoreFocus() {
    var sel = S.focus;
    S.focus = null;
    if (!sel) return;
    var el = document.querySelector(sel);
    if (el) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
  }
  function cssq(v) { return String(v).replace(/["\\]/g, '\\$&'); }
  function staffSel(code) { return '#vaStaffList [data-person="' + cssq(code) + '"]'; }
  function chipSel(code) { return '.gw-va-chip[data-person="' + cssq(code) + '"] .gw-va-cname'; }
  function onlyOne(key) { return keyCodes(key)[0]; }
  function ownCars() {                // 고정 차 주인 → 그 차(YG 카니발·MW 스파크)
    var own = {};
    S.cars.forEach(function (c) { if (c.usage === 'fixed' && c.fixed_user && S.names[c.fixed_user]) own[c.fixed_user] = c; });
    return own;
  }
  function staffNoCar() {             // 방문은 있는데 아직 차 없는 사람(자기 차를 아무도 안 빌렸으면 자기 차로 본다)
    var own = ownCars();
    return Object.keys(S.visits).filter(function (c) {
      return S.names[c] && !S.off[c] && !whereIs(c) && !(own[c] && !ridersOf(own[c]).length);
    }).map(name);
  }
  function riders() {
    var seen = {};
    S.cars.forEach(function (c) { ridersOf(c).forEach(function (p) { seen[p] = 1; }); });
    return Object.keys(seen).length;
  }

  // ── 그리기 ──
  // 명단 이름 옆 = 방문 수와 휴가 꼬리표. 어디 가는지는 위 '그날 방문 일정' 칸과 차 카드에서 본다(11-21 부장님)
  function visitTag(code) {
    var n = S.visits[code];
    var tag = n ? '<small class="gw-va-visit">방문 ' + n + '</small>' : '';
    var leave = S.off[code] || S.half[code];
    var duty = S.online[code];
    return tag + (leave ? '<small class="gw-va-leave">' + esc(leave) + '</small>' : '') +
      (duty ? '<small class="gw-va-leave">' + esc(duty + (n ? '·방문 후' : '')) + '</small>' : '');
  }
  function groupLabels(list) {   // 같은 동(구)끼리 묶기 — [{label, n}] 처음 나온 순서대로
    var at = {}, out = [];
    list.forEach(function (p) {
      if (at[p.label] === undefined) { at[p.label] = out.length; out.push({ label: p.label, n: 0 }); }
      out[at[p.label]].n++;
    });
    return out;
  }
  // 차 카드 = 그 차에 탄 사람들이 갈 곳(동 단위로 묶은 한 줄). 주소·연락처·담당자는 위 '방문 일정' 칸에서 본다
  //   (펼치면 주소마다 같은 이름이 되풀이돼 너무 길었음 — 11-21 부장님)
  function carPlaces(c) {
    var seen = {}, list = [];
    ridersOf(c).forEach(function (code) {
      (S.places[code] || []).forEach(function (p) {
        if (!seen[p.addr]) { seen[p.addr] = 1; list.push(p); }
      });
    });
    if (!list.length) return '';
    var sum = groupLabels(list).map(function (g) { return g.label + (g.n > 1 ? ' ' + g.n + '곳' : ''); }).join(' · ');
    return '<div class="gw-va-places"><i class="fas fa-location-dot" aria-hidden="true"></i><span>' + esc(sum) + '</span></div>';
  }
  function whereText(code) {       // '레이 6854' / '레이 6854 오전 · 포터 4904 오후' / '레이 6854 오전만'
    var am = whereAt(code, 'am'), pm = whereAt(code, 'pm');
    if (!am && !pm) return '';
    if (am && pm && am === pm) return am.label;
    if (am && pm) return am.label + ' 오전 · ' + pm.label + ' 오후';
    return am ? am.label + ' 오전만' : pm.label + ' 오후만';
  }
  function person(code, extra) {   // 왼쪽 명단 한 사람
    var at = whereIs(code);
    var half = at && !(whereAt(code, 'am') && whereAt(code, 'pm'));
    var cls = 'gw-va-person' + (at ? ' placed' : '') + (half ? ' halfonly' : '') + (S.picked === code ? ' picked' : '');
    return '<button type="button" class="' + cls + '" data-person="' + esc(code) + '"' +
      (S.picked === code ? ' aria-pressed="true"' : ' aria-pressed="false"') + '>' +
      '<span class="gw-va-pname">' + esc(name(code)) + '</span>' + visitTag(code) +
      (at ? '<small class="gw-va-where">' + esc(whereText(code)) + '</small>' : '') + (extra || '') + '</button>';
  }
  function offPerson(code) {      // 휴무(종일) — 차에 못 넣음. 고르기·끌기 없이 이름과 꼬리표만
    return '<div class="gw-va-person blocked" aria-disabled="true"><span class="gw-va-pname">' + esc(name(code)) + '</span>' +
      visitTag(code) + '</div>';
  }
  function courseMode() { return S.teams.some(function (t) { return t.label; }); }
  function placeOf(no) {           // 코스 칸 현장 → 지역 꼬리표
    var hit = '';
    (S.courses ? S.courses.courses : []).forEach(function (c) {
      c.sites.forEach(function (s) { if (s.lead_no === no && s.place) hit = s.place; });
    });
    return hit;
  }
  function route(t) {              // 코스가 도는 곳 '화성 동탄구 → 오산 → 평택'(같은 곳 이어지면 한 번)
    var out = [];
    (t.leads || []).forEach(function (n) { var p = placeOf(n); if (p && out[out.length - 1] !== p) out.push(p); });
    return out.join(' → ');
  }
  // 배정됨 — 코스 순서대로 묶어 보기(11-23 부장님). 두 코스에 든 사람은 앞 코스에, 코스에 없는 사람은 맨 아래 '그 밖'
  function placedByCourse(placed) {
    var groups = [], seen = {};
    var src = S.courses ? S.courses.courses.filter(function (c) { return c.kind !== 'work'; })
      : S.teams.filter(function (t) { return t.label; }).map(function (t) { return { label: t.label, members: t.members }; });
    src.forEach(function (c) {
      var list = c.members.filter(function (m) { return placed.indexOf(m) !== -1 && !seen[m]; });
      list.forEach(function (m) { seen[m] = 1; });
      if (list.length) groups.push({ label: c.label, list: list });
    });
    var rest = placed.filter(function (m) { return !seen[m]; });
    if (!groups.length) return '<div class="gw-va-plist">' + placed.map(function (c) { return person(c); }).join('') + '</div>';
    if (rest.length) groups.push({ label: '그 밖', list: rest });
    return groups.map(function (g) {
      return '<div class="gw-va-psub">' + esc(g.label) + '</div><div class="gw-va-plist">' + g.list.map(function (c) { return person(c); }).join('') + '</div>';
    }).join('');
  }
  function staffPanel() {
    var panel = $('vaStaff');
    if (!panel) return;              // 서버 재시작 전 옛 화면 — 명단 칸이 없음
    panel.hidden = !S.canEdit;
    if (!S.canEdit) return;
    var codes = S.staff.map(function (p) { return p.code; });
    // 그날 종일 쉬는 사람(그룹웨어 휴가·토요일 휴무표)은 '배정 없음'에서 빼 '휴무'로 따로 — 차에 넣었으면 '배정됨'에 남는다
    var offList = codes.filter(function (c) { return S.off[c] && !whereIs(c); });
    // 온라인 당번은 사무실에 남아 차가 필요 없다 — 방문도 있는 겸업(방문 다녀와 합류)만 명단에 남긴다
    var dutyList = codes.filter(function (c) { return S.online[c] && !S.visits[c] && !S.off[c] && !whereIs(c); });
    // 사내 업무(방문 일정에서 정한 회사 안 일)도 차가 필요 없다 — 방문도 있으면 명단에 남긴다
    var workList = codes.filter(function (c) { return S.work[c] && !S.visits[c] && !S.off[c] && !whereIs(c) && dutyList.indexOf(c) === -1; });
    var free = codes.filter(function (c) { return !whereIs(c) && !S.off[c] && dutyList.indexOf(c) === -1 && workList.indexOf(c) === -1; });
    // 일정확정 팀(그날 방문을 같이 가는 사람) 중 아직 차 없는 사람이 있는 팀
    var inTeam = {};
    var teams = S.teams.map(function (t) {
      var left = t.members.filter(function (m) { return S.names[m] && free.indexOf(m) !== -1; });
      left.forEach(function (m) { inTeam[m] = 1; });
      return { left: left, visits: t.visits, big: !!t.big, label: t.label || '', memo: t.memo || '', also: t.also || {}, leads: t.leads || [] };
    }).filter(function (t) { return t.left.length; });
    // 함께 가는 팀(2명 이상)을 먼저, 혼자 가는 사람은 그 아래(혼자가 팀 사이에 끼어 있던 것, 11-21 부장님)
    //   같은 쪽 안에서는 서버 순서(방문 많은 순) 그대로. 방문 일정 코스(11-23)는 코스 번호 순서 그대로
    if (!courseMode()) {
      teams = teams.filter(function (t) { return t.left.length > 1; })
        .concat(teams.filter(function (t) { return t.left.length === 1; }));
    }
    // 자기 차(고정 차 주인 — YG 카니발·MW 스파크, 11-22 부장님): 2명 이상 팀에 끼지 않았고, 자기 차를 아무도 빌리지 않았으면
    //   '배정 없음' 대신 '본인 차량'으로(11-22 이름). 팀으로 같이 가면(예: YG 가 4041 팀) 팀 상자에 그대로 둔다
    var own = ownCars();
    var multiTeam = {};
    teams.forEach(function (t) { if (t.left.length > 1) t.left.forEach(function (m) { multiTeam[m] = 1; }); });
    var ownList = free.filter(function (c) { return own[c] && !multiTeam[c] && !ridersOf(own[c]).length; });
    teams = teams.map(function (t) {
      if (t.left.length > 1) return t;
      var u = {}; Object.keys(t).forEach(function (k) { u[k] = t[k]; });
      u.left = t.left.filter(function (m) { return ownList.indexOf(m) === -1; });
      return u;
    }).filter(function (t) { return t.left.length; });
    var freeOther = free.filter(function (c) { return !inTeam[c] && ownList.indexOf(c) === -1; });
    var placed = codes.filter(function (c) { return whereIs(c); });
    function group(title, list, cls) {
      if (!list.length) return '';
      return '<div class="gw-va-group ' + (cls || '') + '"><div class="gw-va-gtitle">' + esc(title) + ' <span>' + list.length + '</span></div>' +
        '<div class="gw-va-plist">' + list.map(function (c) { return person(c); }).join('') + '</div></div>';
    }
    function teamBox(t) {
      var key = t.left.join(',');
      if (t.left.length === 1 && !t.label) return person(key);   // 혼자 가는 사람은 상자 없이(코스는 혼자여도 코스 이름이 보이게)
      var on = S.picked === key;
      // 방문 일정 코스(11-23 부장님) — JW·SJ 가 같은 말로 '코스 1'. 그 밖(슬랙 확정 날)은 '함께'
      var btn = '<button type="button" class="gw-va-teambtn' + (on ? ' picked' : '') + '" data-team="' + esc(key) +
        '" aria-pressed="' + on + '" title="' + (t.left.length > 1 ? '팀 그대로 함께 옮기기' : '옮기기') + '">' +
        '<i class="fas ' + (t.label ? 'fa-route' : 'fa-people-group') + '" aria-hidden="true"></i> ' + esc(t.label || '함께') +
        ' <small>방문 ' + t.visits + '</small></button>';
      var way = t.label ? route(t) : '';
      var info = (way ? '<small class="gw-va-teamroute">' + esc(way) + '</small>' : '') +
        (t.memo ? '<small class="gw-va-teammemo" title="' + esc(t.memo) + '">메모 ' + esc(t.memo) + '</small>' : '');
      // 큰 팀(3명 이상)은 방문을 함께 돈 뒤 현장별로 갈라질 수 있다(11-22 부장님) — 한 차에 다 넣기 전에 한 번 더 보게
      var hint = t.big && t.left.length > 2
        ? '<small class="gw-va-teamhint">함께 방문한 뒤 갈라질 수 있어요. 차를 나눠 넣어 주세요.</small>' : '';
      // 두 코스에 든 사람(오전 코스·오후 코스) — 오전·오후 나눠 탈 수 있게(11-23)
      hint += Object.keys(t.also).filter(function (m) { return t.left.indexOf(m) !== -1; }).map(function (m) {
        return '<small class="gw-va-teamhint">' + esc(name(m)) + ' 님은 ' + esc(t.also[m].join(', ')) +
          '에도 있어요. 차의 [오전·오후 나누기]로 넣어 주세요.</small>';
      }).join('');
      return '<div class="gw-va-team">' + btn + info + hint + t.left.map(function (c) { return person(c); }).join('') + '</div>';
    }
    var nTeam = teams.reduce(function (n, t) { return n + t.left.length; }, 0);
    $('vaStaffList').innerHTML =
      (teams.length ? '<div class="gw-va-group need"><div class="gw-va-gtitle">' + (courseMode() ? '코스' : '일정확정 팀') +
        ' · 배정 없음 <span>' + nTeam + '</span></div>' +
        '<div class="gw-va-plist">' + teams.map(teamBox).join('') + '</div></div>' : '') +
      group('배정 없음', freeOther) +
      (ownList.length ? '<div class="gw-va-group own"><div class="gw-va-gtitle">본인 차량 <span>' + ownList.length + '</span></div>' +
        '<div class="gw-va-plist">' + ownList.map(function (c) {
          return person(c, '<small class="gw-va-where">' + esc(own[c].label) + '</small>');
        }).join('') + '</div></div>' : '') +
      (placed.length ? '<details class="gw-va-group done" id="vaDoneGroup"' + (S.doneOpen ? ' open' : '') + '><summary class="gw-va-gtitle">' +
        '<i class="fas fa-chevron-right gw-va-gfold" aria-hidden="true"></i> 배정됨 <span>' + placed.length + '</span></summary>' +
        placedByCourse(placed) + '</details>' : '') +
      group('온라인 당번', dutyList, 'off') +
      (workList.length ? '<div class="gw-va-group off"><div class="gw-va-gtitle">사내 업무 <span>' + workList.length + '</span></div>' +
        '<div class="gw-va-plist">' + workList.map(function (c) {
          return person(c, '<small class="gw-va-where">' + esc(S.work[c]) + '</small>');
        }).join('') + '</div></div>' : '') +
      (offList.length ? '<div class="gw-va-group off"><div class="gw-va-gtitle">휴무 <span>' + offList.length + '</span></div>' +
        '<div class="gw-va-plist">' + offList.map(offPerson).join('') + '</div></div>' : '');
    var how = $('vaStaffHow');
    how.className = 'gw-va-staff-how' + (S.picked ? ' picking' : '');
    how.textContent = S.picked
      ? keyName(S.picked) + ' 님을 넣을 차에서 [여기에 … 넣기]를 누르세요.' + (anyPlaced(S.picked) ? ' 명단 빈 곳을 누르면 차에서 빠져요.' : '') +
        (TOUCH ? '' : ' Esc 를 누르면 취소돼요.')
      : (TOUCH ? '이름이나 ' + (courseMode() ? '[코스]' : '팀의 [함께]') + '를 누른 뒤 넣을 차를 누르세요. 운전한 사람은 그날 저녁 슬랙으로 물어봐요.'
         : '이름이나 ' + (courseMode() ? '[코스]' : '팀의 [함께]') + '를 끌어 차에 놓으세요. 누른 뒤 차를 눌러도 돼요. 운전한 사람은 그날 저녁 슬랙으로 물어봐요.');
    $('vaStaffCount').textContent = nTeam ? '방문 있는데 배정 없음 ' + nTeam + '명' : '';
    $('vaStaffCount').className = nTeam ? 'gw-va-warn' : '';
  }
  // 칩 = 이름 단추 + [×] 단추 나란히(누를 수 있는 칩 안에 또 단추가 들어 있던 것, 11-21 점검)
  function chip(c, code, i, half) {   // half = 'all'(나누지 않음) / 'am' / 'pm'
    half = half || 'all';
    var lead = '';   // 운전은 배정 때 정하지 않는다 — 그날 저녁 슬랙 DM 으로 실제 운전자를 받음(11-22 부장님)
    var cls = 'gw-va-chip' + (S.picked === code ? ' picked' : '') + (S.off[code] ? ' leave' : '');
    if (!S.canEdit) return '<span class="' + cls + '">' + lead + esc(name(code)) + '</span>';
    var hl = half === 'am' ? ' 오전' : half === 'pm' ? ' 오후' : '';
    var x = '<button type="button" class="gw-va-x" data-rm="' + esc(code) + '" data-plate="' + esc(c.plate) + '" data-half="' + half +
      '" aria-label="' + esc(c.label + hl) + '에서 ' + esc(name(code)) + ' 빼기"><i class="fas fa-xmark" aria-hidden="true"></i></button>';
    return '<span class="' + cls + '" data-person="' + esc(code) + '" data-in="' + esc(c.plate) + '" data-half="' + half + '">' + lead +
      '<button type="button" class="gw-va-cname" aria-pressed="' + (S.picked === code) + '" aria-label="' +
      esc(name(code) + ' — 눌러서 옮기기') + '">' + esc(name(code)) + '</button>' + x + '</span>';
  }
  function fixedSelect(c) {   // 고정 차의 평소 사용자 — 고르면 바로 저장(배정과 별개)
    var has = false;
    var opts = S.staff.map(function (p) {
      if (p.code === c.fixed_user) has = true;
      return '<option value="' + esc(p.code) + '"' + (p.code === c.fixed_user ? ' selected' : '') + '>' + esc(p.name) + '</option>';
    }).join('');
    if (c.fixed_user && !has) opts = '<option value="' + esc(c.fixed_user) + '" selected>' + esc(c.fixed_name || c.fixed_user) + '</option>' + opts;
    return '<select class="form-select form-select-sm gw-va-fixedsel" data-fixed="' + esc(c.plate) + '" aria-label="' +
      esc(c.label) + ' 평소 사용자"><option value="">정해진 사람 없음</option>' + opts + '</select>';
  }
  function card(c) {
    var sub = [c.plate, c.fuel].filter(Boolean).join(' · ');
    var fixed = '';
    if (c.usage === 'fixed') {
      fixed = '<div class="gw-va-fixed">평소: ' + (S.canEdit ? fixedSelect(c)
        : (c.fixed_name ? esc(c.fixed_name) : '<span class="gw-va-none">정해진 사람 없음</span>')) + '</div>';
    }
    // 한 칸(종일 또는 오전·오후 한쪽) = 칩 + [여기에 넣기] + [운전] 고르기
    function zone(list, half) {
      var chips = list.map(function (p, i) { return chip(c, p, i, half); }).join('');
      var hl = half === 'am' ? ' 오전' : half === 'pm' ? ' 오후' : '';
      var put = S.canEdit && S.picked
        ? '<button type="button" class="btn btn-sm btn-outline-secondary gw-va-put" data-put="' + esc(c.plate) + '" data-half="' + half +
          '" aria-label="' + esc(c.label + hl) + '에 ' + esc(keyName(S.picked)) + ' 넣기"><i class="fas fa-plus" aria-hidden="true"></i> 여기에 ' +
          esc(keyName(S.picked)) + ' 넣기</button>'
        : '';
      if (!chips && !put) {
        chips = '<span class="gw-va-none">' + (S.canEdit ? '여기에 사람을 놓으세요'
          : (c.usage === 'fixed' ? '평소대로' : '배정 없음')) + '</span>';
      }
      // 운전할 사람은 고르지 않는다(출발할 때 동행자끼리 정함 → 저녁 DM 으로 받음, 11-22 부장님)
      var drv = false && list.length > 1
        ? '<label class="gw-va-driver"><span>운전</span><select class="form-select form-select-sm" data-driver="' + esc(c.plate) +
          '" data-half="' + half + '" aria-label="' + esc(c.label + hl) + ' 운전할 사람">' + list.map(function (p, i) {
            return '<option value="' + esc(p) + '"' + (i === 0 ? ' selected' : '') + '>' + esc(name(p)) + '</option>';
          }).join('') + '</select></label>'
        : '';
      return '<div class="gw-va-people">' + chips + put + '</div>' + drv;
    }
    var body;
    if (isSplit(c)) {
      body = ['am', 'pm'].map(function (h) {
        return '<div class="gw-va-half"' + (S.canEdit ? ' data-drop="car" data-plate="' + esc(c.plate) + '" data-half="' + h + '"' : '') + '>' +
          '<div class="gw-va-halfhead">' + (h === 'am' ? '오전' : '오후') + '</div>' + zone(h === 'am' ? c.people : c.pm, h) + '</div>';
      }).join('');
    } else {
      body = zone(c.people, 'all');
    }
    // 메모는 적었거나 [메모]를 눌렀을 때만 칸이 보인다(늘 붙어 있어 카드가 길었음, 11-23)
    var showNote = S.canEdit && (c.note || S.noteOpen[c.plate]);
    var note = showNote
      ? '<input type="text" class="form-control form-control-sm gw-va-note" maxlength="60" placeholder="메모" data-note="' +
        esc(c.plate) + '" value="' + esc(c.note || '') + '" aria-label="' + esc(c.label) + ' 메모" autocomplete="off">'
      : (c.note ? '<div class="gw-va-notetext">' + esc(c.note) + '</div>' : '');
    var tools = '';
    if (S.canEdit) {
      if (!showNote) tools += '<button type="button" class="gw-va-tool" data-noteopen="' + esc(c.plate) + '" aria-label="' + esc(c.label) + ' 메모 쓰기">' +
        '<i class="fas fa-pen" aria-hidden="true"></i> 메모</button>';
      if (isSplit(c)) tools += '<button type="button" class="gw-va-tool" data-unsplit="' + esc(c.plate) + '">오후도 오전과 같게</button>';
      else if (c.people.length) tools += '<button type="button" class="gw-va-tool" data-split="' + esc(c.plate) + '">오전·오후 나누기</button>';
    }
    var drop = S.canEdit && !isSplit(c) ? ' data-drop="car" data-half="all"' : '';
    return '<div class="gw-va-card' + (S.picked ? ' can-drop' : '') + (isSplit(c) ? ' split' : '') + '" data-plate="' + esc(c.plate) + '"' + drop + '>' +
      '<div class="gw-va-cardhead"><div class="gw-va-car"><strong>' + esc(c.label) + '</strong><span>' + esc(sub) + '</span></div>' +
      (tools ? '<div class="gw-va-tools">' + tools + '</div>' : '') + '</div>' + fixed +
      (S.courses ? '' : carPlaces(c)) + body + note + '</div>';
  }
  // 빈 차 — 작은 칸. 사람을 놓으면 위 큰 카드로 올라간다(11-23)
  function tile(c) {
    var fixed = c.usage === 'fixed'
      ? '<div class="gw-va-tilefixed">평소 ' + (S.canEdit ? fixedSelect(c) : esc(c.fixed_name || '없음')) + '</div>' : '';
    var put = S.canEdit && S.picked ? '<span class="gw-va-tileput"><i class="fas fa-plus" aria-hidden="true"></i> 넣기</span>' : '';
    return '<div class="gw-va-tile' + (S.picked ? ' can-drop' : '') + '" data-plate="' + esc(c.plate) + '"' +
      (S.canEdit ? ' data-drop="car" data-half="all"' : '') + '>' +
      '<div class="gw-va-car"><strong>' + esc(c.label) + '</strong><span>' + esc(c.plate) + '</span></div>' + fixed + put + '</div>';
  }
  // ── 그날 방문 일정 칸(차 목록 위, 11-21 부장님) ──
  // 캔버스1(봇 자동)과 같은 묶음·표시를 리드 시트에서 바로 그린다. 줄마다 /일정확정으로 들어간 영업 담당자.
  // 맨 위 한 줄 = 그날 쉬는 사람(그룹웨어 휴가·토요일 휴무표). 펼침 상태는 이 브라우저에 기억(PC 펼침·휴대폰 접힘이 기본)
  var SCHED_KEY = 'gw.va.schedOpen';
  function schedOpenPref() {
    try { var v = localStorage.getItem(SCHED_KEY); if (v === '1' || v === '0') return v === '1'; } catch (e) { /* 무시 */ }
    return !(window.matchMedia && window.matchMedia('(max-width: 767.98px)').matches);
  }
  function schedEl() {
    var el = $('vaSched');
    if (el) return el;
    var host = document.querySelector('.gw-va-cars');     // 서버 재시작 전 옛 화면엔 칸이 없어 만들어 넣는다
    if (!host) return null;
    el = document.createElement('details');
    el.className = 'gw-va-sched';
    el.id = 'vaSched';
    host.insertBefore(el, host.firstChild);
    return el;
  }
  function schedItem(it) {
    // 플랫폼 꼬리표(줄마다, 테두리만) + 본인 방문 필수·공사확정(채움) — 11-21 부장님
    var marks = (it.platform ? '<span class="gw-va-smark plat">' + esc(it.platform) + '</span>' : '') +
      it.marks.map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('');
    var who = it.who.length
      ? '<span class="gw-va-swho">' + esc(it.who.join(', ')) + '</span>'
      : '<span class="gw-va-swho none">배분 전</span>';
    var phone = it.phone
      ? '<a href="tel:' + esc(it.phone.replace(/[^\d+]/g, '')) + '">' + esc(it.phone) + '</a>' : '';
    return '<li class="gw-va-sitem">' +
      '<div class="gw-va-sline">' + marks + '<strong>' + esc(it.date) + '</strong>' + (phone ? '<span>' + phone + '</span>' : '') +
        (it.lead_no ? '<span class="gw-va-sno">' + esc(it.lead_no) + '</span>' : '') + who + '</div>' +
      (it.addr ? '<div class="gw-va-saddr">' + esc(it.addr) + '</div>' : '') +
      (it.text ? '<div class="gw-va-stext">' + esc(it.text) + '</div>' : '') + '</li>';
  }
  // ── 그룹웨어에서 확정한 날은 같은 칸을 코스로(11-23 부장님) — 코스마다 사람·지금 넣은 차·메모·현장 ──
  function courseSite(s) {
    if (s.missing) {
      return '<li class="gw-va-sitem"><div class="gw-va-sline"><span class="gw-va-sno">' + esc(s.lead_no) +
        '</span><span class="gw-va-swho none">그날 현장 목록에 없어요(날짜가 바뀌었을 수 있어요)</span></div></li>';
    }
    var marks = (s.platform ? '<span class="gw-va-smark plat">' + esc(s.platform) + '</span>' : '') +
      (s.marks || []).map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('') +
      (s.contact ? '<span class="gw-va-smark plat">' + esc(s.contact + (s.must ? ' · 본인 방문 필수' : '')) + '</span>' : '');
    var phone = s.phone ? '<a href="tel:' + esc(s.phone.replace(/[^\d+]/g, '')) + '">' + esc(s.phone) + '</a>' : '';
    return '<li class="gw-va-sitem"><div class="gw-va-sline">' + (s.place ? '<strong>' + esc(s.place) + '</strong>' : '') + marks +
      (phone ? '<span>' + phone + '</span>' : '') + '<span class="gw-va-sno">' + esc(s.lead_no) + '</span></div>' +
      (s.addr ? '<div class="gw-va-saddr">' + esc(s.addr) + '</div>' : '') +
      (s.text ? '<div class="gw-va-stext">' + esc(s.text) + '</div>' : '') + '</li>';
  }
  function courseCars(members) {   // 코스 사람들이 지금 넣어 둔 차 — 차 없는 사람은 주황
    var cars = [], none = [];
    var own = ownCars();
    members.forEach(function (m) {
      if (S.off[m]) return;
      var w = whereText(m);
      // 고정 차 주인(MW 스파크·YG 카니발)이 자기 차를 아무도 안 빌렸으면 본인 차로 간 것(11-23)
      if (!w && own[m] && !ridersOf(own[m]).length) w = own[m].label + ' (본인)';
      if (!w) none.push(name(m)); else if (cars.indexOf(w) === -1) cars.push(w);
    });
    return cars.map(function (w) { return '<span class="gw-va-ccar"><i class="fas fa-car" aria-hidden="true"></i> ' + esc(w) + '</span>'; }).join('') +
      (none.length ? '<span class="gw-va-ccar none">차 없음 ' + esc(none.join(', ')) + '</span>' : '');
  }
  function refreshCourseCars() {
    document.querySelectorAll('[data-ccars]').forEach(function (el) {
      el.innerHTML = courseCars(String(el.getAttribute('data-ccars')).split(',').filter(Boolean));
    });
  }
  function siteRoute(sites) {
    var out = [];
    sites.forEach(function (s) { var p = s.place; if (p && out[out.length - 1] !== p) out.push(p); });
    return out.join(' → ');
  }
  function courseBody() {
    var C = S.courses;
    var body = C.courses.map(function (c) {
      var who = c.members.map(name).join(' + ') || '사람 없음';
      if (c.kind === 'work') {
        return '<div class="gw-va-cline work"><div class="gw-va-csum"><span class="gw-va-clabel">' + esc(c.label) + '</span>' +
          '<span class="gw-va-cwho">' + esc(who) + '</span>' + (c.task ? '<span class="gw-va-croute">' + esc(c.task) + '</span>' : '') +
          '<span class="gw-va-ccars"><span class="gw-va-ccar">차 필요 없음</span></span></div></div>';
      }
      var route = siteRoute(c.sites);
      return '<details class="gw-va-cline"><summary class="gw-va-csum">' +
        '<i class="fas fa-chevron-right gw-va-cfold" aria-hidden="true"></i><span class="gw-va-clabel">' + esc(c.label) + '</span>' +
        '<span class="gw-va-cwho">' + esc(who) + '</span>' + (route ? '<span class="gw-va-croute">' + esc(route) + '</span>' : '') +
        '<span class="gw-va-ccars" data-ccars="' + esc(c.members.join(',')) + '"></span>' +
        (c.memo ? '<span class="gw-va-cmemoline" title="' + esc(c.memo) + '"><strong>메모</strong> ' + esc(c.memo) + '</span>' : '') +
        '</summary><ul>' + c.sites.map(courseSite).join('') + '</ul></details>';
    }).join('');
    if (C.left.length) {
      body += '<details class="gw-va-cline left"><summary class="gw-va-csum">' +
        '<i class="fas fa-chevron-right gw-va-cfold" aria-hidden="true"></i><span class="gw-va-clabel">코스에 없는 현장</span>' +
        '<span class="gw-va-cwho">' + C.left.length + '건</span><span class="gw-va-croute">' + esc(siteRoute(C.left)) + '</span>' +
        '</summary><ul>' + C.left.map(courseSite).join('') + '</ul></details>';
    }
    return body || '<p class="gw-va-sempty">이날 방문 일정이 없어요.</p>';
  }
  function schedPanel(j) {
    var el = schedEl();
    if (!el) return;
    if (!('schedule' in j)) { el.hidden = true; return; }   // 서버 재시작 전(일정 칸 자료 없음) — '방문 0건'으로 보이지 않게
    var n = S.schedule.reduce(function (a, g) { return a + g.items.length; }, 0);
    var todo = S.schedule.reduce(function (a, g) { return a + g.items.filter(function (it) { return !it.who.length; }).length; }, 0);
    var offs = (j.off || []).map(function (o) { return o.name + '(' + o.why + ')'; });
    var halves = (j.half || []).map(function (o) { return o.name + '(' + o.why + ')'; });
    var duties = (j.online || []).map(function (o) { return o.name + (o.shift ? '(' + o.shift + ')' : ''); });
    var sum = '방문 ' + n + '건' + (todo ? ' · 배분 전 ' + todo + '건' : '') + (offs.length ? ' · 휴무 ' + offs.length + '명' : '');
    var C = S.courses;
    if (C) {                         // 코스로 볼 때 — 방문 수는 코스 현장 + 코스에 없는 현장
      var nv = C.courses.filter(function (c) { return c.kind !== 'work'; });
      var ns = nv.reduce(function (a, c) { return a + c.sites.length; }, 0) + C.left.length;
      sum = '코스 ' + nv.length + '개 · 방문 ' + ns + '건' + (C.left.length ? ' · 코스에 없는 현장 ' + C.left.length + '건' : '') +
        (offs.length ? ' · 휴무 ' + offs.length + '명' : '');
    }
    var bits = [];
    if (offs.length) bits.push('<strong>휴무</strong> ' + esc(offs.join(', ')));
    if (halves.length) bits.push('<strong>반차</strong> ' + esc(halves.join(', ')));
    if (duties.length) bits.push('<strong>온라인 당번</strong> ' + esc(duties.join(', ')));
    var leave = bits.length ? '<p class="gw-va-soff">' + bits.join(' · ') + '</p>' : '';
    var body = C ? courseBody() : (S.schedule.map(function (g) {
      return '<div class="gw-va-sgroup"><h3>' + esc(g.title) + ' <span>' + g.items.length + '</span></h3>' +
        '<ul>' + g.items.map(schedItem).join('') + '</ul></div>';
    }).join('') || '<p class="gw-va-sempty">이날 방문 일정이 없어요.</p>');
    var wasOpen = el.dataset.ready ? el.open : schedOpenPref();
    el.innerHTML = '<summary><i class="fas fa-calendar-day" aria-hidden="true"></i><span class="gw-va-stitle">' +
      esc(j.day_label || '') + ' 방문 일정</span><span class="gw-va-ssum">' + esc(sum) + '</span>' +
      '<i class="fas fa-chevron-down gw-va-fold" aria-hidden="true"></i></summary>' +
      '<div class="gw-va-sbody' + (C ? ' courses' : '') + '">' + leave + body + '</div>';
    el.open = wasOpen;
    el.hidden = false;
    el.dataset.ready = '1';
    refreshCourseCars();
  }
  function render() {
    var used = S.cars.filter(function (c) { return ridersOf(c).length || c.note; });
    var empty = S.cars.filter(function (c) { return used.indexOf(c) === -1; });
    $('vaShared').innerHTML = used.map(card).join('') ||
      '<p class="gw-va-none gw-va-noused">' + (S.canEdit ? '아직 넣은 차가 없어요. 아래 빈 차에 사람을 놓으세요.' : '이날 차량 배정이 없어요.') + '</p>';
    $('vaFixedHead').innerHTML = '빈 차 <span class="gw-va-sec-n">' + empty.length + '대</span>' +
      (S.canEdit ? '<span class="gw-va-sec-sub">사람을 놓으면 위로 올라가요. 고정 차는 평소엔 정해진 사람이 타요.</span>' : '');
    $('vaFixedHead').hidden = !empty.length;
    $('vaFixed').className = 'gw-va-tiles';
    $('vaFixed').innerHTML = empty.map(tile).join('');
    staffPanel();
    if ($('vaBoard')) $('vaBoard').classList.toggle('editing', S.canEdit);
    $('vaBar').hidden = !S.canEdit;
    // 확정 뒤에도 '알리지 않고 저장'을 남긴다 — 메모 오타 하나에 그 차 사람 모두에게 DM 이 가던 것(11-21 점검)
    $('vaConfirm').textContent = S.confirmed ? '저장하고 바뀐 사람에게 알리기' : '확정하고 알리기';
    $('vaSave').textContent = S.confirmed ? '알리지 않고 저장' : '저장만';
    $('vaSave').hidden = false;
    setDirty(S.dirty);
    refreshCourseCars();
    statusLine();
    restoreFocus();
  }
  // 방문 일정(영업 > 방문 일정)을 그룹웨어에서 확정했으면 한 줄 — 팀이 어디서 왔는지(11-22 부장님)
  function courseMembers() {
    var seen = {};
    S.teams.forEach(function (t) {
      t.members.forEach(function (m) { seen[m] = 1; });
      Object.keys(t.also || {}).forEach(function (m) { seen[m] = 1; });
    });
    return seen;
  }
  // 차량 배정을 저장한 뒤 방문 일정이 다시 확정됐고, 지금 배정과 코스가 어긋나면(차 없는 코스 사람·코스에서 빠진 사람) 경고 글
  function staleText() {
    var p = S.plan;
    if (!S.canEdit || !p || !p.confirmed_at || !p.cars_saved_at || !(p.confirmed_at > p.cars_saved_at)) return '';
    var inCourse = courseMembers();
    var none = staffNoCar();
    var out = [];
    S.cars.forEach(function (c) { ridersOf(c).forEach(function (m) { if (!inCourse[m] && !S.visits[m] && out.indexOf(name(m)) === -1) out.push(name(m)); }); });
    if (!none.length && !out.length) return '';
    return '차량 배정 뒤에 방문 일정이 다시 확정됐어요.' + (none.length ? ' 차 없음: ' + none.join(', ') + '.' : '') +
      (out.length ? ' 코스에서 빠짐: ' + out.join(', ') + '.' : '');
  }
  function status(j) {           // 불러올 때 — 한 줄에 쓸 것만 담아 둔다
    S.carWhen = j.confirmed_at ? j.confirmed_at.slice(5, 16).replace('-', '/') + (j.confirmed_name ? ' · ' + j.confirmed_name : '') : '';
    var drafted = j.cars.some(function (c) { return c.people.length; });
    S.draftNote = '';
    if (j.draft_kind === 'teams' && drafted) S.draftNote = (courseMode() ? '코스대로' : '일정확정 팀대로') + ' 최근 탄 차에 미리 넣어 두었어요 — 바뀐 곳만 고쳐 확정해 주세요';
    else if (j.draft_from) S.draftNote = j.draft_from.slice(5).replace('-', '/') + ' 배정을 미리 넣어 두었어요 — 바뀐 곳만 고쳐 확정해 주세요';
    statusLine();
  }
  function statusLine() {        // 날짜 · 확정 상태 · 방문 일정 · 초안 안내 · (경고) — 한 줄(11-23 부장님: 위 안내가 세 줄이라 어지러움)
    var el = $('vaStatus');
    if (!el) return;
    var pl = $('vaPlanLine');
    if (pl) { pl.hidden = true; pl.innerHTML = ''; }
    var head = '<strong>' + esc(S.dayLabel) + '</strong> ' + (S.confirmed
      ? '<span class="gw-badge gw-badge-ok">확정</span>' + (S.carWhen ? ' <span class="gw-va-when">' + esc(S.carWhen) + '</span>' : '')
      : '<span class="gw-badge gw-badge-wait">확정 전</span>');
    var bits = [];
    var p = S.plan;
    if (p && p.confirmed_at) {
      var t = '방문 일정 확정 ' + p.confirmed_at.slice(11, 16) + (p.confirmed_name ? ' · ' + p.confirmed_name : '');
      bits.push('<i class="fas fa-route" aria-hidden="true"></i> ' +
        (p.can_open ? '<a href="/visits/plan?day=' + encodeURIComponent(S.day) + '">' + esc(t) + '</a>' : esc(t)));
    } else if (S.courses && S.courses.source === 'sheet') {
      bits.push('<i class="fas fa-route" aria-hidden="true"></i> 슬랙으로 정한 날 — 시트 담당자대로 코스');
    }
    if (S.draftNote && S.canEdit && !S.confirmed) bits.push(esc(S.draftNote));
    var warn = staleText();
    el.innerHTML = head + (bits.length ? '<span class="gw-va-sinfo">' + bits.join('<span class="gw-va-dot">·</span>') + '</span>' : '') +
      (warn ? '<span class="gw-va-stale" role="status">' + esc(warn) + '</span>' : '');
  }

  // ── 불러오기·저장 ──
  function load(day) {
    msg('');
    var my = ++S.seq;     // ◀ ▶ 를 빨리 누르면 늦게 온 앞 날짜 응답이 덮지 않게
    return api('/api/vehicles/assign' + (day ? '?day=' + encodeURIComponent(day) : '')).then(function (j) {
      if (my !== S.seq) return;
      S.day = j.day; S.cars = j.cars; S.staff = j.staff; S.canEdit = !!j.can_edit;
      S.cars.forEach(function (c) { c.pm = Array.isArray(c.pm) ? c.pm : null; });
      S.confirmed = !!j.confirmed_at; S.picked = null;
      S.names = {}; S.visits = {}; S.places = {}; S.off = {}; S.half = {}; S.online = {}; S.work = {};
      S.noteOpen = {};
      j.staff.forEach(function (p) { S.names[p.code] = p.name; });
      // 확정 전 토요 휴무는 바뀔 수 있어 반차처럼 꼬리표만(방문 일정과 같게 — 코스에 넣을 수 있음, 11-23)
      (j.off || []).forEach(function (o) { if (o.soft) S.half[o.code] = o.why; else S.off[o.code] = o.why; });
      (j.half || []).forEach(function (o) { S.half[o.code] = o.why; });
      (j.online || []).forEach(function (o) { S.online[o.code] = '온라인 당번' + (o.shift ? '(' + o.shift + ')' : ''); });
      (j.work || []).forEach(function (o) { S.work[o.code] = o.task; });     // 방문 일정의 사내 업무(회사 안 일, 11-23)
      S.schedule = j.schedule || [];
      S.courses = j.courses || null;
      S.plan = j.plan || null;
      S.dayLabel = j.day_label || '';
      (j.visitors || []).forEach(function (v) { S.visits[v.code] = v.visits; S.places[v.code] = v.places || []; });
      S.teams = j.teams || [];
      schedPanel(j);
      S.dirty = !!(j.draft_kind && S.canEdit && j.cars.some(function (c) { return c.people.length; }));
      $('vaDay').value = j.day;
      try { history.replaceState(null, '', '/vehicles/assign?day=' + j.day); } catch (e) { /* 무시 */ }
      status(j);
      render();
      S.placeTries = 0;
      placesLater(j.resolving);
    }).catch(function (e) { if (my === S.seq) msg(e.message, 'err'); });
  }
  // 갈 곳의 동을 서버가 뒤에서 찾는 중이면 잠시 뒤 갈 곳만 다시 받는다(배정 고친 것은 그대로)
  function placesLater(resolving) {
    clearTimeout(S.placeTimer);
    if (!resolving || S.placeTries >= 3) return;
    var my = S.seq, day = S.day;
    S.placeTimer = setTimeout(function () {
      S.placeTries++;
      api('/api/vehicles/assign?day=' + encodeURIComponent(day)).then(function (j) {
        if (my !== S.seq || j.day !== S.day) return;
        S.places = {};
        (j.visitors || []).forEach(function (v) { S.places[v.code] = v.places || []; });
        var act = document.activeElement;
        // 메모를 쓰는 중이거나 끄는 중이면 다시 그리지 않는다 — 다음에 그릴 때 들어감
        var inBoard = act && act.closest && act.closest('.gw-va-board');
        if (!D && !(inBoard && act.matches('input, select'))) {
          if (inBoard && act.dataset.team) focusAfter('[data-team="' + cssq(act.dataset.team) + '"]');
          else if (inBoard && act.closest('.gw-va-chip[data-person]') && act.matches('.gw-va-cname')) focusAfter(chipSel(act.closest('.gw-va-chip').dataset.person));
          else if (inBoard && act.dataset.person) focusAfter(staffSel(act.dataset.person));          render();
        }
        placesLater(j.resolving);
      }).catch(function () { /* 다음에 불러올 때 */ });
    }, 15000);
  }
  function rowsForSave() {
    return S.cars.map(function (c) { return { plate: c.plate, people: c.people, pm: c.pm || null, note: c.note || '' }; });
  }
  // notify: 슬랙으로 알릴지. 확정 전 [저장만]·확정 뒤 [알리지 않고 저장]은 false
  function save(notify) {
    var label = $('vaStatus').querySelector('strong');
    var dayText = label ? label.textContent : S.day;
    if (notify && !S.confirmed) {
      var miss = staffNoCar();
      if (!window.confirm(dayText + ' 배정을 확정하고, 차에 넣은 ' + riders() + '명에게 슬랙으로 알릴까요?' +
          (miss.length ? '\n\n방문이 있는데 아직 배정 없는 사람: ' + miss.join(', ') : '') +
          '\n\n보낸 알림은 되돌릴 수 없어요.')) return;
    } else if (notify) {
      if (!window.confirm('고친 내용을 저장하고, 차나 같이 타는 사람이 바뀐 사람에게만 슬랙으로 알릴까요?')) return;
    }
    var btns = [$('vaSave'), $('vaConfirm')];
    btns.forEach(function (b) { b.disabled = true; });
    msg(notify ? '저장하고 알리는 중…' : '저장하는 중…', 'info', true);
    post('/api/vehicles/assign', { day: S.day, rows: rowsForSave(), confirm: !!notify && !S.confirmed, notify: !!notify })
      .then(function (j) {
        S.dirty = false;
        var t = j.changed.length ? '저장했어요.' : '바뀐 내용이 없어요.';
        if (j.notified) t += ' ' + j.notified + '명에게 슬랙으로 알렸어요.';
        else if (notify && j.changed.length) t += ' 슬랙으로 알릴 사람이 없었어요.';
        return load(S.day).then(function () { msg(t, 'ok', true); });
      }).catch(function (e) { msg(e.message, 'err', true); })
      .then(function () { btns.forEach(function (b) { b.disabled = false; }); setDirty(S.dirty); });
  }

  // ── 옮기기 ──
  function unplace(code) {         // 하루 통째로 뺀다(명단으로 끌어 놓을 때)
    var at = whereIs(code);
    removeHalf(code, 'am');
    removeHalf(code, 'pm');
    return at;
  }
  // half = 'all'(나누지 않은 차 — 종일) / 'am' / 'pm'. before = 그 사람 앞자리로(없으면 맨 뒤)
  function place(code, plate, before, half) {
    half = half || 'all';
    if (code === before || S.off[code]) return;     // 휴무(종일)는 차에 못 넣음(11-23)
    var c = carOf(plate);
    if (half === 'all' && isSplit(c)) half = 'am';
    if (half === 'all') { removeHalf(code, 'am'); removeHalf(code, 'pm'); }
    else removeHalf(code, half);
    var list = half === 'pm' ? (c.pm || (c.pm = c.people.slice())) : c.people;
    var i = before ? list.indexOf(before) : -1;
    if (i === -1) list.push(code); else list.splice(i, 0, code);
    tidy(c);
    setDirty(true);
  }
  function placeMany(key, plate, before, half) {   // 팀이면 적힌 순서대로 같은 자리에
    keyCodes(key).forEach(function (c) { place(c, plate, before, half); });
  }
  function pick(code, fromEl) {
    S.picked = S.picked === code ? null : code;
    // 누른 단추로 포커스를 돌려준다(명단·팀·차 안 칩 어디서 눌렀든)
    if (fromEl && fromEl.closest('[data-team]')) focusAfter('[data-team="' + cssq(code) + '"]');
    else if (fromEl && fromEl.closest('.gw-va-chip')) focusAfter(chipSel(code));
    else focusAfter(staffSel(code));
    render();      // 안내는 명단 위 한 줄에서 바뀐다(위에 띄우면 화면이 밀려 다음 누르기가 빗나감)
  }
  function drop(key, target) {   // key = 사람 또는 팀('YM,LSH'), target = 놓은 자리의 요소
    var chipEl = target && target.closest('.gw-va-chip[data-person]');
    var cardEl = target && target.closest('[data-drop="car"]');
    var outEl = target && target.closest('[data-drop="out"]');
    S.picked = null;
    if (cardEl) {
      var half = cardEl.dataset.half || 'all';
      var before = chipEl && (chipEl.dataset.half || 'all') === half ? chipEl.dataset.person : null;
      placeMany(key, cardEl.dataset.plate, before, half);
      focusAfter(chipSel(onlyOne(key)));
    } else if (outEl && anyPlaced(key)) {
      keyCodes(key).forEach(unplace);
      setDirty(true);
      focusAfter(staffSel(onlyOne(key)));
    }
    msg('');
    render();
  }

  // 마우스 끌기 — 손가락(터치)은 스크롤과 겹쳐서 끌지 않고 누르기로 한다
  var D = null;
  document.addEventListener('pointerdown', function (e) {
    if (!S.canEdit || e.pointerType !== 'mouse' || e.button !== 0) return;
    var p = e.target.closest('[data-team], [data-person]');
    if (!p || e.target.closest('[data-rm]')) return;
    D = { code: p.dataset.team || p.dataset.person, x: e.clientX, y: e.clientY, ghost: null, over: null };
  });
  document.addEventListener('pointermove', function (e) {
    if (!D) return;
    if (!D.ghost) {
      if (Math.abs(e.clientX - D.x) + Math.abs(e.clientY - D.y) < 6) return;
      D.ghost = document.createElement('div');
      D.ghost.className = 'gw-va-ghost';
      D.ghost.textContent = keyName(D.code);
      document.body.appendChild(D.ghost);
      document.body.classList.add('gw-va-dragging');
    }
    D.ghost.style.left = e.clientX + 'px';
    D.ghost.style.top = e.clientY + 'px';
    var t = document.elementFromPoint(e.clientX, e.clientY);
    var over = t && (t.closest('.gw-va-chip[data-person]') || t.closest('[data-drop]'));
    if (over !== D.over) {
      if (D.over) D.over.classList.remove('over');
      if (over) over.classList.add('over');
      D.over = over;
    }
  });
  function dropGhost() {             // 끌기가 중간에 끊기면(창 밖·브라우저 끌기 등) 이름표를 지운다
    if (!D) return;
    if (D.ghost) D.ghost.remove();
    if (D.over) D.over.classList.remove('over');
    document.body.classList.remove('gw-va-dragging');
    D = null;
  }
  document.addEventListener('pointercancel', dropGhost);
  window.addEventListener('blur', dropGhost);
  document.addEventListener('pointerup', function (e) {
    if (!D) return;
    var d = D;
    D = null;
    if (!d.ghost) return;            // 움직이지 않았으면 누르기(click)로 처리
    d.ghost.remove();
    document.body.classList.remove('gw-va-dragging');
    if (d.over) d.over.classList.remove('over');
    // 끌기 직후 브라우저가 click 을 보낼 수도, 안 보낼 수도 있다 — 그 한 번만 막고 곧 풀어 다음 누르기는 살린다
    S.suppressClick = true;
    setTimeout(function () { S.suppressClick = false; }, 0);
    drop(d.code, document.elementFromPoint(e.clientX, e.clientY));
  });

  // 누르기 — 사람 고르기 / 차에 넣기 / 명단으로 빼기
  document.addEventListener('click', function (e) {
    if (S.suppressClick) { S.suppressClick = false; return; }
    if (!S.canEdit) return;
    var rm = e.target.closest('[data-rm]');
    if (rm) {
      var rc = carOf(rm.dataset.plate), rh = rm.dataset.half || 'all';
      if (rh === 'all') { removeHalf(rm.dataset.rm, 'am'); removeHalf(rm.dataset.rm, 'pm'); }
      else removeHalf(rm.dataset.rm, rh);
      setDirty(true);
      var nextC = (rh === 'pm' ? pmOf(rc) : rc.people)[0];
      focusAfter(nextC ? chipSel(nextC) : staffSel(rm.dataset.rm));
      render();
      return;
    }
    var no = e.target.closest('[data-noteopen]');
    if (no) {
      S.noteOpen[no.dataset.noteopen] = 1;
      focusAfter('[data-note="' + cssq(no.dataset.noteopen) + '"]');
      render();
      return;
    }
    var sp = e.target.closest('[data-split]');
    if (sp) {                                   // 오전·오후 나누기 — 오후에 오전 사람을 그대로 옮겨 놓고 시작
      var spc = carOf(sp.dataset.split);
      spc.pm = spc.people.slice();
      setDirty(true);
      focusAfter('[data-unsplit="' + cssq(spc.plate) + '"]');
      render();
      return;
    }
    var us = e.target.closest('[data-unsplit]');
    if (us) {
      var usc = carOf(us.dataset.unsplit);
      var pmOnly = (usc.pm || []).filter(function (p) { return usc.people.indexOf(p) === -1; });
      if (pmOnly.length && !window.confirm('오후에만 탄 ' + pmOnly.map(name).join(', ') + ' 님은 이 차 오후에서 빠져요. 오후도 오전과 같게 할까요?')) return;
      // 오전 사람이 다른 차 오후에 들어가 있으면 거기서 빼야 종일이 된다
      usc.people.forEach(function (p) { var o = whereAt(p, 'pm'); if (o && o !== usc) removeHalf(p, 'pm'); });
      usc.pm = null;
      setDirty(true);
      focusAfter('[data-split="' + cssq(usc.plate) + '"]');
      render();
      return;
    }
    if (e.target.closest('input, select, details')) return;   // 갈 곳 펼치기는 차에 넣기가 아님
    var tb = e.target.closest('[data-team]');
    if (tb) { pick(tb.dataset.team, tb); return; }
    var p = e.target.closest('[data-person]');
    if (p && S.picked && S.picked !== p.dataset.person && p.dataset.in) {   // 고른 사람을 차 안 이 사람 앞자리로
      drop(S.picked, p);
      return;
    }
    if (p) { pick(p.dataset.person, p); return; }
    if (S.picked) {
      var putBtn = e.target.closest('[data-put]');
      var target = putBtn ? (putBtn.closest('[data-drop="car"]') || putBtn) : e.target.closest('[data-drop]');
      if (target) drop(S.picked, target);
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && S.picked) {
      var was = document.activeElement;
      pick(S.picked, was && was.closest && (was.closest('[data-team]') || was.closest('.gw-va-chip')) ? was : null);
    }
  });
  document.addEventListener('toggle', function (e) {
    if (e.target && e.target.id === 'vaDoneGroup') S.doneOpen = e.target.open;
  }, true);
  // 방문 일정 칸을 직접 접고 펼친 것만 기억(다시 그릴 때 코드로 여닫는 것은 빼고)
  document.addEventListener('click', function (e) {
    var s = e.target.closest && e.target.closest('#vaSched > summary');
    if (!s) return;
    setTimeout(function () {
      try { localStorage.setItem(SCHED_KEY, s.parentNode.open ? '1' : '0'); } catch (er) { /* 무시 */ }
    }, 0);
  });
  document.addEventListener('change', function (e) {
    var t = e.target;
    if (t.dataset.note) {
      carOf(t.dataset.note).note = t.value.trim();
      setDirty(true);
    } else if (t.dataset.driver) {               // 운전할 사람을 맨 앞으로
      var dc = carOf(t.dataset.driver), dh = t.dataset.half || 'all';
      var dl = dh === 'pm' ? dc.pm : dc.people;
      var at = dl ? dl.indexOf(t.value) : -1;
      if (at > 0) {
        dl.splice(at, 1);
        dl.unshift(t.value);
        tidy(dc);
        setDirty(true);
        focusAfter('[data-driver="' + cssq(dc.plate) + '"][data-half="' + dh + '"]');
        render();
      }
    } else if (t.dataset.fixed) {
      var fc = carOf(t.dataset.fixed);
      post('/api/vehicles/fleet/' + encodeURIComponent(fc.plate), { fixed_user: t.value }).then(function (j) {
        fc.fixed_user = j.car.fixed_user; fc.fixed_name = name(j.car.fixed_user);
        msg(fc.label + '의 평소 사용자를 ' + (fc.fixed_user ? fc.fixed_name : '없음') + '(으)로 바꿨어요.', 'ok');
      }).catch(function (e2) { msg(e2.message, 'err'); t.value = fc.fixed_user || ''; });
    }
  });
  // 초기화(11-22 부장님) — 지금 화면의 배정을 모두 비운다. 저장해야 반영되고, 저장 전엔 다른 날로 갔다 오면 되돌아간다
  if ($('vaReset')) {
    $('vaReset').addEventListener('click', function () {
      if (!riders() && !S.cars.some(function (c) { return c.note; })) { msg('비울 배정이 없어요.', 'info', true); return; }
      if (!window.confirm('지금 화면의 차량 배정을 모두 비울까요?\n저장하기 전까지는 반영되지 않아요.')) return;
      S.cars.forEach(function (c) { c.people = []; c.pm = null; c.note = ''; });
      S.picked = null;
      setDirty(true);
      render();
      msg('모두 비웠어요. 다시 넣은 뒤 저장해 주세요.', 'info', true);
    });
  }
  $('vaSave').addEventListener('click', function () { save(false); });
  $('vaConfirm').addEventListener('click', function () { save(true); });

  // 날짜는 달력으로만 고른다(11-17 결정 — 글자를 치면 '0001-11-23' 같은 날을 바로 불러왔음, 11-21 점검)
  (function pickerOnly(inp) {
    if (!inp) return;
    inp.addEventListener('click', function () { try { inp.showPicker(); } catch (e) { /* 무시 */ } });
    inp.addEventListener('keydown', function (e) {
      if (e.key === 'Tab' || e.key === 'Escape') return;   // 칸 이동·닫기만
      e.preventDefault();
      if (e.key === 'Enter' || e.key === ' ') { try { inp.showPicker(); } catch (er) { /* 무시 */ } }
    });
    inp.addEventListener('paste', function (e) { e.preventDefault(); });
  })($('vaDay'));
  function go(day) {
    if (S.dirty && S.canEdit && !window.confirm('저장하지 않은 배정이 있어요. 다른 날로 갈까요?')) { $('vaDay').value = S.day; return; }
    load(day);
  }
  $('vaPrev').addEventListener('click', function () { go(addDays(S.day, -1)); });
  $('vaNext').addEventListener('click', function () { go(addDays(S.day, 1)); });
  $('vaDay').addEventListener('change', function () { if (this.value) go(this.value); });
  window.addEventListener('beforeunload', function (e) {
    if (S.dirty && S.canEdit) { e.preventDefault(); e.returnValue = ''; }
  });

  var q = new URLSearchParams(location.search).get('day');
  load(q || '');
})();
