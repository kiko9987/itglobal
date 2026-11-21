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
            off: {}, half: {}, schedule: [] };
  // picked = 'YM' 또는 팀 'YM,LSH' · places = {이니셜: [{label, addr}]} 그날 갈 곳(차 카드)
  // off = {이니셜: 까닭} 그날 종일 쉬는 사람, half = 반차(그룹웨어 휴가·토요일 휴무표, 11-21)

  function name(code) { return S.names[code] || code; }
  function carOf(plate) { return S.cars.filter(function (c) { return c.plate === plate; })[0]; }
  function keyCodes(key) { return String(key || '').split(',').filter(Boolean); }
  function keyName(key) { return keyCodes(key).map(name).join(' + '); }
  function anyPlaced(key) { return keyCodes(key).some(function (c) { return whereIs(c); }); }
  function whereIs(code) {
    for (var i = 0; i < S.cars.length; i++) if (S.cars[i].people.indexOf(code) !== -1) return S.cars[i];
    return null;
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
  function staffNoCar() {             // 방문은 있는데 아직 차 없는 사람
    return Object.keys(S.visits).filter(function (c) { return S.names[c] && !S.off[c] && !whereIs(c); }).map(name);
  }
  function riders() {
    var n = 0;
    S.cars.forEach(function (c) { n += c.people.length; });
    return n;
  }

  // ── 그리기 ──
  // 명단 이름 옆 = 방문 수와 휴가 꼬리표. 어디 가는지는 위 '그날 방문 일정' 칸과 차 카드에서 본다(11-21 부장님)
  function visitTag(code) {
    var n = S.visits[code];
    var tag = n ? '<small class="gw-va-visit">방문 ' + n + '</small>' : '';
    var leave = S.off[code] || S.half[code];
    return tag + (leave ? '<small class="gw-va-leave">' + esc(leave) + '</small>' : '');
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
    c.people.forEach(function (code) {
      (S.places[code] || []).forEach(function (p) {
        if (!seen[p.addr]) { seen[p.addr] = 1; list.push(p); }
      });
    });
    if (!list.length) return '';
    var sum = groupLabels(list).map(function (g) { return g.label + (g.n > 1 ? ' ' + g.n + '곳' : ''); }).join(' · ');
    return '<div class="gw-va-places"><i class="fas fa-location-dot" aria-hidden="true"></i><span>' + esc(sum) + '</span></div>';
  }
  function person(code, extra) {   // 왼쪽 명단 한 사람
    var at = whereIs(code);
    var cls = 'gw-va-person' + (at ? ' placed' : '') + (S.picked === code ? ' picked' : '');
    return '<button type="button" class="' + cls + '" data-person="' + esc(code) + '"' +
      (S.picked === code ? ' aria-pressed="true"' : ' aria-pressed="false"') + '>' +
      '<span class="gw-va-pname">' + esc(name(code)) + '</span>' + visitTag(code) +
      (at ? '<small class="gw-va-where">' + esc(at.label) + '</small>' : '') + (extra || '') + '</button>';
  }
  function staffPanel() {
    var panel = $('vaStaff');
    if (!panel) return;              // 서버 재시작 전 옛 화면 — 명단 칸이 없음
    panel.hidden = !S.canEdit;
    if (!S.canEdit) return;
    var codes = S.staff.map(function (p) { return p.code; });
    // 그날 종일 쉬는 사람(그룹웨어 휴가·토요일 휴무표)은 '차 없음'에서 빼 '휴무'로 따로 — 차에 넣었으면 '배정됨'에 남는다
    var offList = codes.filter(function (c) { return S.off[c] && !whereIs(c); });
    var free = codes.filter(function (c) { return !whereIs(c) && !S.off[c]; });
    // 일정확정 팀(그날 방문을 같이 가는 사람) 중 아직 차 없는 사람이 있는 팀
    var inTeam = {};
    var teams = S.teams.map(function (t) {
      var left = t.members.filter(function (m) { return S.names[m] && free.indexOf(m) !== -1; });
      left.forEach(function (m) { inTeam[m] = 1; });
      return { left: left, visits: t.visits };
    }).filter(function (t) { return t.left.length; });
    var freeOther = free.filter(function (c) { return !inTeam[c]; });
    var placed = codes.filter(function (c) { return whereIs(c); });
    function group(title, list, cls) {
      if (!list.length) return '';
      return '<div class="gw-va-group ' + (cls || '') + '"><div class="gw-va-gtitle">' + esc(title) + ' <span>' + list.length + '</span></div>' +
        '<div class="gw-va-plist">' + list.map(function (c) { return person(c); }).join('') + '</div></div>';
    }
    function teamBox(t) {
      var key = t.left.join(',');
      if (t.left.length === 1) return person(key);   // 혼자 가는 사람은 상자 없이
      var on = S.picked === key;
      var btn = t.left.length > 1
        ? '<button type="button" class="gw-va-teambtn' + (on ? ' picked' : '') + '" data-team="' + esc(key) +
          '" aria-pressed="' + on + '" title="팀 그대로 함께 옮기기">' +
          '<i class="fas fa-people-group" aria-hidden="true"></i> 함께 <small>방문 ' + t.visits + '</small></button>'
        : '';
      return '<div class="gw-va-team">' + btn + t.left.map(function (c) { return person(c); }).join('') + '</div>';
    }
    var nTeam = teams.reduce(function (n, t) { return n + t.left.length; }, 0);
    $('vaStaffList').innerHTML =
      (teams.length ? '<div class="gw-va-group need"><div class="gw-va-gtitle">일정확정 팀 · 차 없음 <span>' + nTeam + '</span></div>' +
        '<div class="gw-va-plist">' + teams.map(teamBox).join('') + '</div></div>' : '') +
      group('차 없음', freeOther) +
      group('배정됨', placed, 'done') +
      group('휴무', offList, 'off');
    var how = $('vaStaffHow');
    how.className = 'gw-va-staff-how' + (S.picked ? ' picking' : '');
    how.textContent = S.picked
      ? keyName(S.picked) + ' 님을 넣을 차에서 [여기에 … 넣기]를 누르세요.' + (anyPlaced(S.picked) ? ' 명단 빈 곳을 누르면 차에서 빠져요.' : '') +
        (TOUCH ? '' : ' Esc 를 누르면 취소돼요.')
      : (TOUCH ? '이름이나 팀의 [함께]를 누른 뒤 넣을 차를 누르세요. 차에서 맨 앞 사람이 운전해요.'
         : '이름이나 팀의 [함께]를 끌어 차에 놓으세요. 누른 뒤 차를 눌러도 돼요. 차에서 맨 앞 사람이 운전하고, 사람 위에 놓으면 그 앞자리로 가요.');
    $('vaStaffCount').textContent = nTeam ? '방문 있는데 차 없음 ' + nTeam + '명' : '';
    $('vaStaffCount').className = nTeam ? 'gw-va-warn' : '';
  }
  // 칩 = 이름 단추 + [×] 단추 나란히(누를 수 있는 칩 안에 또 단추가 들어 있던 것, 11-21 점검)
  function chip(c, code, i) {
    var lead = i === 0 ? '<span class="gw-va-lead">운전</span>' : '';
    var cls = 'gw-va-chip' + (S.picked === code ? ' picked' : '');
    if (!S.canEdit) return '<span class="' + cls + '">' + lead + esc(name(code)) + '</span>';
    var x = '<button type="button" class="gw-va-x" data-rm="' + esc(code) + '" data-plate="' + esc(c.plate) +
      '" aria-label="' + esc(c.label) + '에서 ' + esc(name(code)) + ' 빼기"><i class="fas fa-xmark" aria-hidden="true"></i></button>';
    return '<span class="' + cls + '" data-person="' + esc(code) + '" data-in="' + esc(c.plate) + '">' + lead +
      '<button type="button" class="gw-va-cname" aria-pressed="' + (S.picked === code) + '" aria-label="' +
      esc(name(code) + (i === 0 ? '(운전)' : '') + ' — 눌러서 옮기기') + '">' + esc(name(code)) + '</button>' + x + '</span>';
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
    var people = c.people.map(function (p, i) { return chip(c, p, i); }).join('');
    // 사람을 고른 동안은 차마다 [여기에 넣기] 단추(카드 통째를 단추로 만들면 안의 메모 칸·칩까지 단추 속에 갇혔음)
    var put = S.canEdit && S.picked
      ? '<button type="button" class="btn btn-sm btn-outline-secondary gw-va-put" data-put="' + esc(c.plate) + '" aria-label="' +
        esc(c.label) + '에 ' + esc(keyName(S.picked)) + ' 넣기"><i class="fas fa-plus" aria-hidden="true"></i> 여기에 ' +
        esc(keyName(S.picked)) + ' 넣기</button>'
      : '';
    if (!people && !put) {
      people = '<span class="gw-va-none">' + (S.canEdit ? '여기에 사람을 놓으세요'
        : (c.usage === 'fixed' ? '평소대로' : '배정 없음')) + '</span>';
    }
    var note = S.canEdit && (c.people.length || c.note)
      ? '<input type="text" class="form-control form-control-sm gw-va-note" maxlength="60" placeholder="메모(예: 열쇠는 사무실)" data-note="' +
        esc(c.plate) + '" value="' + esc(c.note || '') + '" aria-label="' + esc(c.label) + ' 메모" autocomplete="off">'
      : (c.note ? '<div class="gw-va-notetext">' + esc(c.note) + '</div>' : '');
    var drop = S.canEdit ? ' data-drop="car"' : '';
    return '<div class="gw-va-card' + (S.picked ? ' can-drop' : '') + '" data-plate="' + esc(c.plate) + '"' + drop + '>' +
      '<div class="gw-va-car"><strong>' + esc(c.label) + '</strong><span>' + esc(sub) + '</span>' + fixed + '</div>' +
      carPlaces(c) + '<div class="gw-va-people">' + people + put + '</div>' + note + '</div>';
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
  function schedPanel(j) {
    var el = schedEl();
    if (!el) return;
    if (!('schedule' in j)) { el.hidden = true; return; }   // 서버 재시작 전(일정 칸 자료 없음) — '방문 0건'으로 보이지 않게
    var n = S.schedule.reduce(function (a, g) { return a + g.items.length; }, 0);
    var todo = S.schedule.reduce(function (a, g) { return a + g.items.filter(function (it) { return !it.who.length; }).length; }, 0);
    var offs = (j.off || []).map(function (o) { return o.name + '(' + o.why + ')'; });
    var halves = (j.half || []).map(function (o) { return o.name + '(' + o.why + ')'; });
    var sum = '방문 ' + n + '건' + (todo ? ' · 배분 전 ' + todo + '건' : '') + (offs.length ? ' · 휴무 ' + offs.length + '명' : '');
    var leave = (offs.length || halves.length)
      ? '<p class="gw-va-soff">' + (offs.length ? '<strong>휴무</strong> ' + esc(offs.join(', ')) : '') +
        (offs.length && halves.length ? ' · ' : '') + (halves.length ? '<strong>반차</strong> ' + esc(halves.join(', ')) : '') + '</p>'
      : '';
    var body = S.schedule.map(function (g) {
      return '<div class="gw-va-sgroup"><h3>' + esc(g.title) + ' <span>' + g.items.length + '</span></h3>' +
        '<ul>' + g.items.map(schedItem).join('') + '</ul></div>';
    }).join('') || '<p class="gw-va-sempty">이날 방문 일정이 없어요.</p>';
    var wasOpen = el.dataset.ready ? el.open : schedOpenPref();
    el.innerHTML = '<summary><i class="fas fa-calendar-day" aria-hidden="true"></i><span class="gw-va-stitle">' +
      esc(j.day_label || '') + ' 방문 일정</span><span class="gw-va-ssum">' + esc(sum) + '</span>' +
      '<i class="fas fa-chevron-down gw-va-fold" aria-hidden="true"></i></summary>' +
      '<div class="gw-va-sbody">' + leave + body + '</div>';
    el.open = wasOpen;
    el.hidden = false;
    el.dataset.ready = '1';
  }
  function render() {
    $('vaShared').innerHTML = S.cars.filter(function (c) { return c.usage !== 'fixed'; }).map(card).join('') ||
      '<p class="gw-empty">공용 차가 없어요.</p>';
    var fixedCars = S.cars.filter(function (c) { return c.usage === 'fixed'; });
    $('vaFixed').innerHTML = fixedCars.map(card).join('');
    $('vaFixedHead').hidden = !fixedCars.length;   // 지금은 고정 차 없음(파나메라·벤츠·랜드로버는 배정에서 뺌, 11-20)
    staffPanel();
    if ($('vaBoard')) $('vaBoard').classList.toggle('editing', S.canEdit);
    $('vaBar').hidden = !S.canEdit;
    // 확정 뒤에도 '알리지 않고 저장'을 남긴다 — 메모 오타 하나에 그 차 사람 모두에게 DM 이 가던 것(11-21 점검)
    $('vaConfirm').textContent = S.confirmed ? '저장하고 바뀐 사람에게 알리기' : '확정하고 알리기';
    $('vaSave').textContent = S.confirmed ? '알리지 않고 저장' : '저장만';
    $('vaSave').hidden = false;
    setDirty(S.dirty);
    restoreFocus();
  }
  function status(j) {
    var el = $('vaStatus');
    var parts = ['<strong>' + esc(j.day_label) + '</strong>'];
    if (j.confirmed_at) {
      parts.push('<span class="gw-badge gw-badge-ok">확정</span> <span class="gw-va-when">' + esc(j.confirmed_at.slice(5, 16)) +
        (j.confirmed_name ? ' · ' + esc(j.confirmed_name) : '') + '</span>');
    } else {
      parts.push('<span class="gw-badge gw-badge-wait">확정 전</span>');
    }
    var drafted = j.cars.some(function (c) { return c.people.length; });
    if (j.draft_kind === 'teams' && drafted && S.canEdit) {
      parts.push('<span class="gw-va-draft">일정확정 팀대로, 각 팀이 최근에 탄 차에 미리 넣어 두었어요. 바뀐 곳만 고치고 확정해 주세요.</span>');
    } else if (j.draft_from && S.canEdit) {
      parts.push('<span class="gw-va-draft">아직 배정이 없어서 ' + esc(j.draft_from.slice(5).replace('-', '/')) +
        ' 배정을 미리 넣어 두었어요. 바뀐 곳만 고치고 확정해 주세요.</span>');
    }
    el.innerHTML = parts.join(' ');
  }

  // ── 불러오기·저장 ──
  function load(day) {
    msg('');
    var my = ++S.seq;     // ◀ ▶ 를 빨리 누르면 늦게 온 앞 날짜 응답이 덮지 않게
    return api('/api/vehicles/assign' + (day ? '?day=' + encodeURIComponent(day) : '')).then(function (j) {
      if (my !== S.seq) return;
      S.day = j.day; S.cars = j.cars; S.staff = j.staff; S.canEdit = !!j.can_edit;
      S.confirmed = !!j.confirmed_at; S.picked = null;
      S.names = {}; S.visits = {}; S.places = {}; S.off = {}; S.half = {};
      j.staff.forEach(function (p) { S.names[p.code] = p.name; });
      (j.off || []).forEach(function (o) { S.off[o.code] = o.why; });
      (j.half || []).forEach(function (o) { S.half[o.code] = o.why; });
      S.schedule = j.schedule || [];
      schedPanel(j);
      (j.visitors || []).forEach(function (v) { S.visits[v.code] = v.visits; S.places[v.code] = v.places || []; });
      S.teams = j.teams || [];
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
    return S.cars.map(function (c) { return { plate: c.plate, people: c.people, note: c.note || '' }; });
  }
  // notify: 슬랙으로 알릴지. 확정 전 [저장만]·확정 뒤 [알리지 않고 저장]은 false
  function save(notify) {
    var label = $('vaStatus').querySelector('strong');
    var dayText = label ? label.textContent : S.day;
    if (notify && !S.confirmed) {
      var miss = staffNoCar();
      if (!window.confirm(dayText + ' 배정을 확정하고, 차에 넣은 ' + riders() + '명에게 슬랙으로 알릴까요?' +
          (miss.length ? '\n\n방문이 있는데 아직 차가 없는 사람: ' + miss.join(', ') : '') +
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
  function unplace(code) {
    var at = whereIs(code);
    if (at) at.people.splice(at.people.indexOf(code), 1);
    return at;
  }
  function place(code, plate, before) {   // before = 그 사람 앞자리로(없으면 맨 뒤)
    if (code === before) return;
    unplace(code);
    var c = carOf(plate);
    var i = before ? c.people.indexOf(before) : -1;
    if (i === -1) c.people.push(code); else c.people.splice(i, 0, code);
    setDirty(true);
  }
  function placeMany(key, plate, before) {   // 팀이면 적힌 순서대로 같은 자리에
    keyCodes(key).forEach(function (c) { place(c, plate, before); });
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
      placeMany(key, cardEl.dataset.plate, chipEl ? chipEl.dataset.person : null);
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
      var c = carOf(rm.dataset.plate);
      var left = c.people.indexOf(rm.dataset.rm);
      c.people.splice(left, 1);
      setDirty(true);
      // 뺀 자리 다음 사람(없으면 명단의 그 사람)으로 포커스
      focusAfter(c.people[left] ? chipSel(c.people[left]) : staffSel(rm.dataset.rm));
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
      var target = e.target.closest('[data-drop]');
      if (target) drop(S.picked, target);
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && S.picked) {
      var was = document.activeElement;
      pick(S.picked, was && was.closest && (was.closest('[data-team]') || was.closest('.gw-va-chip')) ? was : null);
    }
  });
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
    } else if (t.dataset.fixed) {
      var fc = carOf(t.dataset.fixed);
      post('/api/vehicles/fleet/' + encodeURIComponent(fc.plate), { fixed_user: t.value }).then(function (j) {
        fc.fixed_user = j.car.fixed_user; fc.fixed_name = name(j.car.fixed_user);
        msg(fc.label + '의 평소 사용자를 ' + (fc.fixed_user ? fc.fixed_name : '없음') + '(으)로 바꿨어요.', 'ok');
      }).catch(function (e2) { msg(e2.message, 'err'); t.value = fc.fixed_user || ''; });
    }
  });
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
