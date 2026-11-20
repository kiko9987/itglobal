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
  function msg(text, kind) {
    var el = $('vaMsg');
    el.className = 'gw-vh-msg' + (text ? ' ' + (kind || 'info') : '');
    el.textContent = text || '';
  }

  var TOUCH = !!(window.matchMedia && window.matchMedia('(pointer: coarse)').matches);   // 휴대폰: 끌기 대신 누르기
  var S = { day: '', cars: [], staff: [], names: {}, canEdit: false, confirmed: false, dirty: false,
            visits: {}, teams: [], picked: null };   // picked = 'YM' 또는 팀 'YM,LSH'

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
    if (b) b.disabled = !v && S.confirmed;
  }

  // ── 그리기 ──
  function visitTag(code) {
    var n = S.visits[code];
    return n ? '<small class="gw-va-visit">방문 ' + n + '</small>' : '';
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
    var free = codes.filter(function (c) { return !whereIs(c); });
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
      group('배정됨', placed, 'done');
    var how = $('vaStaffHow');
    how.className = 'gw-va-staff-how' + (S.picked ? ' picking' : '');
    how.textContent = S.picked
      ? keyName(S.picked) + ' → 넣을 차를 누르세요.' + (anyPlaced(S.picked) ? ' 여기 빈 곳을 누르면 차에서 빠져요.' : '') + ' (Esc: 취소)'
      : (TOUCH ? '이름(또는 팀의 [함께])을 누르고 → 넣을 차를 누르세요.'
         : '이름이나 팀의 [함께]를 끌어 차에 놓으세요. 누르고 → 차를 눌러도 돼요.');
    $('vaStaffCount').textContent = nTeam ? '방문 있는데 차 없음 ' + nTeam + '명' : '';
    $('vaStaffCount').className = nTeam ? 'gw-va-warn' : '';
  }
  function chip(c, code, i) {
    var lead = i === 0 ? '<span class="gw-va-lead">운전</span>' : '';
    var x = S.canEdit ? '<button type="button" class="gw-va-x" data-rm="' + esc(code) + '" data-plate="' + esc(c.plate) +
      '" aria-label="' + esc(name(code)) + ' 빼기">×</button>' : '';
    var cls = 'gw-va-chip' + (S.picked === code ? ' picked' : '');
    var attrs = S.canEdit ? ' data-person="' + esc(code) + '" data-in="' + esc(c.plate) + '" role="button" tabindex="0"' : '';
    return '<span class="' + cls + '"' + attrs + '>' + lead + esc(name(code)) + x + '</span>';
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
    if (!people) {
      people = '<span class="gw-va-none">' + (S.canEdit ? (S.picked ? '여기를 누르면 ' + esc(keyName(S.picked)) + ' 님을 넣어요'
        : '여기에 사람을 놓으세요') : (c.usage === 'fixed' ? '평소대로' : '배정 없음')) + '</span>';
    }
    var note = S.canEdit && (c.people.length || c.note)
      ? '<input type="text" class="form-control form-control-sm gw-va-note" maxlength="60" placeholder="메모" data-note="' +
        esc(c.plate) + '" value="' + esc(c.note || '') + '" aria-label="' + esc(c.label) + ' 메모">'
      : (c.note ? '<div class="gw-va-notetext">' + esc(c.note) + '</div>' : '');
    var drop = S.canEdit ? ' data-drop="car"' + (S.picked ? ' tabindex="0" role="button" aria-label="' +
      esc(c.label) + '에 ' + esc(keyName(S.picked)) + ' 넣기"' : '') : '';
    return '<div class="gw-va-card' + (S.picked ? ' can-drop' : '') + '" data-plate="' + esc(c.plate) + '"' + drop + '>' +
      '<div class="gw-va-car"><strong>' + esc(c.label) + '</strong><span>' + esc(sub) + '</span>' + fixed + '</div>' +
      '<div class="gw-va-people">' + people + '</div>' + note + '</div>';
  }
  function render() {
    $('vaShared').innerHTML = S.cars.filter(function (c) { return c.usage !== 'fixed'; }).map(card).join('') ||
      '<p class="gw-empty">공용 차가 없어요.</p>';
    var fixedCars = S.cars.filter(function (c) { return c.usage === 'fixed'; });
    $('vaFixed').innerHTML = fixedCars.map(card).join('');
    $('vaFixedHead').hidden = !fixedCars.length;   // 지금은 고정 차 없음(파나메라·벤츠·랜드로버는 배정에서 뺌, 11-20)
    staffPanel();
    if ($('vaBoard')) $('vaBoard').classList.toggle('editing', S.canEdit);
    else if (S.canEdit) msg('배정 화면이 새로 바뀌었어요. 서버 재시작 뒤 새로 고치면 사람 명단이 보여요.', 'info');
    $('vaBar').hidden = !S.canEdit;
    $('vaConfirm').textContent = S.confirmed ? '고친 내용 저장 · 바뀐 사람에게 알림' : '확정하고 알리기';
    $('vaSave').hidden = S.confirmed;
    setDirty(S.dirty);
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
    return api('/api/vehicles/assign' + (day ? '?day=' + encodeURIComponent(day) : '')).then(function (j) {
      S.day = j.day; S.cars = j.cars; S.staff = j.staff; S.canEdit = !!j.can_edit;
      S.confirmed = !!j.confirmed_at; S.picked = null;
      S.names = {}; S.visits = {};
      j.staff.forEach(function (p) { S.names[p.code] = p.name; });
      (j.visitors || []).forEach(function (v) { S.visits[v.code] = v.visits; });
      S.teams = j.teams || [];
      S.dirty = !!(j.draft_kind && S.canEdit && j.cars.some(function (c) { return c.people.length; }));
      $('vaDay').value = j.day;
      try { history.replaceState(null, '', '/vehicles/assign?day=' + j.day); } catch (e) { /* 무시 */ }
      status(j);
      render();
    }).catch(function (e) { msg(e.message, 'err'); });
  }
  function rowsForSave() {
    return S.cars.map(function (c) { return { plate: c.plate, people: c.people, note: c.note || '' }; });
  }
  function save(confirm) {
    var btns = [$('vaSave'), $('vaConfirm')];
    btns.forEach(function (b) { b.disabled = true; });
    msg(confirm ? '저장하고 알리는 중…' : '저장하는 중…', 'info');
    post('/api/vehicles/assign', { day: S.day, rows: rowsForSave(), confirm: !!confirm }).then(function (j) {
      S.dirty = false;
      var t = j.changed.length ? '저장했어요.' : '바뀐 내용이 없어요.';
      if (j.notified) t += ' ' + j.notified + '명에게 슬랙으로 알렸어요.';
      return load(S.day).then(function () { msg(t, 'ok'); });
    }).catch(function (e) { msg(e.message, 'err'); })
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
  function pick(code) {
    S.picked = S.picked === code ? null : code;
    render();      // 안내는 명단 위 한 줄에서 바뀐다(위에 띄우면 화면이 밀려 다음 누르기가 빗나감)
  }
  function drop(key, target) {   // key = 사람 또는 팀('YM,LSH'), target = 놓은 자리의 요소
    var chipEl = target && target.closest('.gw-va-chip[data-person]');
    var cardEl = target && target.closest('[data-drop="car"]');
    var outEl = target && target.closest('[data-drop="out"]');
    S.picked = null;
    if (cardEl) {
      placeMany(key, cardEl.dataset.plate, chipEl ? chipEl.dataset.person : null);
    } else if (outEl && anyPlaced(key)) {
      keyCodes(key).forEach(unplace);
      setDirty(true);
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
      c.people.splice(c.people.indexOf(rm.dataset.rm), 1);
      setDirty(true);
      render();
      return;
    }
    if (e.target.closest('input, select')) return;
    var tb = e.target.closest('[data-team]');
    if (tb) { pick(tb.dataset.team); return; }
    var p = e.target.closest('[data-person]');
    if (p && S.picked && S.picked !== p.dataset.person && p.dataset.in) {   // 고른 사람을 차 안 이 사람 앞자리로
      drop(S.picked, p);
      return;
    }
    if (p) { pick(p.dataset.person); return; }
    if (S.picked) {
      var target = e.target.closest('[data-drop]');
      if (target) drop(S.picked, target);
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && S.picked) { pick(S.picked); return; }
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('[data-drop="car"], .gw-va-chip[data-person]')) {
      e.preventDefault();
      e.target.click();
    }
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
