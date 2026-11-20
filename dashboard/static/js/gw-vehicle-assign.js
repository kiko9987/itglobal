/* 그룹웨어 차량 배정 (2026-11-20) — 다음 날 차마다 탈 사람. 고치기는 배정 담당자(SJ)·인사 관리자만.
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

  var S = { day: '', cars: [], staff: [], names: {}, canEdit: false, confirmed: false, dirty: false, unsure: {} };

  function name(code) { return S.names[code] || code; }
  function carOf(plate) { return S.cars.filter(function (c) { return c.plate === plate; })[0]; }
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
  function chip(c, code, i) {
    var lead = i === 0 ? '<span class="gw-va-lead">운전</span>' : '';
    var x = S.canEdit ? '<button type="button" class="gw-va-x" data-rm="' + esc(code) + '" data-plate="' + esc(c.plate) +
      '" aria-label="' + esc(name(code)) + ' 빼기">×</button>' : '';
    return '<span class="gw-va-chip">' + lead + esc(name(code)) + x + '</span>';
  }
  function addSelect(c) {
    var opts = S.staff.map(function (p) {
      var at = whereIs(p.code);
      if (at && at.plate === c.plate) return '';
      var tail = at ? ' — ' + at.label + '에서 옮김' : '';
      return '<option value="' + esc(p.code) + '">' + esc(p.name + ' (' + p.code + ')' + tail) + '</option>';
    }).join('');
    var label = c.usage === 'fixed' ? '+ 빌려 쓰기' : '+ 사람';
    return '<select class="form-select form-select-sm gw-va-add" data-add="' + esc(c.plate) + '" aria-label="' +
      esc(c.label) + '에 사람 넣기"><option value="">' + label + '</option>' + opts + '</select>';
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
  function swapSelect(c) {
    var choices = S.unsure[c.plate];
    if (!choices || choices.length < 2) return '';
    var opts = choices.map(function (p) {
      var o = carOf(p);
      return '<option value="' + esc(p) + '"' + (p === c.plate ? ' selected' : '') + '>' + esc(o ? o.label : p) + '</option>';
    }).join('');
    return '<div class="gw-va-unsure"><i class="fas fa-circle-question" aria-hidden="true"></i> 카톡 글만으로는 어느 차인지 모호해요. 맞는 차를 골라 주세요: ' +
      '<select class="form-select form-select-sm" data-swap="' + esc(c.plate) + '" aria-label="맞는 차">' + opts + '</select></div>';
  }
  function card(c) {
    var sub = [c.plate, c.fuel].filter(Boolean).join(' · ');
    var fixed = '';
    if (c.usage === 'fixed') {
      fixed = '<div class="gw-va-fixed">평소: ' + (S.canEdit ? fixedSelect(c)
        : (c.fixed_name ? esc(c.fixed_name) : '<span class="gw-va-none">정해진 사람 없음</span>')) + '</div>';
    }
    var people = c.people.map(function (p, i) { return chip(c, p, i); }).join('');
    if (!people && !S.canEdit) people = '<span class="gw-va-none">' + (c.usage === 'fixed' ? '평소대로' : '배정 없음') + '</span>';
    var note = S.canEdit && (c.people.length || c.note)
      ? '<input type="text" class="form-control form-control-sm gw-va-note" maxlength="60" placeholder="메모(예: 유리용접)" data-note="' +
        esc(c.plate) + '" value="' + esc(c.note || '') + '" aria-label="' + esc(c.label) + ' 메모">'
      : (c.note ? '<div class="gw-va-notetext">' + esc(c.note) + '</div>' : '');
    return '<div class="gw-va-card' + (S.unsure[c.plate] ? ' unsure' : '') + '" data-plate="' + esc(c.plate) + '">' +
      '<div class="gw-va-car"><strong>' + esc(c.label) + '</strong><span>' + esc(sub) + '</span>' + fixed + '</div>' +
      '<div class="gw-va-people">' + people + (S.canEdit ? addSelect(c) : '') + '</div>' + note + swapSelect(c) + '</div>';
  }
  function render() {
    $('vaShared').innerHTML = S.cars.filter(function (c) { return c.usage !== 'fixed'; }).map(card).join('') ||
      '<p class="gw-empty">공용 차가 없어요.</p>';
    $('vaFixed').innerHTML = S.cars.filter(function (c) { return c.usage === 'fixed'; }).map(card).join('');
    $('vaBar').hidden = !S.canEdit;
    $('vaPasteBox').hidden = !S.canEdit;
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
    if (j.draft_from && S.canEdit) {
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
      S.confirmed = !!j.confirmed_at; S.unsure = {};
      S.names = {};
      j.staff.forEach(function (p) { S.names[p.code] = p.name; });
      S.dirty = !!(j.draft_from && S.canEdit);
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
    var unsure = Object.keys(S.unsure).length;
    if (confirm && unsure && !window.confirm('어느 차인지 모호한 줄이 ' + unsure + '개 있어요. 이대로 확정할까요?')) return;
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

  // ── 고치기 ──
  function moveTo(code, plate) {
    var at = whereIs(code);
    if (at) at.people.splice(at.people.indexOf(code), 1);
    carOf(plate).people.push(code);
    setDirty(true);
    if (at && at.plate !== plate) msg(name(code) + '님을 ' + at.label + '에서 옮겼어요.', 'info');
  }
  document.addEventListener('change', function (e) {
    var t = e.target;
    if (t.dataset.add) {
      if (t.value) moveTo(t.value, t.dataset.add);
      render();
    } else if (t.dataset.note) {
      carOf(t.dataset.note).note = t.value.trim();
      setDirty(true);
    } else if (t.dataset.fixed) {
      var fc = carOf(t.dataset.fixed);
      post('/api/vehicles/fleet/' + encodeURIComponent(fc.plate), { fixed_user: t.value }).then(function (j) {
        fc.fixed_user = j.car.fixed_user; fc.fixed_name = name(j.car.fixed_user);
        msg(fc.label + '의 평소 사용자를 ' + (fc.fixed_user ? fc.fixed_name : '없음') + '(으)로 바꿨어요.', 'ok');
      }).catch(function (e) { msg(e.message, 'err'); t.value = fc.fixed_user || ''; });
    } else if (t.dataset.swap) {
      var from = carOf(t.dataset.swap), to = carOf(t.value);
      if (!from || !to || from === to) return;
      // 두 차의 사람·메모를 맞바꾼다. 고른 줄은 확인 끝, 밀려난 줄이 모호했으면 그 표시는 따라간다
      var tmp = { people: to.people, note: to.note };
      to.people = from.people; to.note = from.note;
      from.people = tmp.people; from.note = tmp.note;
      var other = S.unsure[to.plate];
      delete S.unsure[from.plate];
      delete S.unsure[to.plate];
      if (other && from.people.length) S.unsure[from.plate] = other;
      setDirty(true);
      render();
    }
  });
  document.addEventListener('click', function (e) {
    var rm = e.target.closest('[data-rm]');
    if (!rm) return;
    var c = carOf(rm.dataset.plate);
    c.people.splice(c.people.indexOf(rm.dataset.rm), 1);
    setDirty(true);
    render();
  });
  $('vaParse').addEventListener('click', function () {
    var text = $('vaPaste').value.trim();
    if (!text) { msg('카톡 글을 붙여 넣어 주세요.', 'err'); return; }
    post('/api/vehicles/assign/parse', { day: S.day, text: text }).then(function (j) {
      if (!j.rows.length) { msg('읽을 수 있는 줄이 없어요. 이니셜과 차 이름(또는 번호 4자리)을 확인해 주세요.', 'err'); return; }
      S.cars.forEach(function (c) { if (c.usage !== 'fixed') { c.people = []; c.note = ''; } });
      S.unsure = {};
      j.rows.forEach(function (r) {
        r.people.forEach(function (p) { var at = whereIs(p); if (at) at.people.splice(at.people.indexOf(p), 1); });
        var c = carOf(r.plate);
        c.people = c.people.concat(r.people);
        if (r.note) c.note = r.note;
        if (r.unsure) S.unsure[r.plate] = r.choices;
      });
      setDirty(true);
      render();
      var t = j.rows.length + '줄을 채웠어요.';
      if (Object.keys(S.unsure).length) t += ' 노란 칸은 어느 차인지 확인해 주세요.';
      if (j.skipped.length) t += ' 못 읽은 줄: ' + j.skipped.join(' / ');
      msg(t, j.skipped.length ? 'err' : 'ok');
    }).catch(function (e) { msg(e.message, 'err'); });
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
