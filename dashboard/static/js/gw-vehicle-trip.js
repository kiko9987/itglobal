/* 오늘 내 차 [운행 종료](2026-11-22 부장님) — 운행을 마친 사람이 차량 배정 화면 위에서 바로 운전자를 적는다.
   저녁 슬랙 DM(services/vehicle_driver.send_asks)은 그대로 — 여기서 먼저 적으면 그 차엔 DM 이 안 가고,
   이미 갔으면 '답했어요'로 바뀐다. 오늘 탄 차가 없으면 칸을 아예 안 보인다.
   서버에서 온 글자는 모두 esc() 로만 넣는다. */
(function () {
  'use strict';

  var box = document.getElementById('vaTrip');
  if (!box) return;
  var trips = [];
  var open = {};      // plate → 'pick' | 'split' (운전자 고르는 중)
  var split = {};     // plate → {am, pm}

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function api(url, init) {
    init = init || {};
    init.credentials = 'same-origin';
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error('로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.');
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || '저장하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')');
        }
        return j;
      });
    });
  }

  function who(t) {
    if (t.pm_people) {
      return '오전 ' + (t.am_people.join(', ') || '-') + ' · 오후 ' + (t.pm_people.join(', ') || '-');
    }
    return t.riders.map(function (r) { return r.name; }).join(', ');
  }
  function drove(a) {
    if (a.am && a.am === a.pm) return a.am_name;
    return '오전 ' + (a.am_name || '-') + ' · 오후 ' + (a.pm_name || '-');
  }

  function doneLine(t) {
    var a = t.answer;
    var head = a.ended_at ? '운행 종료 ' + a.ended_at : '운전자 확인';
    var by = a.via === 'auto' ? '혼자 탄 차라 자동' : (a.by_name ? '적은 사람 ' + a.by_name : '');
    return '<p class="gw-vt-done"><i class="fas fa-circle-check" aria-hidden="true"></i> ' + esc(head) +
      ' · 운전: <strong>' + esc(drove(a)) + '</strong>' + (by ? ' <span>(' + esc(by) + ')</span>' : '') + '</p>';
  }

  function nameBtns(t, half) {
    var cur = split[t.plate] || {};
    var only = half && t[half + '_codes'] && t[half + '_codes'].length ? t[half + '_codes'] : null;   // 그 반나절에 탄 사람만
    return t.riders.filter(function (r) { return !only || only.indexOf(r.code) !== -1; }).map(function (r) {
      var on = half && cur[half] === r.code;
      return '<button type="button" class="btn ' + (on ? 'btn-primary' : 'btn-outline-secondary') + ' gw-vt-name"' +
        ' data-plate="' + esc(t.plate) + '" data-code="' + esc(r.code) + '"' + (half ? ' data-half="' + half + '"' : '') +
        (half ? ' aria-pressed="' + (on ? 'true' : 'false') + '"' : '') + '>' + esc(r.name) + '</button>';
    }).join('');
  }

  function picker(t) {
    var mode = open[t.plate];
    if (mode === 'split') {
      var cur = split[t.plate] || {};
      return '<div class="gw-vt-pick">' +
        '<p class="gw-vt-q">오전(13시 전) 운전</p><div class="gw-vt-names">' + nameBtns(t, 'am') + '</div>' +
        '<p class="gw-vt-q">오후(13시부터) 운전</p><div class="gw-vt-names">' + nameBtns(t, 'pm') + '</div>' +
        '<div class="gw-vt-acts"><button type="button" class="btn btn-primary gw-vt-save" data-plate="' + esc(t.plate) + '"' +
        (cur.am && cur.pm ? '' : ' disabled') + '>저장</button>' +
        '<button type="button" class="btn btn-link gw-vt-back" data-plate="' + esc(t.plate) + '">한 사람이 운전</button></div></div>';
    }
    return '<div class="gw-vt-pick"><p class="gw-vt-q">누가 운전했나요?</p><div class="gw-vt-names">' + nameBtns(t, '') +
      '<button type="button" class="btn btn-outline-secondary gw-vt-split" data-plate="' + esc(t.plate) + '">나눠서 운전(오전·오후)</button>' +
      '</div><button type="button" class="btn btn-link gw-vt-cancel" data-plate="' + esc(t.plate) + '">취소</button></div>';
  }

  function card(t) {
    var solo = t.riders.length === 1;
    var body;
    if (open[t.plate]) body = picker(t);
    else if (t.answer) {
      body = doneLine(t) + (solo ? '' : '<button type="button" class="btn btn-link gw-vt-edit" data-plate="' + esc(t.plate) + '">운전자 고치기</button>');
    } else {
      body = '<button type="button" class="btn btn-primary gw-vt-end" data-plate="' + esc(t.plate) + '">운행 종료</button>';
    }
    return '<div class="gw-vt-card">' +
      '<div class="gw-vt-car"><strong>' + esc(t.car) + '</strong><small>' + esc(t.plate) + '</small>' +
      (solo ? '' : '<span>함께: ' + esc(who(t)) + '</span>') + '</div>' + body + '</div>';
  }

  function render(dayLabel) {
    if (!trips.length) { box.hidden = true; box.innerHTML = ''; return; }
    box.hidden = false;
    box.innerHTML = '<h2 class="gw-vt-title">오늘 내 차 <span>' + esc(dayLabel || '') + '</span></h2>' +
      '<p class="gw-vt-how">운행을 마치면 [운행 종료]를 누르고 운전한 사람을 골라 주세요. 못 누르면 저녁에 슬랙으로 물어봐요.</p>' +
      trips.map(card).join('') + '<p class="gw-vh-msg" id="vtMsg" role="status" aria-live="polite"></p>';
  }

  var label = '';
  function say(text, kind) {
    var el = document.getElementById('vtMsg');
    if (el) { el.textContent = text || ''; el.className = 'gw-vh-msg' + (kind ? ' ' + kind : ''); }
  }
  function find(plate) { return trips.filter(function (t) { return t.plate === plate; })[0]; }

  function save(t, am, pm) {
    box.querySelectorAll('button').forEach(function (b) { b.disabled = true; });
    api('/api/vehicles/trip/end', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ day: t.day, plate: t.plate, am: am, pm: pm })
    }).then(function (j) {
      trips = j.trips || [];
      delete open[t.plate]; delete split[t.plate];
      render(label);
      say('운행일지에 적었어요. 수고하셨어요.', 'ok');
    }).catch(function (e) {
      render(label);
      say(e.message, 'err');
    });
  }

  box.addEventListener('click', function (ev) {
    var b = ev.target.closest('button');
    if (!b || b.disabled) return;
    var t = find(b.getAttribute('data-plate'));
    if (!t) return;
    if (b.classList.contains('gw-vt-end')) {
      if (t.riders.length === 1) { save(t, t.riders[0].code, t.riders[0].code); return; }
      open[t.plate] = 'pick';
    } else if (b.classList.contains('gw-vt-edit')) {
      open[t.plate] = 'pick';
    } else if (b.classList.contains('gw-vt-split')) {
      open[t.plate] = 'split';
      split[t.plate] = split[t.plate] || {};
    } else if (b.classList.contains('gw-vt-back')) {
      open[t.plate] = 'pick';
    } else if (b.classList.contains('gw-vt-cancel')) {
      delete open[t.plate]; delete split[t.plate];
    } else if (b.classList.contains('gw-vt-name')) {
      var half = b.getAttribute('data-half');
      var code = b.getAttribute('data-code');
      if (!half) { save(t, code, code); return; }
      split[t.plate][half] = code;
    } else if (b.classList.contains('gw-vt-save')) {
      var s = split[t.plate] || {};
      if (s.am && s.pm) save(t, s.am, s.pm);
      return;
    } else return;
    render(label);
    // 다시 그린 뒤 키보드 위치 — 그 차 칸의 첫 단추로
    var i = trips.indexOf(t);
    var cardEl = box.querySelectorAll('.gw-vt-card')[i];
    var first = cardEl && cardEl.querySelector('button:not([disabled])');
    if (first) first.focus();
  });

  api('/api/vehicles/trip').then(function (j) {
    trips = j.trips || [];
    label = j.day_label;
    render(label);
  }).catch(function () { box.hidden = true; });
})();
