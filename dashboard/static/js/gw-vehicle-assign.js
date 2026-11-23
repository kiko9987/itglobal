/* 그룹웨어 차량 배정 (2026-11-20) — 다음 날 차마다 탈 사람. 고치기는 배정 담당자(SJ)·인사 관리자만.
   카톡 배정을 대신한다(카톡 글 붙여넣기는 넣지 않음 — 부장님 11-20).
   코스 중심(11-23 부장님 1안): 방문 일정과 같은 코스 카드에서 차를 고른다. 코스 전체 한 차면 카드 머리에서 한 번,
   갈라 타면 사람 칩마다(오전·오후 따로도). 왼쪽은 차 목록(누가 탔는지) — 방문 일정처럼 줄을 눌러 체크한 뒤 코스를 누름 — 차를 끌어 코스나 사람에 놓아도 된다.
   저장은 예전 그대로 차마다 사람(오전 people · 오후 pm). 서버에서 온 글자는 모두 esc() 로만 넣는다. */
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

  var TOUCH = !!(window.matchMedia && window.matchMedia('(pointer: coarse)').matches);
  var S = { day: '', cars: [], staff: [], names: {}, canEdit: false, confirmed: false, dirty: false,
            visits: {}, places: {}, teams: [], seq: 0, focus: null, placeTries: 0, placeTimer: null,
            off: {}, half: {}, online: {}, work: {}, schedule: [], courses: null, plan: null, dayLabel: '',
            splitOpen: {}, splitCourse: {}, memoOpen: {}, pick: '', extra: [], carWhen: '', draftNote: '' };
  // pick = 체크한 차(방문 일정처럼 줄을 눌러 체크 → 코스를 누르거나 위 줄에서 코스 고르기, 11-24 부장님) · splitCourse = [갈라 타기]를 연 코스
  // courses = 그룹웨어에서 확정한 방문 일정의 코스 — 슬랙으로 확정한 날은 null(그땐 일정확정 팀을 카드로)
  // splitOpen = {이니셜: 1} 오전·오후 따로 고르는 중 · extra = '그 밖'에 더한 사람(코스에 없는 사람 — 미팅 등)
  // off = 종일 휴무(차에 못 넣음), half = 반차·확정 전 토요 휴무(꼬리표만), online = 온라인 당번, work = 사내 업무

  function name(code) { return S.names[code] || code; }
  function carOf(plate) { return S.cars.filter(function (c) { return c.plate === plate; })[0]; }
  function isOwn(c) { return c.usage === 'own'; }
  // 시공팀 차(11-23 부장님) — 포터 2대·전기트럭은 평소 시공팀이 끌고 나가고, 영업은 현장 지원·물품 운반 때만.
  //   차 종류 이름(포터·트럭)으로 알아본다(차량 관리 용도는 그대로 '공용'). 영업 차 대수에서 빼고 목록 맨 아래 따로
  function isCrew(c) { return c.usage === 'shared' && /포터|트럭/.test((c.model || '') + ' ' + c.label); }
  function carsInOrder() {      // 영업 차 → 시공팀 차. 자차 칸은 목록에서 뺐다(11-24 부장님 — 차 없는 코스는 메모 한 줄로)
    return S.cars.filter(function (c) { return !isCrew(c) && !isOwn(c); }).concat(S.cars.filter(isCrew));
  }
  function ownCar() { return S.cars.filter(isOwn)[0] || null; }           // 자차 칸(11-23 — 차가 모자라면 자기 차로)
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
  function uniq(list) { return list.filter(function (x, i) { return list.indexOf(x) === i; }); }
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
  // 한 사람의 차를 정한다. half = 'all'(종일) / 'am' / 'pm', plate '' = 뺀다. 휴무(종일)는 못 넣음
  function setCar(code, plate, half) {
    if (S.off[code]) return;
    if (half === 'all') { removeHalf(code, 'am'); removeHalf(code, 'pm'); } else removeHalf(code, half);
    if (plate) {
      var c = carOf(plate);
      if (half === 'all') { c.people.push(code); if (c.pm) c.pm.push(code); }
      else {
        if (!c.pm) c.pm = c.people.slice();
        (half === 'am' ? c.people : c.pm).push(code);
      }
      tidy(c);
    }
    setDirty(true);
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
  function focusAfter(sel) { S.focus = sel; }
  function restoreFocus() {
    var sel = S.focus;
    S.focus = null;
    if (!sel) return;
    var el = document.querySelector(sel);
    if (el) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
  }
  function cssq(v) { return String(v).replace(/["\\]/g, '\\$&'); }
  function ownCars() {                // 고정 차 주인 → 그 차(YG 카니발·MW 스파크)
    var own = {};
    S.cars.forEach(function (c) { if (c.usage === 'fixed' && c.fixed_user && S.names[c.fixed_user]) own[c.fixed_user] = c; });
    return own;
  }
  // 고정 차 주인을 '배정 없어도 자기 차로 간다'고 짐작하던 것 없앰(11-24 부장님 — 빈 데서 차를 넣는 화면).
  //   고정 차도 다른 차처럼 코스에 넣어야 들어간다. (운행일지 운전자 추정의 '고정 사용자'는 그대로)
  function byOwnFixed() { return false; }
  function noCar(code) { return !S.off[code] && !whereIs(code) && !byOwnFixed(code); }
  function riders() {
    var seen = {};
    S.cars.forEach(function (c) { ridersOf(c).forEach(function (p) { seen[p] = 1; }); });
    return Object.keys(seen).length;
  }
  function whereText(code) {       // '레이 6854' / '레이 6854 오전 · 포터 4904 오후' / '레이 6854 오전만'
    var am = whereAt(code, 'am'), pm = whereAt(code, 'pm');
    if (!am && !pm) return byOwnFixed(code) ? ownCars()[code].label : '';
    if (am && pm && am === pm) return am.label;
    if (am && pm) return am.label + ' 오전 · ' + pm.label + ' 오후';
    return am ? am.label + ' 오전만' : pm.label + ' 오후만';
  }
  function splitPerson(code) { return whereAt(code, 'am') !== whereAt(code, 'pm') || !!S.splitOpen[code]; }

  // ── 코스(카드 한 장 = 함께 가는 사람 묶음) ──
  function courseMode() { return !!S.courses; }
  function groups() {
    if (S.courses) {
      return S.courses.courses.filter(function (c) { return c.kind !== 'work'; }).map(function (c, i) {
        return { key: 'c' + i, label: c.label, members: c.members.filter(function (m) { return S.names[m]; }),
                 memo: c.memo || '', sites: c.sites || [] };
      });
    }
    var n = 0;
    return S.teams.map(function (t, i) {
      var label = t.label || (t.members.length > 1 ? '함께 ' + (++n) : '혼자');
      return { key: 't' + i, label: label, members: t.members.filter(function (m) { return S.names[m]; }),
               memo: t.memo || '', sites: [], visits: t.visits };
    }).filter(function (g) { return g.members.length; });
  }
  function groupOf() {             // 이니셜 → 든 코스 이름들(두 코스면 둘)
    var out = {};
    groups().forEach(function (g) { g.members.forEach(function (m) { (out[m] = out[m] || []).push(g.label); }); });
    return out;
  }
  function others(inG) {           // 코스 밖인데 차에 넣었거나 '그 밖'에 더한 사람
    return S.staff.map(function (p) { return p.code; }).filter(function (c) {
      return !inG[c] && (whereIs(c) || S.extra.indexOf(c) !== -1);
    });
  }
  function noCarList() {          // 방문 있는데(또는 코스에 있는데) 차 없음 — '이름(코스 N)'
    var inG = groupOf();
    var codes = uniq(Object.keys(inG).concat(Object.keys(S.visits).filter(function (c) { return S.names[c]; })));
    return codes.filter(noCar).map(function (c) { return name(c) + (inG[c] ? '(' + inG[c][0] + ')' : ''); });
  }
  function staffNoCar() { return noCarList(); }

  // ── 차 고르기 칸 ──
  // 고르는 칸의 차 이름 옆 = 지금 그 차에 탄 사람(고르는 사람 빼고) — 비었는지 바로 보이게
  function carOptions(sel, except) {
    function opt(c) {
      var who = isOwn(c) ? [] : ridersOf(c).filter(function (p) { return except.indexOf(p) === -1; }).map(name);
      // 고른 차는 이름만(같이 탄 사람은 왼쪽 차 목록에), 나머지는 지금 탄 사람 또는 '빈 차'
      var text = isOwn(c) || sel === c.plate ? c.label
        : c.label + (who.length ? ' · ' + who.join(', ') : isCrew(c) ? '' : ' · 빈 차');   // 시공팀 차는 '빈 차'가 아님
      return '<option value="' + esc(c.plate) + '"' + (sel === c.plate ? ' selected' : '') + '>' + esc(text) + '</option>';
    }
    var shared = S.cars.filter(function (c) { return c.usage === 'shared' && !isCrew(c); });
    var fixed = S.cars.filter(function (c) { return c.usage === 'fixed'; });
    var crew = S.cars.filter(isCrew);
    return (shared.length ? '<optgroup label="회사 차">' + shared.map(opt).join('') + '</optgroup>' : '') +
      (fixed.length ? '<optgroup label="고정 차(빌려 쓰기)">' + fixed.map(opt).join('') + '</optgroup>' : '') +
      (crew.length ? '<optgroup label="시공팀 차(현장 지원·물품 운반 때)">' + crew.map(opt).join('') + '</optgroup>' : '');
  }
  function personSel(code, half) {
    var at = whereAt(code, half === 'pm' ? 'pm' : 'am');
    var own = null;
    var none = '차 없음';
    var hl = half === 'am' ? ' 오전' : half === 'pm' ? ' 오후' : '';
    return '<select class="form-select form-select-sm gw-vc-sel' + (!at && !own ? ' none' : '') + '" data-pcar="' + esc(code) +
      '" data-half="' + half + '" aria-label="' + esc(name(code) + hl) + ' 차">' +
      '<option value=""' + (at ? '' : ' selected') + '>' + esc(none) + '</option>' + carOptions(at ? at.plate : '', [code]) +
      '</select>';
  }
  // 코스의 차 — 모두 같은 차(종일)면 그 번호판, 아무도 없으면 '', 섞였으면 '__mixed', 모두 자기 고정 차면 '__own'
  function courseCar(g) {
    var ms = g.members.filter(function (m) { return !S.off[m]; });
    if (!ms.length) return '__none';
    var inG = groupOf();
    var one = ms.filter(function (m) { return (inG[m] || []).length < 2; });   // 두 코스에 든 사람은 다른 코스 차일 수 있어 빼고 본다
    if (one.length) ms = one;
    var vals = uniq(ms.map(function (m) {
      if (splitPerson(m)) return '__mixed';
      var at = whereAt(m, 'am');
      return at ? at.plate : (byOwnFixed(m) ? '__own' : '');
    }));
    return vals.length === 1 ? vals[0] : '__mixed';
  }
  // 사람마다 차 고르기(갈라 타기·오전·오후 따로)는 없앴다(11-24 부장님 — 드문 경우라 코스 메모 한 줄로).
  //   예전에 사람마다 다르게 저장한 날은 차 칸에 실제 차 이름만 나열해 보여 주고, 차를 다시 넣으면 코스 모두 같은 차로
  // 코스 머리의 차 칸(11-23 부장님: 코스·사람 묶음은 끝났고 차를 코스에 넣는 일) — 왼쪽 차를 누르면 여기로
  function courseSlot(g) {
    var cur = courseCar(g);
    if (cur === '__none') return '';
    var x = function (label) {
      return S.canEdit ? '<button type="button" class="gw-va-x" data-clearc="' + esc(g.key) + '" aria-label="' + esc(g.label) + '에서 ' +
        esc(label) + ' 빼기"><i class="fas fa-xmark" aria-hidden="true"></i></button>' : '';
    };
    // 갈라 탄 코스 — '사람마다 다름' 대신 실제 차를 나열하고, 차 없는 사람은 주황으로 따로(11-23 부장님)
    //   예) '쏘렌토 8021 · 자차' / '쏘렌토 4041' + '차 없음 1명' / '쏘렌토 8021 · 레이 5358(오후)'
    if (cur === '__mixed') {
      var ms = g.members.filter(function (m) { return !S.off[m]; });
      var names = [];
      var add = function (t) { if (t && names.indexOf(t) === -1) names.push(t); };
      var lab = function (c) { return c.label; };
      ms.forEach(function (m) {
        var am = whereAt(m, 'am'), pm = whereAt(m, 'pm');
        if (am) add(lab(am));
        if (pm && pm !== am) add(lab(pm) + '(오후)');
        if (!am && !pm && byOwnFixed(m)) add(ownCars()[m].label);
      });
      var miss = ms.filter(noCar).length;
      return (names.length ? '<span class="gw-vc-slot on"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(names.join(' · ')) + '</span>' : '') +
        (miss ? '<span class="gw-vc-slot empty"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> 차 없음 ' + miss + '명</span>' : '');
    }
    if (cur === '__own') {
      var one = g.members.filter(function (m) { return !S.off[m]; });
      var own = one.length === 1 ? ownCars()[one[0]] : null;
      return '<span class="gw-vc-slot own"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(own ? own.label : '각자 고정 차') + '</span>';
    }
    if (cur && isOwn(carOf(cur))) {     // 차 없는 코스에 메모를 적은 것(예전 자차 칸) — 경고 대신 '회사 차 없이'
      return '<span class="gw-vc-slot own"><i class="fas fa-note-sticky" aria-hidden="true"></i> 회사 차 없이' + x('회사 차 없이') + '</span>';
    }
    if (cur) {
      var c = carOf(cur);
      return '<span class="gw-vc-slot on"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(c.label) + x(c.label) + '</span>';
    }
    return '<span class="gw-vc-slot empty"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> 차 없음</span>';
  }
  function tagsOf(code) {
    return (S.half[code] ? '<small class="gw-va-leave">' + esc(S.half[code]) + '</small>' : '') +
      (S.online[code] ? '<small class="gw-va-leave">' + esc(S.online[code] + (S.visits[code] ? '·방문 후' : '')) + '</small>' : '');
  }
  function pchip(code, detail) {
    if (S.off[code]) {      // 종일 휴무 — 차에 못 넣음
      return '<span class="gw-va-chip gw-vp-chip gw-vc-p off"><span class="gw-va-cname">' + esc(name(code)) + '</span>' +
        '<small class="gw-va-leave gw-vp-offtag">' + esc(S.off[code]) + '</small></span>';
    }
    var w = detail && noCar(code);
    var ctl;
    if (!detail) {
      ctl = '';
    } else if (!S.canEdit) {
      ctl = '<small class="gw-vc-ptext">' + esc(whereText(code) || '차 없음') + '</small>';
    } else if (splitPerson(code)) {
      ctl = '<span class="gw-vc-half"><small>오전</small>' + personSel(code, 'am') + '</span>' +
        '<span class="gw-vc-half"><small>오후</small>' + personSel(code, 'pm') + '</span>' +
        '<button type="button" class="gw-va-tool gw-vc-join" data-joinp="' + esc(code) + '">하루 같게</button>';
    } else {
      ctl = personSel(code, 'all');
    }
    return '<span class="gw-va-chip gw-vp-chip gw-vc-p' + (w ? ' w' : '') + (detail && splitPerson(code) && S.canEdit ? ' split' : '') +
      '"' + (detail ? ' data-cperson="' + esc(code) + '"' : '') + '>' + (w ? '<i class="fas fa-triangle-exclamation" aria-hidden="true"></i>' : '') +
      '<span class="gw-va-cname">' + esc(name(code)) + '</span>' + tagsOf(code) + ctl +
      (w ? '<span class="visually-hidden"> — 차 없음</span>' : '') + '</span>';
  }
  // 코스가 가는 지역 — 현장을 넣은 순서일 뿐 이동 순서가 아니라 화살표 없이, 같은 지역은 한 번만 쓰고 곳 수를 붙임
  //   예) '서울 동대문구 2 · 서울 용산구 · 서울 마포구' (11-23 부장님)
  function siteRoute(sites) {
    var out = [], n = {};
    sites.forEach(function (s) {
      var p = s.place;
      if (!p) return;
      if (!n[p]) out.push(p);
      n[p] = (n[p] || 0) + 1;
    });
    return out.map(function (p) { return n[p] > 1 ? p + ' ' + n[p] : p; }).join(' · ');
  }
  function placeRoute(members) {   // 슬랙으로 확정한 날 — 사람들이 갈 곳(도로명 기준 시·구, 11-24)을 코스 줄과 같은 모양으로
    var seen = {}, sites = [];
    members.forEach(function (m) {
      (S.places[m] || []).forEach(function (p) {
        if (!seen[p.addr]) { seen[p.addr] = 1; sites.push({ place: p.label }); }   // 같은 현장(주소)은 한 번만
      });
    });
    return siteRoute(sites);
  }
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
  function courseCard(g, inG) {
    var route = g.sites.length ? siteRoute(g.sites) : placeRoute(g.members);
    var count = g.sites.length ? '현장 ' + g.sites.length : (g.visits ? '방문 ' + g.visits : '');
    // 두 코스에 든 사람(오전 코스·오후 코스) — 차는 먼저 넣은 코스 것, 다르면 메모로
    var hint = g.members.filter(function (m) { return (inG[m] || []).length > 1 && !S.off[m]; }).map(function (m) {
      var other = inG[m].filter(function (l) { return l !== g.label; });
      return '<p class="gw-vc-hint">' + esc(name(m)) + ' 님은 ' + esc(other.join(', ')) + '에도 있어요.' +
        (S.canEdit ? ' 차가 다르면 메모에 적어 주세요.' : '') + '</p>';
    }).join('');
    var head = '<div class="gw-vp-chead"><strong>' + esc(g.label) + '</strong>' + courseSlot(g) +
      '<div class="gw-va-people">' + (g.members.map(function (m) { return pchip(m, false); }).join('') ||
        '<span class="gw-va-none">사람 없음</span>') + '</div>' + memoBtn(g) +
      (count ? '<span class="gw-vp-ccount">' + esc(count) + '</span>' : '') + '</div>';
    var memo = g.memo ? '<p class="gw-vp-memotext"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(g.memo) + '</p>' : '';
    var sites = g.sites.length
      ? '<details class="gw-vc-sites"><summary><i class="fas fa-chevron-right gw-va-cfold" aria-hidden="true"></i>' +
        '<span>' + esc(route || '현장') + '</span></summary><ul>' + g.sites.map(courseSite).join('') + '</ul></details>'
      : (route ? '<p class="gw-vc-route"><i class="fas fa-location-dot" aria-hidden="true"></i> ' + esc(route) + '</p>' : '');
    var pk = S.canEdit && S.pick ? carOf(S.pick) : null;     // 차를 체크했으면 코스 카드를 눌러 그 차를 넣는다
    return '<section class="gw-va-card gw-vp-course gw-vc-course' + (pk ? ' pickable' : '') +
      '" data-ccourse="' + esc(g.key) + '"' + (pk ? ' tabindex="0" role="button" aria-label="' + esc(g.label + '에 ' + pk.label + ' 넣기') + '"' : '') +
      '>' + head + carMemo(g) + hint + memo + sites + '</section>';
  }
  // 코스 메모 한 줄(11-24 부장님) — 자차로 가는 사람 같은 드문 예외만 적는다. 그 코스 차의 메모로 저장(배정 DM 에 같이 감)
  //   차 없는 코스(자차로 가는 한 사람 등)도 메모를 적으면 그 코스는 '회사 차 없이'로 보고 경고에서 뺀다(안에선 자차 칸에 저장)
  function memoCar(g) {           // 메모를 붙일 차 — 차 한 대인 코스는 그 차, 차 없는 코스는 자차 칸
    var cur = courseCar(g);
    if (cur === '') return ownCar();
    return cur && cur.indexOf('__') !== 0 ? carOf(cur) : null;
  }
  function memoText(g) { return courseCar(g) === '' ? '' : ((memoCar(g) || {}).note || ''); }
  function memoBtn(g) {           // 평소엔 [메모] 단추만(칸이 늘 보이면 복잡, 11-24) — 눌러야 한 줄 칸이 열린다
    var c = memoCar(g);
    if (!S.canEdit || !c || memoText(g) || S.memoOpen[g.key]) return '';
    return '<button type="button" class="gw-va-tool gw-vc-memobtn" data-memoopen="' + esc(g.key) + '" aria-label="' + esc(g.label) +
      ' 메모 쓰기"><i class="fas fa-pen" aria-hidden="true"></i> 메모</button>';
  }
  function carMemo(g) {
    var c = memoCar(g), text = memoText(g);
    if (!c) return '';
    if (!S.canEdit || !(text || S.memoOpen[g.key])) return text ? '<p class="gw-vc-cmemo"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(text) + '</p>' : '';
    return '<div class="gw-vc-cmemo"><i class="fas fa-note-sticky" aria-hidden="true"></i>' +
      '<input type="text" class="form-control form-control-sm" maxlength="60" data-cmemo="' + esc(c.plate) + '" data-ckey="' + esc(g.key) +
      '" value="' + esc(text) + '" placeholder="' + (isOwn(c) ? '메모 (예: 본인 차량으로 이동)' : '메모 (예외만 한 줄 — 자차로 가는 사람 등)') +
      '" aria-label="' + esc(g.label) + ' 메모" autocomplete="off"></div>';
  }
  function otherCard(inG) {        // 그 밖 — 코스에 없는데 차를 타는 사람(미팅·자재 심부름 등)
    var list = others(inG);
    var cand = S.staff.map(function (p) { return p.code; }).filter(function (c) {
      return !inG[c] && list.indexOf(c) === -1 && !S.off[c];
    });
    if (!list.length && !S.canEdit) return '';
    var add = S.canEdit && cand.length
      ? '<select class="form-select form-select-sm gw-vc-sel gw-vc-add" data-add="1" aria-label="그 밖에 사람 더하기">' +
        '<option value="">+ 사람 더하기</option>' + cand.map(function (c) {
          var why = S.work[c] ? ' (사내 업무)' : S.online[c] ? ' (' + S.online[c] + ')' : '';
          return '<option value="' + esc(c) + '">' + esc(name(c) + why) + '</option>';
        }).join('') + '</select>' : '';
    return '<section class="gw-va-card gw-vp-course gw-vc-course gw-vc-other" data-ccourse="other" aria-label="그 밖">' +
      '<div class="gw-vp-chead"><strong>그 밖</strong><div class="gw-va-people">' +
      (list.map(function (m) { return pchip(m, true); }).join('') || '<span class="gw-va-none">코스에 없는데 차를 탈 사람(미팅 등)</span>') + '</div>' + add + '</div></section>';
  }
  function leftCard() {            // 코스에 없는 현장(방문 일정에서 코스를 안 정한 현장)
    var left = S.courses ? S.courses.left : [];
    if (!left.length) return '';
    return '<section class="gw-va-card gw-vp-course gw-vc-course gw-vc-left"><details class="gw-vc-sites"><summary>' +
      '<i class="fas fa-chevron-right gw-va-cfold" aria-hidden="true"></i><strong>코스에 없는 현장 ' + left.length + '건</strong>' +
      '<span>' + esc(siteRoute(left)) + '</span></summary><ul>' + left.map(courseSite).join('') + '</ul></details></section>';
  }
  function nocarStrip(inG) {       // 차가 필요 없는 사람 — 사내 업무·온라인 당번(방문 없음)·휴무
    var codes = S.staff.map(function (p) { return p.code; });
    var work = codes.filter(function (c) { return S.work[c] && !S.visits[c] && !inG[c] && !whereIs(c); });
    var duty = codes.filter(function (c) { return S.online[c] && !S.visits[c] && !inG[c] && !whereIs(c) && !S.work[c]; });
    var off = codes.filter(function (c) { return S.off[c] && !inG[c]; });
    var bits = [];
    if (work.length) bits.push('<strong>사내 업무</strong> ' + esc(work.map(function (c) { return name(c) + '(' + S.work[c] + ')'; }).join(', ')));
    if (duty.length) bits.push('<strong>온라인 당번</strong> ' + esc(duty.map(name).join(', ')));
    if (off.length) bits.push('<strong>휴무</strong> ' + esc(off.map(function (c) { return name(c) + '(' + S.off[c] + ')'; }).join(', ')));
    return bits.length ? '<p class="gw-vc-strip"><span>차 필요 없음</span> ' + bits.join('<span class="gw-va-dot">·</span>') + '</p>' : '';
  }

  // ── 왼쪽 차 목록 ──
  function carRow(c, inG) {
    var rs = ridersOf(c);
    var labels = uniq(rs.map(function (m) { return (inG[m] || ['그 밖'])[0]; })).filter(function (l) { return l !== '혼자'; });
    var who;
    if (isSplit(c)) who = '오전 ' + (c.people.map(name).join('·') || '없음') + ' / 오후 ' + (c.pm.map(name).join('·') || '없음');
    else who = rs.map(name).join('·');
    var line = rs.length
      ? '<span class="gw-vc-cwho"><span class="gw-vc-clabel">' + esc(labels.join('·')) + '</span> ' + esc(who) + '</span>'
      : '<span class="gw-vc-cwho none">' + (isOwn(c) ? '차가 모자라면 본인 차량으로' : isCrew(c) ? '평소 시공팀'
        : c.usage === 'fixed' ? '평소 ' + esc(c.fixed_name || '정해진 사람 없음') : '빈 차') + '</span>';
    // 차 줄의 메모 칸은 없앴다(11-24 부장님 — 예외는 코스 메모 한 줄로)
    // 고정 차 '평소 [사람 고르기]' 칸은 뺐다(11-24 부장님) — 평소 사용자는 줄 글('평소 ○○')로만 보인다
    var on = S.canEdit && S.pick === c.plate;
    var check = S.canEdit ? '<input type="checkbox" class="form-check-input gw-vp-check" data-carcheck="' + esc(c.plate) + '"' +
      (on ? ' checked' : '') + ' aria-label="' + esc(c.label) + ' 고르기">' : '';
    return '<div class="gw-vc-car' + (rs.length ? ' used' : '') + (isOwn(c) ? ' own' : '') + (on ? ' on' : '') + '" data-car="' + esc(c.plate) + '">' +
      '<div class="gw-vc-carhead">' + check +
      '<strong>' + (isOwn(c) ? '<i class="fas fa-user" aria-hidden="true"></i> ' : '') + esc(c.label) + '</strong>' +
      (isOwn(c) ? '' : '<small>' + esc(c.plate) + '</small>') + '</div>' + line + '</div>';
  }
  // 차 목록 위 한 줄 — 안내 글 ↔ '레이 5358 → [코스에 넣기]' 가 같은 자리(방문 일정 사람 칸과 같게)
  function carBar(gs) {
    if (!S.canEdit) return '누가 어느 차를 타는지 보기만 해요.';
    var c = S.pick ? carOf(S.pick) : null;
    if (!c) return '<span class="gw-vp-how">차를 체크한 뒤 코스를 누르세요.' + (TOUCH ? '' : ' 끌어 놓아도 돼요.') + '</span>';
    return '<strong>' + esc(c.label) + '</strong> → <select class="form-select form-select-sm gw-vp-liftsel" data-carto="1" aria-label="' +
      esc(c.label) + ' 넣을 코스"><option value="">코스에 넣기</option>' + gs.map(function (g) {
        var cur = courseCar(g);
        var now = cur === '' ? ' (차 없음)' : cur === c.plate ? ' (지금 이 차)' : '';
        return '<option value="' + esc(g.key) + '">' + esc(g.label + now) + '</option>';
      }).join('') + '</select><button type="button" class="btn btn-sm btn-link" data-carunchk="1">풀기</button>';
  }

  // ── 그리기 ──
  function board() {               // 서버 재시작 전 옛 화면이면 판을 새 모양으로 바꿔 끼운다
    var b = $('vaBoard');
    if (b && !$('vcCourses')) {
      b.className = 'gw-va-board gw-vc-board';
      b.innerHTML = '<aside class="gw-va-staff gw-vc-cars" aria-label="차 목록"><div class="gw-va-staff-head"><strong>차</strong>' +
        '<span id="vcCarCount"></span></div><p class="gw-va-staff-how" id="vcHow"></p><div id="vcCarList"></div></aside>' +
        '<section class="gw-vc-main"><details class="gw-va-sched" id="vaSched" hidden></details>' +
        '<div id="vcStrip"></div><div class="gw-vp-courses" id="vcCourses"><p class="gw-empty">불러오는 중…</p></div></section>';
    }
    if ($('vaBar') && !$('vcWarn')) {
      var w = document.createElement('span');
      w.id = 'vcWarn';
      w.className = 'gw-vc-warn';
      $('vaBar').insertBefore(w, $('vaBar').firstChild);
    }
  }
  function render() {
    board();
    var inG = groupOf();
    var gs = groups();
    $('vcCourses').innerHTML = gs.map(function (g) { return courseCard(g, inG); }).join('') + otherCard(inG) + leftCard() ||
      '<p class="gw-va-none gw-va-noused">이날 방문 일정이 없어요.</p>';
    $('vcStrip').innerHTML = nocarStrip(inG);
    var sales = S.cars.filter(function (c) { return !isOwn(c) && !isCrew(c); });     // 영업 차만 센다(시공팀 차 빼고)
    var used = sales.filter(function (c) { return ridersOf(c).length; }).length;
    var crewUsed = S.cars.filter(function (c) { return isCrew(c) && ridersOf(c).length; }).length;
    $('vcCarCount').textContent = '쓰는 차 ' + used + ' / ' + sales.length + '대' + (crewUsed ? ' + 시공팀 차 ' + crewUsed : '');
    if (S.pick && !carOf(S.pick)) S.pick = '';
    $('vcHow').className = 'gw-va-staff-how gw-vc-bar' + (S.pick ? ' on' : '');
    $('vcHow').innerHTML = carBar(gs);
    var crewRows = S.cars.filter(isCrew);
    $('vcCarList').innerHTML = carsInOrder().filter(function (c) { return !isCrew(c); }).map(function (c) { return carRow(c, inG); }).join('') +
      (crewRows.length ? '<div class="gw-vc-sub">시공팀 차 <small>현장 지원·물품 운반 때만</small></div>' +
        crewRows.map(function (c) { return carRow(c, inG); }).join('') : '');
    // 아래 바 경고 — 코스 단위로 짧게('차 없는 코스: 코스 1, 코스 3'), 코스 밖 사람만 이름으로
    var missG = gs.filter(function (g) { return g.members.some(noCar); }).map(function (g) { return g.label; });
    var missP = Object.keys(S.visits).filter(function (c) { return S.names[c] && !inG[c] && noCar(c); }).map(name);
    var warn = $('vcWarn');
    if (warn) {
      var wt = (missG.length ? '차 없는 코스: ' + missG.join(', ') : '') + (missP.length ? (missG.length ? ' · ' : '') + '차 없음: ' + missP.join(', ') : '');
      warn.innerHTML = wt && S.canEdit ? '<i class="fas fa-triangle-exclamation" aria-hidden="true"></i> ' + esc(wt) : '';
      warn.hidden = !(wt && S.canEdit);
    }
    $('vaBar').hidden = !S.canEdit;
    $('vaBoard').classList.toggle('editing', S.canEdit);
    $('vaConfirm').textContent = S.dmOff ? (S.confirmed ? '저장(DM 꺼짐)' : '확정(DM 꺼짐)')
      : (S.confirmed ? '저장하고 바뀐 사람에게 알리기' : '확정하고 알리기');
    $('vaSave').textContent = S.confirmed ? '알리지 않고 저장' : '저장만';
    $('vaSave').hidden = false;
    setDirty(S.dirty);
    statusLine();
    restoreFocus();
  }
  // 슬랙으로 확정한 날(코스 없음) — 위에 그날 방문 일정(캔버스1 묶음)을 접어서 둔다
  var SCHED_KEY = 'gw.va.schedOpen';
  function schedOpenPref() {
    try { var v = localStorage.getItem(SCHED_KEY); if (v === '1' || v === '0') return v === '1'; } catch (e) { /* 무시 */ }
    return !(window.matchMedia && window.matchMedia('(max-width: 767.98px)').matches);
  }
  function schedItem(it) {
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
    var el = $('vaSched');
    if (!el) return;
    if (S.courses || !('schedule' in j)) { el.hidden = true; return; }   // 코스로 볼 때는 현장이 코스 카드 안에
    var n = S.schedule.reduce(function (a, g) { return a + g.items.length; }, 0);
    var todo = S.schedule.reduce(function (a, g) { return a + g.items.filter(function (it) { return !it.who.length; }).length; }, 0);
    var sum = '방문 ' + n + '건' + (todo ? ' · 배분 전 ' + todo + '건' : '');
    var body = S.schedule.map(function (g) {
      return '<div class="gw-va-sgroup"><h3>' + esc(g.title) + ' <span>' + g.items.length + '</span></h3>' +
        '<ul>' + g.items.map(schedItem).join('') + '</ul></div>';
    }).join('') || '<p class="gw-va-sempty">이날 방문 일정이 없어요.</p>';
    var wasOpen = el.dataset.ready ? el.open : schedOpenPref();
    el.innerHTML = '<summary><i class="fas fa-calendar-day" aria-hidden="true"></i><span class="gw-va-stitle">' +
      esc(j.day_label || '') + ' 방문 일정</span><span class="gw-va-ssum">' + esc(sum) + '</span>' +
      '<i class="fas fa-chevron-down gw-va-fold" aria-hidden="true"></i></summary><div class="gw-va-sbody">' + body + '</div>';
    el.open = wasOpen;
    el.hidden = false;
    el.dataset.ready = '1';
  }
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
    S.cars.forEach(function (c) {
      if (isOwn(c)) return;
      ridersOf(c).forEach(function (m) { if (!inCourse[m] && !S.visits[m] && out.indexOf(name(m)) === -1) out.push(name(m)); });
    });
    if (!none.length && !out.length) return '';
    return '차량 배정 뒤에 방문 일정이 다시 확정됐어요.' + (none.length ? ' 차 없음: ' + none.join(', ') + '.' : '') +
      (out.length ? ' 코스에서 빠짐: ' + out.join(', ') + '.' : '');
  }
  function status(j) {
    S.carWhen = j.confirmed_at ? j.confirmed_at.slice(5, 16).replace('-', '/') + (j.confirmed_name ? ' · ' + j.confirmed_name : '') : '';
    S.draftNote = '';      // 초안 없앰(11-23 부장님 — 타는 차·인원이 매일 달라서). 빈 날은 비어서 열린다
  }
  function statusLine() {        // 날짜 · 확정 상태 · 방문 일정 · 초안 안내 · (경고) — 한 줄
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
      board();
      S.day = j.day; S.cars = j.cars; S.staff = j.staff; S.canEdit = !!j.can_edit; S.dmOff = !!j.dm_off;
      S.cars.forEach(function (c) {
        c.pm = Array.isArray(c.pm) ? c.pm : null;
        if (isOwn(c)) c.label = '본인 차량 (자차)';     // 화면에서 부르는 이름(11-23 부장님 — '자차'만으론 어색)
      });
      S.confirmed = !!j.confirmed_at;
      S.names = {}; S.visits = {}; S.places = {}; S.off = {}; S.half = {}; S.online = {}; S.work = {};
      S.splitOpen = {}; S.splitCourse = {}; S.memoOpen = {}; S.extra = []; S.pick = '';
      j.staff.forEach(function (p) { S.names[p.code] = p.name; });
      // 확정 전 토요 휴무는 바뀔 수 있어 반차처럼 꼬리표만(방문 일정과 같게 — 코스에 넣을 수 있음, 11-23)
      (j.off || []).forEach(function (o) { if (o.soft) S.half[o.code] = o.why; else S.off[o.code] = o.why; });
      (j.half || []).forEach(function (o) { S.half[o.code] = o.why; });
      (j.online || []).forEach(function (o) { S.online[o.code] = '온라인 당번' + (o.shift ? '(' + o.shift + ')' : ''); });
      (j.work || []).forEach(function (o) { S.work[o.code] = o.task; });
      S.schedule = j.schedule || [];
      S.courses = j.courses || null;
      S.plan = j.plan || null;
      S.dayLabel = j.day_label || '';
      (j.visitors || []).forEach(function (v) { S.visits[v.code] = v.visits; S.places[v.code] = v.places || []; });
      S.teams = j.teams || [];
      schedPanel(j);
      S.dirty = false;
      $('vaDay').value = j.day;
      try { history.replaceState(null, '', '/vehicles/assign?day=' + j.day); } catch (e) { /* 무시 */ }
      status(j);
      var moved = S.canEdit ? followCourses() : [];
      render();
      if (moved.length) msg('방문 일정 코스대로 차에 넣었어요 — ' + moved.join(', ') + '. 저장해 주세요.', 'info', true);
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
        // 고르는 중이거나 메모를 쓰는 중이면 다시 그리지 않는다 — 다음에 그릴 때 들어감
        var inBoard = act && act.closest && act.closest('.gw-va-board');
        if (!D && !(inBoard && act.matches('input, select'))) render();
        placesLater(j.resolving);
      }).catch(function () { /* 다음에 불러올 때 */ });
    }, 15000);
  }
  function rowsForSave() {
    return S.cars.map(function (c) {   // 코스 메모(11-24) — 자차 칸은 탄 사람이 없으면 메모도 비움
      return { plate: c.plate, people: c.people, pm: c.pm || null, note: isOwn(c) && !ridersOf(c).length ? '' : (c.note || '') };
    });
  }
  // notify: 슬랙으로 알릴지. 확정 전 [저장만]·확정 뒤 [알리지 않고 저장]은 false
  function save(notify) {
    var label = $('vaStatus').querySelector('strong');
    var dayText = label ? label.textContent : S.day;
    var miss = staffNoCar();
    var missText = miss.length ? '\n\n아직 차 없는 사람: ' + miss.join(', ') : '';
    if (notify && S.dmOff) {          // 차량 DM 꺼 둠(11-23 부장님) — 확정 기록만 남고 슬랙은 안 나간다
      if (!window.confirm(dayText + ' 배정을 ' + (S.confirmed ? '저장할까요?' : '확정할까요?') + missText +
          '\n\n지금은 차량 DM 을 꺼 두어서 슬랙으로는 알리지 않아요.')) return;
    } else if (notify && !S.confirmed) {
      if (!window.confirm(dayText + ' 배정을 확정하고, 차에 넣은 ' + riders() + '명에게 슬랙으로 알릴까요?' + missText +
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
        else if (notify && j.dm_off) t += ' 차량 DM 은 꺼 두어서 슬랙으로는 알리지 않았어요.';
        else if (notify && j.changed.length) t += ' 슬랙으로 알릴 사람이 없었어요.';
        return load(S.day).then(function () { msg(t, 'ok', true); });
      }).catch(function (e) { msg(e.message, 'err', true); })
      .then(function () { btns.forEach(function (b) { b.disabled = false; }); setDirty(S.dirty); });
  }

  // ── 고르기 ──
  function groupByKey(key) { return groups().filter(function (g) { return g.key === key; })[0]; }
  // 체크한 차를 코스에 넣고 체크를 푼다
  function putPicked(key) {
    var c = carOf(S.pick), g = groupByKey(key);
    if (!c || !g) return;
    setCourse(key, c.plate);
    delete S.splitCourse[key];
    S.pick = '';
    msg(g.label + '에 ' + c.label + ' 넣었어요.', 'ok');
    focusAfter('[data-carcheck="' + cssq(c.plate) + '"]');
    render();
  }
  // 확정한 방문 일정 기준(11-24 부장님) — 코스에 차가 이미 한 대 들어가 있는데 차 없는 사람이 있으면(배정 뒤 코스에 더해진 사람 등)
  //   그 코스 차에 넣는다. 차를 고르는 게 아니라 코스 묶음을 따르는 것. 저장해야 반영. 돌려주는 값: ['김민욱 님 → 쏘렌토 4041', …]
  function followCourses() {
    if (!S.courses) return [];
    var inG = groupOf(), out = [];
    groups().forEach(function (g) {
      var ms = g.members.filter(function (m) { return !S.off[m] && (inG[m] || []).length < 2; });
      var cars = uniq(ms.map(whereIs).filter(Boolean));
      if (cars.length !== 1 || isOwn(cars[0])) return;          // 차가 없거나 여러 대·'회사 차 없이' 코스는 그대로
      ms.filter(noCar).forEach(function (m) {
        setCar(m, cars[0].plate, 'all');
        out.push(name(m) + ' 님 → ' + g.label + ' ' + cars[0].label);
      });
    });
    return out;
  }
  function setCourse(key, plate) {   // 코스 사람 모두(휴무 빼고) 종일 같은 차 / '' = 모두 뺀다
    var g = groupByKey(key);
    if (!g) return;
    var prev = courseCar(g), inG = groupOf();
    g.members.forEach(function (m) {
      if (S.off[m]) return;
      var at = whereIs(m);     // 두 코스에 든 사람 — 다른 코스 차를 이미 탔으면 그대로(다르면 메모로, 11-24)
      if ((inG[m] || []).length > 1 && at && at.plate !== prev) return;
      delete S.splitOpen[m];
      setCar(m, plate, 'all');
    });
  }
  document.addEventListener('change', function (e) {
    var t = e.target;
    if (!t.dataset) return;
    if (t.dataset.cmemo) {                    // 코스 메모 — 다시 그리지 않는다(쓰는 중 칸이 흔들리지 않게)
      var mc = carOf(t.dataset.cmemo), mg = groupByKey(t.dataset.ckey), v = t.value.trim();
      mc.note = v;
      setDirty(true);
      if (isOwn(mc) && mg) {                  // 차 없는 코스 — 메모를 적으면 '회사 차 없이', 지우면 다시 '차 없음'
        var cc = courseCar(mg);
        if (v && cc === '') setCourse(mg.key, mc.plate);
        else if (!v && cc === mc.plate) setCourse(mg.key, '');
        render();
      }
      return;
    }
    if (t.dataset.pcar) {
      var code = t.dataset.pcar, half = t.dataset.half || 'all';
      if (t.value === '__split') {             // 오전·오후 따로 — 지금 차를 양쪽에 둔 채로 칸만 둘로
        S.splitOpen[code] = 1;
        focusAfter('[data-pcar="' + cssq(code) + '"][data-half="am"]');
      } else {
        setCar(code, t.value, half);
        focusAfter('[data-pcar="' + cssq(code) + '"][data-half="' + (splitPerson(code) ? half : 'all') + '"]');
      }
      render();
    } else if (t.dataset.add) {
      if (!t.value) return;
      S.extra.push(t.value);
      focusAfter('[data-pcar="' + cssq(t.value) + '"]');
      render();
    } else if (t.dataset.carto) {
      if (t.value) putPicked(t.value);
    } else if (t.dataset.carcheck) {          // 체크 칸 자체(키보드 Space) — 한 번에 한 대만
      S.pick = t.checked ? t.dataset.carcheck : '';
      focusAfter('[data-carcheck="' + cssq(t.dataset.carcheck) + '"]');
      render();
    } else if (t.dataset.fixed) {
      var fc = carOf(t.dataset.fixed);
      post('/api/vehicles/fleet/' + encodeURIComponent(fc.plate), { fixed_user: t.value }).then(function (j) {
        fc.fixed_user = j.car.fixed_user; fc.fixed_name = name(j.car.fixed_user);
        msg(fc.label + '의 평소 사용자를 ' + (fc.fixed_user ? fc.fixed_name : '없음') + '(으)로 바꿨어요.', 'ok');
        render();
      }).catch(function (e2) { msg(e2.message, 'err'); t.value = fc.fixed_user || ''; });
    }
  });
  document.addEventListener('input', function (e) {
    var t = e.target;
    if (t.dataset && t.dataset.cmemo) { carOf(t.dataset.cmemo).note = t.value.trim(); setDirty(true); }   // 코스에 넣기는 change(칸을 떠날 때)
  });
  document.addEventListener('click', function (e) {
    if (S.suppressClick) { S.suppressClick = false; return; }
    if (!S.canEdit) return;
    var mo = e.target.closest('[data-memoopen]');
    if (mo) {
      S.memoOpen[mo.dataset.memoopen] = 1;
      focusAfter('[data-ckey="' + cssq(mo.dataset.memoopen) + '"]');
      render();
      return;
    }
    if (e.target.closest('[data-carunchk]')) { S.pick = ''; focusAfter('#vcCarList'); render(); return; }
    // 차 줄은 아무 곳이나 눌러도 체크(방문 일정 사람 줄과 같게) — 다른 차를 누르면 그 차로 바뀜
    var crow = e.target.closest('[data-car]');
    if (crow && !e.target.closest('input, select, a, button')) {
      S.pick = S.pick === crow.dataset.car ? '' : crow.dataset.car;
      focusAfter('[data-carcheck="' + cssq(crow.dataset.car) + '"]');
      render();
      return;
    }
    var cl = e.target.closest('[data-clearc]');
    if (cl) {
      setCourse(cl.dataset.clearc, '');
      var oc = ownCar();
      if (oc && !ridersOf(oc).length) oc.note = '';
      focusAfter('[data-ccourse="' + cssq(cl.dataset.clearc) + '"]');
      render();
      return;
    }
    var card = e.target.closest('[data-ccourse]');
    if (S.pick && card && card.dataset.ccourse !== 'other' && !e.target.closest('select, input, button, a, summary')) {
      putPicked(card.dataset.ccourse);
      return;
    }
    var jp = e.target.closest('[data-joinp]');
    if (jp) {                                 // 하루 같게 — 오전 차를 종일로(오전이 없으면 오후 차)
      var code = jp.dataset.joinp;
      var at = whereAt(code, 'am') || whereAt(code, 'pm');
      delete S.splitOpen[code];
      setCar(code, at ? at.plate : '', 'all');
      focusAfter('[data-pcar="' + cssq(code) + '"][data-half="all"]');
      render();
    }
  });

  // 차 끌기(마우스만) — 왼쪽 차를 코스 카드(코스 모두)나 사람 칩(그 사람만)에 놓는다
  var D = null;
  document.addEventListener('pointerdown', function (e) {
    if (!S.canEdit || e.pointerType !== 'mouse' || e.button !== 0) return;
    var row = e.target.closest('[data-car]');
    if (!row || e.target.closest('input, select, a, button')) return;
    D = { plate: row.dataset.car, x: e.clientX, y: e.clientY, ghost: null, over: null };
  });
  document.addEventListener('pointermove', function (e) {
    if (!D) return;
    if (!D.ghost) {
      if (Math.abs(e.clientX - D.x) + Math.abs(e.clientY - D.y) < 6) return;
      D.ghost = document.createElement('div');
      D.ghost.className = 'gw-va-ghost';
      D.ghost.textContent = carOf(D.plate).label;
      document.body.appendChild(D.ghost);
      document.body.classList.add('gw-va-dragging');
    }
    D.ghost.style.left = e.clientX + 'px';
    D.ghost.style.top = e.clientY + 'px';
    var t = document.elementFromPoint(e.clientX, e.clientY);
    var over = t && (t.closest('[data-cperson]') || t.closest('[data-ccourse]'));
    if (over && over.dataset.ccourse === 'other') over = null;   // 그 밖 카드는 사람 칩에만
    if (over !== D.over) {
      if (D.over) D.over.classList.remove('over');
      if (over) over.classList.add('over');
      D.over = over;
    }
  });
  function dropGhost() {
    if (!D) return;
    if (D.ghost) D.ghost.remove();
    if (D.over) D.over.classList.remove('over');
    document.body.classList.remove('gw-va-dragging');
    D = null;
  }
  document.addEventListener('keydown', function (e) {     // 차를 체크한 채 코스 카드에서 Enter·Space = 그 코스에 넣기 · Esc = 체크 풀기
    if (!S.canEdit || !S.pick) return;
    if (e.key === 'Escape') { S.pick = ''; render(); return; }
    var card = e.target.closest && e.target.closest('[data-ccourse]');
    if (!card || card !== e.target || (e.key !== 'Enter' && e.key !== ' ')) return;
    e.preventDefault();
    putPicked(card.dataset.ccourse);
  });
  document.addEventListener('pointercancel', dropGhost);
  window.addEventListener('blur', dropGhost);
  document.addEventListener('pointerup', function () {
    if (!D) return;
    var d = D;
    D = null;
    if (!d.ghost) return;
    d.ghost.remove();
    document.body.classList.remove('gw-va-dragging');
    if (!d.over) return;
    d.over.classList.remove('over');
    S.suppressClick = true;
    setTimeout(function () { S.suppressClick = false; }, 0);
    if (d.over.dataset.cperson) {
      delete S.splitOpen[d.over.dataset.cperson];
      setCar(d.over.dataset.cperson, d.plate, 'all');
    } else {
      setCourse(d.over.dataset.ccourse, d.plate);
    }
    S.pick = '';
    msg('');
    render();
  });

  // 슬랙으로 확정한 날의 방문 일정 칸 — 직접 접고 펼친 것만 기억
  document.addEventListener('click', function (e) {
    var s = e.target.closest && e.target.closest('#vaSched > summary');
    if (!s) return;
    setTimeout(function () {
      try { localStorage.setItem(SCHED_KEY, s.parentNode.open ? '1' : '0'); } catch (er) { /* 무시 */ }
    }, 0);
  });
  // 초기화(11-22 부장님) — 지금 화면의 배정을 모두 비운다. 저장해야 반영되고, 저장 전엔 다른 날로 갔다 오면 되돌아간다
  if ($('vaReset')) {
    $('vaReset').addEventListener('click', function () {
      if (!riders()) { msg('비울 배정이 없어요.', 'info', true); return; }
      if (!window.confirm('지금 화면의 차량 배정을 모두 비울까요?\n저장하기 전까지는 반영되지 않아요.')) return;
      S.cars.forEach(function (c) { c.people = []; c.pm = null; c.note = ''; });
      S.splitOpen = {}; S.pick = '';
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
      if (e.key === 'Tab' || e.key === 'Escape') return;
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
