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
  // focusEl: 창이 열리면 포커스할 칸(없으면 부트스트랩 기본). 닫히면 연 칸으로 — 그 사이 달력이 다시 그려졌으면
  // 같은 날짜 칸을 찾아서(11-21 점검: 다시 그려진 뒤엔 포커스가 돌아오지 않았음)
  function showModal(el, focusEl) {
    if (!window.bootstrap) return null;
    var opener = document.activeElement;
    var openerDay = opener && opener.getAttribute ? opener.getAttribute('data-day') : null;
    el.addEventListener('shown.bs.modal', function s() {
      el.removeEventListener('shown.bs.modal', s);
      if (focusEl) focusEl.focus();
    });
    el.addEventListener('hidden.bs.modal', function h() {
      el.removeEventListener('hidden.bs.modal', h);
      if (opener && opener.focus && document.contains(opener)) opener.focus();
      else if (openerDay) refocus(openerDay);
    });
    var m = window.bootstrap.Modal.getOrCreateInstance(el);
    m.show();
    return m;
  }
  function refocus(day) {
    var el = document.querySelector('#soCal [data-day="' + day + '"][data-mode]');
    if (el) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
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
  var actEl = document.getElementById('soActions');
  if (!actEl) {   // 서버가 옛 템플릿을 들고 있으면(재시작 전) 자리를 직접 만든다 — 없으면 화면 전체가 멈췄음(11-24)
    actEl = document.createElement('div');
    actEl.className = 'gw-so-actions';
    actEl.id = 'soActions';
    var headEl = document.querySelector('#soPage .gw-cal-head');
    if (headEl) headEl.appendChild(actEl);
  }
  var flashEl = document.getElementById('soFlash');

  function monthKey() { return cur.y + '-' + pad(cur.m); }

  // kind: ok · warn(저장은 됐지만 알아 둘 것 — 예: 기본 횟수 넘김) · err(안 됨)
  function flash(text, kind) {
    flashEl.textContent = text;
    flashEl.className = 'gw-lv-flash ' + (kind || 'ok');
    flashEl.hidden = !text;
  }
  // 휴대폰은 달력 대신 토요일 목록(7칸 중 6칸이 빈 평일이고 토요일 칸은 48px — 11-21 점검)
  var NARROW = window.matchMedia ? window.matchMedia('(max-width: 767.98px)') : null;
  function narrow() { return !!(NARROW && NARROW.matches); }

  function satOf(day) { return (data.saturdays || []).filter(function (s) { return s.day === day; })[0]; }

  // 토요일 칸을 누를 수 있나 — 'self'(내 이름 넣고 빼기) · 'pick'(사람 고르기 창) · 'draft'(변경 신청 고르기) · ''
  function clickMode(s) {
    if (!data || s.holiday) return '';
    var me = data.me || {};
    var ed = me.approver || me.editor;     // 바로 고치는 사람 — 승인권자·관리자(11-25 부장님)
    if (s.past && !ed) return '';
    if (data.status === 'confirmed') {
      if (ed) return 'pick';
      return data.can_request_change && me.member ? 'draft' : '';
    }
    if (!data.can_toggle) return '';
    if (ed || data.pick_others) return 'pick';
    return me.member ? 'self' : '';
  }

  function inDraft(list, day) { return list.indexOf(day) !== -1; }

  // 취합 담당 — '빈승정 매니저가'(서버가 재시작 전이면 '취합 담당이')
  function coordName(d) { return (d.coordinators || []).length ? d.coordinators.join('·') + ' 매니저가' : '취합 담당이'; }
  var soHelpOpen = false;   // 접힌 도움말을 열어 둔 채 다시 그려도 열린 그대로
  document.addEventListener('toggle', function (e) {
    if (e.target && e.target.classList && e.target.classList.contains('gw-so-help')) soHelpOpen = e.target.open;
  }, true);
  function render() {
    var d = data;
    document.getElementById('soTitle').textContent = cur.y + '년 ' + cur.m + '월';
    var st = document.getElementById('soStatus');
    st.className = 'gw-so-status ' + d.status;
    st.textContent = d.status_label;

    // 위쪽 안내·단추
    var me = d.me || {};
    var h = '';
    var acts = '';   // 그달 결재 단추 — 달·상태 줄 오른쪽(actEl). 변경 신청 단추는 고른 내용 옆(bar)에 그대로
    var tips = [];   // 확정 뒤 고치는 방법·버튼 안내 — 늘 깔지 않고 접힌 도움말로(11-28 부장님, 토스 원칙). 문구는 그대로
    if (d.why) h += '<div class="gw-so-note warn">' + esc(d.why) + '</div>';
    // 기술영업·취합 담당·승인권자가 아닌 직원 — 보기만(11-25 시공·미화 빼고 모두에게 엶)
    var viewer = !(me.member || me.manager || me.approver || me.editor);
    if (viewer) {
      h += '<p class="gw-so-how">기술영업의 토요일 근무 스케줄이에요. 보기만 할 수 있어요.</p>';
    } else if (d.status === 'open') {
      if (d.note && d.decided_by) h += '<div class="gw-so-note warn">↩️ ' + esc(d.decided_by) + ' 님이 돌려보냈어요: ' + esc(d.note) + '</div>';
      if (d.can_toggle) {
        var nw = narrow();   // 휴대폰은 칸이 아니라 줄마다 단추
        h += '<p class="gw-so-how">' + (me.member && !d.pick_others
          ? (nw ? '쉬고 싶은 토요일의 [휴무 신청]을 누르세요. 꼭 쉬어야 하는 날이면 사유를 함께 적어 주세요.'
            : '쉬고 싶은 토요일 칸을 눌러 휴무를 신청하세요. 꼭 쉬어야 하는 날이면 사유를 함께 적어 주세요.')
          : (nw ? '토요일마다 [휴무자 선택]으로 휴무자를 체크하세요.' : '토요일을 선택해 휴무자를 체크하세요.') +
            ' 전체 인원이 다 선택 후 [결재 올리기]를 눌러 주세요.') +
          (d.max_off != null ? ' 한 토요일에 ' + d.max_off + '명(영업인원의 절반)까지 쉴 수 있어요.' : '') + '</p>';
      } else {
        h += '<p class="gw-so-how">아직 고를 수 없는 달이에요.</p>';
      }
      // 선택 완료(11-25 부장님) — 각자 다 고른 뒤 누름. 모두 누르면 자동으로 결재가 올라감. 서버가 재시작 전이면 can_done 없음
      if (d.can_done) {
        var mon = cur.m + '월';
        if (d.my_done) {
          h += '<p class="gw-so-how gw-so-done-note"><i class="fas fa-circle-check" aria-hidden="true"></i> ' + mon +
            ' 휴무 선택을 완료했어요. 결재가 올라가기 전까지는 고칠 수 있어요.</p>';
          acts = '<button type="button" class="btn btn-outline-secondary btn-sm" data-act="undone">다시 선택</button>';
        } else {
          h += '<p class="gw-so-how">' + mon + ' 휴무를 다 고른 뒤 [선택 완료]를 눌러 주세요. 쉬는 날이 없어도 눌러 주세요. ' +
            '모두 완료하면 ' + coordName(d) + ' 검토한 뒤 결재를 올려요.</p>';
          acts = '<button type="button" class="btn btn-primary btn-sm" data-act="done">선택 완료</button>';
        }
      }
      // 취합 담당(SJ) — 모두 선택 완료하면 검토하고 올림(11-27 회의). 아직 남은 사람이 있으면 옅은 버튼
      if (d.can_submit) {
        var allDone = !!d.done && (d.members || []).every(function (m) { return d.done.indexOf(m.name) !== -1; });
        if (allDone) h += '<p class="gw-so-how gw-so-done-note"><i class="fas fa-circle-check" aria-hidden="true"></i> ' +
          '대상자 모두 선택을 완료했어요. 검토한 뒤 [결재 올리기]를 눌러 주세요.</p>';
        acts += '<button type="button" class="btn ' + (d.can_done && !allDone ? 'btn-outline-secondary' : 'btn-primary') + ' btn-sm" data-act="submit">결재 올리기</button>';
      }
    } else if (d.status === 'wait') {
      h += '<p class="gw-so-how">' + (d.submitted_by === d.auto_submitter ? '대상자 모두 선택을 완료해 자동으로 결재가 올라갔어요('
        : esc(d.submitted_by) + ' 님이 결재를 올렸어요(') + esc((d.submitted_at || '').slice(5, 16)) +
        '). 승인권자 중 한 분이 승인하면 확정돼요.</p>';
      if (d.can_decide) {
        acts = '<button type="button" class="btn btn-outline-secondary btn-sm" data-act="return">돌려보내기</button>' +
          '<button type="button" class="btn btn-primary btn-sm" data-act="approve">승인</button>';
      }
    } else {
      h += '<p class="gw-so-how">' + esc(d.decided_by) + ' 님이 승인해 확정됐어요(' + esc((d.decided_at || '').slice(5, 16)) + ').</p>';
      if (me.approver || me.editor) {
        tips.push(narrow() ? '[휴무자 선택]으로 바로 고칠 수 있어요. 기록이 남아요.' : '토요일 칸을 눌러 바로 고칠 수 있어요. 기록이 남아요.');
        if (!me.approver) tips.push('변경은 본인이 [변경 신청]으로 하는 게 원칙이에요. 관리자 수정은 급할 때만 써 주세요.');
      } else if (d.can_request_change) {
        tips.push(narrow() ? '쉬는 날을 바꾸려면 취소할 내 휴무는 [휴무 취소], 새로 쉴 토요일은 [휴무 신청]으로 고른 뒤 [변경 신청]을 눌러 주세요.'
          : '쉬는 날을 바꾸려면 취소할 내 휴무 칸과 새로 쉴 토요일 칸을 눌러 고른 뒤 [변경 신청]을 눌러 주세요.');
      }
      if (draft.add.length || draft.remove.length) {
        // 바꿀 내용 — 달력 칩과 같은 색 꼬리표(빼기 빨강·넣기 초록, 11-25 부장님 '잘 인식이 안돼')
        var parts = draft.remove.map(function (x) { return '<span class="gw-so-dtag gone">' + esc(md(x)) + '(토) 휴무 취소</span>'; })
          .concat(draft.add.map(function (x) { return '<span class="gw-so-dtag add">' + esc(md(x)) + '(토) 휴무 신청</span>'; }));
        h += '<div class="gw-so-draft"><span class="gw-so-dlabel">바꿀 내용</span> ' + parts.join(' ') + ' ' +
          '<button type="button" class="btn btn-primary btn-sm" data-act="change">변경 신청</button> ' +
          '<button type="button" class="btn btn-light btn-sm" data-act="undo">되돌리기</button></div>';
      }
    }
    // 휴무 출근(11-24) — 결재 올린 뒤(대기·확정). 서버가 재시작 전이면 can_mark_worked 가 없어 안 보임
    if (d.can_mark_worked) {
      // 미리 변경 신청을 못 했을 때만 쓰는 기능이라는 걸 먼저(11-25 부장님 문구). 보통은 휴무 취소·신청으로 조절
      tips.push('휴무 토요일에 급한 업무로 출근을 하게 되어 미리 변경 신청을 하지 못했을 경우 구두 보고 후 [휴무 출근 기록]을 눌러 주세요.');
      acts = '<button type="button" class="btn btn-outline-secondary btn-sm" data-act="worked">휴무 출근 기록</button>' + acts;
    }
    // 전체 휴무 지정·풀기 — 관리자(11-25)
    if (d.can_full_off) {
      acts = '<button type="button" class="btn btn-outline-secondary btn-sm" data-act="full">전체 휴무</button>' + acts;
      // 관리자가 '나한테만 보이는 기능'인 줄 알게(11-25 부장님)
      tips.push('<i class="fas fa-lock" aria-hidden="true"></i> [전체 휴무] 버튼은 관리자에게만 보여요. 토요일 하루를 모두 쉬는 날로 정할 때 써요.');
    }
    if (tips.length) {
      h += '<details class="gw-help sm gw-so-help"' + (soHelpOpen ? ' open' : '') + '><summary><i class="fas fa-circle-question" aria-hidden="true"></i> 고치는 방법·버튼 안내' +
        '<i class="fas fa-chevron-down gw-help-fold" aria-hidden="true"></i></summary><div class="gw-help-body"><ul>' +
        tips.map(function (t) { return '<li>' + t + '</li>'; }).join('') + '</ul></div></details>';
    }
    barEl.innerHTML = h;
    actEl.innerHTML = acts;

    if (narrow()) renderList(d, me); else renderGrid(d, me);
    renderTable(d, me);
    renderSide(d, me);
  }

  // 쉬는 사람 n/최대(11-25 — 한 토요일 절반까지). 재시작 전 서버엔 max_off 가 없어 '' (예전처럼 인원만)
  // 전체 휴무(11-25) — 그날은 이름 대신 '전체 휴무' 한 칸
  function fullOff(s) { return !!(data && data.full_days && Object.prototype.hasOwnProperty.call(data.full_days, s.day)); }
  function fullChip(s, tag) {
    var note = (data.full_days || {})[s.day] || '';
    return '<' + tag + ' class="gw-cal-chip gw-so-chip fulloff"' + (note ? ' title="' + esc(note) + '"' : '') + '>전체 휴무 ' +
      s.names.length + '명' + (note ? '<span class="visually-hidden"> — ' + esc(note) + '</span>' : '') + '</' + tag + '>';
  }
  function capHtml(s) {
    var cap = data && data.max_off;
    if (fullOff(s)) return '<span class="gw-so-cap full">전체</span>';
    if (cap == null || s.holiday) return '';
    var n = s.names.length;
    // '5/5 가득'은 뜻을 알아보기 어려웠다(11-25 부장님) → '휴무 3/5명', 다 차면 '휴무 마감'
    return '<span class="gw-so-cap' + (n >= cap ? ' full' : '') + '" title="쉬는 사람 ' + n + '명 · 최대 ' + cap + '명">' +
      (n >= cap ? '휴무 마감' : '휴무 ' + n + '/' + cap + '명') + '</span>';
  }
  function isFull(s) { return data && data.max_off != null && s.names.length >= data.max_off; }

  // 휴무 출근 칩 — 달력 칸·휴대폰 목록 공용
  function workedChips(s, tag) {
    return (s.worked || []).map(function (n) {
      return '<' + tag + ' class="gw-cal-chip gw-so-chip worked" title="' + esc(n + ' — 휴무 출근') + '">' + esc(n) +
        ' <small>출근</small><span class="visually-hidden"> (쉬기로 했다가 휴무 출근)</span></' + tag + '>';
    }).join('');
  }


  // ── 휴무 메모(11-25 부장님) — 엑셀 메모란 대신. 근무표 오른쪽 빈자리에 ──
  // 꼭 쉬어야 하는 토요일·못 쉬는 토요일과 사유를 먼저 적어 두면 다른 매니저가 보고 양보하고 피해서 잡는다.
  // 서버가 재시작 전이면(memos 없음) 메모 칸을 그리지 않는다.
  function memoMap(d) {
    var m = {};
    (d.memos || []).forEach(function (x) { m[x.day + '|' + x.name] = x; });
    return m;
  }
  function memoIcon(x) {
    if (!x) return '';
    var t = x.kind_label + ': ' + x.reason;
    return ' <i class="fas ' + (x.kind === 'must' ? 'fa-star gw-so-mi must' : 'fa-ban gw-so-mi cant') + '" title="' + esc(t) +
      '" aria-hidden="true"></i><span class="visually-hidden"> (' + esc(t) + ')</span>';
  }
  // 그날 남이 '꼭 쉬어야 해요'라고 적었는데 아직 못 넣은 사람 — 내가 넣으면 그 사람 자리가 모자라는지
  function yieldNeeded(s, me) {
    var cap = data && data.max_off;
    if (cap == null) return [];
    var waiting = (data.memos || []).filter(function (x) {
      return x.day === s.day && x.kind === 'must' && x.name !== me && s.names.indexOf(x.name) === -1;
    });
    return (cap - s.names.length - 1 < waiting.length) ? waiting : [];
  }
  var memoDraft = null;       // 다시 그릴 때 쓰던 칸을 지키려고(불러오기·알림 확인으로 다시 그려짐)
  function keepMemoDraft() {
    var f = document.getElementById('soMemoForm');
    if (!f) return;
    memoDraft = { day: f.day.value, name: f.who ? f.who.value : '', kind: (f.querySelector('input[name="mkind"]:checked') || {}).value,
      reason: f.reason.value, month: data && data.month };
  }
  function memoHtml(d, me) {
    if (!d.memos) return '';
    var h = '<h3 class="gw-so-h" id="soMemoTitle">휴무 메모</h3><div class="gw-so-memo">' +
      // 적는 사람(기술영업 본인)과 보기만 하는 사람(대표님·관리자 등)에게 안내를 다르게(11-25 부장님 '어드민이라 이렇게 보여?')
      '<p class="gw-so-memo-how">' + (d.can_memo
        ? '꼭 쉬어야 하는 토요일이 있으면 달력에서 그 토요일을 눌러 사유와 함께 신청하세요. 다른 분들이 보고 배려해 드려요.'
        : '기술영업 매니저가 본인 휴무 사정을 적는 곳이에요. 여기서는 보기만 할 수 있어요.') + '</p>';
    // 토요일별 메모
    var by = {};
    (d.memos || []).forEach(function (x) { (by[x.day] = by[x.day] || []).push(x); });
    var days = Object.keys(by).sort();
    if (days.length) {
      h += '<ul class="gw-so-memos">';
      days.forEach(function (day) {
        var s = satOf(day);
        h += '<li><div class="gw-so-mday">' + md(day) + '(토)' + (s ? ' <span class="gw-so-sub">' + (s.names.length) + '명 휴무</span>' : '') + '</div><ul>';
        by[day].forEach(function (x) {
          var mine = x.name === me.name;
          var canDel = false;                // 적기·지우기는 달력 신청 창에서(11-25 부장님 — 메모 상자는 결과만)
          var picked = s && s.names.indexOf(x.name) !== -1;
          h += '<li class="' + x.kind + (mine ? ' mine' : '') + '">' + memoIcon(x).replace(/^ /, '') +
            '<span class="gw-so-mwho">' + esc(x.name) + '</span><span class="gw-so-mwhy">' + esc(x.reason) + '</span>' +
            (x.kind === 'must' && !picked ? '<span class="gw-so-mtag">아직 못 넣음</span>' : '') +
            (canDel ? '<button type="button" class="btn btn-link btn-sm gw-so-mdel" data-mdel="' + esc(x.day) + '" data-name="' + esc(x.name) +
              '" aria-label="' + esc(x.name + ' ' + md(x.day) + ' 메모 지우기') + '">지우기</button>' : '') + '</li>';
        });
        h += '</ul></li>';
      });
      h += '</ul>';
    } else {
      h += '<p class="gw-so-none">' + cur.m + '월에 적힌 메모가 없어요.</p>';
    }
    return h + '</div>';
  }
  // ── 토요일 근무표(11-24 부장님) — 사람 × 토요일 표. 칸마다 근무·휴무·휴무 출근, 끝에 휴무 횟수·휴무 출근 횟수 ──
  // 표 자리는 JS 가 만든다(서버가 옛 템플릿이어도 보이게 — 11-24 템플릿 캐시 사고)
  var tableEl = document.getElementById('soTable');
  if (!tableEl) {
    tableEl = document.createElement('section');
    tableEl.className = 'gw-panel gw-so-table';
    tableEl.id = 'soTable';
    calEl.parentNode.insertBefore(tableEl, calEl.nextSibling);
  }
  function mIcon(x) { return memoIcon(x).replace(/^ /, ''); }
  function renderTable(d, me) {
    var sats = d.saturdays || [];
    var mem = d.members || [];
    if (!sats.length || !mem.length) { tableEl.hidden = true; return; }
    var mm = memoMap(d);
    var off = {}, wk = {};
    sats.forEach(function (s) {
      s.names.forEach(function (n) { off[s.day + '|' + n] = true; });
      (s.worked || []).forEach(function (n) { wk[s.day + '|' + n] = true; });
    });
    var h = '<h3 class="gw-so-h" id="soTableTitle">토요일 근무표</h3>' +
      // 휴대폰에선 옆으로 밀어 봄 — 키보드로도 밀 수 있게 초점을 받는 영역으로
      '<div class="gw-so-tablewrap" tabindex="0" role="region" aria-labelledby="soTableTitle"><table class="gw-so-matrix">' +
      '<caption class="visually-hidden">' + cur.y + '년 ' + cur.m + '월 토요일마다 근무·휴무·휴무 출근</caption><thead><tr><th scope="col">이름</th>';
    sats.forEach(function (s) {
      h += '<th scope="col" class="' + (s.day === todayIso ? 'today' : '') + '">' + md(s.day) +
        (s.holiday ? '<small class="gw-so-hol">' + esc(s.holiday) + '</small>' : '') + '</th>';
    });
    h += '<th scope="col" class="n">휴무</th><th scope="col" class="n">휴무 출근</th></tr></thead><tbody>';
    mem.forEach(function (m) {
      h += '<tr' + (m.name === me.name ? ' class="mine"' : '') + '><th scope="row">' + esc(m.name) + '</th>';
      sats.forEach(function (s) {
        var k = s.day + '|' + m.name;
        if (s.holiday) h += '<td class="hol">공휴일</td>';
        // 별·금지 표시는 글자 기준 가운데 정렬에 끼지 않게 오른쪽에 붙임(11-25 부장님)
        else if (wk[k]) h += '<td><span class="gw-so-cw"><span class="gw-so-cell worked">휴무 출근</span>' + mIcon(mm[k]) + '</span></td>';
        else if (off[k]) h += '<td><span class="gw-so-cw"><span class="gw-so-cell off">휴무</span>' + mIcon(mm[k]) + '</span></td>';
        else h += '<td class="work"><span class="gw-so-cw">근무' + mIcon(mm[k]) + '</span></td>';
      });
      h += '<td class="n">' + (m.count || 0) + '</td><td class="n">' + (m.worked || 0) + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><th scope="row">근무</th>';
    sats.forEach(function (s) {
      var n = mem.filter(function (m) { return !off[s.day + '|' + m.name]; }).length;
      h += '<td>' + (s.holiday ? '-' : n + '명') + '</td>';
    });
    h += '<td></td><td></td></tr></tfoot></table></div>';
    // 기술영업 인센티브는 표에 넣지 않는다(격주휴무 없음). 표 아래 '대상 아님' 줄은 뺐다(11-25 부장님 '없어도 돼')
    tableEl.innerHTML = h;
    tableEl.hidden = false;
  }

  // 휴대폰 — 그달 토요일만 한 줄씩. 누를 일은 줄마다 단추 하나(44px)
  function actLabel(mode, s, di) {
    if (mode === 'self') return s.mine ? '내 휴무 바꾸기' : '휴무 신청';
    if (mode === 'pick') return '휴무자 선택';
    if (mode === 'draft') {
      if (s.mine) return inDraft(draft.remove, di) ? '취소 되돌리기' : '휴무 취소';
      return inDraft(draft.add, di) ? '신청 되돌리기' : '휴무 신청';
    }
    return '';
  }
  function renderList(d, me) {
    var rows = (d.saturdays || []).map(function (s) {
      var di = s.day;
      var mode = clickMode(s);
      var names = s.names.slice();
      var chips = fullOff(s) ? fullChip(s, 'span') : names.map(function (n) {
        var gone = me.name === n && inDraft(draft.remove, di);
        return '<span class="gw-cal-chip gw-so-chip' + (n === me.name ? ' mine' : '') + (gone ? ' gone' : '') + '">' + (gone ? '− <s class="gw-so-strike">' + esc(n) + '</s>' : esc(n)) +
          (gone ? '<small class="gw-so-ctag">취소</small>' : '') + '</span>';
      }).join('');
      if (inDraft(draft.add, di)) chips += '<span class="gw-cal-chip gw-so-chip mine add">+ ' + esc(me.name) + '<small class="gw-so-ctag">신청</small></span>';
      chips += workedChips(s, 'span');
      var label = mode ? actLabel(mode, s, di) : '';
      var btn = label
        ? '<button type="button" class="btn btn-outline-secondary gw-so-act" data-day="' + di + '" data-mode="' + mode + '" aria-label="' +
          esc(md(di) + ' 토요일 ' + label) + '">' + esc(label) + '</button>'
        : '';
      return '<li class="gw-so-row' + (s.past ? ' past' : '') + (di === todayIso ? ' today' : '') + '">' +
        '<div class="gw-so-rowhead"><strong>' + md(di) + '(토)</strong>' +
          (s.holiday ? ' <span class="gw-so-hol">' + esc(s.holiday) + '</span>' : '') +
          (s.past ? ' <span class="gw-so-rowsub">지난 날</span>' : '') +
          '<span class="gw-so-rowcount">' + (capHtml(s) || (names.length ? names.length + '명' : '')) + '</span></div>' +
        '<div class="gw-so-rownames">' + (chips || '<span class="gw-so-empty">쉬는 사람 없음</span>') + '</div>' + btn + '</li>';
    }).join('');
    calEl.innerHTML = '<ul class="gw-so-list" aria-label="' + cur.y + '년 ' + cur.m + '월 토요일">' +
      (rows || '<li class="gw-so-row"><span class="gw-so-empty">이 달은 토요일이 없어요.</span></li>') + '</ul>';
  }

  function renderGrid(d, me) {
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
        esc(label + (mode === 'self' ? (s.mine ? ' — 눌러서 내 휴무 바꾸기' : ' — 눌러서 휴무 신청') : ' — 눌러서 고치기')) + '"' : '';
      var mmap = memoMap(d);
      g += '<div class="' + cls + ' gw-so-sat' + (mode ? ' can' : '') + (s.past ? ' past' : '') + '"' + attrs + '>' +
        '<div class="gw-cal-num">' + dd + (s.holiday ? ' <small class="gw-so-hol">' + esc(s.holiday) + '</small>' : '') + capHtml(s) + '</div>';
      if (fullOff(s)) g += fullChip(s, 'div');
      else names.forEach(function (n) {
        var gone = me.name === n && inDraft(draft.remove, di);
        var mx = mmap[di + '|' + n];
        g += '<div class="gw-cal-chip gw-so-chip' + (n === me.name ? ' mine' : '') + (gone ? ' gone' : '') + '"' +
          (mx ? ' title="' + esc(mx.kind_label + ': ' + mx.reason) + '"' : '') + '>' + (gone ? '− <s class="gw-so-strike">' + esc(n) + '</s>' : esc(n)) +
          (mx && mx.kind === 'must' ? ' <i class="fas fa-star gw-so-mi must" aria-hidden="true"></i><span class="visually-hidden"> (꼭 쉬어야 해요: ' + esc(mx.reason) + ')</span>' : '') +
          (gone ? '<small class="gw-so-ctag">취소</small>' : '') + '</div>';
      });
      // 꼭 쉬어야 한다고 적었지만 자리가 없어 아직 못 넣은 사람 — 점선(양보 대기)
      (d.memos || []).forEach(function (x) {
        if (x.day === di && x.kind === 'must' && names.indexOf(x.name) === -1 && !fullOff(s)) {
          g += '<div class="gw-cal-chip gw-so-chip waiting" title="' + esc('꼭 쉬어야 해요: ' + x.reason + ' · 자리가 나면 넣어요') + '">' +
            '<i class="fas fa-star gw-so-mi must" aria-hidden="true"></i> ' + esc(x.name) + '<small class="gw-so-ctag">대기</small></div>';
        }
      });
      if (inDraft(draft.add, di)) g += '<div class="gw-cal-chip gw-so-chip mine add">+ ' + esc(me.name) + '<small class="gw-so-ctag">신청</small></div>';
      g += workedChips(s, 'div');
      if (mode === 'self' && !s.mine) g += isFull(s) ? '<div class="gw-so-plus" aria-hidden="true">휴무 마감 · ★ 적기</div>'
        : '<div class="gw-so-plus" aria-hidden="true">+ 휴무 신청</div>';
      // '+ 고르기'는 뺐다 — 점선 테두리·위 안내 줄과 겹침(11-25 부장님). 직원용 '+ 내 이름'·'+ 휴무 변경하기'는 눌렀을 때 일을 알려 줘 남김
      if (mode === 'draft' && !s.mine && !inDraft(draft.add, di)) g += '<div class="gw-so-plus" aria-hidden="true">+ 휴무 신청</div>';
      if (mode === 'draft' && s.mine && !inDraft(draft.remove, di)) g += '<div class="gw-so-plus minus" aria-hidden="true">− 휴무 취소</div>';   // 내 휴무는 눌러서 취소할 수 있다는 표시(11-25 부장님 — '휴무 취소/휴무 신청' 짝)
      g += '</div>';
    }
    var tail = (7 - ((lead + lastDay) % 7)) % 7;
    for (var t = 0; t < tail; t++) g += '<div class="gw-cal-day out"></div>';
    calEl.innerHTML = g + '</div>';
  }

  // 오른쪽 칸 — 휴무 출근 기록·변경 신청·수정 기록. 사람별 횟수는 근무표 끝 칸으로 옮겼다(11-24, 2번 기준도 뺌)
  var sideEl = document.querySelector('#soPage .gw-so-side');
  var oldCounts = document.getElementById('soCounts');            // 재시작 전 옛 템플릿의 횟수 상자 — 숨김
  if (oldCounts && oldCounts.closest('section')) oldCounts.closest('section').hidden = true;
  var workedBox = document.getElementById('soWorkedBox');
  if (!workedBox && sideEl) {
    workedBox = document.createElement('section');
    workedBox.className = 'gw-panel gw-so-box';
    workedBox.id = 'soWorkedBox';
    workedBox.hidden = true;
    workedBox.innerHTML = '<h3 class="gw-so-h">휴무 출근</h3><ul class="gw-so-changes" id="soWorked"></ul>';
    sideEl.insertBefore(workedBox, sideEl.firstChild);
  }

  // 휴무 선택 현황(11-25 부장님) — 누가 골랐고 누가 아직인지. 오른쪽 칸 맨 아래(수정 기록 밑)
  var pickBox = document.getElementById('soPickStat');
  if (!pickBox && sideEl) {
    pickBox = document.createElement('section');
    pickBox.className = 'gw-panel gw-so-box';
    pickBox.id = 'soPickStat';
    pickBox.hidden = true;
    sideEl.appendChild(pickBox);
  }
  // 휴무 메모(11-25 부장님) — 오른쪽 칸 맨 아래, 휴무 선택 현황 밑
  var memoBox = document.getElementById('soMemoBox');
  if (!memoBox && sideEl) {
    memoBox = document.createElement('section');
    memoBox.className = 'gw-panel gw-so-box gw-so-memobox';
    memoBox.id = 'soMemoBox';
    memoBox.setAttribute('aria-labelledby', 'soMemoTitle');
    memoBox.hidden = true;
    sideEl.appendChild(memoBox);
  }
  // 사람별로 고른 토요일 수 — 전체 휴무·공휴일은 본인이 고른 게 아니라 뺀다
  function pickCounts(d) {
    var c = {};
    (d.saturdays || []).forEach(function (s) {
      if (s.holiday || fullOff(s)) return;
      s.names.forEach(function (n) { c[n] = (c[n] || 0) + 1; });
    });
    return c;
  }
  function notPicked(d) {
    var c = pickCounts(d);
    return (d.members || []).filter(function (m) { return !c[m.name]; }).map(function (m) { return m.name; });
  }

  // 오른쪽 칸은 달마다 숨겼다 보였다 하지 않는다 — 달을 넘기면 배치가 바뀌던 것(11-25 부장님). 비면 '없어요' 한 줄
  var NONE = function (t) { return '<li class="gw-so-none">' + t + '</li>'; };
  function renderSide(d, me) {
    // 휴무 출근 기록
    var wk = d.worked || [];
    if (workedBox) {
      workedBox.hidden = false;
      document.getElementById('soWorked').innerHTML = wk.map(function (w) {
        return '<li><div><strong>' + esc(w.name) + '</strong> ' + esc(w.label) + '</div>' +
          '<div class="gw-so-sub">' + esc(w.reason || '') + ' · 구두 승인 ' + esc(w.verbal_by || '') +
          (w.recorded_by && w.recorded_by !== w.name ? ' · 기록 ' + esc(w.recorded_by) : '') + '</div>' +
          (w.can_undo ? '<div class="gw-so-btns"><button type="button" class="btn btn-light btn-sm" data-wundo="' + esc(w.day) +
            '" data-name="' + esc(w.name) + '">되돌리기</button></div>' : '') + '</li>';
      }).join('') || NONE('이 달 휴무 출근 기록이 없어요.');
    }

    // 변경 신청
    var ch = d.changes || [];
    document.getElementById('soChangesBox').hidden = false;
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
    }).join('') || NONE('이 달 변경 신청이 없어요.');

    // 기록
    var log = d.log || [];
    document.getElementById('soLogBox').hidden = !(me.manager || me.approver || me.editor);   // 보는 사람(역할)으로만 — 달마다 같게
    document.getElementById('soLog').innerHTML = log.map(function (l) {
      return '<li><span>' + esc((l.at || '').slice(5, 16)) + '</span> ' + esc(l.actor || '') + ' · ' + esc(l.detail || '') + '</li>';
    }).join('') || NONE('이 달 수정 기록이 없어요.');

    // 휴무 선택 현황 — 작성 중엔 '미선택'(11-25 부장님 — '고름'이 어색), 결재 뒤엔 '휴무 없음'
    if (pickBox) {
      var mem = d.members || [], pc = pickCounts(d), miss = notPicked(d);
      var done = mem.filter(function (m) { return pc[m.name]; });
      var open = d.status === 'open';
      if (open && d.done) {                // 선택 완료(11-25) — 누른 사람이 '선택 완료', 안 누른 사람이 '미선택'
        done = mem.filter(function (m) { return d.done.indexOf(m.name) !== -1; });
        miss = mem.filter(function (m) { return d.done.indexOf(m.name) === -1; }).map(function (m) { return m.name; });
      }
      var chip = function (n, extra, cls) {
        return '<li class="' + cls + (n === me.name ? ' mine' : '') + '">' + esc(n) + (extra ? ' <b>' + extra + '</b>' : '') + '</li>';
      };
      pickBox.hidden = !mem.length;
      pickBox.innerHTML = '<h3 class="gw-so-h">휴무 선택 현황 <small>' + done.length + '/' + mem.length + '명 선택' + (open && d.done ? ' 완료' : '') + '</small></h3>' +
        '<div class="gw-so-pslabel">' + (open ? '미선택' : '휴무 없음') + ' ' + miss.length + '명</div>' +
        (miss.length ? '<ul class="gw-so-pslist">' + miss.map(function (n) { return chip(n, '', 'miss'); }).join('') + '</ul>'
          : '<p class="gw-so-none">' + (open ? '모두 선택했어요.' : '모두 휴무가 있어요.') + '</p>') +
        (done.length ? '<div class="gw-so-pslabel">선택 완료 ' + done.length + '명</div><ul class="gw-so-pslist">' +
          done.map(function (m) { return chip(m.name, pc[m.name] ? pc[m.name] + '번' : '휴무 없음', 'ok'); }).join('') + '</ul>' : '');
    }

    // 휴무 메모 — 서버가 재시작 전이면(memos 없음) 숨김
    if (memoBox) {
      keepMemoDraft();
      var mh = memoHtml(d, me);
      memoBox.hidden = !mh;
      memoBox.innerHTML = mh;
    }

    // 오른쪽 칸은 늘 둔다(휴무 출근·변경 신청 상자는 늘 보임) — 넓은 화면은 늘 두 칸
    if (sideEl) {
      sideEl.hidden = false;
      document.getElementById('soPage').classList.add('has-side');
    }
  }

  function load(keepDraft) {
    var k = monthKey();
    document.getElementById('soTitle').textContent = cur.y + '년 ' + cur.m + '월';
    if (!keepDraft) draft = { add: [], remove: [] };
    var my = ++seq;
    if (!data) calEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    // 다른 달로 넘길 때 제목만 먼저 바뀌고 앞 달 달력이 그대로 보이던 것 → 흐리게 + 못 누르게(11-21 점검)
    else if (data.month !== k) { [calEl, barEl, actEl].forEach(function (el) { el.classList.add('is-loading'); }); }
    calEl.setAttribute('aria-busy', 'true');
    function done() { [calEl, barEl, actEl].forEach(function (el) { el.classList.remove('is-loading'); }); calEl.removeAttribute('aria-busy'); }
    return api('/api/satoff?month=' + k).then(function (j) {
      if (my !== seq) return;
      done();
      data = j;
      render();
      if (j.notify_fail) flash('⚠ ' + j.notify_fail, 'err');   // 이 달 마지막 처리의 슬랙 알림을 못 보냄(11-24)
      try { history.replaceState(null, '', '?month=' + k); } catch (e) { /* 무시 */ }
    }).catch(function (e) {
      if (my !== seq) return;
      done();
      data = null;
      calEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
      barEl.innerHTML = '';
      actEl.innerHTML = '';
    });
  }
  // 슬랙 알림은 뒤에서 나가므로 몇 초 뒤 다시 읽어 '못 보냄' 기록이 생겼으면 알려 준다(11-24 조용한 실패 점검).
  // 다시 읽기 자체가 실패하면 그냥 둔다 — 못 보낸 알림은 서버가 이미 에러 알림으로 보냈고 다음 불러오기에도 보인다
  function watchNotify(k) {
    setTimeout(function () {
      if (monthKey() !== k) return;
      api('/api/satoff?month=' + k).then(function (j) {
        if (monthKey() !== k || !j.notify_fail) return;
        data = j;
        render();
        flash('⚠ ' + j.notify_fail, 'err');
      }).catch(function () { /* 위 주석 */ });
    }, 6000);
  }

  if (NARROW) {
    var reflow = function () { if (data) render(); };
    if (NARROW.addEventListener) NARROW.addEventListener('change', reflow); else if (NARROW.addListener) NARROW.addListener(reflow);
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
  // '이번 달'은 ◀ ▶ 와 같은 테두리 버튼으로 — 밑줄 링크라 어색했음(11-25 부장님). 재시작 전 옛 템플릿에도 바로 적용
  document.getElementById('soToday').className = 'btn btn-sm btn-outline-secondary';
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
      openSelf(s);
      return;
    }
    if (mode === 'draft') {
      var name = data.me.name;
      if (s.mine) {
        if (inDraft(draft.remove, day)) draft.remove.splice(draft.remove.indexOf(day), 1); else draft.remove.push(day);
      } else if (s.names.indexOf(name) === -1) {
        if (inDraft(draft.add, day)) draft.add.splice(draft.add.indexOf(day), 1); else draft.add.push(day);
      }
      draft.add.sort(); draft.remove.sort();
      render();
      refocus(day);
    } else if (mode === 'pick') {
      openPick(s);
    }
  }
  calEl.addEventListener('click', function (e) {
    var cell = e.target.closest('[data-day][data-mode]');      // PC 달력 칸 · 휴대폰 목록 단추
    if (cell) onDay(cell);
  });
  calEl.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var cell = e.target.closest('.gw-so-sat[data-day]');
    if (cell) { e.preventDefault(); onDay(cell); }
  });

  // 휴무 메모 적기·지우기(11-25)
  var memoHost = memoBox || tableEl;
  memoHost.addEventListener('submit', function (ev) {
    var f = ev.target;
    if (f.id === 'soMemoForm') {
      ev.preventDefault();
      var err = document.getElementById('soMemoErr');
      err.textContent = '';
      var reason = f.reason.value.trim();
      if (!reason) { err.textContent = '사유를 적어 주세요.'; f.reason.focus(); return; }
      var kind = (f.querySelector('input[name="mkind"]:checked') || {}).value || 'must';
      var who = data.me.name;            // 본인만 — 서버도 로그인 이름으로만 저장
      var btn = f.querySelector('button[type="submit"]');
      btn.disabled = true;
      api('/api/satoff/memo', { method: 'POST', body: { day: f.day.value, name: who, kind: kind, reason: reason } }).then(function (j) {
        memoDraft = null;
        var said = md(j.day) + ' 토요일 메모를 남겼어요.' + (j.kind === 'must' && j.picked ? ' 휴무에 이름도 넣었어요.' : '');
        flash(j.warn ? said + ' ' + j.warn : said, j.warn ? 'warn' : 'ok');
        return load(true);
      }).catch(function (e) { err.textContent = e.message; })
        .then(function () { btn.disabled = false; });
    }
  });
  memoHost.addEventListener('click', function (ev) {
    var del = ev.target.closest('[data-mdel]');
    if (del) {
      var day = del.getAttribute('data-mdel'), who = del.getAttribute('data-name');
      if (!window.confirm(who + ' 님 ' + md(day) + ' 토요일 메모를 지울까요? 휴무 이름은 그대로 남아요.')) return;
      api('/api/satoff/memo', { method: 'POST', body: { day: day, name: who, on: false } }).then(function () {
        flash(md(day) + ' 토요일 메모를 지웠어요.', 'ok');
        return load(true);
      }).catch(function (e) { flash(e.message, 'err'); });
      return;
    }
  });


  // ── 본인 휴무 신청 창(11-25 부장님) — 달력 토요일을 누르면: 휴무 신청 / 꼭 쉬어야 해요 / 못 쉬어요 + 사유.
  //    메모를 따로 적지 않고 여기서 함께. 결과는 오른쪽 '휴무 메모'에 보임. 남의 이름은 관리자만(고르기 창)
  var selfEl = document.getElementById('soSelfModal');
  if (!selfEl) {
    selfEl = document.createElement('div');
    selfEl.className = 'modal fade gw-so-modal';
    selfEl.id = 'soSelfModal';
    selfEl.tabIndex = -1;
    selfEl.setAttribute('aria-labelledby', 'soSelfTitle');
    selfEl.setAttribute('aria-hidden', 'true');
    selfEl.innerHTML = '<div class="modal-dialog modal-dialog-centered modal-fullscreen-sm-down"><form class="modal-content" id="soSelfForm" novalidate>' +
      '<div class="modal-header"><h2 class="modal-title fs-5" id="soSelfTitle" tabindex="-1"></h2>' +
      '<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="닫기"></button></div>' +
      '<div class="modal-body"><p class="gw-so-hint" id="soSelfNow"></p><div class="gw-so-yield" id="soSelfYield" hidden></div>' +
      '<div class="gw-so-opts" role="radiogroup" aria-labelledby="soSelfTitle" id="soSelfOpts"></div>' +
      '<div class="gw-so-why" id="soSelfWhyBox" hidden><label class="form-label" for="soSelfWhy">사유</label>' +
      '<input type="text" class="form-control" id="soSelfWhy" maxlength="60" autocomplete="off" placeholder="가족여행, 결혼식, 현장 일정"></div>' +
      '<div class="gw-lv-err" id="soSelfErr" role="alert"></div></div>' +
      '<div class="modal-footer"><button type="button" class="btn btn-light" data-bs-dismiss="modal">닫기</button>' +
      '<button type="submit" class="btn btn-primary" id="soSelfSave">저장</button></div></form></div>';
    document.body.appendChild(selfEl);
  }
  var selfDay = null, selfModal = null;
  var SAVE_LABEL = { off: '휴무 신청', must: '꼭 쉬어야 해요로 신청', cant: '못 쉬어요로 저장', none: '' };
  function selfChoice() { var r = selfEl.querySelector('input[name="soSelfKind"]:checked'); return r ? r.value : ''; }
  function syncSelf() {
    var k = selfChoice(), s = satOf(selfDay), x = memoMap(data)[selfDay + '|' + data.me.name];
    document.getElementById('soSelfWhyBox').hidden = !(k === 'must' || k === 'cant');
    var lab = SAVE_LABEL[k] || '저장';
    if (k === 'none') lab = s && s.mine ? '휴무 취소' : '메모 지우기';
    if (k === 'off' && s && s.mine && !x) lab = '그대로 두기';
    var btn = document.getElementById('soSelfSave');
    btn.textContent = lab;
    btn.className = 'btn ' + (k === 'none' ? 'btn-outline-danger' : 'btn-primary');
    var need = (k === 'off' || k === 'must') && s && !s.mine ? yieldNeeded(s, data.me.name) : [];
    var y = document.getElementById('soSelfYield');
    y.hidden = !need.length;
    if (need.length) {
      y.innerHTML = '<i class="fas fa-star" aria-hidden="true"></i> ' + need.map(function (n) { return '<strong>' + esc(n.name) + '</strong> 님(' + esc(n.reason) + ')'; }).join(', ') +
        '이 이날 꼭 쉬어야 한다고 적어 두었어요. 지금 신청하면 그 자리가 모자라요. 다른 토요일을 먼저 생각해 주세요.';
    }
  }
  function openSelf(s) {
    selfDay = s.day;
    var me = data.me.name, x = memoMap(data)[s.day + '|' + me], full = isFull(s) && !s.mine;
    document.getElementById('soSelfTitle').textContent = md(s.day) + '(토) 휴무';
    document.getElementById('soSelfNow').textContent = '쉬는 사람 ' + s.names.length + (data.max_off != null ? '/' + data.max_off : '') + '명' +
      (s.names.length ? ' · ' + s.names.join(', ') : '') + (s.mine ? ' · 나는 이미 쉬기로 했어요' : '');
    var opts = [
      { v: 'off', icon: 'fa-umbrella-beach', t: '휴무 신청', sub: full ? '휴무가 마감됐어요.' : '이 토요일에 쉬어요.', off: full },
      { v: 'must', icon: 'fa-star', t: '꼭 쉬어야 해요', sub: full ? '지금은 메모만 남아요. 누가 양보하면 그때 이름을 넣을 수 있어요.' : '사유를 적으면 다른 분들이 보고 배려해 드려요.' },
      { v: 'cant', icon: 'fa-ban', t: '이 날은 못 쉬어요', sub: '현장 일정 등으로 이 토요일엔 쉴 수 없어요.' }
    ];
    if (s.mine || x) opts.push({ v: 'none', icon: 'fa-xmark', t: s.mine ? '휴무 취소' : '메모 지우기', sub: s.mine ? '이 토요일에 쉬지 않아요.' : '적어 둔 메모를 지워요.' });
    var cur = x ? x.kind : (s.mine ? 'off' : (full ? 'must' : 'off'));
    document.getElementById('soSelfOpts').innerHTML = opts.map(function (o, i) {
      return '<label class="gw-so-opt' + (o.off ? ' is-off' : '') + ' k-' + o.v + '"><input type="radio" class="form-check-input" name="soSelfKind" value="' + o.v + '"' +
        (o.v === cur ? ' checked' : '') + (o.off ? ' disabled' : '') + '><span class="gw-so-opt-body"><span class="gw-so-opt-t"><i class="fas ' + o.icon +
        '" aria-hidden="true"></i> ' + o.t + '</span><small>' + o.sub + '</small></span></label>';
    }).join('');
    document.getElementById('soSelfWhy').value = x ? x.reason : '';
    document.getElementById('soSelfErr').textContent = '';
    syncSelf();
    selfModal = showModal(selfEl, document.getElementById('soSelfTitle'));
  }
  selfEl.addEventListener('change', function (ev) { if (ev.target.name === 'soSelfKind') { syncSelf(); if (selfChoice() === 'must' || selfChoice() === 'cant') document.getElementById('soSelfWhy').focus(); } });
  document.getElementById('soSelfForm').addEventListener('submit', function (ev) {
    ev.preventDefault();
    var k = selfChoice(), s = satOf(selfDay), day = selfDay, me = data.me.name;
    var x = memoMap(data)[day + '|' + me];
    var why = document.getElementById('soSelfWhy').value.trim();
    var err = document.getElementById('soSelfErr');
    err.textContent = '';
    if (!k || !s) return;
    if ((k === 'must' || k === 'cant') && !why) { err.textContent = '사유를 적어 주세요.'; document.getElementById('soSelfWhy').focus(); return; }
    var toggle = function (on) { return api('/api/satoff/toggle', { method: 'POST', body: { day: day, on: on } }); };
    var memo = function (kind) { return api('/api/satoff/memo', { method: 'POST', body: { day: day, kind: kind, reason: why } }); };
    var unmemo = function () { return x ? api('/api/satoff/memo', { method: 'POST', body: { day: day, on: false } }) : Promise.resolve(); };
    var chain, said = md(day) + ' 토요일 ';
    if (k === 'off') {
      chain = (s.mine ? Promise.resolve({}) : toggle(true)).then(function (r) { return unmemo().then(function () { return r; }); });
      said += s.mine ? (x ? '메모를 지우고 휴무는 그대로 두었어요.' : '휴무는 그대로예요.') : '휴무를 신청했어요.';
    } else if (k === 'must') {
      chain = memo('must');
      said += '꼭 쉬어야 하는 날로 신청했어요.';
    } else if (k === 'cant') {
      chain = (s.mine ? toggle(false) : Promise.resolve({})).then(function () { return memo('cant'); });
      said += '못 쉬는 날로 적었어요.';
    } else {
      chain = (s.mine ? toggle(false) : Promise.resolve({})).then(function (r) { return unmemo().then(function () { return r; }); });
      said += s.mine ? '휴무를 취소했어요.' : '메모를 지웠어요.';
    }
    var btn = document.getElementById('soSelfSave');
    btn.disabled = true;
    chain.then(function (r) {
      if (selfModal) selfModal.hide();
      var warn = (r && r.warn) || '';
      if (k === 'must' && r && r.picked === false) said = md(day) + ' 토요일을 꼭 쉬어야 하는 날로 적었어요.';
      flash(warn ? said + ' ' + warn : said, warn ? 'warn' : 'ok');
      return load();
    }).catch(function (e) { err.textContent = e.message; load(); })
      .then(function () { btn.disabled = false; });
  });

  // ── 사람 고르기 창(관리자) ──
  var pickEl = document.getElementById('soPickModal');
  var pickDay = null;
  var pickModal = null;
  function openPick(s) {
    pickDay = s.day;
    document.getElementById('soPickTitle').textContent = md(s.day) + ' 토요일 휴무자 선택';
    document.getElementById('soPickHint').textContent = data.status === 'confirmed' || data.status === 'wait'
      ? (data.status === 'confirmed' ? '확정된 달이에요. 저장하면 바로 바뀌고 기록이 남아요.' : '결재 대기 중인 달이에요. 승인권자만 고칠 수 있어요.')
      : '휴무자를 고르세요. 이름 옆 숫자는 이번 달 휴무 횟수예요.';
    // 전체 휴무 날은 관리자만 이름을 뺀다(11-27 — SJ 가 10명을 모두 빼서 실제 쉬는 사람이 0명이 됨). 들어 있는 이름은 잠금
    var me = data.me || {};
    var lockFull = fullOff(s) && !(me.approver || me.editor);
    if (lockFull) document.getElementById('soPickHint').textContent = '전체 휴무 날이라 이름은 관리자만 뺄 수 있어요. 그날 나와야 하면 [휴무 출근]으로 남겨 주세요.';
    var counts = {};
    (data.members || []).forEach(function (m) { counts[m.name] = m.count; });
    var mmap = memoMap(data);
    document.getElementById('soPickList').innerHTML = (data.members || []).map(function (m, i) {
      var on = s.names.indexOf(m.name) !== -1;
      var mx = mmap[s.day + '|' + m.name];
      return '<div class="form-check"><input class="form-check-input" type="checkbox" id="soPk' + i + '" value="' + esc(m.name) + '"' + (on ? ' checked' : '') + (on && lockFull ? ' disabled' : '') + '>' +
        '<label class="form-check-label" for="soPk' + i + '">' + esc(m.name) + ' <small class="text-muted">' + (counts[m.name] || 0) + '번</small>' +
        (mx ? ' <span class="gw-so-pmemo ' + mx.kind + '"><i class="fas ' + (mx.kind === 'must' ? 'fa-star gw-so-mi must' : 'fa-ban gw-so-mi cant') +
          '" aria-hidden="true"></i>' + esc(mx.kind_label + ' · ' + mx.reason) + '</span>' : '') +
        '</label></div>';
    }).join('');
    document.getElementById('soPickErr').textContent = '';
    // 열 때 초점을 제목에 — 첫 체크박스로 가면 초점 테두리 때문에 첫 사람이 고른 것처럼 보였다(11-25 부장님)
    var pickTitle = document.getElementById('soPickTitle');
    pickTitle.tabIndex = -1;
    pickModal = showModal(pickEl, pickTitle);
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
    // 저장 한 번 = 요청 한 번(11-26) — 확정된 달이면 #휴가_일정 공지가 한 줄로 묶인다
    function oneByOne() {
      var chain = Promise.resolve();
      jobs.forEach(function (j) {
        chain = chain.then(function () {
          return api('/api/satoff/toggle', { method: 'POST', body: { day: pickDay, name: j.name, on: j.on } })
            .then(function (r) { if (r.warn) warns.push(r.warn); });
        });
      });
      return chain;
    }
    var run = !jobs.length ? Promise.resolve() :
      api('/api/satoff/pick', { method: 'POST', body: { day: pickDay, changes: jobs } })
        .then(function (r) { (r.warns || []).forEach(function (w) { warns.push(w); }); })
        .catch(function (e) {
          if (/오류 40[45]\)/.test(e.message)) return oneByOne();      // 서버 재시작 전(새 경로 없음) — 예전처럼
          throw e;
        });
    run.then(function () {
      if (pickModal) pickModal.hide();
      var said = jobs.length ? md(pickDay) + ' 토요일을 고쳤어요.' : '';
      flash(warns.length ? (said + ' ' + warns.join(' ')).trim() : said, warns.length ? 'warn' : 'ok');
      return load();
    }).catch(function (e) {
      document.getElementById('soPickErr').textContent = e.message;
      load();
    }).then(function () { btn.disabled = false; });
  });

  // ── 사유 창 ──
  var noteEl = document.getElementById('soNoteModal');
  if (noteEl) noteEl.classList.add('gw-so-modal');   // 창 안 칸 크기·정렬을 화면과 같게(11-25)
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
    noteModal = showModal(noteEl, document.getElementById('soNoteText'));
  }
  document.getElementById('soNoteForm').addEventListener('submit', function (ev) {
    ev.preventDefault();
    var input = document.getElementById('soNoteText');
    var text = input.value.trim();
    if (!text) { document.getElementById('soNoteErr').textContent = '사유를 적어 주세요.'; input.focus(); return; }
    var btn = document.getElementById('soNoteOk');
    btn.disabled = true;
    noteJob(text).then(function (msg) {
      if (noteModal) noteModal.hide();
      flash(msg, 'ok');
      return load();
    }).catch(function (e) { document.getElementById('soNoteErr').textContent = e.message; })
      .then(function () { btn.disabled = false; });
  });

  // ── 위쪽 단추(결재 단추는 actEl, 변경 신청·되돌리기는 barEl) ──
  function onAct(ev) {
    var b = ev.target.closest('[data-act]');
    if (!b || !data) return;
    var act = b.getAttribute('data-act');
    var k = data.month;
    var title = cur.y + '년 ' + cur.m + '월';
    if (act === 'done' || act === 'undone') {
      b.disabled = true;
      api('/api/satoff/done', { method: 'POST', body: { month: k, on: act === 'done' } }).then(function (j) {
        if (act === 'undone') flash(cur.m + '월 휴무를 다시 선택할 수 있어요. 다 고른 뒤 [선택 완료]를 눌러 주세요.', 'ok');
        else if (j.submitted) { flash('대상자 모두 선택을 완료해 결재를 올렸어요. 승인권자께 슬랙 알림이 가요.', 'ok'); watchNotify(k); }
        else if (j.all_done) { flash('대상자 모두 선택을 완료했어요. ' + coordName(data) + ' 검토한 뒤 결재를 올려요.', 'ok'); watchNotify(k); }
        else flash(cur.m + '월 휴무 선택을 완료했어요.' + ((j.left || []).length ? ' 아직 ' + j.left.length + '명이 남았어요.' : ''), 'ok');
        return load();
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
      return;
    }
    if (act === 'submit') {
      // 아직 선택 완료하지 않은 사람이 있으면 이름을 함께 묻는다(11-25 부장님 — 빠진 채 올라가지 않게)
      var miss = data.done ? (data.members || []).map(function (m) { return m.name; }).filter(function (n) { return data.done.indexOf(n) === -1; })
        : notPicked(data);
      if (!window.confirm(title + ' 스케줄 근무표를 결재에 올릴까요? 올린 뒤에는 승인권자만 고칠 수 있어요.' +
        (miss.length ? '\n\n아직 선택 완료하지 않은 사람 ' + miss.length + '명: ' + miss.join(', ') : ''))) return;
      b.disabled = true;
      // 슬랙 알림은 뒤에서 따로 나가 결과를 여기서 알 수 없다 → '보냈어요'라고 단정하지 않는다(11-21 점검).
      // 못 보냈으면 서버가 기록에 남기고, watchNotify 가 잠시 뒤 다시 읽어 알려 준다(11-24 조용한 실패 점검)
      api('/api/satoff/submit', { method: 'POST', body: { month: k } }).then(function () {
        flash('결재를 올렸어요. 승인권자께 슬랙 알림이 가요.', 'ok');
        watchNotify(k);
        return load();
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    } else if (act === 'approve') {
      // 승인하면 바로 #휴가_일정 공지·휴가 일수 반영 — 결재함 승인처럼 한 번 묻는다(11-21 점검)
      var picked = (data.members || []).filter(function (m) { return m.count; }).length;
      if (!window.confirm(title + ' 스케줄 근무표를 승인할까요?\n\n쉬는 사람 ' + picked + '명 · 승인하면 #휴가_일정에 공지되고 휴가 일수에 바로 반영돼요.')) return;
      b.disabled = true;
      api('/api/satoff/decide', { method: 'POST', body: { month: k, approve: true } }).then(function (j) {
        var extra = (j.affected || []).length ? ' 휴가 일수가 달라지는 신청 ' + j.affected.length + '건은 경영지원에 알림이 가요.' : '';
        flash('승인했어요. #휴가_일정에 공지가 올라가요.' + extra, 'ok');
        watchNotify(k);
        return load();
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    } else if (act === 'return') {
      openNote('돌려보내기', title + ' 스케줄 근무표를 다시 작성 중으로 돌려요.', '돌려보내는 이유', '돌려보내기', function (note) {
        return api('/api/satoff/decide', { method: 'POST', body: { month: k, approve: false, note: note } })
          .then(function () { watchNotify(k); return '돌려보냈어요. 올린 분께 슬랙 알림이 가요.'; });
      });
    } else if (act === 'full') {
      openFull();
    } else if (act === 'worked') {
      openWorked();
    } else if (act === 'undo') {
      draft = { add: [], remove: [] };
      render();
    } else if (act === 'change') {
      var add = draft.add.slice(), remove = draft.remove.slice();
      var what = remove.map(function (x) { return md(x) + '(토) 휴무 취소'; }).concat(add.map(function (x) { return md(x) + '(토) 휴무 신청'; })).join(', ');
      openNote('변경 신청', what, '바꾸는 이유', '신청', function (reason) {
        return api('/api/satoff/change', { method: 'POST', body: { month: k, add: add, remove: remove, reason: reason } })
          .then(function () { draft = { add: [], remove: [] }; watchNotify(k); return '변경 신청을 했어요. 승인권자 중 한 분이 승인하면 바뀌어요.'; });
      });
    }
  }
  barEl.addEventListener('click', onAct);
  actEl.addEventListener('click', onAct);

  // ── 변경 신청 처리 ──
  document.getElementById('soChanges').addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-cact]');
    if (!b) return;
    var id = b.getAttribute('data-id');
    var act = b.getAttribute('data-cact');
    var c = (data.changes || []).filter(function (x) { return String(x.id) === id; })[0];
    if (!c) return;
    if (act === 'approve') {
      if (!window.confirm(c.name + ' 님 변경을 승인할까요?\n' + c.what + '\n\n승인하면 #휴가_일정에 공지돼요.')) return;
      b.disabled = true;
      api('/api/satoff/change/' + id + '/decide', { method: 'POST', body: { approve: true } }).then(function () {
        flash(c.name + ' 님 변경을 승인했어요.', 'ok'); watchNotify(data.month); return load(true);
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    } else if (act === 'reject') {
      openNote('변경 반려', c.name + ' 님 ' + c.what, '반려하는 이유', '반려', function (note) {
        return api('/api/satoff/change/' + id + '/decide', { method: 'POST', body: { approve: false, note: note } })
          .then(function () { watchNotify(data.month); return c.name + ' 님 변경 신청을 반려했어요.'; });
      });
    } else if (act === 'cancel') {
      if (!window.confirm('변경 신청을 취소할까요?')) return;
      b.disabled = true;
      api('/api/satoff/change/' + id + '/cancel', { method: 'POST' }).then(function () {
        flash('변경 신청을 취소했어요.', 'ok'); return load(true);
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    }
  });

  // ── 전체 휴무 창(11-25 부장님) — 관리자. 토요일 하나를 고르고 지정하거나 풀기 ──
  var fullEl = null, fullModal = null;
  function fullModalEl() {
    if (fullEl) return fullEl;
    fullEl = document.createElement('div');
    fullEl.className = 'modal fade gw-so-modal';
    fullEl.id = 'soFullModal';
    fullEl.tabIndex = -1;
    fullEl.setAttribute('aria-labelledby', 'soFullTitle');
    fullEl.setAttribute('aria-hidden', 'true');
    fullEl.innerHTML =
      '<div class="modal-dialog modal-dialog-centered"><form class="modal-content" id="soFullForm" novalidate>' +
      '<div class="modal-header"><h2 class="modal-title fs-5" id="soFullTitle">전체 휴무</h2>' +
      '<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="닫기"></button></div>' +
      '<div class="modal-body">' +
      '<p class="gw-so-hint">고른 토요일엔 스케줄 근무 대상 모두가 쉬는 사람으로 들어가요. 각자 쉰 횟수에 하루씩 들어가요(격주 1일 차감).</p>' +
      '<p class="gw-so-admin-note small mb-3"><i class="fas fa-lock" aria-hidden="true"></i> 전체 휴무 설정은 관리자에게만 보여요.</p>' +
      '<div class="mb-3"><label class="form-label" for="soFullDay">토요일</label><select class="form-select" id="soFullDay" name="day"></select></div>' +
      '<div class="mb-1" id="soFullNoteWrap"><label class="form-label" for="soFullNote">메모</label>' +
      '<input type="text" class="form-control" id="soFullNote" name="note" maxlength="100" autocomplete="off" placeholder="예: 연말 전체 휴무"></div>' +
      '<div class="gw-lv-err" id="soFullErr" role="alert"></div></div>' +
      '<div class="modal-footer"><button type="button" class="btn btn-light" data-bs-dismiss="modal">닫기</button>' +
      '<button type="submit" class="btn btn-primary" id="soFullSave">전체 휴무로</button></div></form></div>';
    document.body.appendChild(fullEl);
    document.getElementById('soFullDay').addEventListener('change', syncFull);
    document.getElementById('soFullForm').addEventListener('submit', saveFull);
    return fullEl;
  }
  function isFullDay(day) { return !!(data && data.full_days && Object.prototype.hasOwnProperty.call(data.full_days, day)); }
  function syncFull() {
    var on = isFullDay(document.getElementById('soFullDay').value);
    var b = document.getElementById('soFullSave');
    b.textContent = on ? '전체 휴무 풀기' : '전체 휴무로';
    b.className = 'btn ' + (on ? 'btn-outline-danger' : 'btn-primary');
    document.getElementById('soFullNoteWrap').hidden = on;
  }
  function openFull() {
    if (!data) return;
    var el = fullModalEl();
    var sats = (data.saturdays || []).filter(function (s) { return !s.holiday; });
    document.getElementById('soFullDay').innerHTML = sats.map(function (s) {
      return '<option value="' + esc(s.day) + '">' + esc(md(s.day)) + '(토)' + (fullOff(s) ? ' · 전체 휴무' : '') + '</option>';
    }).join('');
    document.getElementById('soFullNote').value = '';
    document.getElementById('soFullErr').textContent = '';
    syncFull();
    fullModal = showModal(el, document.getElementById('soFullDay'));
  }
  function saveFull(ev) {
    ev.preventDefault();
    var day = document.getElementById('soFullDay').value;
    if (!day) return;
    var on = !isFullDay(day);
    if (!on && !window.confirm(md(day) + ' 토요일 전체 휴무를 풀까요? 전체 휴무로 넣은 이름만 빠지고, 각자 고른 사람은 남아요.')) return;
    var btn = document.getElementById('soFullSave');
    btn.disabled = true;
    api('/api/satoff/full', { method: 'POST', body: { day: day, on: on, note: document.getElementById('soFullNote').value.trim() } })
      .then(function (j) {
        if (fullModal) fullModal.hide();
        flash(md(day) + ' 토요일을 ' + (on ? '전체 휴무로 정했어요 (' + j.count + '명 넣음).' : '전체 휴무에서 풀었어요.'), 'ok');
        return load(true);
      })
      .catch(function (e) { document.getElementById('soFullErr').textContent = e.message; })
      .then(function () { btn.disabled = false; });
  }

  // ── 휴무 출근 기록 창(11-24 부장님) — 구두 승인 뒤 본인이. 창은 JS 가 만든다(템플릿 캐시와 무관하게) ──
  var wkEl = null, wkModal = null;
  function workedModal() {
    if (wkEl) return wkEl;
    wkEl = document.createElement('div');
    wkEl.className = 'modal fade gw-so-modal';
    wkEl.id = 'soWorkedModal';
    wkEl.tabIndex = -1;
    wkEl.setAttribute('aria-labelledby', 'soWkTitle');
    wkEl.setAttribute('aria-hidden', 'true');
    wkEl.innerHTML =
      '<div class="modal-dialog modal-dialog-centered modal-fullscreen-sm-down"><form class="modal-content" id="soWkForm" novalidate>' +
      '<div class="modal-header"><h2 class="modal-title fs-5" id="soWkTitle">휴무 출근 기록</h2>' +
      '<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="닫기"></button></div>' +
      '<div class="modal-body">' +
      '<p class="gw-so-hint">쉬기로 한 토요일에 급한 공사로 나오는 날을 남겨요. 남기면 그날은 근무로 바뀌어 차량·방문 배정에도 근무로 보여요.</p>' +
      '<div class="mb-3" id="soWkNameWrap"><label class="form-label" for="soWkName">누가</label><select class="form-select" id="soWkName" name="name"></select></div>' +
      '<div class="mb-3"><label class="form-label" for="soWkDay">쉬기로 했던 토요일</label><select class="form-select" id="soWkDay" name="day"></select></div>' +
      '<div class="mb-3"><label class="form-label" for="soWkBy">구두로 승인해 주신 분</label><select class="form-select" id="soWkBy" name="verbal_by"></select></div>' +
      '<div class="mb-1"><label class="form-label" for="soWkReason">어떤 일로 나오나요</label>' +
      '<input type="text" class="form-control" id="soWkReason" name="reason" maxlength="200" autocomplete="off" placeholder="예: ○○ 현장 공사"></div>' +
      '<div class="gw-lv-err" id="soWkErr" role="alert"></div></div>' +
      '<div class="modal-footer"><button type="button" class="btn btn-light" data-bs-dismiss="modal">닫기</button>' +
      '<button type="submit" class="btn btn-primary" id="soWkSave">기록</button></div></form></div>';
    document.body.appendChild(wkEl);
    document.getElementById('soWkName').addEventListener('change', fillWorkedDays);
    document.getElementById('soWkForm').addEventListener('submit', saveWorked);
    return wkEl;
  }
  function fillWorkedDays() {
    var who = document.getElementById('soWkName').value;
    var days = (data.saturdays || []).filter(function (s) { return !s.holiday && s.names.indexOf(who) !== -1; });
    var sel = document.getElementById('soWkDay');
    sel.innerHTML = days.length
      ? days.map(function (s) { return '<option value="' + esc(s.day) + '">' + esc(md(s.day)) + '(토)' + (s.past ? ' · 지난 날' : '') + '</option>'; }).join('')
      : '<option value="">이번 달 쉬기로 한 토요일이 없어요</option>';
    sel.disabled = !days.length;
    document.getElementById('soWkSave').disabled = !days.length;
  }
  function openWorked() {
    if (!data) return;
    var me = data.me || {};
    var el = workedModal();
    var others = me.manager || me.approver;
    var names = others ? (data.members || []).map(function (m) { return m.name; }) : [me.name];
    var nameSel = document.getElementById('soWkName');
    nameSel.innerHTML = names.map(function (n) { return '<option value="' + esc(n) + '">' + esc(n) + '</option>'; }).join('');
    if (names.indexOf(me.name) !== -1) nameSel.value = me.name;
    document.getElementById('soWkNameWrap').hidden = !others;
    document.getElementById('soWkBy').innerHTML = '<option value="">고르기</option>' + (data.approvers || []).map(function (n) {
      return '<option value="' + esc(n) + '"' + (n === me.name ? ' selected' : '') + '>' + esc(n) + '</option>';
    }).join('');
    document.getElementById('soWkReason').value = '';
    document.getElementById('soWkErr').textContent = '';
    fillWorkedDays();
    wkModal = showModal(el, document.getElementById(others ? 'soWkName' : 'soWkDay'));
  }
  function saveWorked(ev) {
    ev.preventDefault();
    var err = document.getElementById('soWkErr');
    var body = {
      on: true,
      name: document.getElementById('soWkName').value,
      day: document.getElementById('soWkDay').value,
      verbal_by: document.getElementById('soWkBy').value,
      reason: document.getElementById('soWkReason').value.trim()
    };
    if (!body.day) { err.textContent = '쉬기로 했던 토요일을 골라 주세요.'; return; }
    if (!body.verbal_by) { err.textContent = '구두로 승인해 주신 분을 골라 주세요.'; document.getElementById('soWkBy').focus(); return; }
    if (!body.reason) { err.textContent = '어떤 일로 나오는지 적어 주세요.'; document.getElementById('soWkReason').focus(); return; }
    var btn = document.getElementById('soWkSave');
    btn.disabled = true;
    api('/api/satoff/worked', { method: 'POST', body: body }).then(function () {
      if (wkModal) wkModal.hide();
      flash(body.name + ' 님 ' + md(body.day) + ' 토요일을 휴무 출근으로 남겼어요. 그날은 근무로 바뀌었어요.', 'ok');
      return load(true);
    }).catch(function (e) { err.textContent = e.message; })
      .then(function () { btn.disabled = false; });
  }
  if (sideEl) {
    sideEl.addEventListener('click', function (ev) {
      var b = ev.target.closest('[data-wundo]');
      if (!b) return;
      var day = b.getAttribute('data-wundo'), who = b.getAttribute('data-name');
      if (!window.confirm(who + ' 님 ' + md(day) + ' 토요일 휴무 출근을 되돌릴까요? 그날은 다시 쉬는 날이 돼요.')) return;
      b.disabled = true;
      api('/api/satoff/worked', { method: 'POST', body: { on: false, day: day, name: who } }).then(function (j) {
        var said = who + ' 님 ' + md(day) + ' 토요일을 다시 쉬는 날로 되돌렸어요.';
        flash(j && j.warn ? said + ' ' + j.warn : said, j && j.warn ? 'warn' : 'ok');
        return load(true);
      }).catch(function (e) { flash(e.message, 'err'); b.disabled = false; });
    });
  }

  load();
})();
