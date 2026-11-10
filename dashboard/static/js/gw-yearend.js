/* 그룹웨어 연말정산 — 직원(mode 'me')·경영지원 현황판(mode 'admin').
   간소화 자료는 ① 홈택스 일괄제공 동의 ② PDF 직접 업로드 둘 다 받는다. 계산·신고는 세무사. */
(function () {
  'use strict';

  var CFG = window.GW_YE || {};

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function api(url, opts) {
    opts = opts || {};
    var init = { method: opts.method || 'GET', credentials: 'same-origin' };
    if (opts.form) { init.body = opts.form; }
    else if (opts.body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(opts.body); }
    return fetch(url, init).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok || j.success === false) throw new Error(j.message || ('처리하지 못했습니다 (' + r.status + ')'));
        return j;
      });
    });
  }

  function md(s) {
    if (!s) return '';
    var p = s.slice(0, 10).split('-');
    return Number(p[1]) + '월 ' + Number(p[2]) + '일';
  }

  function kb(n) {
    if (!n) return '';
    return n >= 1048576 ? (Math.round(n / 104857.6) / 10) + 'MB' : Math.max(1, Math.round(n / 1024)) + 'KB';
  }

  function badge(status, label) {
    var cls = { draft: 'gw-badge-resigned', submitted: 'gw-badge-leave', returned: 'gw-badge-reject', done: 'gw-badge-active' }[status] || 'gw-badge-resigned';
    return '<span class="gw-badge ' + cls + '">' + esc(label) + '</span>';
  }

  function docList(docs, deletable) {
    if (!docs.length) return '';
    return '<ul class="gw-ye-files">' + docs.map(function (d) {
      return '<li><i class="fas fa-file-alt" aria-hidden="true"></i> <a href="/api/hr/documents/' + d.id + '?view=1" target="_blank" rel="noopener">' +
        esc(d.file_name) + '</a> <span class="text-muted">' + kb(d.size) + '</span>' +
        (deletable ? ' <button type="button" class="btn btn-sm btn-link text-danger p-0 ms-1" data-ye-del="' + d.id + '">삭제</button>' : '') + '</li>';
    }).join('') + '</ul>';
  }

  function uploader(kind, label, multiple) {
    return '<form class="gw-doc-upload" data-ye-up="' + kind + '">' +
      '<input type="file" class="form-control form-control-sm" name="file" ' + (multiple ? 'multiple ' : '') +
      'accept=".pdf,.jpg,.jpeg,.png,.heic" aria-label="' + esc(label) + '">' +
      '<button type="submit" class="btn btn-sm btn-outline-primary">업로드</button>' +
      '<span class="gw-form-error" role="alert"></span></form>';
  }

  // ── 직원 ─────────────────────────────────────────
  if (CFG.mode === 'me') {
    var box = document.getElementById('yeBox');
    var data = null;

    var load = function () {
      return api('/api/yearend/me').then(function (j) { data = j.ye; render(); })
        .catch(function (e) { box.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>'; });
    };

    var save = function (patch) {
      return api('/api/yearend/me', { method: 'POST', body: patch }).then(function (j) { data = j.ye; render(); })
        .catch(function (e) { window.alert(e.message); load(); });
    };

    var guideBulk = '<ol class="gw-ye-steps">' +
      '<li>회사가 1월 초에 홈택스에 연말정산 대상자 명단을 등록해요.</li>' +
      '<li>홈택스(또는 손택스 앱) → 연말정산 간소화 → <b>간소화자료 일괄제공</b>에서 회사 제공에 <b>동의</b>해 주세요.</li>' +
      '<li>부양가족 공제를 받으면 부양가족도 홈택스에서 <b>자료 제공 동의</b>를 해 둬야 함께 넘어가요.</li>' +
      '<li>동의했으면 아래에 체크해 주세요. 자료는 국세청이 회사·세무사에 바로 보내요.</li></ol>';
    var guidePdf = '<ol class="gw-ye-steps">' +
      '<li>홈택스(또는 손택스 앱) → 연말정산 간소화에서 자료를 조회해요.</li>' +
      '<li><b>한 번에 내려받기</b>로 PDF를 받아 아래에 업로드해 주세요(부양가족 자료 포함).</li></ol>';

    var render = function () {
      var v = data, s = v.submission, season = v.season;
      var open = season.status === 'open';
      var editable = open && (s.status === 'draft' || s.status === 'returned');
      document.getElementById('yeTitle').textContent = v.year + '년 귀속 연말정산';
      var h = '';

      h += '<div class="gw-panel gw-ye-head"><div>' + badge(s.status, s.status_label) +
        ' <span class="gw-ye-season">' + esc(season.status_label) + (season.due_date ? ' · 마감 ' + esc(md(season.due_date)) : '') + '</span></div>' +
        (season.notice ? '<div class="gw-ye-notice">' + esc(season.notice) + '</div>' : '') + '</div>';

      if (s.status === 'returned' && s.returned_reason) {
        h += '<div class="gw-callout gw-callout-warn"><b>보완 요청</b> · ' + esc(s.returned_reason) + '</div>';
      }

      if (!open) {
        h += '<div class="gw-panel"><h2 class="gw-panel-title">' + (season.status === 'closed' ? '연말정산 자료 받기가 끝났어요' : '1월에 시작해요') + '</h2>' +
          '<p class="gw-panel-desc">' + (season.status === 'closed' ? '결과(원천징수영수증)가 나오면 아래와 내 정보 → 서류에 올라와요.' :
            '홈택스 간소화 자료는 1월 15일쯤 열려요. 미리 홈택스(손택스) 로그인 방법을 확인해 두고, 부양가족 공제를 받을 분은 부양가족 자료 제공 동의를 받아 두세요.') + '</p>' +
          '<details><summary class="gw-log">간소화 자료 내는 두 가지 방법 미리 보기</summary><div class="gw-sec-title">① 홈택스 일괄제공 동의</div>' + guideBulk +
          '<div class="gw-sec-title">② PDF 직접 업로드</div>' + guidePdf + '</details></div>';
      } else {
        // 1. 방법
        h += '<div class="gw-panel"><h2 class="gw-panel-title">1. 간소화 자료 내는 방법</h2><div class="gw-ye-methods">' +
          '<label class="gw-ye-method' + (s.method === 'bulk' ? ' on' : '') + '"><input type="radio" name="yeMethod" value="bulk"' + (s.method === 'bulk' ? ' checked' : '') + (editable ? '' : ' disabled') + '>' +
          '<span><b>홈택스 일괄제공 동의</b> <span class="gw-badge gw-badge-active">추천</span><br><small>동의만 하면 국세청이 회사로 바로 보내요</small></span></label>' +
          '<label class="gw-ye-method' + (s.method === 'pdf' ? ' on' : '') + '"><input type="radio" name="yeMethod" value="pdf"' + (s.method === 'pdf' ? ' checked' : '') + (editable ? '' : ' disabled') + '>' +
          '<span><b>간소화 PDF 직접 업로드</b><br><small>홈택스에서 받은 PDF를 올려요</small></span></label></div>';
        if (s.method === 'bulk') {
          h += guideBulk +
            '<label class="gw-ye-check"><input type="checkbox" id="yeBulk"' + (s.bulk_consent ? ' checked' : '') + (editable ? '' : ' disabled') + '> 홈택스에서 회사 자료 제공에 동의했어요</label>' +
            '<label class="gw-ye-check"><input type="checkbox" id="yeFamCon"' + (s.family_consent ? ' checked' : '') + (editable ? '' : ' disabled') + '> 부양가족도 자료 제공 동의를 받았어요 <span class="text-muted">(부양가족 공제를 받는 경우)</span></label>';
        } else if (s.method === 'pdf') {
          h += guidePdf + docList(v.simple_docs, editable) + (editable ? uploader('simple', '간소화 PDF', false) : '');
        }
        h += '</div>';

        // 2. 부양가족
        var fams = v.families || [];
        h += '<div class="gw-panel"><h2 class="gw-panel-title">2. 부양가족 확인</h2>';
        if (fams.length) {
          h += '<div class="gw-table-wrap"><table class="gw-table gw-appt-table"><thead><tr><th>이름</th><th>관계</th><th>기본공제</th><th>건보 피부양</th></tr></thead><tbody>' +
            fams.map(function (f) {
              return '<tr><td>' + esc(f.name) + '</td><td>' + esc(f.relation) + '</td><td>' + (f.default_deduction ? '예' : '-') + '</td><td>' + (f.health_dependent ? '예' : '-') + '</td></tr>';
            }).join('') + '</tbody></table></div>';
        } else {
          h += '<p class="gw-panel-desc">등록된 부양가족이 없어요.</p>';
        }
        h += '<div class="gw-ye-methods">' +
          '<label class="gw-ye-method sm' + (s.family_status === 'none' ? ' on' : '') + '"><input type="radio" name="yeFam" value="none"' + (s.family_status === 'none' ? ' checked' : '') + (editable ? '' : ' disabled') + '> 작년과 같아요(변경 없음)</label>' +
          '<label class="gw-ye-method sm' + (s.family_status === 'changed' ? ' on' : '') + '"><input type="radio" name="yeFam" value="changed"' + (s.family_status === 'changed' ? ' checked' : '') + (editable ? '' : ' disabled') + '> 바뀐 게 있어요</label></div>';
        if (s.family_status === 'changed') {
          h += '<textarea class="form-control form-control-sm mt-2" id="yeFamNote" rows="2" maxlength="1000" placeholder="예: 2026년 6월 자녀 출생, 어머니 기본공제 추가"' + (editable ? '' : ' disabled') + '>' + esc(s.family_note || '') + '</textarea>';
        }
        h += '</div>';

        // 3. 추가 증빙
        h += '<div class="gw-panel"><h2 class="gw-panel-title">3. 추가 증빙 <span class="gw-panel-sub">해당될 때만</span></h2>' +
          '<p class="gw-panel-desc">간소화에 안 나오는 것: 월세 계약서·이체 내역, 안경·교복 영수증, 간소화에 빠진 의료비·기부금 영수증 등</p>' +
          docList(v.extra_docs, editable) + (editable ? uploader('extra', '추가 증빙', true) : '') +
          '<textarea class="form-control form-control-sm mt-2" id="yeNote" rows="2" maxlength="1000" placeholder="경영지원·세무사에 전할 말(선택)"' + (editable ? '' : ' disabled') + '>' + esc(s.note || '') + '</textarea></div>';

        // 제출
        h += '<div class="gw-panel">';
        if (editable) {
          var probs = v.problems || [];
          h += probs.length ? '<ul class="gw-ye-probs">' + probs.map(function (p) { return '<li>' + esc(p) + '</li>'; }).join('') + '</ul>' :
            '<p class="gw-panel-desc">다 됐어요. 제출하면 경영지원이 확인해요.</p>';
          h += '<button type="button" class="btn btn-primary" id="yeSubmit"' + (probs.length ? ' disabled' : '') + '>제출하기</button>';
        } else if (s.status === 'submitted') {
          h += '<p class="gw-panel-desc">' + esc(md(s.submitted_at)) + '에 제출했어요. 경영지원이 확인 중이에요.</p>' +
            '<button type="button" class="btn btn-outline-secondary btn-sm" id="yeReopen">수정하기</button>';
        } else if (s.status === 'done') {
          h += '<p class="gw-panel-desc">경영지원이 확인을 마쳤어요.</p>';
        }
        h += '</div>';
      }

      if (v.result_docs && v.result_docs.length) {
        h += '<div class="gw-panel"><h2 class="gw-panel-title">결과 · 근로소득 원천징수영수증</h2>' + docList(v.result_docs, false) + '</div>';
      }
      box.innerHTML = h;
      bind(editable);
    };

    var bind = function (editable) {
      if (!editable) {
        var ro = document.getElementById('yeReopen');
        if (ro) ro.addEventListener('click', function () {
          api('/api/yearend/me/reopen', { method: 'POST' }).then(function (j) { data = j.ye; render(); })
            .catch(function (e) { window.alert(e.message); });
        });
        return;
      }
      box.querySelectorAll('input[name=yeMethod]').forEach(function (r) {
        r.addEventListener('change', function () { save({ method: r.value }); });
      });
      box.querySelectorAll('input[name=yeFam]').forEach(function (r) {
        r.addEventListener('change', function () { save({ family_status: r.value }); });
      });
      var b = document.getElementById('yeBulk');
      if (b) b.addEventListener('change', function () { save({ bulk_consent: b.checked }); });
      var fc = document.getElementById('yeFamCon');
      if (fc) fc.addEventListener('change', function () { save({ family_consent: fc.checked }); });
      ['yeFamNote', 'yeNote'].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.addEventListener('change', function () {
          var patch = {}; patch[id === 'yeNote' ? 'note' : 'family_note'] = el.value; save(patch);
        });
      });
      box.querySelectorAll('[data-ye-del]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          if (!window.confirm('이 파일을 지울까요?')) return;
          api('/api/yearend/me/documents/' + btn.getAttribute('data-ye-del'), { method: 'DELETE' })
            .then(function (j) { data = j.ye; render(); }).catch(function (e) { window.alert(e.message); });
        });
      });
      box.querySelectorAll('form[data-ye-up]').forEach(function (form) {
        form.addEventListener('submit', function (e) {
          e.preventDefault();
          var err = form.querySelector('.gw-form-error');
          var files = Array.prototype.slice.call(form.elements.file.files);
          if (!files.length) { err.textContent = '파일을 골라 주세요.'; return; }
          var big = files.filter(function (f) { return f.size > 8 * 1024 * 1024; });
          if (big.length) { err.textContent = '8MB까지 업로드할 수 있습니다.'; return; }
          var btn = form.querySelector('button'); btn.disabled = true; err.textContent = '';
          var chain = Promise.resolve();
          files.forEach(function (f) {
            chain = chain.then(function () {
              var fd = new FormData(); fd.append('file', f); fd.append('kind', form.getAttribute('data-ye-up'));
              return api('/api/yearend/me/documents', { method: 'POST', form: fd }).then(function (j) { data = j.ye; });
            });
          });
          chain.then(render).catch(function (x) { err.textContent = x.message; btn.disabled = false; load(); });
        });
      });
      var sub = document.getElementById('yeSubmit');
      if (sub) sub.addEventListener('click', function () {
        // 메모 칸에 쓰다 바로 누른 경우도 저장
        var patch = {};
        var n = document.getElementById('yeNote'); if (n) patch.note = n.value;
        var fn = document.getElementById('yeFamNote'); if (fn) patch.family_note = fn.value;
        sub.disabled = true;
        api('/api/yearend/me', { method: 'POST', body: patch })
          .then(function () { return api('/api/yearend/me/submit', { method: 'POST' }); })
          .then(function (j) { data = j.ye; render(); })
          .catch(function (e) { window.alert(e.message); load(); });
      });
    };

    load();
    return;
  }

  // ── 경영지원 현황판 ───────────────────────────────
  if (CFG.mode !== 'admin') return;
  var root = document.getElementById('yeAdmin');
  var A = null;

  var loadAdmin = function () {
    return api('/api/yearend/admin').then(function (j) { A = j; renderAdmin(); })
      .catch(function (e) { root.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>'; });
  };

  var renderAdmin = function () {
    var s = A.season, rows = A.rows;
    document.getElementById('yeAdminTitle').textContent = A.year + '년 귀속 연말정산 현황';
    var cnt = { draft: 0, submitted: 0, returned: 0, done: 0 };
    rows.forEach(function (r) { cnt[r.status] = (cnt[r.status] || 0) + 1; });
    var h = '<div class="gw-panel"><div class="gw-ye-season-form">' +
      '<label>진행 <select class="form-select form-select-sm" id="yeSt">' +
      [['prep', '준비 중(직원에게 안내만)'], ['open', '진행 중(자료 받기)'], ['closed', '마감']].map(function (o) {
        return '<option value="' + o[0] + '"' + (s.status === o[0] ? ' selected' : '') + '>' + o[1] + '</option>';
      }).join('') + '</select></label>' +
      '<label>직원 마감일 <input type="date" class="form-control form-control-sm" id="yeDue" value="' + esc(s.due_date || '') + '"></label>' +
      '<label class="gw-ye-wide">직원 안내 문구 <input class="form-control form-control-sm" id="yeNotice" maxlength="300" value="' + esc(s.notice || '') + '" placeholder="예: 1월 31일까지 꼭 제출해 주세요."></label>' +
      '<button type="button" class="btn btn-sm btn-primary" id="yeSaveSeason">저장</button></div>' +
      '<div class="gw-ye-actions">' +
      '<a class="btn btn-sm btn-outline-secondary" href="/api/yearend/admin/roster.xlsx?year=' + A.year + '" id="yeRoster"><i class="fas fa-file-excel me-1" aria-hidden="true"></i>홈택스 등록 명단(엑셀)</a>' +
      '<a class="btn btn-sm btn-outline-secondary" href="/api/yearend/admin/bundle.zip?year=' + A.year + '"><i class="fas fa-file-archive me-1" aria-hidden="true"></i>전체 자료 묶음(세무사 전달용)</a>' +
      '<button type="button" class="btn btn-sm btn-outline-primary" id="yeNotify"' + (s.status === 'open' ? '' : ' disabled title="진행 중일 때만"') + '><i class="fab fa-slack me-1" aria-hidden="true"></i>미제출·보완 대상 슬랙 알림</button>' +
      '</div><div class="gw-log mt-1">명단에는 주민등록번호 전체가 들어가요. 명단·묶음을 받으면 기록이 남습니다.</div></div>';

    h += '<div class="gw-stats">' +
      '<div class="gw-stat"><div class="k">전체</div><div class="v">' + rows.length + '명</div></div>' +
      '<div class="gw-stat"><div class="k">제출</div><div class="v">' + cnt.submitted + '명</div></div>' +
      '<div class="gw-stat"><div class="k">완료</div><div class="v">' + cnt.done + '명</div></div>' +
      '<div class="gw-stat"><div class="k">작성 중·보완</div><div class="v">' + (cnt.draft + cnt.returned) + '명</div></div></div>';

    h += '<div class="gw-panel gw-table-wrap"><table class="gw-table gw-appt-table"><thead><tr><th>이름</th><th>상태</th><th>방식</th><th>홈택스 동의</th><th>부양가족</th><th>간소화 PDF</th><th>추가 증빙</th><th>원천징수영수증</th><th>제출</th><th></th></tr></thead><tbody>' +
      rows.map(function (r) {
        var method = { bulk: '일괄제공', pdf: 'PDF' }[r.method] || '-';
        var consent = r.method === 'bulk' ? (r.bulk_consent ? '예' : '<span class="text-danger">아직</span>') + (r.family_consent ? ' · 가족 예' : '') : '-';
        var fam = { none: '변경 없음', changed: '<span title="' + esc(r.family_note || '') + '">변경 있음</span>' }[r.family_status] || '-';
        var acts = '<a class="btn btn-sm btn-outline-secondary" href="/hr/people#' + r.person_id + '">자료</a>';
        if (r.status === 'submitted') acts += ' <button type="button" class="btn btn-sm btn-outline-success" data-done="' + r.person_id + '">완료</button>' +
          ' <button type="button" class="btn btn-sm btn-outline-danger" data-return="' + r.person_id + '">보완 요청</button>';
        if (r.status === 'done') acts += ' <button type="button" class="btn btn-sm btn-outline-secondary" data-undone="' + r.person_id + '">완료 취소</button>';
        acts += ' <label class="btn btn-sm btn-outline-primary mb-0">영수증 업로드<input type="file" hidden accept=".pdf,.jpg,.jpeg,.png" data-result="' + r.person_id + '"></label>';
        return '<tr><td><b>' + esc(r.name) + '</b><br><span class="text-muted small">' + esc(r.department) + '</span></td><td>' + badge(r.status, r.status_label) +
          (r.returned_reason && r.status === 'returned' ? '<br><span class="small text-muted">' + esc(r.returned_reason) + '</span>' : '') + '</td><td>' + method + '</td><td>' + consent +
          '</td><td>' + fam + '</td><td>' + (r.simple || '-') + '</td><td>' + (r.extra || '-') + '</td><td>' + (r.result || '-') + '</td><td>' + esc(md(r.submitted_at)) + '</td><td class="gw-doc-act">' + acts + '</td></tr>';
      }).join('') + '</tbody></table></div>';
    root.innerHTML = h;
    bindAdmin();
  };

  var post = function (url, body) { return api(url + (url.indexOf('?') < 0 ? '?year=' + A.year : ''), { method: 'POST', body: body }); };

  var bindAdmin = function () {
    document.getElementById('yeSaveSeason').addEventListener('click', function () {
      var st = document.getElementById('yeSt').value;
      if (st === 'open' && A.season.status !== 'open' && !window.confirm('직원들이 연말정산 자료를 낼 수 있게 엽니다. 계속할까요?')) return;
      post('/api/yearend/admin/season', { status: st, due_date: document.getElementById('yeDue').value, notice: document.getElementById('yeNotice').value })
        .then(loadAdmin).catch(function (e) { window.alert(e.message); });
    });
    var nb = document.getElementById('yeNotify');
    if (nb) nb.addEventListener('click', function () {
      if (!window.confirm('아직 제출하지 않았거나 보완 요청을 받은 직원에게 슬랙 알림을 보냅니다.')) return;
      nb.disabled = true;
      post('/api/yearend/admin/notify', {}).then(function (j) {
        window.alert('보냄 ' + j.sent.length + '명' + (j.skipped.length ? '\n못 보냄(슬랙 연결 없음): ' + j.skipped.join(', ') : ''));
        nb.disabled = false;
      }).catch(function (e) { window.alert(e.message); nb.disabled = false; });
    });
    root.querySelectorAll('[data-done]').forEach(function (b) {
      b.addEventListener('click', function () {
        post('/api/yearend/admin/' + b.getAttribute('data-done') + '/status', { status: 'done' }).then(loadAdmin).catch(function (e) { window.alert(e.message); });
      });
    });
    root.querySelectorAll('[data-undone]').forEach(function (b) {
      b.addEventListener('click', function () {
        post('/api/yearend/admin/' + b.getAttribute('data-undone') + '/status', { status: 'submitted' }).then(loadAdmin).catch(function (e) { window.alert(e.message); });
      });
    });
    root.querySelectorAll('[data-return]').forEach(function (b) {
      b.addEventListener('click', function () {
        var reason = window.prompt('보완 요청 사유(직원에게 슬랙으로 갑니다)');
        if (!reason || !reason.trim()) return;
        post('/api/yearend/admin/' + b.getAttribute('data-return') + '/status', { status: 'returned', reason: reason }).then(loadAdmin).catch(function (e) { window.alert(e.message); });
      });
    });
    root.querySelectorAll('input[data-result]').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var f = inp.files[0];
        if (!f) return;
        var fd = new FormData(); fd.append('file', f);
        api('/api/yearend/admin/' + inp.getAttribute('data-result') + '/result?year=' + A.year, { method: 'POST', form: fd })
          .then(function (j) { if (!j.notified) window.alert('업로드했어요. 슬랙 알림은 보내지 못했어요(슬랙 연결 없음).'); loadAdmin(); })
          .catch(function (e) { window.alert(e.message); });
      });
    });
  };

  loadAdmin();
})();
