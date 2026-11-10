// 플렉스 인사·휴가 다시 받기 (2026-11-10) — 브라우저 창에서 flex.team 에 로그인한 채로 개발자 도구/자동화로 실행.
// 결과는 window.__flexR. 끝나면(window.__progR === 'done') 맨 아래 「내려받기」 부분을 따로 실행해 JSON 으로 저장한다.
// 급여는 플렉스에서 안 써서 뺐다. 결재 문서는 SINCE 이후 바뀐 것만 상세를 받는다.
// 받은 파일은 flex_export_20261106/원본/flex_refresh_YYYYMMDD.json 으로 옮긴 뒤
//   scripts/flex/flex_summary_refresh.py (요약 엑셀) → scripts/hr_import_flex.py (인사·휴가 기록 이관) 순서로.
const SINCE = '2026-11-05';   // 이 날짜(UTC) 이후 바뀐 결재 문서만 상세
window.__flexR = {meta: {started: new Date().toISOString(), source: 'flex.team', customer: 'rM0AKmYbEl', kind: 'refresh-hr-leave'},
                  company: {}, users: {}, timeOffUses: [], approvalDocs: [], approvalList: [], errors: []};
window.__progR = 'start';
(async () => {
  const F = window.__flexR, C = 'rM0AKmYbEl';
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const req = async (u, b, method) => {
    for (let i = 0; i < 3; i++) {
      try {
        const opt = (b === undefined && !method) ? {} : {method: method || 'POST', headers: {'content-type': 'application/json'},
                                                          body: b === undefined ? undefined : JSON.stringify(b)};
        const r = await fetch(u, opt); await sleep(60);
        const t = await r.text(); let d; try { d = JSON.parse(t); } catch (e) { d = t; }
        if (r.status >= 500 || r.status === 429) { await sleep(1500 * (i + 1)); continue; }
        if (r.status >= 400) { F.errors.push({u, s: r.status, d: String(t).slice(0, 300)}); return null; }
        return d;
      } catch (e) { await sleep(1000); }
    }
    F.errors.push({u, s: 'fail'}); return null;
  };
  const statuses = ['COMPANY_JOIN_SCHEDULED', 'RESIGNATION', 'LEAVE_OF_ABSENCE', 'LEAVE_OF_ABSENCE_SCHEDULED',
                    'RESIGNATION_SCHEDULED', 'IN_APPRENTICESHIP', 'IN_EMPLOY'];
  const co = F.company;
  window.__progR = 'company';
  co.info = await req(`/api/v2/core/customers/${C}/info`);
  co.departmentsTree = await req('/action/v2/core/departments/search/tree', {customerIdHashes: [C]});
  for (const k of ['job-ranks', 'job-roles', 'job-titles']) co[k] = await req(`/api/v2/core/customers/${C}/${k}`);
  co.timeOffForms = await req(`/api/v2/time-off/customers/${C}/time-off-forms?includeInactivatedTimeOffForms=true`);
  co.userDocumentTypes = await req(`/action/v2/core/customers/${C}/user-documents/get`, {});
  co.notices = await req('/api/v3/notice/posts:search', {size: 500});
  const list = await req('/action/v2/search/customers/' + C + '/time-series/search-users?size=200', {filter: {userStatuses: statuses}});
  co.userList = list;
  const docTypeIds = ((co.userDocumentTypes && (co.userDocumentTypes.userDocuments || co.userDocumentTypes.list || co.userDocumentTypes)) || []);
  const docIds = Array.isArray(docTypeIds) ? docTypeIds.map(x => x.userDocumentIdHash || x.idHash).filter(Boolean) : [];
  const now = Date.now(); let n = 0;
  for (const usr of (list && list.list) || []) {
    const u = usr.userIdHash; n++;
    window.__progR = `user ${n}/${list.list.length}`;
    const o = F.users[u] = {name: usr.basicInfo.name, search: usr};
    const base = `/api/v2/core/customers/${C}/users/${u}`;
    o.employee = await req(`/api/v2/core/user-employees/${u}`);
    o.personalBundle = await req(`/action/v2/core/bundle/user-personals/${u}/get`, undefined, 'POST');
    o.employeeBundle = await req(`/action/v2/core/bundle/user-employees/${u}/get`, undefined, 'POST');
    o.employmentContracts = await req('/action/v2/core/bundle/user-employment-contracts/search', {userIdHash: u});
    o.salaryContracts = await req(`/api/v2/core/user-salary-contracts/by-user/${u}`);
    o.personnelAppointments = await req(`${base}/personnel-appointments?includeWaiting=true`);
    o.resignations = await req(`/action/v2/core/user-resignations/by-user/${u}`);
    o.leaveOfAbsences = await req(`/action/v2/core/user-leave-of-absences/by-user/${u}`);
    o.workExperiences = await req('/action/v2/core/work-experiences/search-with-approval', {userIdHash: u});
    o.educationExperiences = await req('/action/v2/core/education-experiences/search-with-approval', {userIdHash: u});
    o.dependentFamilies = await req('/action/v2/core/dependent-families/search-with-approval', {userIdHash: u});
    o.careers = await req(`${base}/careers/search?sortProperty=END_DATE&sortDirection=DESC&size=200`);
    o.licenses = await req(`${base}/user-license`);
    o.memos = await req(`/action/v2/core/user-memos/search?userIdHash=${u}&size=200`);
    o.documentLegacies = await req(`${base}/user-documents/legacies`);
    o.documentFiles = docIds.length ? await req(`/action/v2/core/customers/${C}/users/${u}/user-documents/files/get`,
                                                {inquiries: docIds.map(id => ({userDocumentIdHash: id}))}) : null;
    o.certificateHistories = await req(`/api/v2/core/customers/${C}/user-certificate-issuance-histories?userIdHash=${u}`);
    o.annualPolicy = await req(`/api/v2/time-off/customers/${C}/annual-time-off-policies/users/${u}`);
    o.usableBuckets = await req(`/api/v2/time-off/users/${u}/usable-grouped-time-off-buckets/by-time-stamp/${now}`);
    o.timeOffContext = await req(`/api/v2/time-off/customers/${C}/users/${u}/context`);
    o.annualStatusByYear = {};
    for (let y = 2024; y <= 2027; y++) o.annualStatusByYear[y] = await req(`/api/v2/time-off/users/${u}/annual-time-off-status/by-search-year/${y}`);
    o.timeOffUsesAll = await req(`/api/v2/time-off/users/${u}/time-off-uses/by-use-date-range/1514732400000..1861887600000?eventTypes=REGISTER&eventTypes=INVALIDATE&eventTypes=CANCEL`);
  }
  window.__progR = 'approval-list';
  let tok = '';
  const abody = {filter: {statuses: [], templateKeys: [], writerHashedIds: [], approverTargets: [], referrerTargets: [], starred: false},
                 search: {keyword: '', type: 'ALL'}};
  let stop = false;
  for (let g = 0; g < 300 && !stop; g++) {
    const d = await req('/action/v3/approval-document/user-boxes/search?size=50' + (tok ? '&continuationToken=' + encodeURIComponent(tok) : '') +
                        '&sortType=LAST_UPDATED_AT&direction=DESC', abody);
    if (!d) break;
    for (const x of d.documents || []) {
      const doc = x.document || {};
      const upd = String(doc.lastUpdatedAt || doc.updatedAt || x.lastUpdatedAt || '');
      F.approvalList.push({key: doc.documentKey, upd, title: doc.title, status: doc.status});
      if (upd && upd < SINCE) stop = true;
    }
    window.__progR = 'approval-list ' + F.approvalList.length;
    if (!d.hasNext || !d.continuationToken) break; tok = d.continuationToken;
  }
  for (const x of F.approvalList.filter(x => x.upd >= SINCE)) {
    const d = await req('/api/v3/approval-document/approval-documents/' + x.key);
    if (d) F.approvalDocs.push(d);
  }
  F.meta.finished = new Date().toISOString();
  window.__progR = 'done';
})().catch(e => { window.__progR = 'crash: ' + e; });

/* ── 내려받기(끝난 뒤 따로 실행) ──
const s = JSON.stringify(window.__flexR); const b = new Blob([s], {type: 'application/json'});
const a = document.createElement('a'); a.href = URL.createObjectURL(b);
a.download = 'flex_refresh_' + new Date().toISOString().slice(0, 10).replace(/-/g, '') + '.json';
document.body.appendChild(a); a.click(); a.remove();
// 브라우저 창에선 저장이 Downloads 의 .tmp 로 멈출 수 있음 → 크기 확인 후 그 파일을 복사해 JSON 으로 쓴다.
*/
