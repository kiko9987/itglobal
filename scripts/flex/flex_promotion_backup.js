// 플렉스 연차 촉진 기록 받기 (2026-11-18) — 읽기만 한다(검색 요청만, 아무것도 안 바꿈).
// flex.team 에 관리자로 로그인한 브라우저 창(휴가 관리 → 연차 촉진 화면)에서 개발자 도구/자동화로 실행.
// 끝나면 window.__promoSlim(JSON 글자)을 flex_export_20261106/원본/flex_promotion_YYYYMMDD.json 으로 저장한다.
// 저장한 파일은 scripts/leave_import_flex.py ④ 가 연차 촉진 기록(leave_promotions)으로 옮긴다.
//
// 이관 날 꼭 다시 받기: 11-23~12-02 사이 플렉스가 김종연·김태현에게 자동 촉구를 보낸다(소멸 2027-05-23).
// 다시 받지 않으면 우리 쪽이 같은 기간 촉구를 한 번 더 보낸다.
(async () => {
  const C = 'rM0AKmYbEl';
  const statuses = ['COMPANY_JOIN_SCHEDULED', 'RESIGNATION', 'LEAVE_OF_ABSENCE', 'LEAVE_OF_ABSENCE_SCHEDULED',
                    'RESIGNATION_SCHEDULED', 'IN_APPRENTICESHIP', 'IN_EMPLOY'];
  const post = (u, b) => fetch(u, {method: 'POST', headers: {'content-type': 'application/json'}, body: JSON.stringify(b)})
    .then(r => r.ok ? r.json() : Promise.reject(new Error(u + ' ' + r.status)));
  const docs = [];
  const thisYear = new Date().getFullYear();
  for (let y = 2021; y <= thisYear + 1; y++) {
    let tok = null;
    for (let i = 0; i < 20; i++) {
      const body = {searchBoostedAtFrom: `${y}-01-01`, searchBoostedAtTo: `${y}-12-31`, departmentIdHashes: null,
                    continuationSize: 50, userStatuses: statuses, documentDisplayStatuses: null};
      if (tok) body.continuationToken = tok;
      const d = await post(`/api/v2/time-off/customers/${C}/annual-time-off-promotion/documents/search-by-department-and-date-range`, body);
      for (const u of d.list || []) for (const doc of u.documents || []) docs.push(doc);
      if (!d.hasNext) break;
      tok = d.continuationToken;
    }
  }
  const L = await post(`/action/v2/search/customers/${C}/time-series/search-users?size=200`, {filter: {userStatuses: statuses}});
  const names = {};
  for (const x of L.list || []) names[x.userIdHash] = (x.basicInfo && (x.basicInfo.name || x.basicInfo.displayName)) || '?';
  const slim = {
    meta: {at: new Date().toISOString(), kind: 'flex-annual-time-off-promotion-backup',
           note: 'boostedAt 은 UTC(23:00 = 한국 다음 날 08:00). days = 촉구 때 남은 연차, planned = 직원이 낸 사용 계획일'},
    users: Object.fromEntries([...new Set(docs.map(d => d.userIdHash))].map(i => [i, names[i]])),
    docs: docs.map(d => ({u: d.userIdHash, task: d.taskKey, from: d.validUsageFrom, to: d.validUsageTo,
      dissipatedAt: d.dissipatedAt, assign: (d.promotionBucket || {}).totalAssignDays, boostedAt: d.boostedAt,
      days: d.boostedTimeOffDays, status: d.displayStatus, approval: d.approvalStatus || null, userBy: d.shouldUserWriteBy,
      mgrBy: d.shouldManagerWriteBy, planned: d.plannedDates || [], written: d.written, byMgr: d.writtenByManager || false,
      writtenAt: d.writtenAtTimestamp || null, canceled: d.canceled})),
  };
  window.__promoSlim = JSON.stringify(slim);
  return {docs: slim.docs.length, users: Object.keys(slim.users).length};
})();
