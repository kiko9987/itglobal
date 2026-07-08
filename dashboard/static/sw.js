/**
 * Service Worker 무력화 (2026-07-08).
 *
 * 이유: 기존 sw.js의 Cache First 전략이 Vite 해시 파일명이 갱신되는
 *   배포마다 stale HTML을 서빙해 매니저 브라우저에서 ERR_FAILED 재현.
 *   PWA/오프라인 지원은 사용 요구 없음 — 안전을 위해 SW 자체 사용 중지.
 *
 * 동작:
 *   - install 즉시 skipWaiting → 대기 없이 새 SW로 교체
 *   - activate 시 모든 캐시 삭제 + 자기 자신 unregister
 *   - 이후 방문 시 SW 없음 (modern_base.html에서 register 호출 제거됨)
 *
 * 기존 매니저 브라우저에 남아 있던 sw.js v2가 이 파일을 새로 받으면
 * 스스로 해제 + 캐시 삭제되므로 별도 안내 불필요.
 */

self.addEventListener('install', () => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        try {
            const cacheNames = await caches.keys();
            await Promise.all(cacheNames.map((name) => caches.delete(name)));
        } catch (err) {
            // 캐시 삭제 실패해도 unregister는 진행
        }
        try {
            await self.registration.unregister();
        } catch (err) {
            // ignore
        }
        try {
            const clients = await self.clients.matchAll({ type: 'window' });
            clients.forEach((client) => {
                // 다음 로드부터는 SW 없이 네트워크에서 직접 받도록 새로고침
                if (client.url && 'navigate' in client) {
                    client.navigate(client.url);
                }
            });
        } catch (err) {
            // ignore
        }
    })());
});

// fetch 이벤트 훅 제거 — 모든 요청은 브라우저가 직접 처리 (SW 통과).
