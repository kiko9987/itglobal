/**
 * Service Worker - 기본 PWA 기능
 * 캐싱 및 오프라인 지원
 */

const CACHE_NAME = 'itg-dashboard-v2';
const STATIC_CACHE_URLS = [
    '/static/css/project-list.css',
    '/static/js/project-list.js',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js',
    'https://code.jquery.com/jquery-3.6.0.min.js'
];

// Service Worker 설치
self.addEventListener('install', (event) => {
    console.log('🔧 Service Worker 설치 중...');

    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('📦 정적 리소스 캐싱 중...');
                // 중복 URL 제거
                const uniqueUrls = [...new Set(STATIC_CACHE_URLS)];
                console.log('캐싱할 URL 목록:', uniqueUrls);

                // 개별 캐싱으로 중복 오류 방지
                return Promise.allSettled(
                    uniqueUrls.map(async (url) => {
                        try {
                            const response = await fetch(url);
                            if (response.ok) {
                                await cache.put(url, response);
                                console.log(`✅ 캐싱 성공: ${url}`);
                            }
                        } catch (err) {
                            console.warn(`❌ 캐싱 실패: ${url}`, err);
                        }
                    })
                );
            })
            .then(() => {
                console.log('[SUCCESS] Service Worker 설치 완료');
                return self.skipWaiting();
            })
    );
});

// Service Worker 활성화
self.addEventListener('activate', (event) => {
    console.log('[ROCKET] Service Worker 활성화 중...');

    event.waitUntil(
        caches.keys()
            .then((cacheNames) => {
                return Promise.all(
                    cacheNames.map((cacheName) => {
                        if (cacheName !== CACHE_NAME) {
                            console.log('🗑️ 이전 캐시 삭제:', cacheName);
                            return caches.delete(cacheName);
                        }
                    })
                );
            })
            .then(() => {
                console.log('[SUCCESS] Service Worker 활성화 완료');
                return self.clients.claim();
            })
    );
});

// 네트워크 요청 가로채기 (Cache First 전략)
self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);

    // API 요청은 캐싱하지 않음
    if (url.pathname.startsWith('/api/')) {
        return;
    }

    // 외부 리소스나 정적 파일만 캐싱
    if (url.origin !== location.origin && !url.hostname.includes('cdn.')) {
        return;
    }

    event.respondWith(
        caches.match(event.request)
            .then((response) => {
                // 캐시에 있으면 캐시에서 반환
                if (response) {
                    return response;
                }

                // 캐시에 없으면 네트워크에서 가져와서 캐싱
                return fetch(event.request)
                    .then((response) => {
                        // 유효한 응답만 캐싱
                        if (!response || response.status !== 200 || response.type !== 'basic') {
                            return response;
                        }

                        const responseToCache = response.clone();
                        caches.open(CACHE_NAME)
                            .then((cache) => {
                                cache.put(event.request, responseToCache);
                            });

                        return response;
                    })
                    .catch(() => {
                        // 네트워크 실패 시 기본 오프라인 페이지 (향후 구현)
                        console.warn('네트워크 요청 실패:', event.request.url);
                        throw new Error('네트워크 오류');
                    });
            })
    );
});

// 백그라운드 동기화 (향후 확장 가능)
self.addEventListener('sync', (event) => {
    console.log('[REFRESH] 백그라운드 동기화:', event.tag);
});

// 푸시 알림 (향후 확장 가능)
self.addEventListener('push', (event) => {
    console.log('📨 푸시 알림 수신:', event.data?.text());
});