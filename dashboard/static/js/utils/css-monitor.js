/**
 * CSS 모니터링 시스템 - 클라이언트 사이드
 * Feature Flag 기반 CSS 로딩 모니터링 및 자동 폴백
 */

class CSSMonitor {
    constructor() {
        this.monitoringData = {
            cssLoadTimes: {},
            errorCount: 0,
            successCount: 0,
            fallbackTriggered: false
        };

        this.callbacks = {
            onError: [],
            onSuccess: [],
            onFallback: []
        };

        this.featureFlagInfo = null;
        this.startTime = Date.now();

        this.init();
    }

    init() {
        console.log('[CSS Monitor] CSS 모니터링 시스템 초기화');
        this.monitorExistingCSS();
        this.setupCSSLoadMonitoring();
        this.setupErrorHandling();
    }

    monitorExistingCSS() {
        const cssLinks = document.querySelectorAll('link[rel="stylesheet"]');
        console.log(`[CSS Monitor] 기존 CSS 링크 ${cssLinks.length}개 모니터링 시작`);

        cssLinks.forEach((link, index) => {
            const href = link.href;
            const startTime = Date.now();

            // CSS 로딩 완료 감지
            if (link.sheet) {
                // 이미 로드됨
                this.recordSuccess(href, Date.now() - this.startTime);
            } else {
                // 로딩 중
                link.addEventListener('load', () => {
                    const loadTime = Date.now() - startTime;
                    this.recordSuccess(href, loadTime);
                });

                link.addEventListener('error', () => {
                    this.recordError(href, 'CSS 파일 로드 실패');
                });
            }
        });
    }

    setupCSSLoadMonitoring() {
        // 새로 추가되는 CSS 링크 모니터링
        const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                mutation.addedNodes.forEach((node) => {
                    if (node.tagName === 'LINK' && node.rel === 'stylesheet') {
                        this.monitorNewCSSLink(node);
                    }
                });
            });
        });

        observer.observe(document.head, {
            childList: true,
            subtree: true
        });
    }

    monitorNewCSSLink(link) {
        const href = link.href;
        const startTime = Date.now();

        console.log(`[CSS Monitor] 새 CSS 링크 모니터링: ${href}`);

        link.addEventListener('load', () => {
            const loadTime = Date.now() - startTime;
            this.recordSuccess(href, loadTime);
        });

        link.addEventListener('error', () => {
            this.recordError(href, 'CSS 파일 로드 실패');
        });

        // 타임아웃 설정 (10초)
        setTimeout(() => {
            if (!link.sheet) {
                this.recordError(href, 'CSS 로딩 타임아웃');
            }
        }, 10000);
    }

    setupErrorHandling() {
        // CSS 관련 에러 감지
        window.addEventListener('error', (event) => {
            if (event.target.tagName === 'LINK' && event.target.rel === 'stylesheet') {
                this.recordError(event.target.href, '링크 태그 로드 오류');
            }
        });

        // CSSOM 에러 감지
        document.addEventListener('DOMContentLoaded', () => {
            this.validateCSSOM();
        });
    }

    validateCSSOM() {
        try {
            const stylesheets = document.styleSheets;
            for (let i = 0; i < stylesheets.length; i++) {
                const stylesheet = stylesheets[i];
                if (stylesheet.href) {
                    try {
                        // CSS 규칙 접근 시도 (CORS 체크)
                        const rules = stylesheet.cssRules || stylesheet.rules;
                        if (!rules) {
                            this.recordError(stylesheet.href, 'CSS 규칙 접근 불가 (CORS?)');
                        }
                    } catch (e) {
                        this.recordError(stylesheet.href, `CSSOM 접근 오류: ${e.message}`);
                    }
                }
            }
        } catch (e) {
            console.warn('[CSS Monitor] CSSOM 검증 중 오류:', e);
        }
    }

    recordSuccess(href, loadTime) {
        this.monitoringData.successCount++;
        this.monitoringData.cssLoadTimes[href] = loadTime;

        console.log(`[CSS Monitor] CSS 로딩 성공: ${href} (${loadTime}ms)`);

        // 성공 콜백 실행
        this.callbacks.onSuccess.forEach(callback => {
            try {
                callback({ href, loadTime }, loadTime);
            } catch (e) {
                console.error('[CSS Monitor] 성공 콜백 오류:', e);
            }
        });

        // 서버에 성공 리포트
        this.reportToServer('success', { href, loadTime });
    }

    recordError(href, errorMessage) {
        this.monitoringData.errorCount++;

        console.error(`[CSS Monitor] CSS 오류: ${href} - ${errorMessage}`);

        const errorInfo = {
            href,
            error: errorMessage,
            timestamp: Date.now(),
            userAgent: navigator.userAgent,
            featureFlag: this.featureFlagInfo
        };

        // 에러 콜백 실행
        this.callbacks.onError.forEach(callback => {
            try {
                callback(errorInfo);
            } catch (e) {
                console.error('[CSS Monitor] 에러 콜백 오류:', e);
            }
        });

        // 서버에 에러 리포트
        this.reportToServer('error', errorInfo);

        // 자동 폴백 체크
        this.checkAutoFallback();
    }

    checkAutoFallback() {
        const errorRate = this.monitoringData.errorCount /
                         (this.monitoringData.errorCount + this.monitoringData.successCount);

        // 에러율이 50% 이상이고 아직 폴백이 트리거되지 않았다면
        if (errorRate >= 0.5 && !this.monitoringData.fallbackTriggered) {
            console.warn('[CSS Monitor] 높은 에러율 감지, 폴백 트리거');
            this.triggerFallback();
        }
    }

    triggerFallback() {
        this.monitoringData.fallbackTriggered = true;

        // 폴백 콜백 실행
        this.callbacks.onFallback.forEach(callback => {
            try {
                callback(this.monitoringData);
            } catch (e) {
                console.error('[CSS Monitor] 폴백 콜백 오류:', e);
            }
        });

        // 서버에 폴백 리포트
        this.reportToServer('fallback', this.monitoringData);

        console.log('[CSS Monitor] 폴백 모드 활성화됨');
    }

    reportToServer(type, data) {
        // 비동기로 서버에 데이터 전송
        setTimeout(() => {
            this.sendToServer(type, data);
        }, 100);
    }

    async sendToServer(type, data) {
        try {
            const payload = {
                type,
                data,
                timestamp: Date.now(),
                url: window.location.href,
                userAgent: navigator.userAgent,
                featureFlag: this.featureFlagInfo
            };

            await fetch('/css-monitor/api/report', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(payload)
            });

        } catch (e) {
            // 서버 리포트 실패는 조용히 처리
            console.debug('[CSS Monitor] 서버 리포트 실패:', e.message);
        }
    }

    // 콜백 등록 메서드들
    onError(callback) {
        this.callbacks.onError.push(callback);
    }

    onSuccess(callback) {
        this.callbacks.onSuccess.push(callback);
    }

    onFallback(callback) {
        this.callbacks.onFallback.push(callback);
    }

    // 현재 상태 반환
    getStatus() {
        return {
            ...this.monitoringData,
            errorRate: this.monitoringData.errorCount /
                      (this.monitoringData.errorCount + this.monitoringData.successCount),
            totalRequests: this.monitoringData.errorCount + this.monitoringData.successCount,
            averageLoadTime: this.calculateAverageLoadTime()
        };
    }

    calculateAverageLoadTime() {
        const loadTimes = Object.values(this.monitoringData.cssLoadTimes);
        if (loadTimes.length === 0) return 0;

        return loadTimes.reduce((sum, time) => sum + time, 0) / loadTimes.length;
    }

    // 수동 테스트용 메서드
    simulateError(href = 'test-error.css') {
        this.recordError(href, 'Simulated error for testing');
    }

    simulateSuccess(href = 'test-success.css', loadTime = 150) {
        this.recordSuccess(href, loadTime);
    }
}

// 전역 인스턴스 생성
window.cssMonitor = new CSSMonitor();

// 개발 모드에서 디버깅 정보 표시
if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
    // 5초 후 상태 출력
    setTimeout(() => {
        console.log('[CSS Monitor] 현재 상태:', window.cssMonitor.getStatus());
    }, 5000);
}

console.log('[CSS Monitor] CSS 모니터링 시스템 로드 완료');