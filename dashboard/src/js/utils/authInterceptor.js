import logger from '../utils/logger.js';

/**
 * 전역 fetch interceptor - 모든 API 요청의 401 응답을 자동 처리
 * Session timeout warning 추가 (세션 만료 5분 전 경고)
 */

// 원본 fetch를 백업
const originalFetch = window.fetch;

// 세션 타임아웃 설정 (서버에서 동적으로 가져옴)
const REDIRECT_LOCK_DURATION = 5000; // 5초 (다중 탭 리다이렉트 방지)

// 동적 타임아웃 설정 (기본값: 4시간 = 240분)
let SESSION_TIMEOUT_MINUTES = 240; // 서버 기본값과 동일
let SESSION_WARNING_TIME = (SESSION_TIMEOUT_MINUTES - 5) * 60 * 1000; // 5분 전 경고
let SESSION_TIMEOUT = SESSION_TIMEOUT_MINUTES * 60 * 1000;

let lastActivityTime = Date.now();
let sessionWarningShown = false;

/**
 * 서버에서 세션 타임아웃 설정을 가져와서 동적으로 적용
 */
async function initializeSessionTimeout() {
  try {
    // localStorage에서 먼저 확인 (페이지 새로고침 시 빠른 초기화)
    const storedTimeout = localStorage.getItem('session_timeout_minutes');
    if (storedTimeout) {
      updateSessionTimeout(parseInt(storedTimeout, 10));
    }

    // 서버에서 최신 값 가져오기 (비동기)
    const response = await originalFetch('/api/auth/ping', { method: 'POST' });
    if (response.ok) {
      const data = await response.json();
      if (data.session_timeout_minutes) {
        updateSessionTimeout(data.session_timeout_minutes);
        localStorage.setItem('session_timeout_minutes', data.session_timeout_minutes);
      }
    }
  } catch (error) {
    // 초기화 실패 시 기본값 사용 (로그인 전이거나 네트워크 오류)
    logger.debug('[Auth] 세션 타임아웃 초기화 실패, 기본값 사용:', error);
  }
}

/**
 * 세션 타임아웃 값 업데이트
 */
function updateSessionTimeout(timeoutMinutes) {
  SESSION_TIMEOUT_MINUTES = timeoutMinutes;
  SESSION_WARNING_TIME = (timeoutMinutes - 5) * 60 * 1000; // 5분 전 경고
  SESSION_TIMEOUT = timeoutMinutes * 60 * 1000;
  logger.debug(`[Auth] 세션 타임아웃 업데이트: ${timeoutMinutes}분 (경고: ${timeoutMinutes - 5}분)`);
}

// 초기화 실행 (페이지 로드 시)
initializeSessionTimeout();

/**
 * 이중 리다이렉트 방지 (다중 탭 대응)
 * localStorage를 사용하여 탭 간 리다이렉트 상태 공유
 */
function isAlreadyRedirecting() {
  try {
    const redirectState = localStorage.getItem('auth_redirecting');
    if (!redirectState) return false;

    const { timestamp } = JSON.parse(redirectState);
    const now = Date.now();

    // 5초 이내에 다른 탭에서 리다이렉트 시작했으면 무시
    if (now - timestamp < REDIRECT_LOCK_DURATION) {
      return true;
    }

    // 5초 지났으면 만료된 것으로 간주하고 제거
    localStorage.removeItem('auth_redirecting');
    return false;
  } catch (error) {
    // localStorage 접근 실패 시 (incognito 모드 등)
    logger.warn('[Auth] localStorage 접근 실패, 이중 리다이렉트 방지 비활성화:', error);
    return false;
  }
}

function setRedirectingState() {
  try {
    localStorage.setItem('auth_redirecting', JSON.stringify({
      timestamp: Date.now(),
      tabId: crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36)
    }));
  } catch (error) {
    logger.warn('[Auth] localStorage 설정 실패:', error);
  }
}

// 세션 활동 갱신
function updateActivity() {
  lastActivityTime = Date.now();
  sessionWarningShown = false;
}

// 세션 경고 표시
function showSessionWarning() {
  if (sessionWarningShown || isAlreadyRedirecting()) return;
  sessionWarningShown = true;

  // 경고 메시지 표시
  const warningDiv = document.createElement('div');
  warningDiv.id = 'session-warning-banner';
  warningDiv.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    background: linear-gradient(135deg, #FFF3CD 0%, #FFF8E1 100%);
    color: #856404;
    padding: 16px 24px;
    text-align: center;
    z-index: 10000;
    border-bottom: 3px solid #FFEB3B;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    font-size: 14px;
    font-weight: 600;
    animation: slideDown 0.3s ease-out;
  `;

  warningDiv.innerHTML = `
    <div style="display: flex; align-items: center; justify-content: center; gap: 16px;">
      <span style="font-size: 20px;">⏰</span>
      <span>세션이 5분 후 만료됩니다. 페이지를 새로고침하거나 작업을 계속하면 세션이 연장됩니다.</span>
      <button id="session-refresh-btn" style="
        background: #856404;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 6px;
        cursor: pointer;
        font-weight: 600;
        transition: all 0.2s;
      " onmouseover="this.style.background='#705c04'" onmouseout="this.style.background='#856404'">
        지금 연장하기
      </button>
      <button id="session-warning-close" style="
        background: transparent;
        border: none;
        color: #856404;
        font-size: 20px;
        cursor: pointer;
        padding: 0 8px;
      ">×</button>
    </div>
  `;

  document.body.appendChild(warningDiv);

  // 연장하기 버튼 이벤트
  document.getElementById('session-refresh-btn').addEventListener('click', async () => {
    try {
      // 세션 연장을 위한 ping 요청 (login_time 갱신)
      const response = await originalFetch('/api/auth/ping', { method: 'POST' });
      const data = await response.json();

      // 서버에서 받은 타임아웃 값으로 업데이트
      if (data.session_timeout_minutes) {
        updateSessionTimeout(data.session_timeout_minutes);
        localStorage.setItem('session_timeout_minutes', data.session_timeout_minutes);
      }

      updateActivity();

      // 성공 메시지 표시
      warningDiv.style.background = 'linear-gradient(135deg, #D4EDDA 0%, #E8F5E9 100%)';
      warningDiv.style.borderBottomColor = '#4CAF50';
      warningDiv.innerHTML = `
        <div style="display: flex; align-items: center; justify-content: center; gap: 16px;">
          <span style="font-size: 20px;">✅</span>
          <span style="color: #155724; font-weight: 600;">세션이 연장되었습니다!</span>
        </div>
      `;

      setTimeout(() => {
        warningDiv.remove();
      }, 2000);
    } catch (error) {
      logger.error('[Auth] 세션 연장 실패:', error);
    }
  });

  // 닫기 버튼 이벤트
  document.getElementById('session-warning-close').addEventListener('click', () => {
    warningDiv.remove();
  });

  // CSS 애니메이션 추가
  if (!document.getElementById('session-warning-styles')) {
    const style = document.createElement('style');
    style.id = 'session-warning-styles';
    style.textContent = `
      @keyframes slideDown {
        from {
          transform: translateY(-100%);
          opacity: 0;
        }
        to {
          transform: translateY(0);
          opacity: 1;
        }
      }
    `;
    document.head.appendChild(style);
  }
}

// 세션 타임아웃 체크 (1분마다)
setInterval(() => {
  const timeSinceLastActivity = Date.now() - lastActivityTime;

  // 타임아웃 5분 전 경고
  if (timeSinceLastActivity >= SESSION_WARNING_TIME && !sessionWarningShown) {
    showSessionWarning();
  }

  // 타임아웃 도달 시 자동 로그아웃 (서버 세션 만료와 동기화)
  if (timeSinceLastActivity >= SESSION_TIMEOUT && !isAlreadyRedirecting()) {
    logger.warn(`[Auth] 세션 타임아웃 (${SESSION_TIMEOUT_MINUTES}분) - 자동 로그아웃`);
    setRedirectingState();

    const currentPath = window.location.pathname + window.location.search;
    if (currentPath !== '/login' && currentPath !== '/') {
      sessionStorage.setItem('returnUrl', currentPath);
    }

    window.location.href = '/login?session_expired=true&reason=timeout';
  }
}, 60000); // 1분마다 체크

// 사용자 활동 감지 (마우스, 키보드, 스크롤)
['mousedown', 'keydown', 'scroll', 'touchstart'].forEach(event => {
  document.addEventListener(event, updateActivity, { passive: true });
});

/**
 * URL이 내부 API 요청인지 확인
 * @param {string|Request} urlOrRequest - URL 문자열 또는 Request 객체
 * @returns {boolean}
 */
function isInternalAPIRequest(urlOrRequest) {
  let url;

  // Request 객체인 경우
  if (urlOrRequest instanceof Request) {
    url = urlOrRequest.url;
  } else if (typeof urlOrRequest === 'string') {
    url = urlOrRequest;
  } else {
    return false;
  }

  // 절대 URL인 경우 (https://, http://)
  if (url.startsWith('http://') || url.startsWith('https://')) {
    // 현재 도메인과 다른 경우 외부 요청으로 간주
    const currentHost = window.location.host;
    const urlObj = new URL(url);
    if (urlObj.host !== currentHost) {
      return false; // 외부 API (Google Analytics, Sentry 등)
    }
    // 같은 도메인이면 경로 체크
    url = urlObj.pathname;
  }

  // 상대 URL 또는 같은 도메인의 절대 URL
  // /api/ 또는 /leads/api/로 시작하는 경우 내부 API로 간주
  return url.startsWith('/api/') || url.startsWith('/leads/api/');
}

// fetch를 오버라이드 (화이트리스트 방식)
window.fetch = async function(...args) {
  const urlOrRequest = args[0];

  // 내부 API 요청인 경우에만 인증 체크
  const isInternalAPI = isInternalAPIRequest(urlOrRequest);

  // fetch 호출 자체도 활동으로 간주 (내부 API 요청만)
  if (isInternalAPI) {
    updateActivity();
  }

  try {
    const response = await originalFetch(...args);

    // 내부 API 요청이 아니면 그대로 반환 (써드파티 SDK 영향 안 받음)
    if (!isInternalAPI) {
      return response;
    }

    // 401 Unauthorized 응답 처리 (내부 API만)
    if (response.status === 401) {
      // 이미 리다이렉트 중이면 중복 처리 방지 (다중 탭 대응)
      if (isAlreadyRedirecting()) {
        logger.debug('[Auth] 다른 탭에서 이미 리다이렉트 진행 중, 현재 탭은 대기');
        return response;
      }

      setRedirectingState();
      logger.warn('[Auth] 세션 만료 감지 (401) - 로그인 페이지로 리다이렉트');

      // 세션 경고 배너 제거
      const warningBanner = document.getElementById('session-warning-banner');
      if (warningBanner) {
        warningBanner.remove();
      }

      // 현재 페이지 URL을 저장 (로그인 후 돌아올 수 있도록)
      const currentPath = window.location.pathname + window.location.search;
      if (currentPath !== '/login' && currentPath !== '/') {
        sessionStorage.setItem('returnUrl', currentPath);
      }

      // 로그인 페이지로 리다이렉트
      window.location.href = '/login?session_expired=true';

      // 응답은 그대로 반환 (리다이렉트 전까지 에러 방지)
      return response;
    }

    // API 요청인데 HTML이 반환된 경우 (세션 만료로 로그인 페이지 리다이렉트됨)
    const contentType = response.headers.get('content-type');
    if (contentType && contentType.includes('text/html')) {
      if (isAlreadyRedirecting()) {
        logger.debug('[Auth] 다른 탭에서 이미 리다이렉트 진행 중, 현재 탭은 대기');
        return response;
      }

      setRedirectingState();
      logger.warn('[Auth] API 요청에 HTML 응답 감지 - 세션 만료로 판단');

      // 세션 경고 배너 제거
      const warningBanner = document.getElementById('session-warning-banner');
      if (warningBanner) {
        warningBanner.remove();
      }

      // 현재 페이지 URL을 저장
      const currentPath = window.location.pathname + window.location.search;
      if (currentPath !== '/login' && currentPath !== '/') {
        sessionStorage.setItem('returnUrl', currentPath);
      }

      // 로그인 페이지로 리다이렉트
      window.location.href = '/login?session_expired=true';

      return response;
    }

    return response;
  } catch (error) {
    // 네트워크 오류 등은 그대로 throw
    throw error;
  }
};

logger.debug('[Auth] 전역 인증 interceptor 활성화 (세션 타임아웃 모니터링 포함)');
