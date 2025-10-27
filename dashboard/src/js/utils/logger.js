/**
 * 전역 로거 유틸리티
 * 개발 모드에서만 로그를 출력합니다.
 */

class Logger {
  constructor() {
    this.isDevelopment = this.checkDevelopmentMode();
  }

  /**
   * 개발 모드 체크
   * @returns {boolean}
   */
  checkDevelopmentMode() {
    // 브라우저 환경이 아닌 경우 (Node.js, SSR 등)
    if (typeof window === 'undefined') {
      return false;
    }

    // 1. hostname 체크 (localhost 또는 127.0.0.1)
    const isLocalhost = window.location?.hostname === 'localhost' ||
                       window.location?.hostname === '127.0.0.1';

    // 2. window.DEBUG 플래그 체크
    const isDebugMode = window.DEBUG === true;

    // 3. URL 파라미터 체크 (?debug=true)
    const urlParams = new URLSearchParams(window.location?.search || '');
    const isDebugParam = urlParams.get('debug') === 'true';

    return isLocalhost || isDebugMode || isDebugParam;
  }

  /**
   * 디버그 레벨 로그 (개발 모드에서만)
   * @param {...any} args
   */
  debug(...args) {
    if (this.isDevelopment) {
      console.log(...args);
    }
  }

  /**
   * 정보 레벨 로그 (개발 모드에서만)
   * @param {...any} args
   */
  info(...args) {
    if (this.isDevelopment) {
      console.info(...args);
    }
  }

  /**
   * 경고 레벨 로그 (항상 출력)
   * @param {...any} args
   */
  warn(...args) {
    console.warn(...args);
  }

  /**
   * 에러 레벨 로그 (항상 출력)
   * @param {...any} args
   */
  error(...args) {
    console.error(...args);
  }

  /**
   * 테이블 형식 로그 (개발 모드에서만)
   * @param {*} data
   */
  table(data) {
    if (this.isDevelopment) {
      console.table(data);
    }
  }

  /**
   * 그룹 시작 (개발 모드에서만)
   * @param {string} label
   */
  group(label) {
    if (this.isDevelopment) {
      console.group(label);
    }
  }

  /**
   * 그룹 종료 (개발 모드에서만)
   */
  groupEnd() {
    if (this.isDevelopment) {
      console.groupEnd();
    }
  }

  /**
   * 시간 측정 시작 (개발 모드에서만)
   * @param {string} label
   */
  time(label) {
    if (this.isDevelopment) {
      console.time(label);
    }
  }

  /**
   * 시간 측정 종료 (개발 모드에서만)
   * @param {string} label
   */
  timeEnd(label) {
    if (this.isDevelopment) {
      console.timeEnd(label);
    }
  }

  /**
   * 개발 모드 상태 확인
   * @returns {boolean}
   */
  isDevMode() {
    return this.isDevelopment;
  }
}

// 싱글톤 인스턴스 생성 및 export
const logger = new Logger();

export default logger;
