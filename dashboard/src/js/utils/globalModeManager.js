/**
 * 전역 ModeManager 인스턴스
 *
 * Singleton 패턴으로 애플리케이션 전체에서 하나의 ModeManager 인스턴스만 사용
 */

import ModeManager from './ModeManager.js';
import logger from './logger.js';

// 전역 인스턴스 (Singleton)
let globalModeManagerInstance = null;

/**
 * 전역 ModeManager 인스턴스 조회
 * 없으면 자동으로 생성
 *
 * @returns {ModeManager} 전역 ModeManager 인스턴스
 */
export function getGlobalModeManager() {
  if (!globalModeManagerInstance) {
    logger.debug('[GlobalModeManager] 전역 ModeManager 인스턴스 생성');

    globalModeManagerInstance = new ModeManager({
      debugMode: false // 프로덕션에서는 false, 개발 중에는 true로 설정
    });

    // 전역 접근 (디버깅용)
    if (typeof window !== 'undefined') {
      window.modeManager = globalModeManagerInstance;
      logger.debug('[GlobalModeManager] window.modeManager로 접근 가능 (디버깅용)');
    }

    // 전역 이벤트 리스너 (모든 모드 변경 로깅)
    globalModeManagerInstance.on('modeChanged', (data) => {
      logger.debug(`[GlobalModeManager] 모드 변경: ${data.eventType}`, data);
    });
  }

  return globalModeManagerInstance;
}

/**
 * 전역 ModeManager 인스턴스 재설정
 * 테스트 또는 특수한 경우에만 사용
 */
export function resetGlobalModeManager() {
  if (globalModeManagerInstance) {
    globalModeManagerInstance.destroy();
    globalModeManagerInstance = null;
    logger.debug('[GlobalModeManager] 전역 인스턴스 재설정 완료');
  }
}

// 기본 export
export default getGlobalModeManager;
