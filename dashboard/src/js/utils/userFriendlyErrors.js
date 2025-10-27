import logger from '../utils/logger.js';

/**
 * 사용자 친화적인 에러 메시지 변환 유틸리티
 * 기술적인 에러를 한국어 + 액션 가이드로 변환
 */

const ERROR_MESSAGES = {
  // HTTP 에러
  400: {
    message: '잘못된 요청입니다',
    action: '입력하신 정보를 다시 확인해주세요'
  },
  401: {
    message: '로그인이 필요합니다',
    action: '다시 로그인해주세요'
  },
  403: {
    message: '접근 권한이 없습니다',
    action: '관리자에게 문의하세요'
  },
  404: {
    message: '요청하신 정보를 찾을 수 없습니다',
    action: '페이지를 새로고침 후 다시 시도해주세요'
  },
  429: {
    message: '요청이 너무 많습니다',
    action: '잠시 후 다시 시도해주세요'
  },
  500: {
    message: '서버에서 오류가 발생했습니다',
    action: '잠시 후 다시 시도하거나 관리자에게 문의하세요'
  },
  502: {
    message: '서버 연결에 실패했습니다',
    action: '네트워크 상태를 확인하거나 잠시 후 다시 시도해주세요'
  },
  503: {
    message: '서비스를 일시적으로 사용할 수 없습니다',
    action: '서버 점검 중일 수 있습니다. 잠시 후 다시 시도해주세요'
  },
  504: {
    message: '서버 응답 시간이 초과되었습니다',
    action: '네트워크 상태를 확인하거나 잠시 후 다시 시도해주세요'
  },

  // 네트워크 에러
  'NetworkError': {
    message: '네트워크 연결이 끊어졌습니다',
    action: '인터넷 연결을 확인하고 다시 시도해주세요'
  },
  'TimeoutError': {
    message: '요청 시간이 초과되었습니다',
    action: '네트워크 상태를 확인하거나 잠시 후 다시 시도해주세요'
  },

  // 데이터 에러
  'ValidationError': {
    message: '입력하신 정보가 올바르지 않습니다',
    action: '입력값을 다시 확인해주세요'
  },
  'DuplicateError': {
    message: '이미 존재하는 데이터입니다',
    action: '다른 값을 입력해주세요'
  },

  // Google Sheets 에러
  'SheetsAPIError': {
    message: 'Google Sheets 연결에 실패했습니다',
    action: '잠시 후 다시 시도하거나 관리자에게 문의하세요'
  },
  'SheetsQuotaError': {
    message: 'Google Sheets 요청 한도를 초과했습니다',
    action: '잠시 후 다시 시도해주세요'
  }
};

/**
 * 에러를 사용자 친화적인 메시지로 변환
 * @param {Error|Object|string} error - 에러 객체 또는 메시지
 * @param {string} context - 에러 발생 컨텍스트 (예: "프로젝트 저장")
 * @returns {Object} - { title, message, action, originalError }
 */
export function getFriendlyErrorMessage(error, context = '') {
  let errorCode = null;
  let errorType = null;
  let originalMessage = '';

  // Error 객체 파싱
  if (error instanceof Error) {
    originalMessage = error.message;
    errorType = error.name;
  } else if (typeof error === 'object') {
    // HTTP Response 객체
    if (error.status) {
      errorCode = error.status;
    }
    // Fetch API Error Response
    if (error.response?.status) {
      errorCode = error.response.status;
    }
    // 커스텀 에러 객체
    if (error.message) {
      originalMessage = error.message;
    }
    if (error.type) {
      errorType = error.type;
    }
  } else if (typeof error === 'string') {
    originalMessage = error;

    // HTTP 에러 코드 추출 (예: "HTTP 500: INTERNAL SERVER ERROR")
    const httpMatch = error.match(/HTTP\s+(\d{3})/i);
    if (httpMatch) {
      errorCode = parseInt(httpMatch[1]);
    }
  }

  // 서버가 보낸 구체적인 메시지 우선 사용 (UX 개선)
  // 원본 메시지에 "필드" 키워드가 있으면 서버의 상세 메시지로 간주
  let errorInfo = null;

  if (originalMessage && originalMessage.includes('필드')) {
    // 서버가 보낸 구체적인 메시지를 그대로 사용
    errorInfo = {
      message: originalMessage,
      action: '입력값을 수정하고 다시 시도해주세요'
    };
  }
  // 1. HTTP 코드로 찾기
  else if (errorCode && ERROR_MESSAGES[errorCode]) {
    errorInfo = ERROR_MESSAGES[errorCode];
  }
  // 2. 에러 타입으로 찾기
  else if (errorType && ERROR_MESSAGES[errorType]) {
    errorInfo = ERROR_MESSAGES[errorType];
  }
  // 3. 원본 메시지에서 키워드 찾기
  else if (originalMessage) {
    const lowerMessage = originalMessage.toLowerCase();

    if (lowerMessage.includes('network') || lowerMessage.includes('연결')) {
      errorInfo = ERROR_MESSAGES['NetworkError'];
    } else if (lowerMessage.includes('timeout') || lowerMessage.includes('시간 초과')) {
      errorInfo = ERROR_MESSAGES['TimeoutError'];
    } else if (lowerMessage.includes('quota') || lowerMessage.includes('rate limit')) {
      errorInfo = ERROR_MESSAGES['SheetsQuotaError'];
    } else if (lowerMessage.includes('validation') || lowerMessage.includes('invalid')) {
      errorInfo = ERROR_MESSAGES['ValidationError'];
    } else if (lowerMessage.includes('duplicate') || lowerMessage.includes('중복')) {
      errorInfo = ERROR_MESSAGES['DuplicateError'];
    } else if (lowerMessage.includes('sheets') || lowerMessage.includes('google')) {
      errorInfo = ERROR_MESSAGES['SheetsAPIError'];
    }
  }

  // 기본 에러 메시지
  if (!errorInfo) {
    errorInfo = {
      message: '예상치 못한 오류가 발생했습니다',
      action: '잠시 후 다시 시도해주세요'
    };
  }

  // 컨텍스트 추가
  const title = context ? `${context} 중 오류 발생` : '오류 발생';

  return {
    title,
    message: errorInfo.message,
    action: errorInfo.action,
    originalError: originalMessage || (typeof error === 'string' ? error : JSON.stringify(error)),
    code: errorCode || errorType || 'UNKNOWN'
  };
}

/**
 * Toast 알림으로 사용자 친화적 에러 표시
 * @param {Error|Object|string} error - 에러
 * @param {string} context - 컨텍스트
 * @param {Object} options - Toast 옵션
 */
export function showFriendlyError(error, context = '', options = {}) {
  const friendlyError = getFriendlyErrorMessage(error, context);

  // Toast 메시지 생성
  const message = `
    <div class="error-toast-content">
      <div class="error-toast-message">${friendlyError.message}</div>
      <div class="error-toast-action">${friendlyError.action}</div>
    </div>
  `;

  // Bootstrap Toast 사용 (있는 경우)
  if (window.bootstrap && typeof window.showToast === 'function') {
    window.showToast({
      title: friendlyError.title,
      message: message,
      type: 'error',
      duration: options.duration || 5000,
      ...options
    });
  }
  // Fallback: Alert
  else {
    alert(`${friendlyError.title}\n\n${friendlyError.message}\n\n💡 ${friendlyError.action}`);
  }

  // 콘솔에 원본 에러 로깅 (디버깅용)
  logger.error('[원본 에러]', friendlyError.originalError);
  logger.error('[에러 코드]', friendlyError.code);

  return friendlyError;
}

/**
 * 프로젝트 저장 관련 특화 에러 메시지
 * @param {Error} error - 에러
 * @returns {Object} - 친화적 에러 정보
 */
export function getProjectSaveError(error) {
  const context = '프로젝트 저장';
  const friendlyError = getFriendlyErrorMessage(error, context);

  // 프로젝트 저장 관련 추가 가이드
  const additionalActions = {
    400: '입력하신 필드 값들을 다시 확인해주세요 (금액, 날짜 형식 등)',
    404: '프로젝트가 삭제되었을 수 있습니다. 프로젝트 목록을 새로고침해주세요',
    429: 'Google Sheets 저장 한도를 초과했습니다. 1분 후 다시 시도해주세요',
    500: '서버에서 저장 중 오류가 발생했습니다. 데이터가 저장되지 않았을 수 있으니 확인 후 다시 시도해주세요'
  };

  if (friendlyError.code && additionalActions[friendlyError.code]) {
    friendlyError.action = additionalActions[friendlyError.code];
  }

  return friendlyError;
}

/**
 * 데이터 로딩 관련 특화 에러 메시지
 * @param {Error} error - 에러
 * @returns {Object} - 친화적 에러 정보
 */
export function getDataLoadError(error) {
  const context = '데이터 불러오기';
  const friendlyError = getFriendlyErrorMessage(error, context);

  const additionalActions = {
    404: 'Google Sheets를 찾을 수 없습니다. 시트가 삭제되었거나 권한이 변경되었을 수 있습니다',
    429: 'Google Sheets 조회 한도를 초과했습니다. 1분 후 자동으로 다시 시도됩니다',
    500: '서버에서 데이터 처리 중 오류가 발생했습니다. 페이지를 새로고침해주세요'
  };

  if (friendlyError.code && additionalActions[friendlyError.code]) {
    friendlyError.action = additionalActions[friendlyError.code];
  }

  return friendlyError;
}

export default {
  getFriendlyErrorMessage,
  showFriendlyError,
  getProjectSaveError,
  getDataLoadError
};
