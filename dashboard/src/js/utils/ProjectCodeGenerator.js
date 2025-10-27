import logger from '../utils/logger.js';

/**
 * ProjectCodeGenerator - 프로젝트 코드 생성 유틸리티
 */
export default class ProjectCodeGenerator {
  constructor() {
    this.previewTimeout = null;
  }

  /**
   * 프로젝트 코드 미리보기 (클라이언트 측 생성)
   * @param {string} company - 사업자
   * @param {string} owner - 담당자
   * @param {function} callback - 결과를 받을 콜백 (code, error)
   * @param {number} delay - 디바운싱 딜레이 (ms)
   */
  previewCode(company, owner, callback, delay = 300) {
    console.log('[DEBUG] ProjectCodeGenerator.previewCode 호출:', { company, owner, delay });

    // 기존 타임아웃 취소
    if (this.previewTimeout) {
      clearTimeout(this.previewTimeout);
      console.log('[DEBUG] 이전 타임아웃 취소됨');
    }

    // 값이 없으면 즉시 반환
    if (!company || !owner) {
      console.log('[DEBUG] company 또는 owner가 비어있음, 에러 콜백 호출');
      callback(null, '사업자와 담당자를 선택하세요');
      return;
    }

    // 로딩 상태 콜백 (짧은 시간이지만 UX를 위해 유지)
    console.log('[DEBUG] 로딩 상태 콜백 호출');
    callback('loading', null);

    // 디바운싱된 클라이언트 측 계산
    this.previewTimeout = setTimeout(() => {
      console.log('[DEBUG] 디바운싱 타임아웃 종료, 클라이언트 측 코드 생성 시작');
      try {
        // 프로젝트 데이터 가져오기
        const projectData = window.ITGlobalApp?.api?.getCurrentData();

        if (!projectData || !Array.isArray(projectData) || projectData.length === 0) {
          console.error('[DEBUG] 프로젝트 데이터를 찾을 수 없음');
          callback(null, '프로젝트 데이터를 불러올 수 없습니다');
          return;
        }

        console.log('[DEBUG] 프로젝트 데이터 로드 완료:', projectData.length, '건');

        // 회사 코드 매핑
        const companyMapping = {
          '플렌트': 'P',
          '글로벌': 'G',
          '글로벌그룹': 'R'
        };
        const companyCode = companyMapping[company] || 'G';

        // 담당자 코드 매핑
        const ownerMapping = {
          '박용구': 'YG',
          '박정우': 'JW',
          '강성환': 'SH',
          '박민우': 'MW',
          '이근혁': 'GH',
          '김호중': 'HJ',
          '아이티': 'IT',
          '김단이': 'DN',
          '권태훈': 'TH',
          '주영민': 'YM',
          '심장원': 'SJW',
          '빈승정': 'SJ',
          '박민재': 'MJ',
          '조성헌': 'JSH',
          '황해승': 'HS',
          '강민석': 'MS'
        };
        const ownerSuffix = ownerMapping[owner] || (owner.length >= 2 ? owner.substring(0, 2).toUpperCase() : owner.toUpperCase());

        // 기존 프로젝트 코드에서 최대 순번 찾기
        let maxSequence = 0;
        projectData.forEach(project => {
          const code = project['프로젝트 코드'];
          if (code && typeof code === 'string' && code.includes('-')) {
            try {
              // G0001-IT 형식에서 숫자 부분 추출
              const parts = code.split('-')[0];
              const numberPart = parts.replace(/[^0-9]/g, ''); // 숫자만 추출
              if (numberPart) {
                const num = parseInt(numberPart, 10);
                if (num > maxSequence) {
                  maxSequence = num;
                }
              }
            } catch (e) {
              // 파싱 실패는 무시
            }
          }
        });

        // 다음 순번
        const sequence = maxSequence + 1;

        // 프로젝트 코드 생성: G0001-IT 형식
        const projectCode = `${companyCode}${sequence.toString().padStart(4, '0')}-${ownerSuffix}`;

        console.log('[DEBUG] 클라이언트 측 코드 생성 완료:', {
          companyCode,
          ownerSuffix,
          maxSequence,
          sequence,
          projectCode
        });

        callback(projectCode, null);

      } catch (error) {
        console.error('[DEBUG] 클라이언트 측 코드 생성 오류:', error);
        logger.error('프로젝트 코드 생성 오류:', error);
        callback(null, '코드 생성 오류');
      }
    }, delay);
    console.log(`[DEBUG] 디바운싱 타임아웃 설정됨 (${delay}ms)`);
  }

  /**
   * 회사명에서 접두사 추출 (레거시 로직)
   * @param {string} company - 회사명
   * @returns {string} 접두사 (예: "G", "B", "C")
   */
  static extractCompanyPrefix(company) {
    const prefixMap = {
      '광림': 'G',
      '거성': 'GS',
      '보성': 'B',
      '청우': 'C'
    };

    for (const [key, prefix] of Object.entries(prefixMap)) {
      if (company.includes(key)) {
        return prefix;
      }
    }

    // 기본값: 첫 글자
    return company.charAt(0).toUpperCase();
  }

  /**
   * 담당자명에서 접미사 추출 (레거시 로직)
   * @param {string} owner - 담당자명
   * @returns {string} 접미사 (예: "KHJ", "LKH")
   */
  static extractOwnerSuffix(owner) {
    const suffixMap = {
      '김현종': 'KHJ',
      '이근혁': 'LKH',
      '황샛별': 'HSB',
      '심장원': 'SJW',
      '김단이': 'KDI'
    };

    if (suffixMap[owner]) {
      return suffixMap[owner];
    }

    // 기본값: 이름의 초성
    return this.getKoreanInitials(owner);
  }

  /**
   * 한글 이름의 초성 추출
   * @param {string} name - 한글 이름
   * @returns {string} 초성 (예: "김현종" → "KHJ")
   */
  static getKoreanInitials(name) {
    const initials = {
      'ㄱ': 'G', 'ㄲ': 'GG', 'ㄴ': 'N', 'ㄷ': 'D', 'ㄸ': 'DD',
      'ㄹ': 'L', 'ㅁ': 'M', 'ㅂ': 'B', 'ㅃ': 'BB', 'ㅅ': 'S',
      'ㅆ': 'SS', 'ㅇ': '', 'ㅈ': 'J', 'ㅉ': 'JJ', 'ㅊ': 'C',
      'ㅋ': 'K', 'ㅌ': 'T', 'ㅍ': 'P', 'ㅎ': 'H'
    };

    let result = '';
    for (let i = 0; i < name.length; i++) {
      const code = name.charCodeAt(i) - 0xAC00;
      if (code >= 0 && code <= 11171) {
        const initialIndex = Math.floor(code / 588);
        const initialChars = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'];
        const initial = initialChars[initialIndex];
        result += initials[initial] || '';
      }
    }
    return result || name.substring(0, 3).toUpperCase();
  }

  /**
   * 프로젝트 코드 형식 검증
   * @param {string} code - 검증할 코드
   * @returns {boolean} 유효 여부
   */
  static validateCodeFormat(code) {
    // 형식: G0001-KHJ, GS0123-LKH 등
    const pattern = /^[A-Z]{1,2}\d{4}-[A-Z]{2,3}$/;
    return pattern.test(code);
  }

  /**
   * 타임아웃 정리
   */
  cleanup() {
    if (this.previewTimeout) {
      clearTimeout(this.previewTimeout);
      this.previewTimeout = null;
    }
  }
}
