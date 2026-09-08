/**
 * Invoice Badge Module
 * 계산서 발행 상태 뱃지. Y 포맷 "{마지막 진행단계} - {상태}" (또는 bare 미발행).
 *
 * 상태별 색 (구글시트 Y열 색상과 일치):
 *   - 그냥 미발행(입금·계산서 없음) → 흰색/연한 회색 (none)
 *   - {단계} - 미발행(입금됨, 계산서 필요) → 회색 (unissued)
 *   - 발행완료  → 초록 (issued)
 *   - N입금     → 빨강/분홍 (cash)
 *   - 카드결제  → 보라 (card)
 *   - 기타      → 노랑 (other)
 */
import BaseBadge from './BaseBadge.js';

export default class InvoiceBadge extends BaseBadge {
  constructor() {
    super();

    // 파스텔 톤(연한 배경 + 중간 텍스트) — PM 사이트 다른 뱃지와 통일.
    this.categoryConfig = {
      'none': {       // 그냥 미발행/없음 → 거의 흰색
        cssClass: 'invoice-none',
        style: 'background-color: #fbfbfc; color: #aeb4bb;',
      },
      'unissued': {   // {단계} - 미발행 → 연한 회색 (부가세 '미포함'과 동일)
        cssClass: 'invoice-unissued',
        style: 'background-color: #f8f9fa; color: #6c757d;',
      },
      'issued': {     // 발행완료 → 연한 초록
        cssClass: 'invoice-issued',
        style: 'background-color: #e7f5e8; color: #3f9142;',
      },
      'cash': {       // N입금 → 연한 빨강/분홍
        cssClass: 'invoice-cash',
        style: 'background-color: #fdeceb; color: #c65c5c;',
      },
      'card': {       // 카드결제 → 연한 보라
        cssClass: 'invoice-card',
        style: 'background-color: #f1ebfa; color: #8368c4;',
      },
      'other': {      // 기타 → 연한 노랑
        cssClass: 'invoice-other',
        style: 'background-color: #fdf7e4; color: #b48a2c;',
      },
    };
  }

  /**
   * 값 → 카테고리 판정. Y = "{단계} - {상태}" 에서 상태 부분으로 판정.
   */
  categorize(billValue) {
    if (billValue === null || billValue === undefined) return 'none';
    const s = String(billValue).trim();
    // 그냥 미발행/없음 (입금·계산서 둘 다 없음) → 흰색
    if (!s || s === '-' || s === '미발행' || s.toLowerCase() === 'false') {
      return 'none';
    }
    if (s === true || s.toLowerCase() === 'true' || s === '발행완료') {
      return 'issued';
    }
    // 새 포맷 "{단계} - {상태}" → 상태 추출 (없으면 전체)
    const status = s.includes(' - ') ? s.split(' - ').pop().trim() : s;
    if (status === '발행완료' || status === '발행') return 'issued';
    if (status === '미발행') return 'unissued';  // {단계} - 미발행 (입금됨, 계산서 필요)
    if (status === 'N입금') return 'cash';
    if (status === '카드결제' || status.includes('카드')) return 'card';
    if (status === '기타') return 'other';
    // 레거시 폴백
    if (s.startsWith('일반') || s === '세금계산서') return 'issued';
    if (s.startsWith('N입금') || s === '현금' || s === '현금거래') return 'cash';
    return 'other';
  }

  /**
   * 뱃지 생성
   */
  create(billValue, options = {}) {
    const category = this.categorize(billValue);
    const config = this.categoryConfig[category] || this.categoryConfig['other'];

    // 표시 텍스트: 빈값/'-'만 '미발행', 나머지는 원문 그대로("계약금 - 미발행" 등)
    const s = String(billValue == null ? '' : billValue).trim();
    const text = (!s || s === '-') ? '미발행' : s;

    return this.createBadgeHtml(text, `badge invoice-badge ${config.cssClass}`, {
      style: config.style,
      title: `계산서: ${text}`,
      dataAttributes: {
        'invoice-category': category,
      },
      ...options,
    });
  }
}
