/**
 * Invoice Badge Module
 * 계산서 발행 상태 뱃지.
 *
 * 값 종류:
 *   - 미발행 / '' / '-'                → 빨강 (danger)
 *   - 일반 계열 (세금계산서 정상 발행)   → 초록 (success)
 *   - N입금 / 현금 (계산서 미발행 현금)  → 파랑 (info)
 *   - 부분 발행 (콤마 있음)              → 노랑 (warning)
 *   - 기타 (카드결제 등)                  → 회색 (secondary)
 */
import BaseBadge from './BaseBadge.js';

export default class InvoiceBadge extends BaseBadge {
  constructor() {
    super();

    this.categoryConfig = {
      'unissued': {
        cssClass: 'invoice-unissued',
        style: 'background-color: #f8d7da; color: #842029;',
      },
      'issued': {
        cssClass: 'invoice-issued',
        style: 'background-color: #d1e7dd; color: #0f5132;',
      },
      'cash': {
        cssClass: 'invoice-cash',
        style: 'background-color: #cfe2ff; color: #084298;',
      },
      'partial': {
        cssClass: 'invoice-partial',
        style: 'background-color: #fff3cd; color: #664d03;',
      },
      'other': {
        cssClass: 'invoice-other',
        style: 'background-color: #e2e3e5; color: #41464b;',
      },
    };
  }

  /**
   * 값 → 카테고리 판정
   */
  categorize(billValue) {
    if (billValue === null || billValue === undefined) return 'unissued';
    const s = String(billValue).trim();
    if (!s || s === '-' || s === '미발행' || s.toLowerCase() === 'false') {
      return 'unissued';
    }
    if (s === true || s.toLowerCase() === 'true' || s === '발행완료') {
      return 'issued';
    }
    // 부분 발행: 콤마 or 여러 카테고리 조합
    if (s.includes(',') || (s.includes('-') && (s.includes('일반') || s.includes('N입금')))) {
      // 하나의 카테고리 (예: '일반-잔금') 는 partial 아님, 여러 개 (예: '일반-계약금, N입금-잔금') 만
      const items = s.split(',').map(x => x.trim()).filter(Boolean);
      if (items.length > 1) return 'partial';
    }
    if (s.startsWith('일반') || s === '세금계산서' || s === '일반') return 'issued';
    if (s.startsWith('N입금') || s === '현금' || s === '현금거래') return 'cash';
    if (s.includes('카드')) return 'other';
    return 'other';
  }

  /**
   * 뱃지 생성
   */
  create(billValue, options = {}) {
    const category = this.categorize(billValue);
    const config = this.categoryConfig[category] || this.categoryConfig['other'];

    // 표시 텍스트
    let text;
    if (category === 'unissued') {
      text = '미발행';
    } else {
      const s = String(billValue || '').trim();
      text = s || '미발행';
    }

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
