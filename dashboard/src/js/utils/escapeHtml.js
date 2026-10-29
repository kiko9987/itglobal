/**
 * innerHTML 템플릿에 넣는 글자 값 이스케이프 (본문·속성 둘 다 안전).
 * 사업자명·특이사항·주소처럼 슬랙 모달·OCR·편집자가 넣은 값이 <img onerror=...> 로 실행되지 않게 (2026-10-29 점검).
 */
export function escHtml(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
