/**
 * 고객 리드 시트 값 해석 — 순수 함수 (리드 관리 페이지·필터 공용)
 */

// 실제 쓰이는 상태값 (시트 실데이터 + 슬랙 흐름 기준). '방문 대기'·'공사 드랍'은
// 사용 0건이라 선택지에서 제외 — 목록 밖 값이 들어있으면 화면이 그대로 보존 표시한다.
export const LEAD_STATUS_OPTIONS = [
    '상담 대기', '유선 상담', '부재중',
    '방문 예약', '방문 완료', '방문 취소',
    '견적 제출', '문의 드랍',
    '공사 확정', '공사 취소',
];

export function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// 공동방문 "권태훈,강정권" → ['권태훈', '강정권'] (leads.api_search_leads_for_project 와 같은 구분자)
export function splitOwners(raw) {
    const v = String(raw || '').trim();
    if (!v || v === '-') return [];
    return v.split(/[,/·&.\s]+/).map((s) => s.trim()).filter(Boolean);
}

const pad2 = (n) => String(n).padStart(2, '0');

function toIso(s) {
    // 연도 2자리("26-01-22")는 20YY 로 해석 — 시트에 수기 입력 몇 건 존재
    const m = String(s || '').trim().match(/^(\d{4}|\d{2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?$/);
    if (!m) return '';
    const y = m[1].length === 2 ? `20${m[1]}` : m[1];
    return `${y}-${pad2(m[2])}-${pad2(m[3])}`;
}

export function localIso(d) {
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

export function shiftIso(iso, days) {
    const d = new Date(`${iso}T00:00:00`);
    d.setDate(d.getDate() + days);
    return localIso(d);
}

export function daysBetween(a, b) {
    return Math.round((new Date(`${b}T00:00:00`) - new Date(`${a}T00:00:00`)) / 86400000);
}

/**
 * 방문 예정일 → {start, end} (ISO). 슬랙과 같은 양식
 * (slack_helpers._split_visit_date_range). '~'는 시간이 아니라 방문 기간(범위)이다.
 *   "2026-06-25" / "2026-06-23~24" / "2026-07-01~08-15" / "2026-12-30~2027-01-02"
 * 시트 escape(') · '-' · 읽을 수 없는 값은 start='' 로 돌려준다.
 */
export function splitVisitRange(raw) {
    const v = String(raw || '').replace(/^'/, '').trim();
    if (!v || v === '-') return { start: '', end: '' };
    const [a, b] = v.split('~', 2).map((x) => (x || '').trim());
    const start = toIso(a);
    if (!start) return { start: '', end: '' };
    if (!b) return { start, end: '' };
    let end = '';
    if (/^\d{1,2}$/.test(b)) {
        end = `${start.slice(0, 8)}${pad2(b)}`;
    } else if (/^\d{1,2}[-.]\d{1,2}$/.test(b)) {
        const [mm, dd] = b.split(/[-.]/);
        end = `${start.slice(0, 4)}-${pad2(mm)}-${pad2(dd)}`;
    } else {
        end = toIso(b);
    }
    return { start, end: end && end > start ? end : '' };
}

/** {start, end} → 시트 양식 (slack_helpers._format_visit_date_range 와 동일). */
export function formatVisitRange(start, end) {
    if (!start) return '';
    if (!end || end <= start) return start;
    const [sy, sm] = start.split('-');
    const [ey, em, ed] = end.split('-');
    if (sy === ey && sm === em) return `${start}~${ed}`;
    if (sy === ey) return `${start}~${em}-${ed}`;
    return `${start}~${end}`;
}

/**
 * 상담 내용(K열) → 회차 배열. slack_bot._parse_consultation_entries 와 같은 규칙.
 * 저장 형식: "[MM.DD HH:MM 이니셜 · 상태] 내용" 을 '─────────'(또는 인라인 ' ─── ')로 누적.
 * 헤더 없는 옛 형식은 {time:'', ini:'', status:'', content} 한 회차로 돌려준다.
 */
const CONSULT_ENTRY_RE = /^\[\s*(\d{2}\.\d{2}\s+\d{2}:\d{2})\s+(\S+)\s*·\s*([^\]]+)\]\s*([\s\S]*)$/;

export function parseConsultEntries(text) {
    const v = String(text || '').trim();
    if (!v || v === '-') return [];
    return v.split(/\s*─{3,}\s*/)
        .map((c) => c.trim())
        .filter(Boolean)
        .map((chunk) => {
            const m = chunk.match(CONSULT_ENTRY_RE);
            return m
                ? { time: m[1].trim(), ini: m[2].trim(), status: m[3].trim(), content: m[4].trim() }
                : { time: '', ini: '', status: '', content: chunk };
        });
}

/** 상담 시간 정렬 키 YYYYMMDDHHmm — 시트에 형식 8종 혼재 ("2026.09.29. 10:00", "2026. 09. 29. 오후 3:05" 등). */
export function consultSortKey(raw) {
    const s = String(raw || '');
    const m = s.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D*?(오전|오후)?\s*(\d{1,2})[:.](\d{2})/);
    if (m) {
        let h = Number(m[5]);
        if (m[4] === '오후' && h < 12) h += 12;
        if (m[4] === '오전' && h === 12) h = 0;
        return `${m[1]}${pad2(m[2])}${pad2(m[3])}${pad2(h)}${m[6]}`;
    }
    const d = s.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    return d ? `${d[1]}${pad2(d[2])}${pad2(d[3])}0000` : '';
}

/** 정렬 키(YYYYMMDDHHmm) → ISO 날짜 'YYYY-MM-DD' ('' 가능) */
export function sortKeyToIso(key) {
    return key && key.length >= 8 ? `${key.slice(0, 4)}-${key.slice(4, 6)}-${key.slice(6, 8)}` : '';
}

/** 상담 시간 표시 — "09.29 14:20" (시각 없으면 날짜만) */
export function consultDisplay(raw) {
    const k = consultSortKey(raw);
    if (!k) return String(raw || '').trim();
    const hm = k.slice(8, 12);
    return `${k.slice(4, 6)}.${k.slice(6, 8)}${hm && hm !== '0000' ? ` ${hm.slice(0, 2)}:${hm.slice(2)}` : ''}`;
}

/** 상담 시간 전체 표시 — "2026-09-30 11:48" (시각 없으면 날짜만, 해석 불가면 원문) */
export function consultFull(raw) {
    const k = consultSortKey(raw);
    if (!k) return String(raw || '').trim();
    const hm = k.slice(8, 12);
    return `${sortKeyToIso(k)}${hm && hm !== '0000' ? ` ${hm.slice(0, 2)}:${hm.slice(2)}` : ''}`;
}

/** 경과 라벨 — 오늘 / 어제 / n일 전 / n개월 전 / n년 전 */
export function agoLabel(iso, today = localIso(new Date())) {
    if (!iso) return '';
    const days = daysBetween(iso, today);
    if (days <= 0) return '오늘';
    if (days === 1) return '어제';
    if (days < 31) return `${days}일 전`;
    if (days < 365) return `${Math.floor(days / 30)}개월 전`;
    return `${Math.floor(days / 365)}년 전`;
}

// 상태 → 배지 클래스 (leads.css 의 lead-status-*)
const STATUS_CLASS = {
    '상담 대기': 'waiting', '유선 상담': 'consult', '부재중': 'absent',
    '방문 예약': 'visit', '방문 완료': 'visited', '방문 취소': 'cancel',
    '견적 제출': 'quote', '문의 드랍': 'drop',
    '공사 확정': 'confirmed', '공사 취소': 'cancel',
};

export function leadStatusClass(status) {
    return `lead-status-${STATUS_CLASS[String(status || '').trim()] || 'etc'}`;
}

/** 방문 모드 대상 — 방문일이 잡혀 있거나 방문 예약 상태인 리드 */
export function isVisitLead(lead) {
    return !!splitVisitRange(lead['방문 예정일']).start || String(lead['상태'] || '').trim() === '방문 예약';
}
