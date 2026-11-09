"""전자결재 엔진 (2026-10-30) — 플렉스 워크플로우 대체, 1차 양식 = 자금집행.

결재선(지출형): 요청 → ① 승인(고광일) → ② 최종 승인(JW·YG 중 한 명, 11-09) → ③ 이체 완료(경영지원 명단) → 끝.
  - 승인자가 직접 요청하면 ① 은 자동 통과. JW·YG 본인 요청이면 ② 는 자동 통과.
  - 반려(사유 필수)·요청 취소(요청자·승인자)는 끝나기 전 언제든.
  - EAPPROVAL_FINAL_ENABLED=false 면 ② 없이 예전 2단계.
상태: pending(승인 대기) → final(최종 승인 대기) → approved(이체 대기) → done(이체 완료)
      pending/approved → rejected(반려) / cancelled(취소)

여기는 슬랙 I/O 없는 장부·상태 전환·카드 모양만. 슬랙 처리는 blueprints/slack_eapproval.py.
설계: docs/e_approval_design.md
"""
from __future__ import annotations

import json
import logging
import os
import re
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

FORM_FUND = 'fund'
FORM_SUPPLY = 'supply'   # 비품 신청(11-01) — 회사 재고(프로젝트0)로 사는 것. ② 는 이체 대신 구매.
FORM_LABELS = {FORM_FUND: '자금집행', FORM_SUPPLY: '비품 신청'}
_STEP2 = {FORM_FUND: '이체', FORM_SUPPLY: '구매'}   # ② 단계 이름(이체 완료 / 구매 완료)

# 11-01 나머지 양식(계약서·보증보험·카드결제·법인서류·법인카드 지출결의=업무요청형, 개인경비=지출형)
from dashboard.services import e_approval_forms as forms  # noqa: E402

for _k, _f in forms.FORMS.items():
    FORM_LABELS[_k] = _f.label
    _STEP2[_k] = '처리' if _f.kind == 'task' else '이체'
TASK_FORMS = forms.TASK_FORMS   # 업무요청형 — 승인 없이 바로 경영지원 처리 대기, ② = 처리 완료

PENDING, APPROVED, DONE, REJECTED, CANCELLED = 'pending', 'approved', 'done', 'rejected', 'cancelled'
REVISE = 'revise'   # 11-05 수정 요청 — 승인자·경영지원이 돌려보냄 → 요청자가 같은 결재를 고쳐 다시 올림
# 11-09 최종 승인 — 지출형은 ① KiKO 승인 뒤 JW·YG 중 한 명이 이유를 보고 최종 승인해야 이체 대기로(이유 모르는 지출 방지)
FINAL = 'final'
OPEN_STATUSES = (PENDING, FINAL, APPROVED, REVISE)
STATUS_LABELS = {
    PENDING: ('🟡', '승인 대기'),
    FINAL: ('🟣', '최종 승인 대기'),
    APPROVED: ('🔵', '이체 대기'),
    REVISE: ('🟠', '수정 요청'),
    DONE: ('🟢', '이체 완료'),
    REJECTED: ('🔴', '반려'),
    CANCELLED: ('⚪', '요청 취소'),
}

EVIDENCE_TYPES = [('전자계산서', '전자계산서'), ('수기계산서', '수기계산서'), ('기타영수증', '기타영수증')]
URGENCY_TYPES = [('일반', '일반'), ('긴급', '긴급')]
GRN_TYPES = [('G', 'G (글로벌)'), ('R', 'R (글로벌그룹)'), ('N', 'N (사업자 없음)')]   # 값은 G/R/N 그대로

# 슬랙 식별자
A_APPROVE, A_REJECT, A_TRANSFER, A_CANCEL = 'eap_approve', 'eap_reject', 'eap_transfer', 'eap_cancel'
A_FINAL = 'eap_final'                 # [최종 승인] (JW·YG)
A_NEW = 'eap_new'                     # 채널 메뉴 [💸 자금집행 요청]
A_NEW_FROM_PROJECT = 'eap_new_project'  # 공사확정 카드 [💸 자금집행 요청]
A_PROJECT = 'eap_project'             # 새 요청 모달 프로젝트 external_select
A_NEW_SUPPLY = 'eap_new_supply'       # 채널 메뉴 [🛒 비품 신청]
A_FORM_PICK = 'eap_form_pick'         # 채널 메뉴 「다른 결재 요청」 고르기 → 양식 창
A_MY_REQUESTS = 'eap_my_requests'     # 채널 메뉴 [📂 내 요청] → 본인에게만 보이는 진행 목록(11-03)
A_FORM_OPEN = 'eap_form_open'       # 버튼 value=양식 key → 양식 창(법인카드 월 안내 등)
A_PROJECT_FORM_PICK = 'eap_project_form_pick'   # 공사확정 카드 「📋 결재 요청」 고르기(공사봇)
A_INVOICE = 'eap_invoice'             # 자금집행 모달 「계산서」 external_select(현장 미지정 계산서 글)
A_PARTNER = 'eap_partner'             # 자금집행 모달 「거래처」 external_select(아는 거래처 + 새 거래처)
CB_NEW, CB_REJECT = 'eap_new_submit', 'eap_reject_submit'
A_REVISE, CB_REVISE = 'eap_revise', 'eap_revise_submit'   # [✏️ 수정 요청] → 무엇을 고칠지 적는 창(11-05)
A_EDIT = 'eap_edit'                                         # [✏️ 수정하기] → 원래 값이 채워진 양식 창
ATTACH_BLOCK = 'attach'                                     # 창 「첨부 파일」 칸(file_input) — 카드 스레드로 옮겨 올림
CB_SUPPLY = 'eap_supply_submit'
SC_HUB = 'eap_sc_hub'   # 전역 바로가기 「전자결재」 하나(입력창 [+]) → 메뉴와 같은 창 — manifest 와 같게(11-05)
A_HUB_BACK = 'eap_hub_back'           # 바로가기 창 「내 요청」 목록 → [← 처음으로]
FN_OPEN_MENU = 'eap_open_menu'        # 워크플로 사용자 지정 단계 「전자결재 열기」 → 메뉴 창 — manifest functions 와 같게(11-05)

_DEFAULT_APPROVERS = ('U04UL2ZLJAX',)  # 고광일


def _env_ids(name: str) -> tuple:
    return tuple(x.strip() for x in os.getenv(name, '').split(',') if x.strip())


def approver_ids() -> tuple:
    """① 승인할 수 있는 사람."""
    return _env_ids('EAPPROVAL_APPROVER_IDS') or _DEFAULT_APPROVERS


def executor_ids() -> tuple:
    """② 이체 완료를 누를 수 있는 사람 = 정산 승인 명단(SB·KiKO·JW·YG, 경영지원 보조 추가 시 거기에)."""
    from dashboard.services.settlement_checkers import checker_ids
    return checker_ids()


_DEFAULT_N_EXECUTORS = ('U050XNG5VBR', 'U051ENK8AAY')   # 박용구(YG)·박정우(JW) — N통장 관리
_N_LABELS = {'U050XNG5VBR': 'YG', 'U051ENK8AAY': 'JW'}


def is_n_account(row: Optional[dict]) -> bool:
    """N통장에서 나가는 지출(자금집행·비품). 개인경비는 사업자 구분이 없어 경영지원이 판단."""
    return bool(row) and row.get('form') in (FORM_FUND, FORM_SUPPLY) and str(row.get('grn') or '')[:1].upper() == 'N'


def n_executor_ids() -> tuple:
    return _env_ids('EAPPROVAL_N_EXECUTOR_IDS') or _DEFAULT_N_EXECUTORS


_DEFAULT_FINAL = ('U051ENK8AAY', 'U050XNG5VBR')   # 박정우(JW)·박용구(YG) — 둘 중 한 명이 최종 승인(11-09)


def final_approver_ids() -> tuple:
    return _env_ids('EAPPROVAL_FINAL_APPROVER_IDS') or _DEFAULT_FINAL


def final_names() -> str:
    return '·'.join(_N_LABELS.get(u, '최종 승인자') for u in final_approver_ids())


def needs_final(form: str) -> bool:
    """최종 승인이 있는 양식 — 지출형(자금집행·비품·개인경비). EAPPROVAL_FINAL_ENABLED=false 면 끔."""
    return form not in TASK_FORMS and os.getenv('EAPPROVAL_FINAL_ENABLED', 'true').strip().lower() != 'false'


def after_first(form: str, requester_id: str, actor_id: str = '') -> Tuple[str, bool]:
    """① 승인 다음 상태와 최종 승인 자동 통과 여부 — JW·YG 본인 요청이면 최종 승인은 자동(본인이 이유를 안다).
    JW·YG 가 ① 을 직접 누르면(KiKO 휴가 등, 11-09) 최종 승인까지 한 번에."""
    if not needs_final(form):
        return APPROVED, False
    if requester_id in final_approver_ids() or actor_id in final_approver_ids():
        return APPROVED, True
    return FINAL, False


def self_approves(form: str, requester_id: str) -> bool:
    """요청자 본인 승인으로 ① 을 건너뛰는지 — KiKO, 그리고 JW·YG(본인 요청은 바로 최종 승인, 11-09 KiKO)."""
    if form in TASK_FORMS:
        return False
    return requester_id in approver_ids() or (needs_final(form) and requester_id in final_approver_ids())


def first_names() -> str:
    """① 승인할 수 있는 사람 — KiKO + JW·YG(KiKO 휴가일 수 있어 11-09)."""
    return '·'.join(['KiKO'] * bool(approver_ids()) + [_N_LABELS.get(u, '최종 승인자') for u in final_approver_ids()
                                                       if u not in approver_ids()])


def executor_dm_ids(row: Optional[dict] = None) -> tuple:
    """이체 요청 DM 받을 사람. 글로벌·글로벌그룹(G·R) 통장 = 경영지원(명단 첫 줄, SB),
    N통장 = YG·JW(11-06 KiKO). 다른 사람도 카드에서 대신 누를 수 있다.
    EAPPROVAL_EXECUTOR_DM_IDS 가 있으면(시험 모드) 통장과 관계없이 그 사람에게만."""
    ids = _env_ids('EAPPROVAL_EXECUTOR_DM_IDS')
    if ids:
        return ids
    if is_n_account(row):
        return n_executor_ids()
    return executor_ids()[:1]


def notice_ids(row: Optional[dict]) -> tuple:
    """이체 요청은 안 받지만 알아야 하는 사람 — N통장 건이면 경영지원(SB)에게 'YG·JW 에게 요청됐다' 알림만."""
    if _env_ids('EAPPROVAL_EXECUTOR_DM_IDS') or not is_n_account(row):
        return ()
    gr = executor_ids()[:1]
    return tuple(u for u in gr if u not in n_executor_ids())


def channel_id() -> str:
    return os.getenv('SLACK_EAPPROVAL_CHANNEL', '').strip()


# ── 장부 ──────────────────────────────────────────────────────────────
_DDL = [
    '''CREATE TABLE IF NOT EXISTS approvals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        form TEXT NOT NULL,
        status TEXT NOT NULL,
        requester_id TEXT,
        requester_name TEXT,
        project_code TEXT,
        cost_category TEXT,
        content TEXT,
        evidence TEXT,
        grn TEXT,
        partner TEXT,
        payee_name TEXT,
        bank TEXT,
        account TEXT,
        amount INTEGER NOT NULL DEFAULT 0,
        urgency TEXT,
        memo TEXT,
        origin_channel TEXT,
        origin_ts TEXT,
        card_channel TEXT,
        card_ts TEXT,
        dm_refs TEXT NOT NULL DEFAULT '[]',
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_approvals_status ON approvals(status)',
    'CREATE INDEX IF NOT EXISTS idx_approvals_origin ON approvals(origin_channel, origin_ts)',
    'CREATE INDEX IF NOT EXISTS idx_approvals_project ON approvals(project_code)',
    '''CREATE TABLE IF NOT EXISTS approval_steps (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        approval_id INTEGER NOT NULL,
        step TEXT NOT NULL,
        actor_id TEXT,
        actor_name TEXT,
        result TEXT NOT NULL,
        reason TEXT,
        at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_approval_steps_aid ON approval_steps(approval_id)',
    '''CREATE TABLE IF NOT EXISTS payee_accounts (
        partner_key TEXT PRIMARY KEY,
        partner TEXT,
        payee_name TEXT,
        bank TEXT,
        account TEXT,
        last_used TIMESTAMP
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


def _ensure_tables(conn) -> None:
    global _table_ready
    if _table_ready:
        return
    for ddl in _DDL:
        conn.execute(ddl)
    # 11-01 비품 신청 — 양식별 추가 칸(상품 링크·수량·단가·사유)은 JSON 한 칸에
    cols = {r[1] for r in conn.execute('PRAGMA table_info(approvals)').fetchall()}
    if 'extra' not in cols:
        conn.execute("ALTER TABLE approvals ADD COLUMN extra TEXT NOT NULL DEFAULT '{}'")
    conn.commit()
    _table_ready = True


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _row(conn, aid: int) -> Optional[dict]:
    r = conn.execute('SELECT * FROM approvals WHERE id=?', (aid,)).fetchone()
    return dict(r) if r else None


def get(aid: int) -> Optional[dict]:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return _row(conn, aid)


def steps(aid: int) -> List[dict]:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM approval_steps WHERE approval_id=? ORDER BY id', (aid,)).fetchall()]


def find_by_origin(channel: str, ts: str) -> Optional[dict]:
    """계산서 글에서 시작한 결재(가장 최근 것)."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute(
            'SELECT * FROM approvals WHERE origin_channel=? AND origin_ts=? ORDER BY id DESC LIMIT 1',
            (channel, ts)).fetchone()
        return dict(r) if r else None


def find_by_project(form: str, project_code: str) -> List[dict]:
    """그 현장의 그 양식 결재(반려·취소 제외, 최근 순) — 보증보험 창에서 계약서 확인용."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM approvals WHERE form=? AND project_code=? AND status NOT IN (?,?) ORDER BY id DESC',
            (form, project_code, REJECTED, CANCELLED)).fetchall()]


def partner_key(name: str) -> str:
    """거래처 이름 비교용 — (주)·㈜·주식회사·공백·기호 제거, 소문자."""
    s = re.sub(r'\(주\)|㈜|주식회사|\(유\)|유한회사', '', name or '')
    return re.sub(r'[\s\.\,\-\(\)·]', '', s).lower()


def find_unlinked_match(grn: str, amount: int, partner: str) -> Optional[dict]:
    """계산서 글과 G/R·금액·거래처가 같은데 아직 계산서 글에 안 묶인 자금집행(반려·취소 제외).

    딱 한 건일 때만 돌려준다(여러 건이면 사람이 고른다)."""
    key = partner_key(partner)
    if not key or not amount:
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            '''SELECT * FROM approvals WHERE form=? AND origin_ts IS NULL AND amount=? AND grn=?
               AND status IN (?,?,?,?,?) ORDER BY id DESC LIMIT 20''',
            # 수정 요청 중인 결재도 후보 — 빼면 같은 건으로 결재를 또 올려 두 번 이체할 수 있다(11-05)
            (FORM_FUND, int(amount), grn, PENDING, FINAL, APPROVED, REVISE, DONE)).fetchall()
    hits = [dict(r) for r in rows if partner_key(r['partner']) == key]
    return hits[0] if len(hits) == 1 else None


def link_origin(aid: int, channel: str, ts: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('UPDATE approvals SET origin_channel=?, origin_ts=?, updated_at=? WHERE id=? AND origin_ts IS NULL',
                     (channel, ts, _now(), aid))
        conn.commit()


_EXTRA_LOCK = threading.Lock()


def add_extra_files(aid: int, files: List[dict]) -> None:
    """창에서 올린 첨부 [{name, url}] 를 extra['files'] 에 쌓는다 — 이체 완료 때 수기계산서 드라이브 저장용(11-09).
    첨부·계산서 사진을 올리는 스레드가 동시에 돌 수 있어 잠금."""
    if not files:
        return
    with _EXTRA_LOCK, _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT extra FROM approvals WHERE id=?', (aid,)).fetchone()
        if not r:
            return
        x = json.loads((r[0] if not hasattr(r, 'keys') else r['extra']) or '{}')
        have = {f.get('url') for f in x.get('files') or []}
        x['files'] = (x.get('files') or []) + [{'name': f.get('name') or 'file', 'url': f.get('url') or ''}
                                               for f in files if f.get('url') and f.get('url') not in have]
        conn.execute('UPDATE approvals SET extra=? WHERE id=?', (json.dumps(x, ensure_ascii=False), aid))
        conn.commit()


def create(form: str, requester_id: str, requester_name: str, fields: Dict) -> dict:
    """결재 생성. 승인자가 직접 요청하면 ① 자동 통과(바로 이체 대기). 업무요청형은 승인 없이 바로 처리 대기."""
    task = form in TASK_FORMS
    auto = self_approves(form, requester_id)
    auto_final = False
    if task:
        status = APPROVED
    elif auto:
        status, auto_final = after_first(form, requester_id)
    else:
        status = PENDING
    now = _now()
    cols = ('project_code', 'cost_category', 'content', 'evidence', 'grn', 'partner', 'payee_name',
            'bank', 'account', 'amount', 'urgency', 'memo', 'origin_channel', 'origin_ts', 'extra')
    vals = [fields.get(c) for c in cols]
    vals[cols.index('amount')] = int(fields.get('amount') or 0)
    vals[cols.index('extra')] = json.dumps(fields.get('extra') or {}, ensure_ascii=False)
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        cur = conn.execute(
            f'''INSERT INTO approvals (form, status, requester_id, requester_name, {", ".join(cols)},
                created_at, updated_at) VALUES (?,?,?,?,{",".join("?" * len(cols))},?,?)''',
            (form, status, requester_id, requester_name, *vals, now, now))
        aid = cur.lastrowid
        conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, at) VALUES (?,?,?,?,?,?)',
                     (aid, 'request', requester_id, requester_name, 'requested', now))
        if auto:
            conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)',
                         (aid, 'approve', requester_id, requester_name, 'auto', '요청자가 승인자', now))
        if auto_final:
            _auto_final_step(conn, aid, requester_id, requester_name, now)
        conn.commit()
        return _row(conn, aid)


def _auto_final_step(conn, aid: int, actor_id: str, actor_name: str, now: str,
                     reason: str = '요청자가 최종 승인자') -> None:
    conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, reason, at) '
                 'VALUES (?,?,?,?,?,?,?)', (aid, 'final', actor_id, actor_name, 'auto', reason, now))


_TRANSITIONS = {
    # 동작: (허용 이전 상태, 다음 상태, 단계 이름, 결과) — approve 의 다음 상태는 after_first 로 정한다
    'approve': ((PENDING,), APPROVED, 'approve', 'approved'),
    'final': ((FINAL,), APPROVED, 'final', 'approved'),
    'reject': (OPEN_STATUSES, REJECTED, 'approve', 'rejected'),
    'transfer': ((APPROVED,), DONE, 'transfer', 'done'),
    'cancel': (OPEN_STATUSES, CANCELLED, 'cancel', 'cancelled'),
    'revise': ((PENDING, FINAL, APPROVED), REVISE, 'revise', 'revise'),
}


def can_act(action: str, row: dict, user_id: str) -> bool:
    """권한 — 승인=KiKO·JW·YG(본인 요청은 본인이 못 누름), 최종 승인=JW·YG, 반려=승인자(승인·최종 승인 대기 중엔
    JW·YG, 이체 대기 중엔 이체 담당도), 이체=이체 담당, 취소=요청자·승인자."""
    if action == 'approve':
        return user_id in approver_ids() or (user_id in final_approver_ids() and user_id != row.get('requester_id'))
    if action == 'final':
        return user_id in final_approver_ids()
    if action == 'reject':
        st = row.get('status')
        return (user_id in approver_ids() or (st in (PENDING, FINAL) and user_id in final_approver_ids())
                or (st == APPROVED and user_id in executor_ids()))
    if action == 'transfer':
        return user_id in executor_ids()
    if action == 'cancel':
        if user_id in approver_ids():
            return True
        return user_id == row.get('requester_id') and requester_can_cancel(row)
    if action == 'revise':   # 수정 요청 = 승인자(KiKO) + 경영지원(SB·JW·YG) + 최종 승인자 — 플렉스 댓글로 돌려보내던 사람들
        return user_id in approver_ids() or user_id in executor_ids() or user_id in final_approver_ids()
    if action == 'edit':
        return user_id == row.get('requester_id')
    return False


def requester_can_cancel(row: dict) -> bool:
    """요청자 취소 — 승인 뒤(이체·구매 대기)엔 경영지원이 이미 은행에서 보내는 중일 수 있어 막는다(11-06 검수).
    그때는 경영지원에게 [반려]를 부탁한다. 업무요청형은 승인 단계가 없어 처리 대기 중에도 취소 가능."""
    st = row.get('status')
    if st in (PENDING, FINAL, REVISE):   # 최종 승인 전까진 아직 돈이 안 나간다
        return True
    return st == APPROVED and row.get('form') in TASK_FORMS


# 다시 올릴 때 고칠 수 있는 장부 칸(양식 값 전부는 extra 에)
EDIT_COLUMNS = ('project_code', 'cost_category', 'content', 'evidence', 'grn', 'partner', 'payee_name',
                'bank', 'account', 'amount', 'urgency', 'memo')


def resubmit(aid: int, actor_id: str, actor_name: str, fields: Dict) -> Optional[Tuple[dict, dict]]:
    """수정 요청 → 요청자가 고쳐 다시 올림. 같은 결재의 칸을 바꾸고 처음 올릴 때와 같은 단계로 되돌린다
    (업무요청형·승인자 본인 요청 = 처리 대기, 나머지 = 승인 대기). (새 row, 바뀐 칸 {칸: (전, 후)}) 또는 None."""
    now = _now()
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        old = _row(conn, aid)
        if not old or old['status'] != REVISE:
            return None
        task = old['form'] in TASK_FORMS
        self_ok = self_approves(old['form'], old['requester_id'])
        auto_final = False
        if task:
            nxt = APPROVED
        elif self_ok:
            nxt, auto_final = after_first(old['form'], old['requester_id'])
        else:
            nxt = PENDING
        sets, vals, changed = [], [], {}
        for c in EDIT_COLUMNS:
            if c not in fields:
                continue
            v = int(fields.get(c) or 0) if c == 'amount' else fields.get(c)
            if (old.get(c) or None) != (v or None):
                changed[c] = (old.get(c), v)
            sets.append(f'{c}=?')
            vals.append(v)
        if 'extra' in fields:
            new_x = dict(fields.get('extra') or {})   # 결재 이름도 고친 값으로(G→R 현장·지출결의 달, 11-05 점검)
            old_x = extra(old)
            ex = {**old_x, **new_x}
            # 바뀐 칸 — 양식 값(f) + 비품 칸(link·qty·unit_price·reason)
            pairs = [(old_x.get('f') or {}, new_x.get('f') or {})] if 'f' in new_x else []
            pairs.append(({k: v for k, v in old_x.items() if k not in ('f', 'payee_saved')},
                          {k: v for k, v in new_x.items() if k not in ('f', 'payee_saved')}))
            for a, b in pairs:
                for k in b:
                    if str(a.get(k) or '') != str(b.get(k) or '') and k not in changed:
                        changed[k] = (a.get(k), b.get(k))
            sets.append('extra=?')
            vals.append(json.dumps(ex, ensure_ascii=False))
        cur = conn.execute(f'UPDATE approvals SET {", ".join(sets + ["status=?", "updated_at=?"])} '
                           f'WHERE id=? AND status=?', (*vals, nxt, now, aid, REVISE))
        if cur.rowcount != 1:
            conn.rollback()
            return None
        # 승인 기록은 새로 받는다 — 예전 승인·처리 줄은 지우지 않고 'resubmit' 줄 뒤로 밀린다(step_line 은 마지막 것)
        conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, at) '
                     'VALUES (?,?,?,?,?,?)', (aid, 'resubmit', actor_id, actor_name, 'resubmitted', now))
        if self_ok:
            conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)', (aid, 'approve', actor_id, actor_name, 'auto', '요청자가 승인자', now))
        if auto_final:
            _auto_final_step(conn, aid, actor_id, actor_name, now)
        conn.commit()
        return _row(conn, aid), changed


def last_revise(step_rows: List[dict]) -> Optional[dict]:
    """가장 최근 수정 요청 줄(다시 올린 뒤면 None)."""
    for s in reversed(step_rows):
        if s['step'] == 'resubmit':
            return None
        if s['step'] == 'revise':
            return s
    return None


def transition(aid: int, action: str, actor_id: str, actor_name: str, reason: str = '') -> Optional[dict]:
    """상태 전환(두 사람이 동시에 눌러도 한 번만 — 이전 상태 조건부 UPDATE). 실패면 None."""
    allowed, nxt, step, result = _TRANSITIONS[action]
    now = _now()
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        auto_final = False
        if action == 'approve':   # ① 다음 — 지출형은 최종 승인 대기(JW·YG 본인 요청이면 자동 통과)
            old = _row(conn, aid)
            if not old:
                return None
            nxt, auto_final = after_first(old['form'], old['requester_id'], actor_id)
        elif action == 'reject':   # 최종 승인 대기 중 반려는 '최종 승인' 칸에 남긴다(① 승인 기록은 그대로)
            old = _row(conn, aid)
            if old and old['status'] == FINAL:
                step = 'final'
        cur = conn.execute(
            f'UPDATE approvals SET status=?, updated_at=? WHERE id=? AND status IN ({",".join("?" * len(allowed))})',
            (nxt, now, aid, *allowed))
        if cur.rowcount != 1:
            conn.rollback()
            return None
        conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)', (aid, step, actor_id, actor_name, result, reason or None, now))
        if auto_final and old['requester_id'] in final_approver_ids():
            _auto_final_step(conn, aid, old['requester_id'], old['requester_name'], now)
        elif auto_final:   # JW·YG 가 ① 을 눌러 최종 승인까지
            _auto_final_step(conn, aid, actor_id, actor_name, now, '최종 승인자가 승인')
        conn.commit()
        return _row(conn, aid)


def set_card(aid: int, channel: str, ts: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('UPDATE approvals SET card_channel=?, card_ts=? WHERE id=?', (channel, ts, aid))
        conn.commit()


def add_dm_ref(aid: int, channel: str, ts: str, role: str, user_id: str) -> None:
    """결재 DM 위치 기록 — 상태가 바뀌면 이 DM 들도 같이 고친다."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT dm_refs FROM approvals WHERE id=?', (aid,)).fetchone()
        refs = json.loads((r['dm_refs'] if r else None) or '[]')
        refs.append({'ch': channel, 'ts': ts, 'role': role, 'user': user_id})
        conn.execute('UPDATE approvals SET dm_refs=? WHERE id=?', (json.dumps(refs), aid))
        conn.commit()


def dm_refs(row: dict) -> List[dict]:
    return json.loads(row.get('dm_refs') or '[]')


def extra(row: dict) -> dict:
    try:
        return json.loads(row.get('extra') or '{}')
    except (TypeError, ValueError):
        return {}


# ── 거래처 계좌 기억 ─────────────────────────────────────────────────
def lookup_payee(partner: str) -> Optional[dict]:
    key = partner_key(partner)
    if not key:
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT * FROM payee_accounts WHERE partner_key=?', (key,)).fetchone()
        return dict(r) if r else None


def remember_payee(partner: str, payee_name: str, bank: str, account: str) -> None:
    key = partner_key(partner)
    if not (key and payee_name and bank and account):
        return
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute(
            '''INSERT INTO payee_accounts (partner_key, partner, payee_name, bank, account, last_used)
               VALUES (?,?,?,?,?,?)
               ON CONFLICT(partner_key) DO UPDATE SET partner=excluded.partner, payee_name=excluded.payee_name,
                 bank=excluded.bank, account=excluded.account, last_used=excluded.last_used''',
            (key, partner, payee_name, bank, account, _now()))
        conn.commit()


def personal_memory(uid: str) -> dict:
    """본인 기억 — 법인카드 번호·개인경비 받을 계좌 {'card','bank','account'}.

    payee_accounts 에 partner_key='card:U…'/'staff:U…', partner='@card'/'@staff' 로 둔다(거래처 목록에선 뺌)."""
    out = {}
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        for kind in ('card', 'staff'):
            r = conn.execute('SELECT * FROM payee_accounts WHERE partner_key=?', (f'{kind}:{uid}',)).fetchone()
            if r and kind == 'card':
                out['card'] = r['account']
            elif r:
                out.update(bank=r['bank'], account=r['account'])
    return out


def remember_personal(kind: str, uid: str, name: str, bank: str = '', account: str = '') -> None:
    if not (uid and account):
        return
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute(
            '''INSERT INTO payee_accounts (partner_key, partner, payee_name, bank, account, last_used)
               VALUES (?,?,?,?,?,?)
               ON CONFLICT(partner_key) DO UPDATE SET payee_name=excluded.payee_name, bank=excluded.bank,
                 account=excluded.account, last_used=excluded.last_used''',
            (f'{kind}:{uid}', f'@{kind}', name, bank, account, _now()))
        conn.commit()


# ── 입력 검사 ─────────────────────────────────────────────────────────
def parse_amount(s) -> int:
    """'533,500원' → 533500. 숫자가 없으면 0.
    시트 값은 420000.0 처럼 소수로 오기도 한다 — 숫자만 남기면 4200000(10배)이 됐다(11-09 R4251-TH)."""
    if isinstance(s, bool):
        return 0
    if isinstance(s, (int, float)):
        return int(round(s))
    s = re.sub(r'\.\d*\s*(원)?\s*$', '', str(s or '').strip())   # '420,000.00원' 의 소수 부분 버림
    digits = re.sub(r'[^\d]', '', s)
    return int(digits) if digits else 0


AMOUNT_HELP = '숫자로만 넣어 주세요. 예: 350000 (「35만」처럼 쓰면 35원이 돼요)'


def bad_amount(s) -> bool:
    """숫자·쉼표·공백·'원' 말고 다른 글자가 있으면 True — '35만원'이 35원으로 들어가던 것(11-06 검수)."""
    return bool(re.search(r'[^\d,\s원]', str(s or '')))


def amount_errors(raws: Dict[str, str]) -> Dict[str, str]:
    """{block_id: 입력 그대로} → {block_id: 오류}."""
    return {k: AMOUNT_HELP for k, v in raws.items() if bad_amount(v)}


def normalize_account(s: str) -> str:
    """계좌번호 — 숫자와 하이픈만."""
    return re.sub(r'[^\d\-]', '', (s or '').strip())


def grn_from_code(code: str) -> str:
    c = (code or '').strip()[:1].upper()
    return c if c in ('G', 'R') else ''


def validate_transfer(payee_name: str, bank: str, account: str) -> Dict[str, str]:
    """이체 정보 오류 {block_id: 문구}."""
    errs = {}
    if not (payee_name or '').strip():
        errs['payee_name'] = '예금주를 넣어 주세요.'
    if not (bank or '').strip():
        errs['bank'] = '은행을 넣어 주세요.'
    acct = normalize_account(account)
    if len(re.sub(r'\D', '', acct)) < 8:
        errs['account'] = '계좌번호를 숫자로 넣어 주세요.'
    return errs


# ── 카드 모양 ─────────────────────────────────────────────────────────
def _fmt_when(s: str) -> str:
    return str(s or '')[5:16].replace('-', '.')


_INITIAL_EXTRA = {'고광일': 'KiKO'}   # 코드 접미사 표에 없는 사람
_suffix_cache = {'at': 0.0, 'map': {}}


def _suffix_map() -> Dict[str, str]:
    """이름 → 이니셜(코드 접미사 표). 카드 한 장에 여러 번 부르므로 5분 캐시."""
    import time
    if time.time() - _suffix_cache['at'] > 300:
        try:
            from dashboard.services.staff_directory import suffix_map
            _suffix_cache['map'] = suffix_map()
        except Exception as exc:
            logger.warning(f'[전자결재] 이니셜 표 조회 실패: {exc}')
        _suffix_cache['at'] = time.time()
    return _suffix_cache['map']


def initial(name: str) -> str:
    """카드·DM·기록에 쓰는 이니셜(11-01) — 강성환→SH, 고광일→KiKO. 표에 없으면 이름 그대로."""
    name = (name or '').strip()
    if not name:
        return '-'
    return _INITIAL_EXTRA.get(name) or _suffix_map().get(name) or name


def executor_label(name: str) -> str:
    """② 처리자 표시 — '경영지원 (SB)'. 대신 누른 사람이면 그 사람 이니셜."""
    return f'경영지원 ({initial(name)})'


def title(row: dict) -> str:
    name = extra(row).get('title') or FORM_LABELS.get(row.get('form'), '결재')
    return f"{name} #{row['id']}"


def step2_word(row: dict) -> str:
    """② 단계 이름 — 자금집행 '이체', 비품 신청 '구매'."""
    return _STEP2.get(row.get('form'), _STEP2[FORM_FUND])


def status_label(row: dict) -> tuple:
    """(이모지, 이름) — ② 단계 이름이 양식마다 달라서(이체 대기 / 구매 대기)."""
    st = row.get('status')
    if st == APPROVED:
        return '🔵', f'{step2_word(row)} 대기'
    if st == DONE:
        return '🟢', f'{step2_word(row)} 완료'
    return STATUS_LABELS.get(st, ('', st))


def step_line(row: dict, step_rows: List[dict], approver_name: str = '고광일') -> str:
    """'결재: ✅ KiKO 승인 10.30 14:20 → ⏳ 이체' 한 줄 (이름 대신 이니셜, ② 는 '경영지원 (SB)').
    수정해 다시 올린 결재는 마지막으로 다시 올린 뒤의 승인·처리만 본다."""
    last = max((i for i, s in enumerate(step_rows) if s['step'] == 'resubmit'), default=-1)
    by = {s['step']: s for s in step_rows[last + 1:]}
    st = row.get('status')
    if row.get('form') in TASK_FORMS:   # 업무요청형 — 경영지원 처리 한 단계
        t = by.get('transfer')
        if t:
            return f"결재: ✅ {executor_label(t['actor_name'])} 처리 {_fmt_when(t['at'])}"
        return '결재: ⏳ 경영지원 처리' if st == APPROVED else '결재: ⏸ 경영지원 처리'
    a = by.get('approve')
    ap = executor_label(approver_name)   # ① 도 '경영지원 (KiKO)' — ② '경영지원 (SB)'와 같은 모양(11-09 KiKO)
    if needs_final(row.get('form')):   # ① 은 KiKO·JW·YG 누구나(11-09)
        ap = first_names()

    def _who(s):   # JW·YG 는 경영지원 명단에도 있지만 승인자로 누른 거라 이니셜만
        return initial(s['actor_name']) if s.get('actor_id') in final_approver_ids() else executor_label(s['actor_name'])
    f = by.get('final')
    if a and a['result'] == 'auto':
        first = '✅ 승인 (요청자가 승인자)'
    elif a and a['result'] == 'approved':
        first = f"✅ {_who(a)} 승인 {_fmt_when(a['at'])}"
    elif a and a['result'] == 'rejected':
        first = f"❌ {_who(a)} 반려 {_fmt_when(a['at'])}"
    else:
        first = f'⏳ {ap} 승인' if st == PENDING else f'⏸ {ap} 승인'
    # 최종 승인(11-09) — 기능을 켜기 전에 이미 이체 대기·완료로 넘어간 결재는 이 칸 없이 예전 모양
    mid = ''
    if f and f['result'] == 'auto' and a and f.get('actor_id') == a.get('actor_id') and a['result'] == 'auto':
        first = f"✅ 최종 승인 ({initial(f['actor_name'])} 본인 요청)"   # JW·YG 본인 요청 — ① 없이 바로(11-09)
    elif f and f['result'] == 'auto' and a and f.get('actor_id') == a.get('actor_id') and a['result'] == 'approved':
        first += ' (최종 승인 겸)'   # JW·YG 가 ① 을 눌러 한 번에
    elif f and f['result'] == 'auto':
        mid = '✅ 최종 승인 (요청자 본인)'
    elif f and f['result'] == 'approved':
        mid = f"✅ {initial(f['actor_name'])} 최종 승인 {_fmt_when(f['at'])}"
    elif f and f['result'] == 'rejected':
        mid = f"❌ {initial(f['actor_name'])} 반려 {_fmt_when(f['at'])}"
    elif needs_final(row.get('form')) and st in (PENDING, FINAL, REVISE, REJECTED, CANCELLED):
        mid = f'⏳ {final_names()} 최종 승인' if st == FINAL else f'⏸ {final_names()} 최종 승인'
    if mid:
        first = f'{first} → {mid}'
    w = step2_word(row)
    t = by.get('transfer')
    if t:
        second = f"✅ {executor_label(t['actor_name'])} {w} {_fmt_when(t['at'])}"
    elif st == APPROVED:
        second = f'⏳ {w}'
    else:
        second = f'⏸ {w}'
    return f'결재: {first} → {second}'


CARD_SEP = '--------------------------------------------'   # 온라인 리드 카드와 같은 구분선(11-06 KiKO)
GRN_NAMES = {'G': '글로벌 통장', 'R': '글로벌그룹 통장', 'N': 'N통장'}
GRN_BIZ = {'G': '글로벌', 'R': '글로벌그룹', 'N': '사업자 없음'}
_KV = re.compile(r'^([^:<>*\n]{1,16}): (.*)$', re.S)


def _kv(line: str) -> str:
    """'거래처: 컴퓨존' → '*거래처* : 컴퓨존' (온라인 리드 카드 모양). 항목 이름이 없는 줄은 그대로."""
    m = _KV.match(line)
    return f'*{m.group(1).strip()}* : {m.group(2)}' if m else line


def _key_line(row: dict) -> str:
    """카드 둘째 줄 — 무엇이, 얼마가, 어디로(이체 담당이 제일 먼저 볼 것)."""
    form = row.get('form')
    amt = int(row.get('amount') or 0)
    money = f'💰 *{amt:,}원*' if amt else ''
    if form == FORM_FUND:   # 통장은 셋째 줄(_tag_line)로 — 금액·거래처만 크게(11-09 SB)
        return f"{money} → *{row.get('partner') or '-'}*"
    if form == FORM_SUPPLY:
        qty = int(extra(row).get('qty') or 0)
        return f"{money} · {row.get('content') or '-'}" + (f' × {qty}개' if qty else '')
    if form == forms.FORM_EXPENSE:
        return f"{money} → {row.get('requester_name') or '-'} 본인 계좌 (경영지원 DM에서 확인)"
    what = forms.summary(form, extra(row).get('f') or {}) if form in forms.FORMS else (row.get('content') or '')
    return ' · '.join(x for x in (money, what) if x) or '-'


def _tag_line(row: dict) -> str:
    """카드 셋째 줄 — 사업자(통장)·프로젝트 코드를 굵게(11-09 SB: 이체할 때 이 둘을 계속 찾아 봐야 했다).
    둘 다 없으면 ''."""
    from dashboard.services.purchase_invoice_assign import NO_SITE_CODE
    g = str(row.get('grn') or '')[:1].upper()
    parts = []
    if g in GRN_NAMES:   # 자금집행은 돈이 나가는 통장, 나머지 양식은 사업자 이름
        name = GRN_NAMES[g] if row.get('form') == FORM_FUND else GRN_BIZ[g]
        parts.append(f'🏦 *{g} · {name}*')
    code = (row.get('project_code') or '').strip()
    if code and code != NO_SITE_CODE:
        parts.append(f'📍 *{code}*')
    return '   '.join(parts)


def _short_cost(row: dict) -> str:
    from dashboard.services.purchase_invoice_assign import cost_label
    c = row.get('cost_category')
    return cost_label(c).replace('공통비 · ', '') if c else ''


def card_text(row: dict, step_rows: List[dict], project_line: str, quote_amount: int = 0) -> str:
    """온라인 리드 카드 모양(11-06 KiKO) — 머리 · 무엇이/얼마/어디로 · 구분선 · 세부 · 구분선 · 요청·결재.
    끝난 결재(완료·반려·취소)는 세 줄로 접는다 — 채널에서 진행 중인 건이 눈에 띄게."""
    emoji, label = status_label(row)
    st = row.get('status')
    content = (row.get('content') or '').strip()
    step = step_line(row, step_rows).replace('결재: ', '', 1)
    closed = st in (DONE, REJECTED, CANCELLED)
    tag = _tag_line(row)
    gap = ['']   # 금액·통장 묶음 위아래 빈 줄 — 먼저 눈에 띄게(11-09 KiKO 예시 D, 끝난 카드 박스 안도 같게)
    top = ([f"{emoji} *{title(row)}* · {label}" + (' · 🚨 긴급' if row.get('urgency') == '긴급' and not closed else '')]
           + gap + [_key_line(row)] + ([tag] if tag else []) + gap + [CARD_SEP])
    head = top + [f'현장: {project_line}']
    lines = []
    if row.get('form') in forms.FORMS:
        lines += forms.card_lines(row['form'], extra(row).get('f') or {})
    elif row.get('form') == FORM_SUPPLY:
        x = extra(row)
        qty, unit = int(x.get('qty') or 0), int(x.get('unit_price') or 0)
        if qty and unit:
            lines.append(f'단가: {unit:,}원 × {qty}개 (VAT 포함)')
        if x.get('link'):
            lines.append(f"상품: <{x['link']}|상품 보기>")
        if x.get('reason'):
            lines.append(f"사유: {x['reason']}")
    else:
        cost = _short_cost(row)
        if cost:
            lines.append(f'분류: {cost}')
        if content:
            lines.append(f'내용: {content}')
        # 예금주가 '(주)컴퓨존'이면 괄호가 겹쳐 '((주)컴퓨존)' — 예금주는 · 뒤로(11-06)
        lines.append(f"계좌: {row.get('bank') or '-'} {row.get('account') or '-'} · 예금주 {row.get('payee_name') or '-'}"
                     + (' · 예전 계좌 자동 채움' if extra(row).get('payee_saved') else ''))
        lines.append(f"증빙: {row.get('evidence') or '-'} (VAT 포함 금액)")
        if quote_amount:
            # 공사현황 「총액 1」 = 공사 전체 금액 — '고객 견적'이라 쓰면 항목 견적으로 오해(11-06 검수)
            lines.append(f'공사 총액: {int(quote_amount):,}원 (VAT 별도)')
    memo = (row.get('memo') or '').strip()
    if memo:
        lines.append(f"{'구매 요청사항' if row.get('form') == FORM_SUPPLY else '비고 (특이사항)'}: {memo}")
    foot = [CARD_SEP, f"요청: {initial(row.get('requester_name'))} {_fmt_when(row.get('created_at'))}",
            f'결재: {step}']
    rv = last_revise(step_rows) if st == REVISE else None
    if rv:
        foot.append(f"✏️ {executor_label(rv['actor_name'])} 수정 요청 {_fmt_when(rv['at'])} — {rv.get('reason') or '-'}")
    body = top + [_kv(ln) for ln in head[len(top):] + lines + foot]
    if closed:
        return _greyed(row, step_rows, emoji, label, body)
    return '\n'.join(f'>{ln}' for ln in '\n'.join(body).split('\n'))


def _greyed(row: dict, step_rows: List[dict], emoji: str, label: str, body: List[str]) -> str:
    """끝난 결재 — 온라인 리드 '상담 완료' 카드처럼(11-06 KiKO): 처리 결과 몇 줄 + 원래 내용은 회색 코드 박스."""
    last = next((s for s in reversed(step_rows) if s['step'] in ('transfer', 'approve', 'final', 'cancel')
                 and s['result'] in ('done', 'rejected', 'cancelled')), None)
    who = ''
    if last:
        # 요청자 본인 취소·최종 승인자(JW·YG) 반려는 이니셜만
        mine = (last['result'] == 'cancelled' and last.get('actor_name') == row.get('requester_name')) \
            or last['step'] == 'final'
        who = initial(last['actor_name']) if mine else executor_label(last['actor_name'])   # 요청자 본인 취소만 이니셜
    # 긴급이었던 건은 끝난 뒤에도 표시(11-09 SB)
    hdr = [f"{emoji} *{title(row)} · {label}*" + (' · 🚨 긴급' if row.get('urgency') == '긴급' else ''),
           '⠀']   # 제목 밑 한 줄 띄움(11-09 KiKO) — 빈 줄은 슬랙이 지워 '⠀' 빈 글자
    if last:
        hdr += [f'처리자 : {who}', f"처리 시간 : {_fmt_when(last['at'])}"]
        if last['result'] in ('rejected', 'cancelled') and last.get('reason'):
            hdr.append(f"{'반려' if last['result'] == 'rejected' else '취소'} 사유 : {last['reason']}")
    # 굵은 글씨 표시만 지운다 — '*'를 다 지우면 법인카드 '4518-****-****-1234'가 '4518---1234'로 깨졌다(11-06 시뮬)
    lines = [re.sub(r'^\*([^*\n]+)\* : ', r'\1 : ', ln) for ln in body]
    lines = [re.sub(r'(?<![*\w])\*([^*\s][^*\n]*?)\*(?![*\w])', r'\1', ln) for ln in lines]
    plain = '\n'.join(lines).replace('```', "'''")
    plain = re.sub(r'<(https?://[^|>]+)\|([^>]+)>', r'\2 (\1)', plain)   # 코드 박스 안에선 링크 모양이 안 바뀐다
    return '\n'.join(hdr) + f'\n```\n{plain}\n```'


def _btn(text: str, action: str, aid: int, style: Optional[str] = None, confirm: Optional[dict] = None) -> dict:
    b = {'type': 'button', 'action_id': action, 'value': str(aid),
         'text': {'type': 'plain_text', 'text': text, 'emoji': True}}
    if style:
        b['style'] = style
    if confirm:
        b['confirm'] = confirm
    return b


def _confirm(title_: str, text: str, ok: str) -> dict:
    if len(text) > 300:   # 슬랙 확인 창 글자 한도 — 넘으면 카드 갱신이 통째로 실패(11-06 검수)
        text = text[:297] + '…'
    return {'title': {'type': 'plain_text', 'text': title_},
            'text': {'type': 'plain_text', 'text': text},
            'confirm': {'type': 'plain_text', 'text': ok},
            'deny': {'type': 'plain_text', 'text': '아니요'}}


def buttons(row: dict, role: str = 'all') -> List[dict]:
    """상태·받는 사람 역할에 맞는 버튼. role: all(채널 카드) / approver / executor / requester."""
    st, aid = row.get('status'), row['id']
    amt = f"{int(row.get('amount') or 0):,}원"
    out = []
    if st == PENDING and role in ('all', 'approver', 'final'):   # JW·YG 도 ① 을 누를 수 있다(11-09)
        out += [_btn('승인', A_APPROVE, aid, 'primary'), _btn('반려', A_REJECT, aid, 'danger')]
    if st == FINAL and role in ('all', 'final'):
        out += [_btn('최종 승인', A_FINAL, aid, 'primary'), _btn('반려', A_REJECT, aid, 'danger')]
    form = row.get('form')
    if st == APPROVED and role in ('all', 'executor') and form in TASK_FORMS:
        out.append(_btn('✅ 처리 완료', A_TRANSFER, aid, 'primary', _confirm(
            '처리 완료', f'{title(row)} 처리를 마쳤나요?', '처리 완료')))
        out.append(_btn('반려', A_REJECT, aid, 'danger'))   # 내용이 모자라면 경영지원이 사유와 함께 돌려보냄
    elif st == APPROVED and role in ('all', 'executor') and form == forms.FORM_EXPENSE:
        # 확인 창은 채널 카드에 실려 누구나 열 수 있다 — 개인 계좌는 빼고(경영지원 DM 에만, 11-06 검수)
        out.append(_btn('💸 이체 완료', A_TRANSFER, aid, 'primary', _confirm(
            '이체 완료', f"{row.get('requester_name')} 개인경비 {amt}\n본인 계좌(경영지원 DM)로 이체를 마쳤나요?",
            '이체 완료')))
    elif st == APPROVED and role in ('all', 'executor') and form == FORM_SUPPLY:
        out.append(_btn('🛒 구매 완료', A_TRANSFER, aid, 'primary', _confirm(
            '구매 완료', f"{row.get('content') or '비품'} ({amt})\n구매를 마쳤나요?", '구매 완료')))
    elif st == APPROVED and role in ('all', 'executor'):
        out.append(_btn('💸 이체 완료', A_TRANSFER, aid, 'primary', _confirm(
            '이체 완료', f"{row.get('bank')} {row.get('account')} · 예금주 {row.get('payee_name')}\n{amt} 이체를 마쳤나요?",
            '이체 완료')))
    if st in (PENDING, FINAL, APPROVED) and role == 'all':   # 고칠 게 있으면 반려 대신 돌려보내기(11-05)
        out.append(_btn('✏️ 수정 요청', A_REVISE, aid))
    if st == REVISE and role in ('all', 'requester'):
        out.append(_btn('✏️ 수정하기', A_EDIT, aid, 'primary'))
    if st in OPEN_STATUSES and role in ('all', 'requester') and requester_can_cancel(row):
        out.append(_btn('요청 취소', A_CANCEL, aid, None, _confirm(
            '요청 취소', f'{title(row)} 요청을 취소할까요?', '요청 취소')))
    return out


def card_blocks(row: dict, step_rows: List[dict], project_line: str, quote_amount: int = 0,
                role: str = 'all') -> List[dict]:
    # 11-01 다른 채널 카드처럼 위·아래, 내용과 버튼 사이에 한 줄씩 여백.
    # section 글자 한도 3,000자 — 넘으면 카드 게시·갱신이 통째로 실패한다(긴 비고·사유, 11-05 점검)
    body = '⠀\n' + card_text(row, step_rows, project_line, quote_amount)
    if len(body) > 2900:
        body = body[:2880] + '\n…(길어서 줄였어요)'
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': body}}]
    btns = buttons(row, role)
    if btns:
        blocks.append(_SPACER)
        blocks.append({'type': 'actions', 'elements': btns})
    if row.get('status') == REVISE and role == 'all':
        hint = '요청자가 [✏️ 수정하기]로 고쳐 다시 올리면 이 카드가 바뀌어요. 간단한 확인은 스레드에 @이름으로 남겨 주세요.'
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': hint}]})
    elif row.get('status') in OPEN_STATUSES and role == 'all':
        form = row.get('form')
        if form in TASK_FORMS:
            hint = ('파일을 더 붙일 땐 이 카드 스레드에 올려 주세요. 경영지원은 결과 파일을 스레드에 올리고 '
                    '[처리 완료]를 눌러 주세요.')
        else:
            hint = '파일을 더 붙일 땐 이 카드 스레드에 올려 주세요. 고칠 게 있으면 [✏️ 수정 요청]으로 돌려보내 주세요.'
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': hint}]})
    blocks.append(_SPACER)
    return blocks


_SPACER = {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]}


def spaced(text: str, elements: Optional[List[dict]] = None, notes: Optional[List[dict]] = None) -> List[dict]:
    """다른 채널 카드와 같은 여백(11-06) — 위 빈 줄 · 본문 · 빈 줄 · 버튼 · (안내) · 아래 빈 줄."""
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': '⠀\n' + text}}]
    if elements:
        blocks += [_SPACER, {'type': 'actions', 'elements': elements}]
    blocks += notes or []
    blocks.append(_SPACER)
    return blocks


def dm_intro(row: dict, role: str) -> str:
    """DM 맨 위 한 줄 — 내 차례면 할 일, 아니면 지금 상태."""
    st = row.get('status')
    w = step2_word(row)
    if role == 'approver' and st == PENDING:
        return f"🔔 {initial(row.get('requester_name'))}님이 {title(row)} 승인을 요청했어요. 카드에서 [승인]이나 [반려]를 눌러 주세요."
    if role == 'final' and st == PENDING:   # KiKO 휴가일 수 있어 JW·YG 도 처음부터 받는다(11-09)
        return (f"🔔 {initial(row.get('requester_name'))}님이 {title(row)} 승인을 요청했어요. "
                f"{first_names()} 중 한 분이 [승인]이나 [반려]를 누르면 돼요. "
                '여기서 승인하면 최종 승인까지 한 번에 끝나요.')
    if role == 'final' and st == FINAL:
        return (f"🔔 {initial(row.get('requester_name'))}님이 올린 {title(row)} — 승인이 끝나 최종 승인 차례예요. "
                f"내용·이유를 확인하고 카드에서 [최종 승인]이나 [반려]를 눌러 주세요. ({final_names()} 중 한 분만 누르면 돼요)")
    if role == 'executor' and st == APPROVED and row.get('form') in TASK_FORMS:
        return (f"🔔 {initial(row.get('requester_name'))}님이 {title(row)}을 올렸어요. "
                '처리하고 카드에서 [처리 완료]를 눌러 주세요.')
    if role == 'executor' and st == APPROVED:
        return f'🔔 {title(row)} 승인이 끝났어요. {w}하고 카드에서 [{w} 완료]를 눌러 주세요.'
    if role == 'notice' and st == APPROVED:   # N통장 — 경영지원(SB)은 알림만(11-06)
        names = '·'.join(_N_LABELS.get(u, 'N통장 담당') for u in n_executor_ids())
        return f'ℹ️ {title(row)}은 N통장 건이라 {names}님에게 {w} 요청이 갔어요. (알림만 — 처리는 안 하셔도 돼요)'
    if role == 'requester' and st == REVISE:
        rv = last_revise(steps(row['id'])) or {}
        return (f"✏️ {initial(rv.get('actor_name'))}님이 {title(row)} 수정을 요청했어요.\n"
                f"*{rv.get('reason') or '-'}*\n[✏️ 수정하기]를 누르면 올린 내용이 채워진 창이 열려요. 고쳐서 다시 올려 주세요.")
    if role == 'requester' and st in OPEN_STATUSES:
        return f'{title(row)} 요청을 다시 올렸어요. 진행되면 이 메시지가 바뀝니다.'
    emoji, label = status_label(row)
    return f'{emoji} {title(row)} — {label}'


def what_of(row: dict) -> str:
    """한 줄 요약의 '무엇' — 양식은 양식 요약, 비품은 품명, 자금집행은 거래처."""
    form = row.get('form')
    if form in forms.FORMS:
        return forms.summary(form, extra(row).get('f') or {})
    return (row.get('content') if form == FORM_SUPPLY else row.get('partner')) or '-'


# ── 내 요청 (11-03) ───────────────────────────────────────────────────
MY_RECENT_DAYS = 14


def my_requests(uid: str, open_limit: int = 20, closed_limit: int = 5) -> Tuple[List[dict], List[dict]]:
    """(진행 중, 최근 끝난 것) — 메뉴 [📂 내 요청]. 끝난 것은 최근 2주 안에서 몇 건만."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        q = 'SELECT * FROM approvals WHERE requester_id=? AND status {} ORDER BY id DESC LIMIT ?'
        ph = ','.join('?' * len(OPEN_STATUSES))
        opened = [dict(r) for r in conn.execute(
            q.format(f'IN ({ph})'), (uid, *OPEN_STATUSES, open_limit)).fetchall()]
        since = (datetime.now() - timedelta(days=MY_RECENT_DAYS)).strftime('%Y-%m-%d %H:%M:%S')
        closed = [dict(r) for r in conn.execute(
            f'SELECT * FROM approvals WHERE requester_id=? AND status NOT IN ({ph}) AND updated_at>=? '
            'ORDER BY updated_at DESC LIMIT ?', (uid, *OPEN_STATUSES, since, closed_limit)).fetchall()]
    return opened, closed


def my_request_line(row: dict, link: str = '') -> str:
    emoji, label = status_label(row)
    amt = int(row.get('amount') or 0)
    s = f"{emoji} *{title(row)}* · {what_of(row)}" + (f' · {amt:,}원' if amt else '') + f' · {label}'
    return s + (f'  <{link}|카드 보기>' if link else '')


def my_requests_text(opened: List[dict], closed: List[dict], links: Dict[int, str]) -> str:
    lines = ['*📂 내 요청*']
    if opened:
        lines.append(f'*진행 중 ({len(opened)}건)*')
        lines += [f'• {my_request_line(r, links.get(r["id"], ""))}' for r in opened]
    else:
        lines.append('진행 중인 요청이 없어요.')
    if closed:
        lines.append(f'*최근 {MY_RECENT_DAYS}일 안에 끝난 요청*')
        lines += [f'• {my_request_line(r, links.get(r["id"], ""))}' for r in closed]
    return '\n'.join(lines)


def dm_blocks(row: dict, role: str, project_line: str, card_link: str = '') -> List[dict]:
    """DM = 알림 + 한 줄 요약 + 카드 링크. 버튼은 채널 카드에만 둔다(누르는 곳 한 군데)."""
    form = row.get('form')
    amt = int(row.get('amount') or 0)
    summary = f"{project_line} · {what_of(row)}" + (f' · *{amt:,}원*' if amt else '')
    if form == forms.FORM_EXPENSE and role == 'executor':   # 개인 계좌는 카드엔 안 보이고 경영지원 DM 에만
        f = extra(row).get('f') or {}
        summary += f"\n이체 계좌: {f.get('bank') or '-'} {f.get('account') or '-'} ({row.get('requester_name') or '-'})"
    link = (f'<{card_link}|👉 결재 카드 열기>' if card_link
            else '#전자_결재 채널의 카드에서 처리해 주세요.')
    text = f'{dm_intro(row, role)}\n\n{summary}\n{step_line(row, steps(row["id"]))}'
    # 수정하기는 DM 에서도 바로(채널 카드까지 안 가도)
    els = [_btn('✏️ 수정하기', A_EDIT, row['id'], 'primary')] if role == 'requester' and row.get('status') == REVISE else None
    return spaced(text[:2900], els, [{'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': link}]}])


def menu_blocks() -> List[dict]:
    """#전자_결재 맨 위에 고정해 둘 메뉴."""
    return [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text':
            '*전자결재 요청*\n'
            '• 거래처에 돈을 보내는 건(자재·외주·현장 비용, 회사 재고 자재 포함) → 💸 자금집행 요청\n'
            '• 경영지원이 대신 주문해 주는 사무용 비품 → 🛒 비품 신청\n'
            '• 계약서·보증보험·카드결제·폐기물 처리·법인서류·개인경비·법인카드 지출결의 → 📋 다른 결재 요청\n'
            '• 재직·경력증명서 → 📜 증명서 신청 (본인과 경영지원만 봐요)\n'
            '• 내가 올린 요청이 어디까지 갔는지 → 📂 내 요청\n'
            '현장 계산서는 #영업_관리 계산서 글의 [📍 현장 지정]에서 바로 올릴 수 있어요.'}},
        {'type': 'actions', 'elements': [
            {'type': 'button', 'action_id': A_NEW,   # 초록(primary) 없이 다른 버튼과 같게(11-05)
             'text': {'type': 'plain_text', 'text': '💸 자금집행 요청', 'emoji': True}},
            {'type': 'button', 'action_id': A_NEW_SUPPLY,
             'text': {'type': 'plain_text', 'text': '🛒 비품 신청', 'emoji': True}},
            {'type': 'static_select', 'action_id': A_FORM_PICK, 'options': forms.menu_options(),
             'placeholder': {'type': 'plain_text', 'text': '📋 다른 결재 요청', 'emoji': True}},
            {'type': 'button', 'action_id': 'cert_open',
             'text': {'type': 'plain_text', 'text': '📜 증명서 신청', 'emoji': True}},
            {'type': 'button', 'action_id': A_MY_REQUESTS,
             'text': {'type': 'plain_text', 'text': '📂 내 요청', 'emoji': True}},
        ]},
    ]


def form_open_action_id(form_key: str) -> str:
    """양식 버튼 action_id — 한 actions 블록 안에서 겹치면 안 돼 양식 key 를 붙인다(핸들러는 앞부분으로 받는다)."""
    return f'{A_FORM_OPEN}_{form_key}'


def _menu_btn(action_id: str, text: str, value: str = '') -> dict:
    """메뉴 버튼 — 모두 같은 모양(한 버튼만 초록이면 그것만 눌러야 하는 것처럼 보여 뺐다, 11-05)."""
    b = {'type': 'button', 'action_id': action_id, 'text': {'type': 'plain_text', 'text': text, 'emoji': True}}
    if value:
        b['value'] = value
    return b


def _form_btns(*keys: str) -> List[dict]:
    out = []
    for k in keys:
        f = forms.get(k)
        if f and f.menu:
            out.append(_menu_btn(form_open_action_id(k), f'{f.emoji} {f.label}', value=k))
    return out


def hub_blocks() -> List[dict]:
    """창(워크플로·바로가기)·홈 탭용 메뉴 — 양식을 고르기 칸 대신 버튼으로 다 보여 준다(11-05).
    채널 고정 메시지(menu_blocks)는 폭이 좁아 고르기 칸을 그대로 둔다. 휴가는 별도 휴가 봇(slack_leave)."""
    return [
        # 한 묶음에 버튼이 6개를 넘으면 슬랙이 [더 보기]로 접는다 — 양식이 7개가 돼 세 묶음으로 나눔(11-09 KiKO)
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': '*지급·구매 요청*'}},
        {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
            '• 자금집행: 거래처 대금 송금 (자재·외주·현장 비용, 회사 재고 자재 포함)\n'
            '• 비품 신청: 사무용 비품 구매 (경영지원 주문)\n'
            '• 개인경비: 본인 돈으로 먼저 쓴 비용 돌려받기'}]},
        {'type': 'actions', 'elements': [
            _menu_btn(A_NEW, '💸 자금집행 요청'),
            _menu_btn(A_NEW_SUPPLY, '🛒 비품 신청'),
            *_form_btns(forms.FORM_EXPENSE)]},
        {'type': 'divider'},
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': '*공사 관련*'}},
        {'type': 'actions', 'elements': _form_btns(forms.FORM_CONTRACT, forms.FORM_INSURANCE, forms.FORM_CARDPAY,
                                                   forms.FORM_WASTE)},
        {'type': 'divider'},
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': '*회사 서류·정산*'}},
        {'type': 'actions', 'elements': _form_btns(forms.FORM_CORPDOCS, forms.FORM_CARDEXP)},
        {'type': 'divider'},
        # 증명서(11-06) — 플렉스 「문서·증명서」 대체. 신청자·경영지원 DM 만(채널 카드 없음). 휴직·사직은 메뉴에 두지 않는다
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': '*인사*'}},
        {'type': 'actions', 'elements': [_menu_btn('cert_open', '📜 증명서 신청 (재직·경력)')]},
        {'type': 'divider'},
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': '*결재 현황*'}},
        {'type': 'actions', 'elements': [_menu_btn(A_MY_REQUESTS, '📂 내 요청')]},
        {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
            '현장 계산서는 #영업_관리 계산서 글의 [📍 현장 지정]에서 바로 올릴 수 있어요.'}]},
    ]


def hub_view() -> dict:
    """바로가기 「전자결재」·워크플로 버튼 창. 고르면 그 입력 창이 위에 쌓인다([취소] → 이 창)."""
    return {
        'type': 'modal',
        'title': {'type': 'plain_text', 'text': '전자결재'},
        'close': {'type': 'plain_text', 'text': '닫기'}, 'notify_on_close': True,
        'blocks': hub_blocks(),
    }


def home_view() -> dict:
    """앱 홈 탭(왼쪽 앱 목록 → 전자결재) — 창과 같은 버튼 메뉴(11-05). `/` 를 몰라도 찾을 수 있게."""
    return {'type': 'home', 'blocks': hub_blocks()}


def my_requests_view(text: str) -> dict:
    """바로가기 창 안에서 보는 「내 요청」 — section 글자 한도(3,000자) 안으로 줄 단위로 나눈다."""
    chunks, cur = [], ''
    for line in text.split('\n'):
        if cur and len(cur) + len(line) + 1 > 2900:
            chunks.append(cur)
            cur = ''
        cur = f'{cur}\n{line}' if cur else line
    if cur:
        chunks.append(cur)
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': c}} for c in chunks]
    blocks.append({'type': 'actions', 'elements': [
        {'type': 'button', 'action_id': A_HUB_BACK, 'text': {'type': 'plain_text', 'text': '← 처음으로'}}]})
    return {
        'type': 'modal',
        'title': {'type': 'plain_text', 'text': '내 요청'},
        'close': {'type': 'plain_text', 'text': '닫기'}, 'notify_on_close': True,
        'blocks': blocks,
    }


# ── 새 요청 모달 ──────────────────────────────────────────────────────
def _opts(pairs) -> List[dict]:
    return [{'text': {'type': 'plain_text', 'text': label}, 'value': v} for v, label in pairs]


def grn_block(initial: str = '') -> dict:
    """사업자 구분 — 어느 사업자 통장에서 돈이 나가는지 정하는 칸이라 필수(11-01)."""
    opts = _opts(GRN_TYPES)
    el = {'type': 'static_select', 'action_id': 'grn', 'options': opts,
          'placeholder': {'type': 'plain_text', 'text': '글로벌 / 글로벌그룹 / 사업자 없음'}}
    pick = next((o for o in opts if o['value'] == initial), None)
    if pick:
        el['initial_option'] = pick
    return {'type': 'input', 'block_id': 'grn', 'label': {'type': 'plain_text', 'text': '사업자 구분'},
            'hint': {'type': 'plain_text', 'text': '어느 사업자 통장에서 돈이 나가는지 골라 주세요.'},
            'element': el}


PAYEE_SAVED_NOTE = '💡 예전에 이 거래처로 보낸 계좌를 채웠어요. 계좌가 바뀌었을 수 있으니 한 번 더 확인해 주세요.'


def payee_saved_note() -> dict:
    return {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': PAYEE_SAVED_NOTE}]}


def transfer_input_blocks(prefill: Optional[dict] = None, urgency: str = '일반', sfx: str = '') -> List[dict]:
    """예금주·은행·계좌·집행 시기 입력 (새 요청 모달·현장 지정 모달 공용).

    sfx: 거래처를 바꿔 계좌를 다시 채울 때 block_id 를 바꾸려고(같은 block_id 면 슬랙이 예전 입력을 유지해
    initial_value 가 안 먹는다). 읽을 때는 field_value() 로 접미사 무관하게."""
    p = prefill or {}

    def _text(block_id, label, placeholder, value=''):
        el = {'type': 'plain_text_input', 'action_id': block_id, 'max_length': 100,
              'placeholder': {'type': 'plain_text', 'text': placeholder}}
        if value:
            el['initial_value'] = str(value)[:100]
        return {'type': 'input', 'block_id': block_id + sfx, 'label': {'type': 'plain_text', 'text': label},
                'element': el}

    urg_opts = _opts(URGENCY_TYPES)
    bank_opts = [{'text': {'type': 'plain_text', 'text': b}, 'value': b} for b in forms.BANKS]
    bank_el = {'type': 'static_select', 'action_id': 'bank', 'options': bank_opts,
               'placeholder': {'type': 'plain_text', 'text': '은행 고르기'}}
    bank = forms.normalize_bank(p.get('bank', ''))   # 예전 직접 입력값('국민', '신한은핸' …) → 항목
    if bank:
        bank_el['initial_option'] = next(o for o in bank_opts if o['value'] == bank)
    return [
        _text('payee_name', '예금주', '예: 윤정민', p.get('payee_name', '')),
        {'type': 'input', 'block_id': 'bank' + sfx, 'label': {'type': 'plain_text', 'text': '은행'},
         'hint': {'type': 'plain_text', 'text': forms.BANK_HINT}, 'element': bank_el},
        _text('account', '계좌번호', '예: 110-511-825643', p.get('account', '')),
        {'type': 'input', 'block_id': 'urgency', 'label': {'type': 'plain_text', 'text': '집행 시기'},
         'element': {'type': 'radio_buttons', 'action_id': 'urgency', 'options': urg_opts,
                     'initial_option': next(o for o in urg_opts if o['value'] == urgency)}},
    ]


def field_block_id(values: dict, base: str) -> str:
    """state 에 실제로 있는 block_id — base 그대로 또는 base_N(계좌 다시 채운 칸)."""
    for k in values:
        if k == base or (k.startswith(base + '_') and k[len(base) + 1:].isdigit()):
            return k
    return base


def field_value(values: dict, base: str) -> str:
    bid = field_block_id(values, base)
    el = (values.get(bid) or {}).get(base) or {}
    return (el.get('value') or (el.get('selected_option') or {}).get('value') or '').strip()   # 글 칸·고르기 칸


PARTNER_NEW = 'new:'


def known_partners() -> List[str]:
    """거래처 목록 — 실제로 이체한 곳(최근 순) + 샛별님 계산서 글에 나온 곳."""
    names, seen = [], set()

    def _add(n):
        k = partner_key(n)
        if k and k not in seen:
            seen.add(k)
            names.append(n)
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        for r in conn.execute("SELECT partner FROM payee_accounts WHERE partner NOT LIKE '@%' "
                              'ORDER BY last_used DESC').fetchall():
            _add(r['partner'])
    try:
        from dashboard.services.purchase_invoice_assign import recent_partners
        for n in recent_partners():
            _add(n)
    except Exception as exc:
        logger.debug(f'[전자결재] 계산서 거래처 목록 실패(무시): {exc}')
    return names


PARTNER_RECENT_N = 15


def _partner_opt(n: str) -> dict:
    return {'text': {'type': 'plain_text', 'text': n[:75]}, 'value': n[:150]}


def partner_options(query: str) -> dict:
    """상호 검색 — ack(**결과) 로 넘긴다(11-01 정렬 통일).

    빈 클릭: '최근 거래한 곳'(최근 순 15곳) + '가나다순'(나머지) 두 묶음.
    입력: 이름이 그 글자로 시작하는 곳 먼저, 나머지는 가나다순. 같은 곳이 없으면 '새 거래처: …' 맨 아래.
    """
    q = (query or '').strip()
    qk = partner_key(q)
    names = known_partners()
    if not qk:
        recent = names[:PARTNER_RECENT_N]
        rest = sorted(names[PARTNER_RECENT_N:], key=partner_key)
        groups = [{'label': {'type': 'plain_text', 'text': '최근 거래한 곳'},
                   'options': [_partner_opt(n) for n in recent]}]
        if rest:
            label = '가나다순' if len(rest) <= 100 else '가나다순 (더 찾으려면 상호를 입력하세요)'
            groups.append({'label': {'type': 'plain_text', 'text': label},
                           'options': [_partner_opt(n) for n in rest[:100]]})
        return {'option_groups': groups} if recent else {'options': []}
    hits = sorted((n for n in names if qk in partner_key(n)),
                  key=lambda n: (not partner_key(n).startswith(qk), partner_key(n)))
    opts = [_partner_opt(n) for n in hits[:99]]
    if not any(partner_key(n) == qk for n in hits):
        # 아는 거래처가 걸리면 그게 먼저(실수로 새로 만들지 않게), '새 거래처'는 맨 아래
        opts.append({'text': {'type': 'plain_text', 'text': f'새 거래처: {q}'[:75]},
                     'value': (PARTNER_NEW + q)[:150]})
    return {'options': opts}


def partner_from_value(value: str) -> str:
    v = value or ''
    return v[len(PARTNER_NEW):] if v.startswith(PARTNER_NEW) else v


INVOICE_NONE = 'none'


def invoice_option(inv: Optional[dict]) -> dict:
    """계산서 고르기 선택지 — inv={'ch','ts','parsed'} / None=계산서 없음."""
    if not inv:
        return {'text': {'type': 'plain_text', 'text': '계산서 없음 (선입금)'}, 'value': INVOICE_NONE}
    from dashboard.services.purchase_invoice_assign import post_summary
    p = inv['parsed']
    value = json.dumps({'ch': inv['ch'], 'ts': inv['ts'], 'd': p.get('date_md', ''), 'g': p.get('grn', ''),
                        'a': int(p.get('amount') or 0), 'p': (p.get('partner') or '')[:40]}, ensure_ascii=False)
    return {'text': {'type': 'plain_text', 'text': post_summary(p)[:75]}, 'value': value}


def invoice_from_value(value: str) -> Optional[dict]:
    if not value or value == INVOICE_NONE:
        return None
    v = json.loads(value)
    return {'ch': v['ch'], 'ts': v['ts'], 'parsed': {'date_md': v.get('d', ''), 'grn': v.get('g', ''),
                                                    'amount': int(v.get('a') or 0), 'partner': v.get('p', '')}}


def _set_initial(blocks: List[dict], block_id: str, value, label: str = '') -> None:
    """수정 창 — 그 칸에 올린 값을 처음 값으로(글 칸·고르기 칸·현장 검색 칸)."""
    if value in (None, ''):
        return
    for b in blocks:
        if b.get('block_id') != block_id:
            continue
        el = b['element']
        if el['type'] == 'plain_text_input':
            el['initial_value'] = str(value)[:el.get('max_length') or 3000]   # 한도보다 길면 슬랙이 창을 거부한다
        elif el['type'] == 'static_select':
            opt = next((o for o in el['options'] if o['value'] == str(value)), None)
            if opt:
                el['initial_option'] = opt
        elif el['type'] == 'external_select':
            el['initial_option'] = {'text': {'type': 'plain_text', 'text': (label or str(value))[:75]},
                                    'value': str(value)[:150]}


def new_request_view(meta: dict, fixed_project: Optional[str] = None, fixed_label: str = '',
                     prefill: Optional[dict] = None) -> dict:
    """자금집행 모달(11-01: 메뉴·바로가기·공사 카드 어디서 열어도 같은 창).

    맨 위 「계산서」에서 #영업_관리에 올라온 현장 미지정 계산서를 고르면 거래처·금액·G/R 은 그 글 값으로
    정해지고(meta['inv']), 제출하면 그 글의 현장 지정까지 함께 끝난다. '계산서 없음'이면 직접 입력(선지급).
    fixed_project 면 현장 고르기 없이 고정(공사확정 카드에서)."""
    from dashboard.services.purchase_invoice_assign import COST_CATEGORIES
    inv = meta.get('inv')
    inv_el = {'type': 'external_select', 'action_id': A_INVOICE, 'min_query_length': 0,
              'placeholder': {'type': 'plain_text', 'text': '#영업_관리에 올라온 계산서 고르기'}}
    if inv:
        inv_el['initial_option'] = invoice_option(inv)
    elif meta.get('no_inv'):
        inv_el['initial_option'] = invoice_option(None)
    edit = bool(meta.get('edit'))
    if edit:
        # 수정 창에서는 계산서 연결을 바꾸지 않는다 — 바꾸면 금액은 새 계산서, 연결·현장 지정은 예전 글로
        # 어긋나고 새 계산서로 결재가 또 올라갈 수 있다(11-05 점검). 계산서가 다르면 취소하고 새로 올린다.
        blocks: List[dict] = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': (
            f"*계산서* {invoice_option(inv)['text']['text']} (바꿀 수 없어요)" if inv else
            '*계산서* 없음 (선입금) — 계산서를 바꾸려면 이 요청을 취소하고 새로 올려 주세요.')}}]
    else:
        blocks = [{
            'type': 'input', 'block_id': 'invoice', 'optional': True, 'dispatch_action': True,
            'label': {'type': 'plain_text', 'text': '계산서'}, 'element': inv_el,
            'hint': {'type': 'plain_text', 'text': '계산서 선택 시 거래처 · 금액 정보 자동 채움. '
                                                   '계산서가 없으면 비워 두세요(선입금).'}}]
    if inv and edit:
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                       '📄 거래처·금액·현장은 계산서 글에 지정된 그대로예요. 계산서 파일은 처음 올릴 때 카드 스레드에 붙었어요.'}]})
    elif inv:
        files = meta.get('inv_files') or []
        link = f" <{meta['inv_link']}|계산서 글 보기>" if meta.get('inv_link') else ''
        if files:
            photo = f"📎 계산서 파일({', '.join(files)})은 요청하면 결재 카드 스레드에 자동으로 붙어요.{link}"
        else:
            photo = f'📎 계산서 글에 붙은 파일이 없어요. 계산서 파일이 있으면 아래 「첨부 파일」에 올려 주세요.{link}'
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                       '📄 거래처·금액은 계산서 글 그대로 올라가고, 그 글의 현장 지정도 함께 끝납니다.\n' + photo}]})
    if fixed_project:
        blocks.append({'type': 'section', 'text': {'type': 'mrkdwn', 'text': f'*{fixed_label or fixed_project}*'}})
    else:
        blocks.append({'type': 'input', 'block_id': 'project', 'label': {'type': 'plain_text', 'text': '현장'},
                       'element': {'type': 'external_select', 'action_id': A_PROJECT, 'min_query_length': 0,
                                   'placeholder': {'type': 'plain_text',
                                                   'text': '코드·상호·주소로 검색 (누르면 내 담당 현장이 최근 순으로 나와요)'}}})
    ev_opts = _opts(EVIDENCE_TYPES)
    blocks += [
        {'type': 'input', 'block_id': 'evidence', 'label': {'type': 'plain_text', 'text': '증빙'},
         'element': {'type': 'static_select', 'action_id': 'evidence', 'options': ev_opts,
                     'initial_option': ev_opts[0]}},
        {'type': 'input', 'block_id': 'cost', 'label': {'type': 'plain_text', 'text': '원가 분류'},
         'element': {'type': 'static_select', 'action_id': 'cost', 'options': _opts(COST_CATEGORIES),
                     'placeholder': {'type': 'plain_text', 'text': '원가 분류'}}},
        # 계산서를 골라도 필수(11-09 KiKO) — 최종 승인(JW·YG)이 '왜 나가는 돈인지' 보고 판단한다
        {'type': 'input', 'block_id': 'content',
         'label': {'type': 'plain_text', 'text': '내용'},
         'element': {'type': 'plain_text_input', 'action_id': 'content', 'max_length': 300,
                     'placeholder': {'type': 'plain_text', 'text': '예: FCU 교체 부품 구매'}},
         'hint': {'type': 'plain_text', 'text': '무엇을 왜 사는지 적어 주세요. 최종 승인할 때 이 내용을 봐요.'}},
    ]
    if not inv:   # 계산서를 고르면 거래처·금액·G/R 은 그 글 값 — 칸을 숨긴다
        p_el = {'type': 'external_select', 'action_id': A_PARTNER, 'min_query_length': 0,
                'placeholder': {'type': 'plain_text', 'text': '상호 검색 (처음이면 이름을 입력해 고르기)'}}
        if meta.get('partner'):
            p_el['initial_option'] = {'text': {'type': 'plain_text', 'text': meta['partner'][:75]},
                                      'value': meta['partner'][:150]}
        blocks += [
            {'type': 'input', 'block_id': 'partner', 'dispatch_action': True,
             'label': {'type': 'plain_text', 'text': '거래처 (상호)'}, 'element': p_el,
             # 고르기 칸이라 직접 못 쓰는 줄 알았다(11-06 KiKO) — 처음 거래처 넣는 법을 칸 아래에
             'hint': {'type': 'plain_text', 'text':
                      '처음 거래하는 곳이면 상호를 끝까지 입력한 뒤, 목록 맨 아래 「새 거래처: ○○」를 고르세요.'}},
            {'type': 'input', 'block_id': 'amount', 'label': {'type': 'plain_text', 'text': '금액 (VAT 포함, 원)'},
             'element': {'type': 'plain_text_input', 'action_id': 'amount',
                         'placeholder': {'type': 'plain_text', 'text': '예: 277200'}}},
        ]
    if not inv or not (inv.get('parsed') or {}).get('grn'):
        blocks.append(grn_block(grn_from_code(fixed_project)))
    # 이체 계좌 — 거래처(또는 계산서)를 고르면 예전에 보낸 계좌를 창에 채워 보여 주고 확인 안내(11-01)
    payee = meta.get('payee')
    if payee:
        head = '*이체 계좌*'
    elif meta.get('partner') or inv:
        head = '*이체 계좌* — 처음 보내는 거래처예요. 예금주·은행·계좌를 넣어 주세요.'
    else:
        head = '*이체 계좌* — 거래처를 고르면 예전에 보낸 계좌가 있을 때 채워집니다.'
    blocks.append({'type': 'section', 'text': {'type': 'mrkdwn', 'text': head}})
    if payee and not meta.get('edit'):   # 수정 창은 올렸던 계좌 그대로라 '예전 계좌' 안내를 안 띄운다
        blocks.append(payee_saved_note())
    pv = int(meta.get('pv') or 0)
    p = prefill or {}
    blocks += transfer_input_blocks(payee, urgency=p.get('urgency') or '일반', sfx=f'_{pv}' if pv else '')
    blocks.append({'type': 'input', 'block_id': 'memo', 'optional': True,
                   'label': {'type': 'plain_text', 'text': '비고 (특이사항)'},
                   'element': {'type': 'plain_text_input', 'action_id': 'memo', 'multiline': True,
                               'max_length': 1000}})   # 길면 카드가 잘려 결재 줄이 안 보였다(11-06 검수)
    blocks.append(forms.attach_block('세금계산서·견적서 등'))
    if p:   # 수정 창(11-05) — 올린 값을 채운다(계산서·거래처·계좌는 meta 로 이미 채움)
        _set_initial(blocks, 'project', p.get('project_code'), p.get('_project_label', ''))
        for bid, key in (('evidence', 'evidence'), ('cost', 'cost_category'), ('content', 'content'),
                         ('grn', 'grn'), ('memo', 'memo')):
            _set_initial(blocks, bid, p.get(key))
        if p.get('amount'):
            _set_initial(blocks, 'amount', str(int(p['amount'])))
    title, submit = forms.edit_texts(meta, '자금집행 요청', '요청')
    return {
        'type': 'modal', 'callback_id': CB_NEW,
        'title': {'type': 'plain_text', 'text': title},
        'submit': {'type': 'plain_text', 'text': submit},
        'close': {'type': 'plain_text', 'text': '취소'}, 'notify_on_close': True,
        # 한글 그대로(\uXXXX 면 6배) — 계산서 파일 이름·거래처·계좌 이름이 길면 3,000자 한도(11-06 검수)
        'private_metadata': json.dumps(meta, ensure_ascii=False),
        'blocks': blocks,
    }


def read_new_request(view: dict) -> Dict:
    """새 요청 모달 값 → 필드(검사 전)."""
    v = view['state']['values']

    def _sel(bid, aid=None):
        return (((v.get(bid) or {}).get(aid or bid) or {}).get('selected_option') or {}).get('value') or ''

    def _txt(bid):
        return (((v.get(bid) or {}).get(bid) or {}).get('value') or '').strip()

    meta = json.loads(view.get('private_metadata') or '{}')
    f = {
        'project_code': meta.get('project') or _sel('project', A_PROJECT),
        'evidence': _sel('evidence'), 'cost_category': _sel('cost'), 'content': _txt('content'),
        'partner': partner_from_value(_sel('partner', A_PARTNER)) or meta.get('partner', ''),
        'amount': parse_amount(_txt('amount')), 'grn': _sel('grn'),
        'payee_name': field_value(v, 'payee_name'), 'bank': field_value(v, 'bank'),
        'account': normalize_account(field_value(v, 'account')),
        'urgency': _sel('urgency') or '일반', 'memo': _txt('memo'),
    }
    inv = meta.get('inv')
    if not inv and not meta.get('edit'):
        # 고른 뒤 창 갱신이 실패하면 meta 엔 없지만 화면엔 고른 계산서가 보인다 — 화면 값으로(11-06 검수)
        sel = _sel('invoice', A_INVOICE)
        if sel and sel != INVOICE_NONE:
            inv = invoice_from_value(sel)
    if inv:   # 계산서 글 값으로 — 그 글 스레드에도 결과가 남게 origin 연결
        p = inv['parsed']
        f.update(partner=p.get('partner', ''), amount=int(p.get('amount') or 0), grn=p.get('grn') or f['grn'],
                 origin_channel=inv['ch'], origin_ts=inv['ts'], _inv_parsed=p)
    if meta.get('edit'):
        f['_edit'] = int(meta['edit'])
    f['_files'] = forms.read_attachments(view)
    if not inv:
        f['_bad'] = amount_errors({'amount': _txt('amount')})
    return f


def complete_new_request(f: Dict) -> Dict[str, str]:
    """새 요청 값 보완(사업자 구분·기억한 계좌) + 오류 {block_id: 문구}. f 를 고친다."""
    errs = {}
    if not (f.get('content') or '').strip():
        errs['content'] = '무엇을 왜 사는지 적어 주세요.'
    if not f.get('project_code'):
        errs['project'] = '현장을 골라 주세요.'
    else:
        from dashboard.services.purchase_invoice_assign import cost_project_error
        e = cost_project_error(f.get('cost_category'), f['project_code'])
        if e:
            errs['cost'] = e
    if f.get('origin_ts'):
        dup = find_by_origin(f['origin_channel'], f['origin_ts'])
        if dup and dup['status'] not in (REJECTED, CANCELLED) and dup['id'] != f.get('_edit'):
            errs['invoice'] = f'이 계산서로 올린 {title(dup)}이 이미 있어요.'
        else:
            # 계산서 없이 먼저 올린(선입금) 같은 건 — G/R·금액·거래처가 같으면 두 번 이체될 수 있다(11-06 검수)
            m = find_unlinked_match(f.get('grn') or '', int(f.get('amount') or 0), f.get('partner') or '')
            if m and m['id'] != f.get('_edit'):
                errs['invoice'] = (f"{initial(m.get('requester_name'))}님이 먼저 올린 {title(m)}"
                                   f"({int(m.get('amount') or 0):,}원)과 같은 건이에요. 새로 올리지 않아도 돼요.")
    if f.get('amount', 0) <= 0:
        errs['amount'] = '금액을 숫자로 넣어 주세요.'
    errs.update(f.pop('_bad', None) or {})
    if not f.get('grn'):
        errs['grn'] = '어느 사업자 통장에서 나가는지 골라 주세요.'
    if not (f.get('payee_name') or f.get('bank') or f.get('account')):
        saved = lookup_payee(f.get('partner', ''))
        if saved:
            f.update(payee_name=saved['payee_name'], bank=saved['bank'], account=saved['account'])
        else:
            errs['account'] = '처음 요청하는 거래처예요. 예금주·은행·계좌를 넣어 주세요.'
    if 'account' not in errs:
        errs.update(validate_transfer(f.get('payee_name'), f.get('bank'), f.get('account')))
    mark_payee_saved(f)
    return errs


def mark_payee_saved(f: Dict) -> None:
    """기억해 둔 계좌와 같은 계좌로 보내면 카드에 표시(경영지원이 이체 전에 한 번 더 보게)."""
    saved = lookup_payee(f.get('partner', ''))
    if saved and normalize_account(saved.get('account')) == normalize_account(f.get('account')):
        f['extra'] = {**(f.get('extra') or {}), 'payee_saved': True}


# ── 비품 신청 모달 (11-01) ────────────────────────────────────────────
def supply_request_view(meta: dict, prefill: Optional[dict] = None) -> dict:
    """비품 신청 — 플렉스 「비품 신청(프로젝트0)」 칸 그대로. 현장은 프로젝트0 고정. prefill = 수정 창 값."""
    def _text(block_id, label, placeholder, optional=False, multiline=False, hint=''):
        # 길이 제한(11-06 검수) — 품명이 길면 [구매 완료] 확인 창(300자)이 깨져 카드가 멈췄다. 상품 주소는 길 수 있어 넉넉히
        mx = 2000 if block_id == 'link' else (1000 if multiline else 200)
        b = {'type': 'input', 'block_id': block_id, 'optional': optional,
             'label': {'type': 'plain_text', 'text': label},
             'element': {'type': 'plain_text_input', 'action_id': block_id, 'multiline': multiline,
                         'max_length': mx}}
        if placeholder:   # 빈 placeholder 는 슬랙이 창 자체를 거부한다
            b['element']['placeholder'] = {'type': 'plain_text', 'text': placeholder}
        if hint:
            b['hint'] = {'type': 'plain_text', 'text': hint}
        return b

    qty = _text('qty', '수량', '예: 2')
    qty['element']['initial_value'] = '1'
    title, submit = forms.edit_texts(meta, '비품 신청', '신청')
    blocks = [
        # 프로젝트0 안내 줄·사업자 구분은 뺐다(11-05) — 회사 재고(프로젝트0)로 기록, 어느 통장인지는 경영지원 판단
        _text('item', '품명', '예: A4 용지 박스'),
        _text('link', '상품 링크', '예: https://www.coupang.com/vp/products/…',
              hint='구매할 상품 페이지 주소를 붙여 넣어 주세요.'),
        qty,
        _text('unit_price', '단가 (VAT 포함, 원)', '예: 25000'),
        _text('amount', '총 금액 (VAT 포함, 원)', '배송비가 있으면 더해서 입력', optional=True,
              hint='비우면 수량 × 단가로 계산합니다.'),
        _text('reason', '사유', '예: 사무실 용지 소진'),
        _text('memo', '구매 요청사항', '예: 배송지, 색상·크기, 쿠폰 코드, 장바구니 링크',
              optional=True, multiline=True),
        forms.attach_block('상품 캡처·견적서 등'),
    ]
    for bid, v in (prefill or {}).items():
        _set_initial(blocks, bid, v)
    return {
        'type': 'modal', 'callback_id': CB_SUPPLY,
        'title': {'type': 'plain_text', 'text': title},
        'submit': {'type': 'plain_text', 'text': submit},
        'close': {'type': 'plain_text', 'text': '취소'}, 'notify_on_close': True,
        # 한글 그대로(\uXXXX 면 6배) — 계산서 파일 이름·거래처·계좌 이름이 길면 3,000자 한도(11-06 검수)
        'private_metadata': json.dumps(meta, ensure_ascii=False),
        'blocks': blocks,
    }


def supply_prefill(row: dict) -> dict:
    """비품 신청 수정 창 값 {block_id: 값}."""
    x = extra(row)
    qty, unit, amt = int(x.get('qty') or 0), int(x.get('unit_price') or 0), int(row.get('amount') or 0)
    return {'item': row.get('content'), 'link': x.get('link'), 'qty': x.get('qty'),
            'unit_price': x.get('unit_price'), 'reason': x.get('reason'), 'memo': row.get('memo'),
            # 총 금액은 직접 넣었던 때(배송비 등)만 — 수량×단가로 계산된 값을 채우면 수량을 고쳐도 금액이 안 바뀐다
            'amount': amt if amt and amt != qty * unit else None}


def read_supply_request(view: dict) -> Dict:
    """비품 신청 모달 값 → 필드(검사 전)."""
    v = view['state']['values']

    def _txt(bid):
        return (((v.get(bid) or {}).get(bid) or {}).get('value') or '').strip()

    meta = json.loads(view.get('private_metadata') or '{}')
    f = {
        'content': _txt('item'), 'link': _txt('link'), 'qty': parse_amount(_txt('qty')),
        'unit_price': parse_amount(_txt('unit_price')), 'amount': parse_amount(_txt('amount')),
        'reason': _txt('reason'), 'memo': _txt('memo'), '_files': forms.read_attachments(view),
        '_bad': amount_errors({k: _txt(k) for k in ('qty', 'unit_price', 'amount')}),
    }
    if meta.get('edit'):
        f['_edit'] = int(meta['edit'])
    return f


def complete_supply_request(f: Dict) -> Dict[str, str]:
    """비품 신청 값 검사 + 총 금액 채움 + 장부 필드로 정리. 오류 {block_id: 문구}. f 를 고친다."""
    from dashboard.services.purchase_invoice_assign import STOCK_CODE
    errs = {}
    if not f.get('content'):
        errs['item'] = '품명을 넣어 주세요.'
    if not re.match(r'https?://\S+', f.get('link') or ''):
        errs['link'] = '상품 페이지 주소(http로 시작)를 넣어 주세요.'
    if f.get('qty', 0) <= 0:
        errs['qty'] = '수량을 숫자로 넣어 주세요.'
    if f.get('unit_price', 0) <= 0:
        errs['unit_price'] = '단가를 숫자로 넣어 주세요.'
    if not f.get('reason'):
        errs['reason'] = '사유를 적어 주세요.'
    errs.update(f.pop('_bad', None) or {})
    if not f.get('amount') and 'qty' not in errs and 'unit_price' not in errs:
        f['amount'] = f['qty'] * f['unit_price']
    f.update(project_code=STOCK_CODE, urgency='일반',
             extra={'link': f.get('link'), 'qty': f.get('qty'), 'unit_price': f.get('unit_price'),
                    'reason': f.get('reason')})
    return errs
