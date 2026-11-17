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
    REVISE: ('🟠', '보완 요청'),
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
A_BRANCH = 'eap_branch'               # 계산서로 거래처가 정해졌을 때 「지점 계좌」 static_select(11-16)
A_COST = 'eap_cost'                   # 자금집행 「원가 분류」 — '법인 간 이체'를 고르면 창이 바뀐다(11-16)
A_MOVE_PICK = 'eap_move_pick'         # 법인 간 이체 — 옮길 입금(대기 중인 자금 이동 요청) 고르기
A_MOVE_OPEN = 'eap_move_open'         # 매니저 DM [📝 결재 올리기] — 채워진 법인 간 이체 창 열기
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


def weekend_test_ids() -> tuple:
    """주말 시험 모드(EAPPROVAL_EXECUTOR_DM_IDS) — 켜져 있으면 이체 요청 DM 이 통장과 상관없이 이 사람에게만 간다."""
    return _env_ids('EAPPROVAL_EXECUTOR_DM_IDS')


TEST_MODE_NOTE = ('⚠️ 시험 모드 — 이체 요청 DM 이 정해진 시험자에게만 갑니다(경영지원 DM 없음). '
                  '평일 운영 전에 .env 의 EAPPROVAL_EXECUTOR_DM_IDS 를 지우고 재시작하세요.')


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
    # 11-15 거래처가 확인해 준 계좌 목록 — 하이엠솔루텍처럼 센터마다 계좌가 따로인 거래처.
    # 목록에 있으면 '예전 계좌와 달라요' ⚠️ 를 띄우지 않고 어느 센터인지 붙이고, 없으면 ⚠️(account=숫자만)
    '''CREATE TABLE IF NOT EXISTS payee_known_accounts (
        partner_key TEXT NOT NULL,
        partner TEXT,
        bank TEXT,
        account TEXT NOT NULL,
        label TEXT,
        added_at TIMESTAMP,
        PRIMARY KEY (partner_key, account)
    )''',
    # 11-11 제출 기록 — 창 제출(ack) 뒤 백그라운드에서 결재를 만드는데, 그 사이 서버가 꺼지면 요청이 사라지고
    # 요청자도 몰랐다. 제출 내용을 먼저 남기고, 끝나지 않은 것은 재시작 뒤 이어서 처리한다(recover_jobs)
    '''CREATE TABLE IF NOT EXISTS approval_jobs (
        id TEXT PRIMARY KEY,
        kind TEXT NOT NULL,
        user_id TEXT,
        payload TEXT NOT NULL,
        status TEXT NOT NULL,
        attempts INTEGER NOT NULL DEFAULT 0,
        approval_id INTEGER,
        error TEXT,
        created_at TIMESTAMP NOT NULL,
        updated_at TIMESTAMP NOT NULL
    )''',
    'CREATE INDEX IF NOT EXISTS idx_approval_jobs_status ON approval_jobs(status)',
    # 11-16 카드를 마지막으로 그린 상태 — '상태|수정 회차'. 슬랙 갱신이 실패하면 장부와 카드가 어긋난 채
    # 남았다(끝난 결재가 '이체 대기'로 보임). 정기 점검이 이 값과 지금 상태를 비교해 다시 그린다
    '''CREATE TABLE IF NOT EXISTS approval_drawn (
        approval_id INTEGER PRIMARY KEY,
        sig TEXT NOT NULL,
        at TIMESTAMP NOT NULL
    )''',
    # 11-16 작은 설정값(메뉴 메시지 ts 등) — 메뉴가 지워지면 다시 올리려고
    '''CREATE TABLE IF NOT EXISTS approval_kv (
        key TEXT PRIMARY KEY,
        value TEXT
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


# ── 제출 기록(11-11) ──────────────────────────────────────────────────
JOB_PENDING, JOB_RUNNING, JOB_DONE, JOB_FAILED = 'pending', 'running', 'done', 'failed'


def job_add(kind: str, user_id: str, payload: dict) -> str:
    """제출 내용을 먼저 남긴다. 기록 id 반환(결재 extra['job'] 에도 실려 같은 제출을 두 번 만들지 않게)."""
    import uuid
    jid = 'job-' + uuid.uuid4().hex[:16]
    now = _now()
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('INSERT INTO approval_jobs (id, kind, user_id, payload, status, created_at, updated_at) '
                     'VALUES (?,?,?,?,?,?,?)',
                     (jid, kind, user_id, json.dumps(payload, ensure_ascii=False, default=str), JOB_PENDING, now, now))
        conn.commit()
    return jid


def job_claim(jid: str, stale_before: str = '') -> bool:
    """처리 시작 — 대기 중이거나(stale_before 가 있으면) 그 전부터 '처리 중'에 멈춘 것만. 동시에 둘이 못 잡는다."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        if stale_before:
            cur = conn.execute("UPDATE approval_jobs SET status=?, attempts=attempts+1, updated_at=? WHERE id=? "
                               "AND (status=? OR (status=? AND updated_at < ?))",
                               (JOB_RUNNING, _now(), jid, JOB_PENDING, JOB_RUNNING, stale_before))
        else:
            cur = conn.execute('UPDATE approval_jobs SET status=?, attempts=attempts+1, updated_at=? WHERE id=? AND status=?',
                               (JOB_RUNNING, _now(), jid, JOB_PENDING))
        conn.commit()
        return cur.rowcount == 1


def job_finish(jid: str, ok: bool, approval_id: Optional[int] = None, error: str = '') -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('UPDATE approval_jobs SET status=?, approval_id=?, error=?, updated_at=? WHERE id=?',
                     (JOB_DONE if ok else JOB_FAILED, approval_id, (error or '')[:500] or None, _now(), jid))
        conn.commit()


def job_get(jid: str) -> Optional[dict]:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT * FROM approval_jobs WHERE id=?', (jid,)).fetchone()
        return dict(r) if r else None


def jobs_unfinished(older_than: str) -> List[dict]:
    """끝나지 않은 제출(대기·처리 중) 중 older_than 보다 오래된 것."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM approval_jobs WHERE status IN (?,?) AND updated_at < ? ORDER BY created_at',
            (JOB_PENDING, JOB_RUNNING, older_than)).fetchall()]


def find_by_job(jid: str) -> Optional[dict]:
    """그 제출로 만든 결재(extra['job'])."""
    if not jid:
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT * FROM approvals WHERE extra LIKE ? ORDER BY id DESC LIMIT 1',
                         (f'%"job": "{jid}"%',)).fetchone()
        return dict(r) if r else None


def open_rows() -> List[dict]:
    """진행 중인 결재(승인·최종 승인·이체 대기·수정 요청) — 리마인드용(11-11)."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(f'SELECT * FROM approvals WHERE status IN ({",".join("?" * 4)}) ORDER BY id',
                            (PENDING, FINAL, APPROVED, REVISE)).fetchall()
        return [dict(r) for r in rows]


def find_by_origin(channel: str, ts: str) -> Optional[dict]:
    """계산서 글에서 시작한 결재(가장 최근 것)."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute(
            'SELECT * FROM approvals WHERE origin_channel=? AND origin_ts=? ORDER BY id DESC LIMIT 1',
            (channel, ts)).fetchone()
        return dict(r) if r else None


class DuplicateOriginError(RuntimeError):
    """같은 계산서 글로 진행·완료된 결재가 이미 있음 — 두 번 이체되지 않게 생성 거절(11-16)."""

    def __init__(self, row: dict):
        self.row = row
        super().__init__(f"이 계산서로 올린 {title(row)}이 이미 있어요.")


def find_live_by_origin(channel: str, ts: str, exclude_id: Optional[int] = None) -> Optional[dict]:
    """계산서 글로 올린 결재 중 반려·취소가 아닌 것(가장 최근). find_by_origin 은 가장 최근 한 건만 봐서,
    그게 반려·취소면 그보다 앞서 이체까지 끝난 결재를 놓쳤다(2026-11-16 상태 머신 점검)."""
    if not (channel and ts):
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            'SELECT * FROM approvals WHERE origin_channel=? AND origin_ts=? AND status NOT IN (?,?) ORDER BY id DESC',
            (channel, ts, REJECTED, CANCELLED)).fetchall()
    for r in rows:
        r = dict(r)
        if exclude_id is None or r['id'] != int(exclude_id):
            return r
    return None


def drawn_sig(row: dict) -> str:
    """카드가 보여 줘야 할 상태 표시 — 상태와 수정 회차."""
    return f"{row.get('status')}|{revision(row['id'])}"


def mark_drawn(aid: int, sig: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('INSERT INTO approval_drawn (approval_id, sig, at) VALUES (?,?,?) '
                     'ON CONFLICT(approval_id) DO UPDATE SET sig=excluded.sig, at=excluded.at',
                     (int(aid), sig, _now()))
        conn.commit()


def drawn(aid: int) -> str:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT sig FROM approval_drawn WHERE approval_id=?', (int(aid),)).fetchone()
        return (r[0] if r else '') or ''


def rows_changed_since(since: str) -> List[dict]:
    """since('YYYY-mm-dd HH:MM:SS') 뒤에 바뀐 결재(끝난 것 포함) — 카드 다시 그리기 점검용."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM approvals WHERE updated_at >= ? ORDER BY id', (since,)).fetchall()]


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


def find_unlinked_match(grn: str, amount: int, partner: str, code: str = '') -> Optional[dict]:
    """계산서 글과 G/R·금액·거래처가 같은데 아직 계산서 글에 안 묶인 자금집행(반려·취소 제외).

    딱 한 건일 때만 돌려준다(여러 건이면 사람이 고른다). 단 후보가 여럿이면 계산서 글에 적힌 공사 코드로
    좁힌다 — 크레인 양중 33만 원처럼 흔한 금액은 90일 안에 쉽게 겹쳐, 짝을 못 찾으면 제출 창 경고도
    [📍 현장 지정]의 자동 연결도 둘 다 안 됐다 (2026-11-17 #45 서울크레인)."""
    key = partner_key(partner)
    if not key or not amount:
        return None
    # 이체 완료된 선입금은 최근 것만 — 기간 없이 두면 반년 전 선입금이 같은 금액의 새 계산서를 목록에서 계속 뺐다
    #   (11-14 보안 점검 4차). LIMIT 은 거래처 필터 뒤에서 걸리게 뺐다(흔한 금액이면 후보를 놓쳤다).
    since = (datetime.now() - timedelta(days=DUP_WINDOW_DAYS)).strftime('%Y-%m-%d %H:%M:%S')
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            '''SELECT * FROM approvals WHERE form=? AND origin_ts IS NULL AND amount=? AND grn=?
               AND status IN (?,?,?,?,?) AND (status != ? OR updated_at >= ?) ORDER BY id DESC''',
            # 수정 요청 중인 결재도 후보 — 빼면 같은 건으로 결재를 또 올려 두 번 이체할 수 있다(11-05)
            (FORM_FUND, int(amount), grn, PENDING, FINAL, APPROVED, REVISE, DONE, DONE, since)).fetchall()
    hits = [dict(r) for r in rows if partner_key(r['partner']) == key]
    if len(hits) > 1 and (code or '').strip():
        same = [h for h in hits if (h.get('project_code') or '').strip().upper() == code.strip().upper()]
        if len(same) == 1:
            return same[0]
    return hits[0] if len(hits) == 1 else None


DUP_WINDOW_DAYS = 90
OLD_PREPAID_DAYS = 365


def find_old_prepaid(grn: str, amount: int, partner: str) -> Optional[dict]:
    """계산서 없이 먼저 보낸(선입금) 같은 금액·거래처 결재 중 이체 완료 90일~1년 지난 것 — 가장 최근 한 건.

    find_unlinked_match 는 90일 지난 선입금을 보지 않는다(새 계산서를 가리지 않게). 늦게 온 계산서면 두 번 이체될 수
    있어, 제출 창에서 한 번 묻고 카드에도 ⚠️ 로 남긴다(11-15). 사업자(G/R)는 보지 않는다 — 사업자를 잘못 고른 같은
    건도 잡게(최근 중복 검사와 같은 기준, 11-15 5차 점검). grn 인자는 호출 호환용."""
    key = partner_key(partner)
    if not key or not amount:
        return None
    now = datetime.now()
    recent = (now - timedelta(days=DUP_WINDOW_DAYS)).strftime('%Y-%m-%d %H:%M:%S')
    oldest = (now - timedelta(days=OLD_PREPAID_DAYS)).strftime('%Y-%m-%d %H:%M:%S')
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            '''SELECT * FROM approvals WHERE form=? AND origin_ts IS NULL AND amount=? AND status=?
               AND updated_at < ? AND updated_at >= ? ORDER BY id DESC''',
            (FORM_FUND, int(amount), DONE, recent, oldest)).fetchall()
    hits = [dict(r) for r in rows if partner_key(r['partner']) == key]
    return hits[0] if hits else None


def old_prepaid_warning(row: dict) -> str:
    when = str(row.get('updated_at') or '')[:10]
    return (f"⚠️ 거래처·금액이 같은 선입금이 있어요: {title(row)}"
            f"({initial(row.get('requester_name'))}·{when} 이체 완료, 계산서 없이 보냄). "
            "그 선입금의 계산서면 올리지 말고, 다른 건이면 [요청]을 한 번 더 눌러 주세요.")


def find_recent_fund_dups(amount, partner: str, exclude_id: Optional[int] = None,
                          days: int = DUP_WINDOW_DAYS) -> List[dict]:
    """최근 N일 같은 거래처·금액 자금집행(반려·취소 제외) — 계산서 유무·글 연결과 상관없이(11-14 보안 점검 4차).

    예전 중복 검사는 계산서를 고른 경우에만, 그것도 글에 안 묶인 선입금 한 건만 봤다 → 「계산서 없음」으로 다시 올리기·
    같은 계산서 글 재게시는 경고 없이 두 번 결재됐다. G/R 은 보지 않는다(사업자를 잘못 고른 같은 건도 잡게)."""
    key = partner_key(partner)
    try:
        amount = int(amount or 0)
    except (TypeError, ValueError):
        amount = 0
    if not key or amount <= 0:
        return []
    since = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            '''SELECT * FROM approvals WHERE form=? AND amount=? AND status IN (?,?,?,?,?) AND created_at >= ?
               ORDER BY id DESC''',
            (FORM_FUND, amount, PENDING, FINAL, APPROVED, REVISE, DONE, since)).fetchall()
    return [dict(r) for r in rows
            if partner_key(r['partner']) == key and (exclude_id is None or r['id'] != int(exclude_id))]


def dup_warning(dups: List[dict]) -> str:
    """제출 창 경고 — 같은 창에서 한 번 더 누르면 통과(정말 같은 금액을 또 보내는 경우도 있어서)."""
    shown = ', '.join(f"{title(d)}({initial(d.get('requester_name'))}·{STATUS_LABELS.get(d['status'], ('', d['status']))[1]})"
                      for d in dups[:2])
    more = f' 외 {len(dups) - 2}건' if len(dups) > 2 else ''
    return (f'⚠️ 최근 {DUP_WINDOW_DAYS}일 안에 같은 거래처·금액 결재가 있어요: {shown}{more}. '
            '같은 건이면 올리지 말고, 다른 건이면 [요청]을 한 번 더 눌러 주세요.')


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


_CREATE_LOCK = threading.Lock()   # 결재 생성(중복 확인 + 저장) — 서버는 프로세스 하나(waitress)


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
    x = dict(fields.get('extra') or {})
    if fields.get('_job'):   # 제출 기록 id — 재시작 뒤 이어서 처리할 때 같은 결재를 또 만들지 않게(11-11)
        x['job'] = fields['_job']
    # 같은 거래처·금액 결재 — 카드에 ⚠️ 로 남겨 승인자·이체 담당이 보게(11-14).
    # 확인과 저장을 한 잠금 안에서: 두 사람이 몇 초 차이로 올려도 뒤 카드에는 ⚠️ 가 붙는다(11-15)
    with _CREATE_LOCK:
        if form == FORM_FUND and fields.get('origin_ts'):
            # 창 검사(complete_new_request)를 둘 다 통과한 두 사람이 몇 초 차이로 같은 계산서를 올려도 한 건만(11-16)
            _live = find_live_by_origin(fields.get('origin_channel') or '', fields['origin_ts'])
            if _live:
                raise DuplicateOriginError(_live)
        if form == FORM_FUND:
            _dups = find_recent_fund_dups(fields.get('amount'), fields.get('partner') or '')
            if _dups:
                x['dup_of'] = [d['id'] for d in _dups[:3]]
            if fields.get('origin_ts'):   # 늦게 온 계산서 — 신청자 창에서만 묻던 것을 카드에도(5차 점검 #7)
                _old = find_old_prepaid(fields.get('grn') or '', int(fields.get('amount') or 0),
                                        fields.get('partner') or '')
                if _old:
                    x['old_prepaid'] = _old['id']
        vals[cols.index('extra')] = json.dumps(x, ensure_ascii=False)
        return _insert_new(form, status, requester_id, requester_name, cols, vals, now, auto, auto_final)


def _insert_new(form, status, requester_id, requester_name, cols, vals, now, auto, auto_final) -> dict:
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
        return user_id == row.get('requester_id') and requester_can_edit(row)
    return False


def requester_can_edit(row: dict) -> bool:
    """요청자가 직접 고치기 — 취소와 같은 선(11-11 KiKO: 올리고 나서 실수를 알아도 고칠 길이 없었다).
    수정 요청을 받았을 때 + 최종 승인 전(지출형) / 처리 전(업무요청형). 고치면 승인을 처음부터 다시 받는다."""
    return requester_can_cancel(row)


def requester_can_cancel(row: dict) -> bool:
    """요청자 취소 — 승인 뒤(이체·구매 대기)엔 경영지원이 이미 은행에서 보내는 중일 수 있어 막는다(11-06 검수).
    그때는 경영지원에게 [반려]를 부탁한다. 업무요청형은 승인 단계가 없어 처리 대기 중에도 취소 가능."""
    st = row.get('status')
    if st in (PENDING, FINAL, REVISE):   # 최종 승인 전까진 아직 돈이 안 나간다
        return True
    return st == APPROVED and row.get('form') in TASK_FORMS


# extra 중 시스템이 붙이는 표시 칸 — 수정 기록('바뀐 칸')에 넣지 않는다
_SYSTEM_EXTRA = ('f', 'payee_saved', 'payee_changed', 'payee_known', 'payee_unlisted', 'dup_of', 'old_prepaid', 'inv_by', 'inv_by_other', 'job',
                 'fund_move')

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
        if not old or not requester_can_edit(old):   # 수정 요청 + 승인 전 요청자 직접 고치기(11-11)
            return None
        if actor_id != old.get('requester_id'):   # 창 열기에서만 막던 것을 여기서도(11-14 보안 점검 3차)
            return None
        prev = old['status']
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
        if old['form'] == FORM_FUND and ('amount' in fields or 'partner' in fields):
            # 금액·거래처를 고치면 같은 거래처·금액 결재를 다시 찾는다(11-14)
            _amt = fields.get('amount') if 'amount' in fields else old.get('amount')
            _ptn = fields.get('partner') if 'partner' in fields else old.get('partner')
            _dups = find_recent_fund_dups(_amt, _ptn or '', exclude_id=aid)
            fields = {**fields, 'extra': {**extra(old), **(fields.get('extra') or {}),
                                          'dup_of': [d['id'] for d in _dups[:3]]}}
        if 'extra' in fields:
            new_x = dict(fields.get('extra') or {})   # 결재 이름도 고친 값으로(G→R 현장·지출결의 달, 11-05 점검)
            old_x = extra(old)
            ex = {**old_x, **new_x}
            # 바뀐 칸 — 양식 값(f) + 비품 칸(link·qty·unit_price·reason). 시스템이 붙이는 표시 칸은 뺀다
            pairs = [(old_x.get('f') or {}, new_x.get('f') or {})] if 'f' in new_x else []
            pairs.append(({k: v for k, v in old_x.items() if k not in _SYSTEM_EXTRA},
                          {k: v for k, v in new_x.items() if k not in _SYSTEM_EXTRA}))
            for a, b in pairs:
                for k in b:
                    if str(a.get(k) or '') != str(b.get(k) or '') and k not in changed:
                        changed[k] = (a.get(k), b.get(k))
            sets.append('extra=?')
            vals.append(json.dumps(ex, ensure_ascii=False))
        if not changed and prev != REVISE:   # 바뀐 게 없으면 받은 승인을 지우지 않는다
            conn.rollback()
            return old, {}
        cur = conn.execute(f'UPDATE approvals SET {", ".join(sets + ["status=?", "updated_at=?"])} '
                           f'WHERE id=? AND status=?', (*vals, nxt, now, aid, prev))
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


def revision(aid: int) -> int:
    """수정해 다시 올린 횟수 — 버튼 값에 실어 옛 버튼을 가려낸다(11-11)."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute("SELECT COUNT(*) FROM approval_steps WHERE approval_id=? AND step='resubmit'",
                         (int(aid),)).fetchone()
        return int(r[0] if r else 0)


def parse_button_value(value: str) -> Tuple[int, Optional[int]]:
    """'12:1' → (12, 1). 옛 버튼 '12' → (12, None)."""
    a, _, r = str(value or '').partition(':')
    return int(a), (int(r) if r.strip().isdigit() else None)


def last_revise(step_rows: List[dict]) -> Optional[dict]:
    """가장 최근 수정 요청 줄(다시 올린 뒤면 None)."""
    for s in reversed(step_rows):
        if s['step'] == 'resubmit':
            return None
        if s['step'] == 'revise':
            return s
    return None


def transition(aid: int, action: str, actor_id: str, actor_name: str, reason: str = '',
               expected_rev: Optional[int] = None, expected_status: Optional[str] = None) -> Optional[dict]:
    """상태 전환(두 사람이 동시에 눌러도 한 번만 — 이전 상태 조건부 UPDATE). 실패면 None.
    expected_rev = 누른 사람이 본 수정 회차. 주면 같은 UPDATE 조건에 넣어, 확인과 승인 사이에 요청자가 고쳐
    다시 올린 결재는 승인되지 않는다(11-14 보안 점검 3차 — 바뀐 계좌로 이체 DM 이 가던 틈)."""
    allowed, nxt, step, result = _TRANSITIONS[action]
    if expected_status is not None:
        # 권한 확인 때 본 상태 그대로일 때만 — 권한은 상태마다 다른데(요청자 취소는 승인 전만 등) 확인과 처리 사이에
        #   상태가 바뀌면 예전엔 열린 상태 전부를 허용해 통과했다(2026-11-16 상태 머신 점검)
        if expected_status not in allowed:
            return None
        allowed = (expected_status,)
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
        sql = f'UPDATE approvals SET status=?, updated_at=? WHERE id=? AND status IN ({",".join("?" * len(allowed))})'
        args = [nxt, now, aid, *allowed]
        if expected_rev is not None:
            sql += (" AND (SELECT COUNT(*) FROM approval_steps"
                    " WHERE approval_steps.approval_id=approvals.id AND step='resubmit')=?")
            args.append(int(expected_rev))
        cur = conn.execute(sql, args)
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


def kv_get(key: str) -> str:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT value FROM approval_kv WHERE key=?', (key,)).fetchone()
        return (r[0] if r else '') or ''


def kv_set(key: str, value: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('INSERT INTO approval_kv (key, value) VALUES (?,?) '
                     'ON CONFLICT(key) DO UPDATE SET value=excluded.value', (key, value))
        conn.commit()


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


def _acct_digits(s: str) -> str:
    return re.sub(r'\D', '', s or '')


def _bank_key(s: str) -> str:
    """'우리은행'·'우리 은행'·'우리' 를 같게."""
    return re.sub(r'\s|은행$', '', (s or '').strip()).replace('은행', '')


def add_known_accounts(partner: str, rows: List[Tuple[str, str, str]]) -> int:
    """거래처 확인 계좌 넣기 — rows=[(은행, 계좌, 이름표)]. 같은 계좌는 은행·이름표만 고친다. 넣은 수."""
    key = partner_key(partner)
    if not key:
        return 0
    n = 0
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        for bank, account, label in rows:
            digits = _acct_digits(account)
            if len(digits) < 8:
                continue
            conn.execute(
                '''INSERT INTO payee_known_accounts (partner_key, partner, bank, account, label, added_at)
                   VALUES (?,?,?,?,?,?)
                   ON CONFLICT(partner_key, account) DO UPDATE SET partner=excluded.partner, bank=excluded.bank,
                     label=excluded.label''',
                (key, partner, (bank or '').strip(), digits, (label or '').strip(), _now()))
            n += 1
        conn.commit()
    return n


def known_accounts(partner: str) -> List[dict]:
    key = partner_key(partner)
    if not key:
        return []
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM payee_known_accounts WHERE partner_key=? ORDER BY account', (key,)).fetchall()]


def match_known_account(partner: str, bank: str, account: str) -> Tuple[bool, Optional[dict]]:
    """(이 거래처에 확인 계좌 목록이 있는가, 맞는 계좌 행 또는 None). 은행이 다르면 다른 계좌로 본다."""
    rows = known_accounts(partner)
    digits = _acct_digits(account)
    for r in rows:
        if r['account'] == digits and (not r.get('bank') or _bank_key(r['bank']) == _bank_key(bank)):
            return True, r
    return bool(rows), None


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


def grn_mismatch(code: str, grn: str) -> str:
    """현장 코드 앞글자(G/R)와 고른 사업자가 다르면 안내 문구, 같거나 판단할 수 없으면 ''.
    공통비·프로젝트0 처럼 G/R 로 시작하지 않는 현장, 사업자 없음(N)은 묻지 않는다."""
    want, got = grn_from_code(code), (grn or '').strip()[:1].upper()
    if not want or got not in ('G', 'R') or want == got:
        return ''
    return (f"현장 {code}은 {want}({GRN_BIZ[want]}) 현장인데 {got}({GRN_BIZ[got]})로 골랐어요. "
            f"맞게 고른 거면 아래 버튼을 한 번 더 눌러 주세요.")


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
        # JW·YG 가 ① 을 눌러 한 번에 — 결과만 '최종 승인'으로('승인 (최종 승인 겸)'은 뜻이 안 들어왔다, 11-11 KiKO)
        first = f"✅ {_who(a)} 최종 승인 {_fmt_when(a['at'])}"
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
        mv = extra(row).get('fund_move')
        if mv:   # 법인 간 이체 — 어느 입금을 옮기는지(11-16)
            lines.append(f'옮기는 입금: {move_summary(mv)}')
        if content:
            lines.append(f'내용: {content}')
        # 예금주가 '(주)컴퓨존'이면 괄호가 겹쳐 '((주)컴퓨존)' — 예금주는 · 뒤로(11-06)
        lines.append(f"계좌: {row.get('bank') or '-'} {row.get('account') or '-'} · 예금주 {row.get('payee_name') or '-'}"
                     + (' · 예전 계좌 자동 채움' if extra(row).get('payee_saved') else '')
                     + (f" · ✅ 확인된 계좌({extra(row)['payee_known']})" if extra(row).get('payee_known') else ''))
        lines += fund_warning_lines(row)
        if row.get('cost_category') != COST_MOVE:   # 법인 간 이체는 증빙·VAT 가 없다
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
        foot.append(f"🔁 {executor_label(rv['actor_name'])} 보완 요청 {_fmt_when(rv['at'])} — {rv.get('reason') or '-'}")
    body = top + [_kv(ln) for ln in head[len(top):] + lines + foot]
    if closed:
        return _greyed(row, step_rows, emoji, label, body)
    return '\n'.join(f'>{ln}' for ln in '\n'.join(body).split('\n'))


def fund_warning_lines(row: dict) -> List[str]:
    """자금집행 카드 경고 줄(11-14 보안 점검 4차) — 승인자·이체 담당이 눈으로만 막던 것을 카드에 띄운다.
    ① 기억해 둔 거래처 계좌와 다른 계좌 ② 최근 같은 거래처·금액 결재 ③ 계산서 글 작성자(경영지원 외면 ⚠️)."""
    if row.get('form') != FORM_FUND:
        return []
    x = extra(row)
    out = []
    if x.get('payee_changed'):
        out.append(f"⚠️ 예전에 보낸 계좌({x['payee_changed']})와 달라요 — 이체 전에 거래처에 확인해 주세요")
    if x.get('payee_unlisted'):
        out.append('⚠️ 거래처가 알려 준 계좌 목록에 없는 계좌예요 — 이체 전에 거래처에 확인해 주세요')
    if '어음' in str(x.get('payee_known') or ''):
        out.append('⚠️ 어음 계좌예요 — 이체할 계좌가 맞는지 확인해 주세요')
    ids = [i for i in (x.get('dup_of') or []) if i != row.get('id')]
    if ids:
        names = []
        for i in ids[:3]:
            d = get(int(i))
            if d and d.get('status') not in (REJECTED, CANCELLED):
                names.append(f"{title(d)}({initial(d.get('requester_name'))}·{status_label(d)[1]})")
        if names:
            out.append(f"⚠️ 최근 같은 거래처·금액 결재: {', '.join(names)} — 두 번 이체가 아닌지 확인해 주세요")
    if x.get('old_prepaid'):
        d = get(int(x['old_prepaid']))
        if d:
            out.append(f"⚠️ 계산서 없이 먼저 보낸 같은 거래처·금액 선입금: {title(d)}"
                       f"({initial(d.get('requester_name'))}·{str(d.get('updated_at') or '')[:10]} 이체 완료)"
                       " — 그 선입금의 계산서면 다시 이체하지 마세요")
    if row.get('origin_ts') and 'inv_by' in x:
        who = x.get('inv_by') or '작성자 확인 안 됨'
        out.append(f"계산서 글: {who} 작성" + (' ⚠️ 경영지원이 올린 글이 아니에요' if x.get('inv_by_other') else ''))
    # 증빙은 계산서인데 계산서 글을 안 골랐으면 — 스레드에 붙을 사진이 없다(2026-11-17 #45 서울크레인).
    #   창에서 직접 올린 첨부가 있으면 증빙은 있는 셈이니 띄우지 않는다. 공통비 비품 구매는 선입금 뒤
    #   계산서를 받는 게 정상이라 공사 현장(G/R) 건만 본다 — 안 그러면 경고가 흔해져 무뎌진다.
    if (not row.get('origin_ts') and '계산서' in str(row.get('evidence') or '')
            and grn_from_code(row.get('project_code') or '')
            and row.get('cost_category') != COST_MOVE and not x.get('files')):
        out.append('⚠️ 계산서 글을 고르지 않았어요 — 선입금이면 그대로 두고, 계산서가 있으면 '
                   '#영업_관리 그 글에서 [📍 현장 지정]을 눌러 주세요')
    return out


def set_extra(aid: int, **kv) -> Optional[dict]:
    """extra 칸 몇 개만 덮어쓴다(카드 표시용, 11-14). 다른 스레드의 첨부 기록과 겹치지 않게 같은 잠금."""
    with _EXTRA_LOCK, _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = _row(conn, aid)
        if not r:
            return None
        x = {**extra(r), **kv}
        conn.execute('UPDATE approvals SET extra=? WHERE id=?', (json.dumps(x, ensure_ascii=False), aid))
        conn.commit()
        return _row(conn, aid)


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
    # 박스 첫 줄 제목은 위 굵은 제목과 겹친다(11-10 KiKO) — 제목과 그 밑 빈 줄은 빼고 금액 줄부터
    body = list(body)
    if body and title(row) in body[0]:
        body.pop(0)
    while body and not body[0].strip(' ⠀'):
        body.pop(0)
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
    # 승인·최종 승인·이체 완료 버튼엔 수정 회차를 싣는다 — 열어 둔 옛 카드·확인 창으로 고친 뒤의 결재를
    # 처리하지 않게(11-11 점검: 옛 계좌로 이체하고 새 결재를 완료 처리)
    rv = f'{aid}:{revision(aid)}' if st in (PENDING, FINAL, APPROVED) else aid
    out = []
    if st == PENDING and role in ('all', 'approver', 'final'):   # JW·YG 도 ① 을 누를 수 있다(11-09)
        out += [_btn('승인', A_APPROVE, rv, 'primary'), _btn('반려', A_REJECT, aid, 'danger')]
    if st == FINAL and role in ('all', 'final'):
        out += [_btn('최종 승인', A_FINAL, rv, 'primary'), _btn('반려', A_REJECT, aid, 'danger')]
    form = row.get('form')
    if st == APPROVED and role in ('all', 'executor') and form in TASK_FORMS:
        out.append(_btn('✅ 처리 완료', A_TRANSFER, rv, 'primary', _confirm(
            '처리 완료', f'{title(row)} 처리를 마쳤나요?', '처리 완료')))
        out.append(_btn('반려', A_REJECT, aid, 'danger'))   # 내용이 모자라면 경영지원이 사유와 함께 돌려보냄
    elif st == APPROVED and role in ('all', 'executor') and form == forms.FORM_EXPENSE:
        # 확인 창은 채널 카드에 실려 누구나 열 수 있다 — 개인 계좌는 빼고(경영지원 DM 에만, 11-06 검수)
        out.append(_btn('💸 이체 완료', A_TRANSFER, rv, 'primary', _confirm(
            '이체 완료', f"{row.get('requester_name')} 개인경비 {amt}\n본인 계좌(경영지원 DM)로 이체를 마쳤나요?",
            '이체 완료')))
    elif st == APPROVED and role in ('all', 'executor') and form == FORM_SUPPLY:
        out.append(_btn('🛒 구매 완료', A_TRANSFER, rv, 'primary', _confirm(
            '구매 완료', f"{row.get('content') or '비품'} ({amt})\n구매를 마쳤나요?", '구매 완료')))
    elif st == APPROVED and role in ('all', 'executor'):
        _x = extra(row) if form == FORM_FUND else {}
        _chg, _known = _x.get('payee_changed'), str(_x.get('payee_known') or '')
        out.append(_btn('💸 이체 완료', A_TRANSFER, rv, 'primary', _confirm(
            '이체 완료', f"{row.get('bank')} {row.get('account')} · 예금주 {row.get('payee_name')}\n{amt} 이체를 마쳤나요?"
            + (f"\n⚠️ 예전에 보낸 계좌({_chg})와 달라요." if _chg else '')
            + ('\n⚠️ 거래처가 알려 준 계좌 목록에 없어요.' if _x.get('payee_unlisted') else '')
            + ((f"\n✅ 확인된 계좌({_known})" + (' — ⚠️ 어음 계좌예요.' if '어음' in _known else '')) if _known else ''),
            '이체 완료')))
    # 순서: 처리하는 사람 버튼 → [보완 요청](돌려보내기) → 요청자 버튼 [수정하기]·[취소하기] 붙여서 맨 뒤(11-11 KiKO)
    if st in (PENDING, FINAL, APPROVED) and role == 'all':   # 고칠 게 있으면 반려 대신 돌려보내기(11-05)
        out.append(_btn('🔁 보완 요청', A_REVISE, aid))
    if st == REVISE and role in ('all', 'requester'):
        out.append(_btn('✏️ 수정하기', A_EDIT, aid, 'primary'))
    elif st in OPEN_STATUSES and role in ('all', 'requester') and requester_can_edit(row):   # 승인 전 직접(11-11)
        out.append(_btn('✏️ 수정하기', A_EDIT, aid))
    if st in OPEN_STATUSES and role in ('all', 'requester') and requester_can_cancel(row):
        out.append(_btn('✖️ 취소하기', A_CANCEL, aid, None, _confirm(
            '요청 취소', f'{title(row)} 요청을 취소할까요?', '취소하기')))
    return out


_ACTION_OF_BUTTON = {A_APPROVE: 'approve', A_FINAL: 'final', A_TRANSFER: 'transfer', A_REJECT: 'reject',
                     A_REVISE: 'revise', A_CANCEL: 'cancel'}


def offered_actions(row: dict) -> set:
    """채널 카드에 지금 그려지는 처리 단추 — 그룹웨어 결재함·서버 검사가 같은 표를 쓴다(11-16).
    예전엔 결재함이 '허용 상태 ∩ 권한'을 그대로 보여 줘 슬랙엔 없는 출구(이체 대기 중 요청 취소·자금집행 반려)가
    웹에만 있었다."""
    return {_ACTION_OF_BUTTON[b['action_id']] for b in buttons(row, 'all')
            if b.get('action_id') in _ACTION_OF_BUTTON}


def card_blocks(row: dict, step_rows: List[dict], project_line: str, quote_amount: int = 0,
                role: str = 'all') -> List[dict]:
    # 11-01 다른 채널 카드처럼 위·아래, 내용과 버튼 사이에 한 줄씩 여백.
    # section 글자 한도 3,000자 — 넘으면 카드 게시·갱신이 통째로 실패한다(긴 비고·사유, 11-05 점검)
    body = '⠀\n' + card_text(row, step_rows, project_line, quote_amount)
    if len(body) > 2900:
        body = body[:2880] + '\n…(길어서 줄였어요)'
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': body}}]
    if role == 'all' and row.get('status') in (PENDING, FINAL, APPROVED) and weekend_test_ids():
        # 시험 모드가 남아 있으면 카드에 바로 보이게 — 예전엔 아무 표시 없이 경영지원 DM 만 빠졌다(11-16 점검)
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': TEST_MODE_NOTE}]})
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
        elif requester_can_edit(row):
            hint = ('파일을 더 붙일 땐 이 카드 스레드에 올려 주세요. 올린 사람은 [✏️ 수정하기]로 바로 고칠 수 있고 '
                    '(승인은 다시 받아요), 다른 사람은 [🔁 보완 요청]으로 돌려보내 주세요.')
        else:
            hint = '파일을 더 붙일 땐 이 카드 스레드에 올려 주세요. 고칠 게 있으면 [🔁 보완 요청]으로 돌려보내 주세요.'
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
    if role == 'executor' and st == APPROVED and is_n_account(row) and len(n_executor_ids()) > 1:
        # N통장은 JW·YG 둘 다 받는다 — 둘 다 보내면 두 번 나간다(11-11 점검, KiKO: 주의 문구로)
        names = '·'.join(_N_LABELS.get(u, 'N통장 담당') for u in n_executor_ids())
        return (f'🔔 {title(row)} 승인이 끝났어요. {w}하고 카드에서 [{w} 완료]를 눌러 주세요.\n'
                f'⚠️ N통장 건이라 {names} 두 분께 같이 갔어요. 보내기 전에 카드에서 이미 [{w} 완료]가 됐는지 확인하고, '
                f'보낸 뒤에는 바로 [{w} 완료]를 눌러 주세요.')
    if role == 'executor' and st == APPROVED:
        return f'🔔 {title(row)} 승인이 끝났어요. {w}하고 카드에서 [{w} 완료]를 눌러 주세요.'
    if role == 'notice' and st == APPROVED:   # N통장 — 경영지원(SB)은 알림만(11-06)
        names = '·'.join(_N_LABELS.get(u, 'N통장 담당') for u in n_executor_ids())
        return f'ℹ️ {title(row)}은 N통장 건이라 {names}님에게 {w} 요청이 갔어요. (알림만 — 처리는 안 하셔도 돼요)'
    if role == 'requester' and st == REVISE:
        rv = last_revise(steps(row['id'])) or {}
        return (f"🔁 {initial(rv.get('actor_name'))}님이 {title(row)} 보완을 요청했어요.\n"
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
    # 카드가 없을 때 '처리해 주세요'는 진행 중인 결재에만 — 끝난 결재엔 할 일이 없다(11-16 #23·#24)
    link = (f'<{card_link}|👉 결재 카드 열기>' if card_link
            else '#전자_결재 채널의 카드에서 처리해 주세요.' if row.get('status') in OPEN_STATUSES else '')
    text = f'{dm_intro(row, role)}\n\n{summary}\n{step_line(row, steps(row["id"]))}'
    # 수정하기는 DM 에서도 바로(채널 카드까지 안 가도)
    els = [_btn('✏️ 수정하기', A_EDIT, row['id'], 'primary')] if role == 'requester' and row.get('status') == REVISE else None
    notes = [{'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': link}]}] if link else None
    return spaced(text[:2900], els, notes)


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
    # 지점마다 계좌가 다른 거래처(확인 계좌 목록이 있는 곳)는 지점을 목록에서 바로 고른다(11-16 KiKO).
    # '하이엠' → 하이엠솔루텍 지점 전부, '하이엠 강남'·'강남' → 그 지점. 고르면 은행·계좌가 채워진다.
    branch_opts = _branch_options(q)
    branch_keys = {partner_key(o['_partner']) for o in branch_opts}
    opts = []
    for o in branch_opts:
        o.pop('_partner')
        opts.append(o)
    opts += [_partner_opt(n) for n in hits if partner_key(n) not in branch_keys]
    opts = opts[:99]
    multi = len(q.split()) > 1 and branch_opts   # '하이엠 강남' — 지점을 찾는 검색이지 새 상호가 아니다
    if not any(partner_key(n) == qk for n in hits) and qk not in branch_keys and not multi:
        # 아는 거래처가 걸리면 그게 먼저(실수로 새로 만들지 않게), '새 거래처'는 맨 아래
        opts.append({'text': {'type': 'plain_text', 'text': f'새 거래처: {q}'[:75]},
                     'value': (PARTNER_NEW + q)[:150]})
    return {'options': opts}


PARTNER_BRANCH = 'acct:'   # 지점 계좌 선택지 값 — 'acct:' + {"p": 거래처, "a": 계좌 숫자(빈 값=직접 입력)}


def _all_known_accounts() -> List[dict]:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM payee_known_accounts ORDER BY partner_key, label, account').fetchall()]


def branch_option(partner: str, row: Optional[dict]) -> dict:
    """지점 계좌 선택지. row=None 이면 '목록에 없는 계좌(직접 입력)'."""
    if row:
        bank = _bank_key(row.get('bank') or '')
        label = row.get('label') or '확인된 계좌'
        if '어음' in label:   # 어음 계좌는 고를 때부터 눈에 띄게(이체 확인 창에도 ⚠️)
            label = f'⚠️ {label} 계좌'
        text = f"{partner} · {label} ({bank} …{row['account'][-4:]})"
        acct = row['account']
    else:
        text, acct = f'{partner} · 목록에 없는 계좌 (직접 입력)', ''
    value = PARTNER_BRANCH + json.dumps({'p': partner[:60], 'a': acct}, ensure_ascii=False)
    return {'text': {'type': 'plain_text', 'text': text[:75]}, 'value': value[:150]}


def _branch_options(query: str) -> List[dict]:
    """검색어 낱말이 모두 '거래처+지점 이름'에 들어 있는 지점 계좌 선택지(+거래처마다 '직접 입력' 하나).

    '_partner' 키에 거래처 이름을 같이 넣어 돌려준다(부르는 쪽이 빼고 쓴다)."""
    words = [partner_key(w) for w in (query or '').split() if partner_key(w)]
    if not words:
        return []
    out, partners = [], []
    for r in _all_known_accounts():
        name = r.get('partner') or r['partner_key']
        hay = partner_key(name) + partner_key(r.get('label') or '')
        if all(w in hay for w in words):
            out.append({**branch_option(name, r), '_partner': name})
            if name not in partners:
                partners.append(name)
    for name in partners:
        out.append({**branch_option(name, None), '_partner': name})
    return out


def branch_from_value(value: str) -> Optional[Tuple[str, str]]:
    """지점 계좌 선택지면 (거래처, 계좌 숫자 — 직접 입력이면 ''), 아니면 None."""
    v = value or ''
    if not v.startswith(PARTNER_BRANCH):
        return None
    try:
        d = json.loads(v[len(PARTNER_BRANCH):])
        return (d.get('p') or '', d.get('a') or '')
    except (ValueError, AttributeError):
        return None


def partner_from_value(value: str) -> str:
    v = value or ''
    br = branch_from_value(v)
    if br:
        return br[0]
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
    if meta.get('move'):   # 법인 간 이체 — 정해진 값은 숨기고 확인만(11-16)
        return move_request_view(meta, prefill)
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
                       '📄 거래처·금액은 계산서 글 그대로예요. 현장을 바꾸면 계산서 글의 현장 지정도 같이 바뀌어요. '
                       '계산서 파일은 처음 올릴 때 카드 스레드에 붙었어요.'}]})
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
    # 증빙 기본값은 계산서 글을 골랐을 때만 '전자계산서' — 안 고르고 올리는 선입금에서도 기본값이 남아
    #   '증빙: 전자계산서'인데 스레드엔 계산서가 없는 카드가 됐다(11-01 검수 개선안, 2026-11-17 #45).
    ev_el = {'type': 'static_select', 'action_id': 'evidence', 'options': ev_opts,
             'placeholder': {'type': 'plain_text', 'text': '증빙 종류 고르기'}}
    if inv:
        ev_el['initial_option'] = ev_opts[0]
    # 법인 간 이체는 증빙이 없다(제출 때 '없음 (법인 간 이체)'로 채워진다) — 기본값을 비웠으니 비워도 넘어가게
    ev_move = meta.get('move_opts') is not None or bool(meta.get('move_manual'))
    blocks += [
        {'type': 'input', 'block_id': 'evidence', 'label': {'type': 'plain_text', 'text': '증빙'},
         'element': ev_el, **({'optional': True} if ev_move else {})},
        # '법인 간 이체'를 고르면 옮길 입금을 고르게 창이 바뀐다(dispatch, 11-16)
        {'type': 'input', 'block_id': 'cost', 'label': {'type': 'plain_text', 'text': '원가 분류'},
         'dispatch_action': True,
         'element': {'type': 'static_select', 'action_id': A_COST, 'options': _opts(fund_cost_categories()),
                     'placeholder': {'type': 'plain_text', 'text': '원가 분류'}}},
    ]
    if meta.get('move_opts') is not None:   # 법인 간 이체 — 옮길 입금 고르기
        opts = [o for o in meta['move_opts']][:20] + [move_option(None)]
        blocks.append({'type': 'input', 'block_id': 'move_pick', 'dispatch_action': True, 'optional': True,
                       'label': {'type': 'plain_text', 'text': '옮길 입금'},
                       'element': {'type': 'static_select', 'action_id': A_MOVE_PICK, 'options': opts,
                                   'placeholder': {'type': 'plain_text', 'text': '잘못 들어온 입금 고르기'}},
                       'hint': {'type': 'plain_text', 'text':
                                '고르면 금액·받는 통장이 채워져요. 목록에 없으면 「직접 입력」을 고르세요.'}})
    elif meta.get('move_manual'):
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                       '🔁 법인 간 이체 — 「사업자 구분」은 돈이 잘못 들어온 법인, 받는 통장은 공사 법인 통장이에요. '
                       '증빙은 없어도 돼요.'}]})
    blocks += [
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
        if meta.get('partner_opt'):   # 지점 계좌로 고른 거래처 — 고른 지점이 그대로 보이게
            p_el['initial_option'] = meta['partner_opt']
        elif meta.get('partner'):
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
    inv_partner = (inv or {}).get('parsed', {}).get('partner') or ''
    if inv and meta.get('branch_list') and not edit:
        # 계산서로 거래처가 정해지면 거래처 칸이 숨겨진다 — 지점은 여기서 고른다
        rows = known_accounts(inv_partner)
        b_opts = ([branch_option(inv_partner, r) for r in rows] + [branch_option(inv_partner, None)])[:100]
        b_el = {'type': 'static_select', 'action_id': A_BRANCH, 'options': b_opts,
                'placeholder': {'type': 'plain_text', 'text': '지점 고르기'}}
        cur = meta.get('partner_opt')
        if cur:
            b_el['initial_option'] = next((o for o in b_opts if o['value'] == cur['value']), cur)
        blocks.append({'type': 'input', 'block_id': 'branch', 'optional': True, 'dispatch_action': True,
                       'label': {'type': 'plain_text', 'text': '지점 계좌'}, 'element': b_el,
                       'hint': {'type': 'plain_text', 'text': '지점마다 계좌가 달라요. 고르면 은행·계좌가 채워져요.'}})
    if meta.get('branch') and payee:
        head = f"*이체 계좌* — ✅ 확인된 계좌({meta['branch']})"
    elif meta.get('branch_list') and not payee:
        where = ('위 「지점 계좌」에서 지점을' if inv else
                 '위 「거래처」에 상호를 다시 입력하면 지점별로 나와요. 지점을')
        head = (f'*이체 계좌* — 지점마다 계좌가 다른 거래처예요. {where} 고르면 은행·계좌가 채워져요. '
                '목록에 없는 계좌면 직접 넣어 주세요.')
    elif payee:
        head = '*이체 계좌*'
    elif meta.get('partner') or inv:
        head = '*이체 계좌* — 처음 보내는 거래처예요. 예금주·은행·계좌를 넣어 주세요.'
    else:
        head = '*이체 계좌* — 거래처를 고르면 예전에 보낸 계좌가 있을 때 채워집니다.'
    blocks.append({'type': 'section', 'text': {'type': 'mrkdwn', 'text': head}})
    if payee and not meta.get('edit') and not meta.get('branch'):   # 수정 창은 올렸던 계좌 그대로라 '예전 계좌' 안내를 안 띄운다
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
        'evidence': _sel('evidence'), 'cost_category': _sel('cost', A_COST) or _sel('cost'), 'content': _txt('content'),
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
    if not inv and not meta.get('move'):
        f['_bad'] = amount_errors({'amount': _txt('amount')})
    apply_move_fields(f, meta)   # 법인 간 이체 — 정해진 값(11-16)
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
        dup = find_live_by_origin(f['origin_channel'], f['origin_ts'], exclude_id=f.get('_edit'))
        if dup:
            errs['invoice'] = f'이 계산서로 올린 {title(dup)}이 이미 있어요.'
        else:
            # 계산서 없이 먼저 올린(선입금) 같은 건 — G/R·금액·거래처가 같으면 두 번 이체될 수 있다(11-06 검수)
            m = find_unlinked_match(f.get('grn') or '', int(f.get('amount') or 0), f.get('partner') or '',
                                    f.get('project_code') or '')
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
    if not errs:
        errs.update(move_errors(f))   # 법인 간 이체 — 보내는 법인·같은 입금 두 번(11-16)
    return errs


def mark_payee_saved(f: Dict) -> None:
    """기억해 둔 계좌와 같은 계좌로 보내면 카드에 표시(경영지원이 이체 전에 한 번 더 보게)."""
    saved = lookup_payee(f.get('partner', ''))
    same = bool(saved and normalize_account(saved.get('account')) == normalize_account(f.get('account'))
                and (saved.get('bank') or '') == (f.get('bank') or ''))   # 은행만 바뀐 것도 '다른 계좌'(11-11)
    x = dict(f.get('extra') or {})
    if same:
        x['payee_saved'] = True
    # 거래처가 확인해 준 계좌 목록이 있으면 그 목록으로 판단(11-15 하이엠솔루텍 — 센터마다 계좌가 달라
    # 다른 센터로 보낼 때마다 '예전 계좌와 달라요'가 떠 보완 요청이 났다, #25)
    has_list, hit = match_known_account(f.get('partner', ''), f.get('bank', ''), f.get('account', ''))
    x.pop('payee_known', None)
    x.pop('payee_unlisted', None)
    if has_list and normalize_account(f.get('account')):
        x.pop('payee_changed', None)
        if hit:
            x['payee_known'] = hit.get('label') or '확인된 계좌'
        else:
            x['payee_unlisted'] = True
        f['extra'] = x
        return
    # 이체한 적 있는 거래처인데 계좌가 다르면 카드·이체 확인 창에 ⚠️ — 계좌 바꿔치기·오타(11-14 보안 점검 4차)
    if saved and not same and normalize_account(f.get('account')):
        acct = normalize_account(saved.get('account'))
        x['payee_changed'] = f"{saved.get('bank') or ''} …{acct[-4:]}".strip()
    else:
        x.pop('payee_changed', None)
    f['extra'] = x


# ── 법인 간 이체 (11-16) ─────────────────────────────────────────────
# 고객이 다른 법인 통장으로 입금 → #입금_관리 [지정] 두 번(자금 이동 요청) → 매니저 DM [📝 결재 올리기] →
# 채워진 창을 확인하고 요청 → 승인 → 샛별님 이체 → 그룹/글로벌 통장 입금 문자로 [법인 간 이체로 처리].
# 결재는 그 자금 이동 요청(시트 메모 줄)과 extra['fund_move'] 로 연결 — 같은 입금으로 두 번 못 올린다.
COST_MOVE = '이동:법인간'
EVIDENCE_MOVE = '없음 (법인 간 이체)'
ENTITY_SHORT = {'G': '글로벌', 'R': '글로벌그룹'}
ENTITY_PAYEE = {   # 받는 통장 — 그룹 기업은행(하도급지킴이)은 관공서 전용이라 하나은행
    'G': {'payee_name': '(주)아이티글로벌', 'bank': 'IBK기업은행', 'account': '452-039388-01-011'},
    'R': {'payee_name': '(주)아이티글로벌그룹', 'bank': '하나은행', 'account': '255-910014-31304'},
}
MOVE_PICK_MANUAL = 'manual'


def fund_cost_categories() -> List[Tuple[str, str]]:
    from dashboard.services.purchase_invoice_assign import COST_CATEGORIES, FUND_ONLY_COST
    return list(COST_CATEGORIES) + list(FUND_ONLY_COST)


def move_from_req(req: dict) -> dict:
    """자금 이동 요청(fund_transfer.pending_requests 한 줄) → 결재에 붙일 연결 정보."""
    m = {k: req.get(k) for k in ('code', 'stage', 'partner', 'bank', 'from', 'to', 'date')}
    m['amount'] = int(req.get('amount') or 0)
    m['key'] = move_key(m)
    return m


def move_key(m: dict) -> str:
    return (f"{m.get('code')}|{m.get('stage')}|{m.get('from')}>{m.get('to')}|{int(m.get('amount') or 0)}|"
            f"{partner_key(m.get('partner') or '')}|{m.get('date') or ''}")


def _md(date: str) -> str:
    """'2026/11/16' → '11/16'."""
    parts = re.split(r'[/\-.]', date or '')
    return '/'.join(parts[-2:]) if len(parts) >= 2 else (date or '')


def move_summary(m: dict) -> str:
    who = ' '.join(x for x in ((m.get('partner') or '').strip(), _md(m.get('date') or '')) if x)
    return (f"{m.get('code')} {m.get('stage')} 입금({who + ' ' if who else ''}{int(m.get('amount') or 0):,}원) · "
            f"{ENTITY_SHORT.get(m.get('from'), m.get('from'))} 통장 → {ENTITY_SHORT.get(m.get('to'), m.get('to'))} 통장")


def move_content(m: dict) -> str:
    who = (m.get('partner') or '').strip()
    return (f"{m.get('code')} {m.get('stage')} 입금" + (f"({who} {_md(m.get('date') or '')})" if who else '')
            + f"이 {ENTITY_SHORT.get(m.get('from'), '')} 통장으로 들어와 {ENTITY_SHORT.get(m.get('to'), '')} 통장으로 옮김")


def find_move_approval(key: str, exclude_id: Optional[int] = None) -> Optional[dict]:
    """그 자금 이동 요청으로 올린 결재(반려·취소 제외)."""
    if not key:
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            "SELECT * FROM approvals WHERE form=? AND status NOT IN (?, ?) AND extra LIKE ? ORDER BY id DESC",
            (FORM_FUND, REJECTED, CANCELLED, '%fund_move%')).fetchall()
    for r in rows:
        d = dict(r)
        if d['id'] != exclude_id and (extra(d).get('fund_move') or {}).get('key') == key:
            return d
    return None


def open_moves(project_code: str = '', reqs: Optional[List[dict]] = None) -> List[dict]:
    """결재가 아직 안 올라간 자금 이동 요청 → [move] (현장을 주면 그 현장만)."""
    if reqs is None:
        try:
            from dashboard.services.fund_transfer import pending_requests
            reqs = pending_requests()
        except Exception as exc:
            logger.warning(f'[전자결재] 자금 이동 요청 조회 실패: {exc}')
            reqs = []
    out = []
    for r in reqs:
        if project_code and r.get('code') != project_code:
            continue
        m = move_from_req(r)
        if m['from'] in ENTITY_PAYEE and m['to'] in ENTITY_PAYEE and not find_move_approval(m['key']):
            out.append(m)
    return out


def move_option(m: Optional[dict]) -> dict:
    if not m:
        return {'text': {'type': 'plain_text', 'text': '목록에 없음 (직접 입력)'}, 'value': MOVE_PICK_MANUAL}
    return {'text': {'type': 'plain_text', 'text': move_summary(m)[:75]}, 'value': m['key'][:150]}


def apply_move_fields(f: Dict, meta: dict) -> None:
    """법인 간 이체 — 정해진 값 채우기(창에 안 보이는 칸). 원가 분류가 '법인 간 이체'면 직접 입력도."""
    mv = meta.get('move')
    if mv:
        f.update(project_code=mv['code'], cost_category=COST_MOVE, grn=mv['from'], amount=int(mv['amount']))
        x = dict(f.get('extra') or {})
        x['fund_move'] = mv
        f['extra'] = x
    if f.get('cost_category') != COST_MOVE:
        return
    f['evidence'] = EVIDENCE_MOVE
    to = (f.get('project_code') or '')[:1]
    if not f.get('partner') and to in ENTITY_PAYEE:
        f['partner'] = ENTITY_PAYEE[to]['payee_name']


def move_errors(f: Dict) -> Dict[str, str]:
    """법인 간 이체 검사 + 직접 입력이면 맞는 자금 이동 요청을 찾아 연결. 오류 {block_id: 문구}."""
    if f.get('cost_category') != COST_MOVE:
        return {}
    code = f.get('project_code') or ''
    to, frm = code[:1], f.get('grn') or ''
    if to not in ENTITY_PAYEE:
        return {'cost': '법인 간 이체는 글로벌(G)·글로벌그룹(R) 공사에서만 쓸 수 있어요.'}
    if frm == to or frm not in ENTITY_PAYEE:
        return {'grn': f'돈이 잘못 들어온 법인을 골라 주세요. 이 공사는 {ENTITY_SHORT[to]}({to}) 공사라 '
                       f'{ENTITY_SHORT["G" if to == "R" else "R"]} 통장에서 보내요.'}
    x = dict(f.get('extra') or {})
    if not x.get('fund_move'):
        hit = next((m for m in open_moves(code) if m['from'] == frm and m['to'] == to
                    and m['amount'] == int(f.get('amount') or 0)), None)
        if hit:
            x['fund_move'] = hit
            f['extra'] = x
    mv = x.get('fund_move')
    if mv:
        dup = find_move_approval(mv.get('key'), exclude_id=f.get('_edit'))
        if dup:
            return {'content': f"이 입금으로 올린 {title(dup)}({initial(dup.get('requester_name'))}·"
                               f"{status_label(dup)[1]})이 이미 있어요."}
    return {}


def move_request_view(meta: dict, prefill: Optional[dict] = None) -> dict:
    """[📝 결재 올리기]·옮길 입금을 고른 창 — 현장·금액·보내는 법인·받는 통장이 정해져 있고 확인만 한다."""
    from dashboard.services.lead_helpers import slack_esc   # 입금자 이름은 은행 문자에서 읽은 글
    mv = meta['move']
    p = prefill or {}
    to = mv['to']
    payee = meta.get('payee') or ENTITY_PAYEE.get(to)
    blocks: List[dict] = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': (
            f"🔁 *법인 간 이체*\n{slack_esc(move_summary(mv))}\n"
            f"*{int(mv['amount']):,}원*을 {ENTITY_SHORT.get(mv['from'])} 통장에서 {ENTITY_SHORT.get(to)} 통장으로 보내요.")}},
        {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
            '현장·금액·보내는 통장은 입금 기록 그대로예요. 내용과 받는 통장을 확인하고 요청해 주세요. '
            '원가에는 안 들어가요.'}]},
        {'type': 'input', 'block_id': 'content', 'label': {'type': 'plain_text', 'text': '내용'},
         'element': {'type': 'plain_text_input', 'action_id': 'content', 'max_length': 300,
                     'initial_value': (p.get('content') or move_content(mv))[:300]}},
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': f'*받는 통장* — {ENTITY_SHORT.get(to)} 통장'}},
    ]
    pv = int(meta.get('pv') or 0)
    blocks += transfer_input_blocks(payee, urgency=p.get('urgency') or '일반', sfx=f'_{pv}' if pv else '')
    blocks.append({'type': 'input', 'block_id': 'memo', 'optional': True,
                   'label': {'type': 'plain_text', 'text': '비고 (특이사항)'},
                   'element': {'type': 'plain_text_input', 'action_id': 'memo', 'multiline': True,
                               'max_length': 1000}})
    if p.get('memo'):
        _set_initial(blocks, 'memo', p['memo'])
    title_, submit = forms.edit_texts(meta, '법인 간 이체', '요청')
    return {
        'type': 'modal', 'callback_id': CB_NEW,
        'title': {'type': 'plain_text', 'text': title_},
        'submit': {'type': 'plain_text', 'text': submit},
        'close': {'type': 'plain_text', 'text': '취소'}, 'notify_on_close': True,
        'private_metadata': json.dumps(meta, ensure_ascii=False),
        'blocks': blocks,
    }


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


_LOGIN_ONLY_SEGS = {'myshop', 'mypage', 'my', 'cart', 'basket', 'order', 'orders', 'orderlist', 'checkout',
                    'login', 'signin', 'member', 'myinfo'}


def login_only_link(url: Optional[str]) -> bool:
    """로그인해야 보이는 쇼핑몰 주소(마이페이지·장바구니·주문)인가 — 경로 조각 이름으로만 본다(11-11 #15
    diaryon …/myshop/index.html). 상품 페이지 주소에 그 낱말이 섞인 것(order-made 등)은 안 걸린다."""
    m = re.match(r'https?://[^/?#]+(/[^?#]*)?', (url or '').strip())
    if not m:
        return False
    segs = [re.sub(r'\.(html?|php|asp|aspx|jsp)$', '', s.lower()) for s in (m.group(1) or '').split('/') if s]
    return any(s in _LOGIN_ONLY_SEGS for s in segs)


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
