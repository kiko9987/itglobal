"""그룹웨어 재무 > 입출금(2026-11-26). 관리자(admin·super_admin)만 — 잔액이 보이는 화면.
규칙은 services/bank_ledger.py. 상태를 바꾸는 API 는 /api/ 로 시작(보안 미들웨어 경로 규칙).
"""
from datetime import date, timedelta

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.blueprints.hr import BUG_ERRORS, _is_admin, _is_field, gw_page
from dashboard.services import bank_ledger as bl
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_finance_bp = Blueprint('gw_finance', __name__)

HIDDEN_CODES = ('농협',)   # 입출금 화면에 안 보이는 계좌(장부엔 남음)


def _admin() -> bool:
    return _is_admin() and not _is_field()


@gw_finance_bp.before_request
def _admin_only():
    if not session.get('user') or _admin():
        return None
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'message': '관리자만 볼 수 있는 화면이에요.'}), 403
    return redirect('/home')


def _by() -> str:
    u = session.get('user') or {}
    return u.get('name') or u.get('email') or ''


def _err(e: Exception):
    if isinstance(e, BUG_ERRORS):
        logger.error(f'[입출금] 처리 실패(코드 오류): {type(e).__name__}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    if isinstance(e, PermissionError):
        return jsonify({'success': False, 'message': str(e)}), 403
    if isinstance(e, LookupError):
        return jsonify({'success': False, 'message': str(e)}), 404
    if isinstance(e, ValueError):
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.error(f'[입출금] 처리 실패: {e}', exc_info=True)
    return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500


def _day(v, default: date) -> str:
    try:
        return date.fromisoformat(str(v or '')).isoformat()
    except ValueError:
        return default.isoformat()


def _appr(a: dict) -> dict:
    """화면에 보일 결재 요약(관리자 전용 화면)."""
    from dashboard.services import e_approval_forms  # noqa: F401 — 양식마다 이름이 FORM_LABELS 에 들어감
    from dashboard.services.e_approval import FORM_LABELS
    return {'id': a['id'], 'form': FORM_LABELS.get(a.get('form'), a.get('form') or ''), 'content': a.get('content') or '',
            'partner': a.get('partner') or a.get('payee_name') or '', 'amount': int(a.get('amount') or 0),
            'requester': a.get('requester_name') or '', 'grn': str(a.get('grn') or '')[:1].upper(),
            'project_code': a.get('project_code') or '', 'done_at': str(a.get('done_at') or '')[:16]}


@gw_finance_bp.route('/finance/bank')
@gw_page
def bank_page():
    return render_template('gw/finance_bank.html')


@gw_finance_bp.route('/api/finance/bank', methods=['GET'])
@gw_page
def api_list():
    try:
        today = date.today()
        start = _day(request.args.get('start'), today - timedelta(days=6))
        end = _day(request.args.get('end'), today)
        if start > end:
            start, end = end, start
        acct = int(request.args.get('account') or 0) or None
        # 출금 ↔ 전자결재 이체 완료 짝(11-26) — 서버가 알아서. 실패해도 목록은 보여 줌
        d0, d1 = date.fromisoformat(start), date.fromisoformat(end)
        approvals = []
        try:
            from dashboard.services import fund_daily as fd
            approvals = fd.transfers_done((d0 - timedelta(days=2)).isoformat(), (d1 + timedelta(days=3)).isoformat())
            bl.link_approvals((d0 - timedelta(days=1)).isoformat(), (d1 + timedelta(days=2)).isoformat(), approvals)
        except Exception as e:
            logger.error(f'[입출금] 전자결재 짝짓기 실패: {e}', exc_info=True)
        rows = bl.list_tx(start, end, acct, request.args.get('direction') or '',
                          request.args.get('unassigned') == '1', (request.args.get('q') or '').strip()[:40])
        # 농협(N통장)은 SB 가 관리하지 않는 계좌(11-26 부장님) — 화면·합계에서 뺀다.
        # 장부엔 그대로 남김: 자금일보 초안이 전자결재 N 출금과 짝지을 때 씀(services/fund_daily.py)
        rows = [r for r in rows if r.get('acct_code') not in HIDDEN_CODES]
        # 자금일보에서 아침마다 옮겨 오는 잔액 줄은 거래가 아니라 거래 내역엔 숨김 — 구분 '잔액 맞추기'를 고르면 보임
        if request.args.get('direction') != bl.BAL:
            rows = [r for r in rows if r.get('source') != bl.REPORT]
        amap = {a['id']: a for a in approvals}
        for r in rows:
            a = amap.get(r.get('approval_id')) if r.get('approval_id') else None
            r['approval'] = _appr(a) if a else ({'id': r['approval_id']} if r.get('approval_id') else None)
            r['no_approval'] = r['direction'] == bl.OUT and r.get('source') not in ('manual', bl.REPORT) and not r.get('approval_id')
        tot = {'in': sum(r['amount'] for r in rows if r['direction'] == bl.IN),
               'out': sum(r['amount'] for r in rows if r['direction'] == bl.OUT),
               'unassigned': sum(1 for r in rows if r['direction'] in (bl.IN, bl.OUT) and not (r['memo'] or '')),
               'no_approval': sum(1 for r in rows if r['no_approval'])}
        appr = request.args.get('appr') or ''
        if appr == 'linked':
            rows = [r for r in rows if r.get('approval_id')]
        elif appr == 'none':
            rows = [r for r in rows if r['no_approval']]
        # 이체 완료했는데 출금 문자가 없는 결재 — 문자 오는 사업자(G 1011·R 9012)만, 장부에 문자가 들어오기 시작한 뒤만.
        # 농협(N)은 관리 안 하는 계좌라 뺌. G·R 이라도 문자 없는 통장(1012 등)에서 나갔으면 여기 뜰 수 있음(화면에 적어 둠)
        since = bl.sms_start()
        linked = bl.linked_approval_ids() if since else set()
        no_out = [_appr(a) for a in approvals
                  if since and start <= str(a.get('done_at') or '')[:10] <= end and str(a.get('done_at')) >= since
                  and str(a.get('grn') or '')[:1].upper() in ('G', 'R') and a['id'] not in linked]
        tot['no_out'] = len(no_out)
        return jsonify({'success': True, 'start': start, 'end': end, 'rows': rows, 'total': tot, 'no_out': no_out,
                        'accounts': [a for a in bl.summary(end) if a['code'] not in HIDDEN_CODES]})
    except Exception as e:
        return _err(e)


@gw_finance_bp.route('/api/finance/bank', methods=['POST'])
@gw_page
def api_add():
    try:
        tid = bl.add_manual(request.get_json(silent=True) or {}, _by(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'id': tid})


@gw_finance_bp.route('/api/finance/bank/<int:tid>', methods=['POST'])
@gw_page
def api_update(tid):
    try:
        bl.update(tid, request.get_json(silent=True) or {}, _by(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})


@gw_finance_bp.route('/api/finance/bank/<int:tid>/delete', methods=['POST'])
@gw_page
def api_delete(tid):
    try:
        bl.delete(tid, _by(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})
