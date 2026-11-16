# -*- coding: utf-8 -*-
"""슬랙 [✏️ 내용 수정] 창 (2026-11-16 상태 머신 점검 A6·A7).

① 창을 연 때 값을 기억해, 직원이 실제로 바꾼 칸만 쓴다(열어 둔 사이 남이 바꾼 칸을 옛 값으로 되돌리지 않음).
② 직원이 바꾼 칸을 그 사이 남도 바꿨으면 아무것도 저장하지 않고 다시 열어 달라고 안내.
③ 취소된 공사는 창을 열지도, 저장하지도 않는다.
④ 재개 카드는 시트를 다시 읽지 않고 캐시+재개 뒤 값으로 그린다.
"""
import json
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb
import dashboard.services.project_service as ps
import dashboard.services.project_slack_actions as psa

BASE_PROJECT = {
    '프로젝트 코드': 'G4300-JW', '거래처명': '메이드바이', '공사 내용': '천장형 2대 설치',
    '도급 구분': '외주', '시공자': '정건우', '총액 1': 5000000, '부가세': True,
    '공사 시작': '2026-11-20', '공사 종료': '2026-11-21', '수금 관련 특이사항': '-',
}


class _Client:
    def __init__(self):
        self.ephemeral, self.views = [], []

    def chat_postEphemeral(self, **k):
        self.ephemeral.append(k['text'])

    def views_open(self, **k):
        self.views.append(('open', k['view']))
        return {'view': {'id': 'V1'}}

    def views_update(self, **k):
        self.views.append(('update', k['view']))


def _state(p):
    """창이 미리 채운 값 그대로의 제출 상태 (직원이 아무것도 안 바꿈)."""
    def _opt(v):
        return {'value': v}
    return {
        'partner_name': {'value': {'selected_option': _opt(p['거래처명'])}},
        'content': {'value': {'value': p['공사 내용']}},
        'contract_type': {'value': {'selected_options': [_opt(x) for x in p['도급 구분'].split(', ') if x]}},
        'contractor': {'value': {'selected_options': [_opt(x) for x in p['시공자'].split(', ') if x]}},
        'amount': {'value': {'value': f"{int(p['총액 1']):,}"}},
        'vat': {'value': {'selected_options': [_opt('sep')] if p['부가세'] else []}},
        'start_date': {'value': {'selected_date': p['공사 시작']}},
        'end_date': {'value': {'selected_date': p['공사 종료']}},
        'reason': {'value': {'value': '현장 변경'}},
    }


@pytest.fixture
def env(monkeypatch):
    cur = {'p': dict(BASE_PROJECT)}
    edits = []
    monkeypatch.setattr(ps, 'get_project_records', lambda *a, **k: [cur['p']])
    monkeypatch.setattr(psa, 'perform_edit', lambda code, upd, reason, ini: edits.append(upd) or {'ok': True})
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')
    monkeypatch.setattr(sb, '_post_project_edit_notice_card', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_notify_project_edit_result', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_find_pending_amount_request', lambda code: None)
    monkeypatch.setattr(sb, '_post_amount_edit_request_card', lambda *a, **k: 'ts')
    import dashboard.services.project_slack_notifier as psn
    monkeypatch.setattr(psn, 'refresh_project_card_license', lambda *a, **k: None)
    return type('E', (), {'cur': cur, 'edits': edits})


def _submit(state, base=True, opened_with=None):
    meta = {'code': 'G4300-JW', 'channel': 'C1', 'message_ts': '1.1'}
    if base:
        meta['base'] = sb._edit_form_base(opened_with or BASE_PROJECT)
    c = _Client()
    sb._process_project_edit_submission(c, {'user': {'id': 'U1'}},
                                        {'private_metadata': json.dumps(meta, ensure_ascii=False),
                                         'state': {'values': state}})
    return c


def test_untouched_field_changed_by_other_is_not_reverted(env):
    env.cur['p'] = dict(BASE_PROJECT, 시공자='이경택')          # 창 연 뒤 PM 에서 시공자 변경
    st = _state(BASE_PROJECT)
    st['content']['value']['value'] = '천장형 3대 설치'          # 직원은 공사 내용만 고침
    _submit(st)
    assert env.edits == [{'공사 내용': '천장형 3대 설치'}]


def test_old_behaviour_without_base_reverts(env):
    """배포 전에 열린 창(base 없음)은 예전처럼 지금 값과 비교 — 비교 기준 확인용."""
    env.cur['p'] = dict(BASE_PROJECT, 시공자='이경택')
    st = _state(BASE_PROJECT)
    st['content']['value']['value'] = '천장형 3대 설치'
    _submit(st, base=False)
    assert env.edits and env.edits[0].get('시공자') == '정건우'


def test_conflict_on_same_field_saves_nothing(env):
    env.cur['p'] = dict(BASE_PROJECT, 시공자='이경택')
    st = _state(BASE_PROJECT)
    st['contractor']['value']['selected_options'] = [{'value': '김민수'}]
    st['content']['value']['value'] = '천장형 3대 설치'
    c = _submit(st)
    assert env.edits == []
    assert '시공자' in c.ephemeral[0] and '저장하지 않았습니다' in c.ephemeral[0]


def test_same_change_by_both_is_no_conflict(env):
    env.cur['p'] = dict(BASE_PROJECT, 시공자='김민수')
    st = _state(BASE_PROJECT)
    st['contractor']['value']['selected_options'] = [{'value': '김민수'}]
    c = _submit(st)
    assert env.edits == [] and '바뀐 내용이 없어' in c.ephemeral[0]


def test_no_change_submit_writes_nothing(env):
    c = _submit(_state(BASE_PROJECT))
    assert env.edits == [] and '바뀐 내용이 없어' in c.ephemeral[0]


def test_amount_change_still_goes_to_request(env, monkeypatch):
    sent = []
    monkeypatch.setattr(sb, '_post_amount_edit_request_card', lambda p, upd, *a: sent.append(upd) or 'ts')
    st = _state(BASE_PROJECT)
    st['amount']['value']['value'] = '5,500,000'
    _submit(st)
    assert sent == [{'총액 1': 5500000}] and env.edits == []


def test_submit_on_cancelled_project_is_refused(env):
    env.cur['p'] = dict(BASE_PROJECT, **{'수금 관련 특이사항': '공사 취소 (사유: 고객 변심)'})
    st = _state(BASE_PROJECT)
    st['content']['value']['value'] = '천장형 3대 설치'
    c = _submit(st)
    assert env.edits == [] and '취소된 공사' in c.ephemeral[0]


def _open(monkeypatch, project):
    monkeypatch.setattr(ps, 'get_project_records', lambda *a, **k: [project])
    monkeypatch.setattr(sb, '_load_active_constructors_flat', lambda: ['정건우'])
    c = _Client()
    sb._open_project_edit_modal(c, {'trigger_id': 't', 'actions': [{'value': 'G4300-JW'}],
                                    'channel': {'id': 'C1'}, 'message': {'ts': '1.1'}})
    return c.views[-1][1]


def test_open_cancelled_project_shows_notice(monkeypatch):
    v = _open(monkeypatch, dict(BASE_PROJECT, **{'수금 관련 특이사항': '공사 취소 (사유: x)'}))
    assert 'submit' not in v and '취소된 공사' in v['blocks'][0]['text']['text']


def test_open_stores_base(monkeypatch):
    v = _open(monkeypatch, BASE_PROJECT)
    meta = json.loads(v['private_metadata'])
    assert meta['base'] == sb._edit_form_base(BASE_PROJECT) and len(v['private_metadata']) < 3000


def test_perform_edit_refuses_cancelled(monkeypatch):
    monkeypatch.setattr(psa, '_load_project', lambda code: {'수금 관련 특이사항': '공사 취소 (사유: x)'})
    monkeypatch.setattr(psa, '_get_sheet_context', lambda: (_ for _ in ()).throw(AssertionError('시트 안 봐야 함')))
    assert psa.perform_edit('G1', {'공사 내용': 'x'}, '사유', 'JW') == {'ok': False, 'reason': 'already_cancelled'}


def test_uncancel_card_uses_resumed_values(monkeypatch):
    """재개 카드는 시트 강제 읽기 대신 캐시+재개 뒤 값(총액 복원)으로 그린다 (A6)."""
    import dashboard.services.project_slack_notifier as psn
    import dashboard.services.business_license_handler as blh
    seen, refresh = [], []
    stale = dict(BASE_PROJECT, **{'총액 1': 0, '공사 확정': '', '수금 관련 특이사항': '공사 취소 (사유: x)'})

    def _records(force_refresh=False):
        refresh.append(force_refresh)
        return [stale]
    monkeypatch.setattr(ps, 'get_project_records', _records)
    monkeypatch.setattr(psn, '_build_blocks', lambda latest, code, **k: seen.append(latest) or [])
    monkeypatch.setattr(psn, '_thread_permalink', lambda *a: '')
    monkeypatch.setattr(blh, 'verify_license_exists', lambda code: False)

    class _P:
        def chat_update(self, **k):
            pass

        def chat_postMessage(self, **k):
            pass
    monkeypatch.setattr(sb, '_project_client', lambda: _P())
    resumed = dict(BASE_PROJECT, **{'총액 1': 5000000, '공사 확정': '2026-11-16', '수금 관련 특이사항': ''})
    sb.apply_project_uncancel_to_slack('G4300-JW', resumed, 'JW', channel='C', ts='1', notify=False)
    assert refresh == [False]
    assert seen[0]['총액 1'] == 5000000 and seen[0]['수금 관련 특이사항'] == ''


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
