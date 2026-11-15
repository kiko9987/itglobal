# -*- coding: utf-8 -*-
"""5차 보안 점검 #9 — 플렉스 보관함은 작성자 '이름'으로만 주인을 가린다.

인사카드가 없는 옛 작성자는 동명이인 검사(hr.same_name_people)에 안 잡혀, 이름이 같은 지금 직원이 자기 근무 기간
안의 그 사람 문서를 보게 된다. 11-15 운영 DB(읽기 전용)로 센 결과: 문서 1,305건·작성자 17명 모두 인사카드 있음,
같은 이름 카드 2쌍은 생년월일이 같은 재입사. 그래서 화면 로직은 그대로 두고, 가져오기에서 막는다."""
import json
import sqlite3
from contextlib import contextmanager

import pytest

from dashboard.services import gw_approval as ga
from dashboard.services import hr

CARDS = {'홍길동': [{'id': 1, 'name': '홍길동', 'birth_date': '1990-01-01', 'hire_date': '2020-01-01'}]}


@pytest.fixture
def env(monkeypatch, tmp_path):
    conn = sqlite3.connect(':memory:')
    conn.execute(ga._ARCHIVE_DDL)

    @contextmanager
    def fake():
        yield conn
        conn.commit()
    monkeypatch.setattr(ga, '_aconn', fake)
    monkeypatch.setattr(hr, 'same_name_people', lambda n: list(CARDS.get((n or '').strip(), [])))

    def export(writers):
        docs = [{'c': f'F-{i}', 't': '지출결의', 's': 'DONE', 'w': w, 'at': '2025-06-01T10:00', 'f': {}}
                for i, w in enumerate(writers, 1)]
        p = tmp_path / 'flex_workflow_export_test.json'
        p.write_text(json.dumps({'exported_at': '2026-11-15', 'docs': docs}, ensure_ascii=False), encoding='utf-8')
        return str(p)

    def count():
        return conn.execute('SELECT COUNT(*) FROM flex_approval_archive').fetchone()[0]
    return export, count


def test_all_writers_have_cards_imports(env):
    export, count = env
    out = ga.import_flex_export(export(['홍길동', '홍길동']))
    assert out['no_card_writers'] == [] and not out['blocked'] and count() == 2


def test_writer_without_card_blocks_import(env):
    export, count = env
    out = ga.import_flex_export(export(['홍길동', '김옛날']))
    assert out['no_card_writers'] == ['김옛날'] and out['blocked'] and count() == 0     # 아무것도 안 씀


def test_preview_reports_but_is_not_blocked(env):
    export, count = env
    out = ga.import_flex_export(export(['김옛날']), apply=False)
    assert out['no_card_writers'] == ['김옛날'] and not out['blocked'] and count() == 0


def test_allow_no_card_after_check(env):
    export, count = env
    out = ga.import_flex_export(export(['홍길동', '김옛날']), allow_no_card=True)
    assert out['no_card_writers'] == ['김옛날'] and not out['blocked'] and count() == 2


def test_same_birth_namesake_cards_are_one_person(monkeypatch):
    """생년월일까지 같은 카드 = 재입사로 본다(운영 동명 카드 2쌍이 모두 이 경우). 다른 생년월일이면 그 기간은 숨김."""
    me = {'id': 1, 'name': '홍길동', 'birth_date': '1990-01-01', 'hire_date': '2025-03-01'}
    old = {'id': 2, 'name': '홍길동', 'birth_date': '1990-01-01', 'hire_date': '2020-01-01', 'resign_date': '2021-12-31'}
    other = {'id': 3, 'name': '홍길동', 'birth_date': '1985-05-05', 'hire_date': '2022-01-01', 'resign_date': '2022-12-31'}
    monkeypatch.setattr(hr, 'same_name_people', lambda n: [me, old, other])
    w = ga._archive_windows(me)
    assert w == {'mine': [('2025-03-01', '9999-12-31'), ('2020-01-01', '2021-12-31')],
                 'others': [('2022-01-01', '2022-12-31')]}
    assert ga._in_windows('2021-06-01 10:00', w) and not ga._in_windows('2022-06-01 10:00', w)
