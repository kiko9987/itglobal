"""프로젝트 번호 — 한 번 등록된 적 있는 번호는 다시 주지 않는다(2026-11-26 R4286 재사용). 임시 DB만 씀."""
import sys

sys.path.insert(0, '.')

from dashboard.utils.user_database import UserDatabase


def _db(tmp_path):
    return UserDatabase(db_path=str(tmp_path / 'users.db'))


def _created(db, code):
    with db._get_connection() as c:
        c.execute("INSERT INTO audit_logs (user_email, action, details, project_code, field_name) "
                  "VALUES ('t@x', 'CREATE_PROJECT', '새 프로젝트 등록', ?, '전체')", (code,))
        c.commit()


def test_normal_sequence(tmp_path):
    db = _db(tmp_path)
    assert db.get_next_project_number('GLOBAL', current_max_from_sheet=4284) == 4285
    assert db.get_next_project_number('GLOBAL', current_max_from_sheet=4284) == 4286


def test_skips_number_used_before_after_counter_lowered(tmp_path):
    """R4286-JSH 를 지우고 카운터를 4285 로 낮춰도 4286 은 다시 안 나간다 → 4287."""
    db = _db(tmp_path)
    assert db.get_next_project_number('GLOBAL', current_max_from_sheet=4285) == 4286
    _created(db, 'R4286-JSH')
    with db._get_connection() as c:      # 중복 정리하며 카운터를 낮춘 상황
        c.execute("UPDATE project_code_sequences SET current_number = 4285 WHERE pattern = 'GLOBAL'")
        c.commit()
    assert db.get_next_project_number('GLOBAL', current_max_from_sheet=4285) == 4287
    assert db.get_next_project_number('GLOBAL', current_max_from_sheet=4285) == 4288   # 카운터도 따라 올라감


def test_other_numbers_not_confused(tmp_path):
    db = _db(tmp_path)
    _created(db, 'G14286-AA')      # 다섯 자리·다른 꼴은 4286 으로 보지 않음
    _created(db, 'L-04286')
    assert db.get_next_project_number('GLOBAL', current_max_from_sheet=4285) == 4286
