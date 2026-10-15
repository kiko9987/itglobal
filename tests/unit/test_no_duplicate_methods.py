# -*- coding: utf-8 -*-
"""클래스 안 같은 이름 메서드 중복 금지 (2026-10-16 사고).

GoogleSheetsManager 에 get_row_values 가 두 번 정의돼 뒤엣것(수금 속도 개선용, 다른 시그니처)이
앞엣것(PM 프로젝트 수정 PUT 이 쓰는 것)을 조용히 덮어써 PM 저장이 전부 500. 파이썬은 경고 없이
뒤엣것을 쓰므로 정적으로 막는다. (property setter·overload 데코레이터는 예외)
"""
import ast
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[2] / 'dashboard'
_ALLOWED_DECOS = {'setter', 'getter', 'deleter', 'overload', 'register'}


def _deco_names(fn):
    out = set()
    for d in fn.decorator_list:
        node = d.func if isinstance(d, ast.Call) else d
        if isinstance(node, ast.Attribute):
            out.add(node.attr)
        elif isinstance(node, ast.Name):
            out.add(node.id)
    return out


def test_no_duplicate_method_names_in_classes():
    dups = []
    for path in ROOT.rglob('*.py'):
        try:
            tree = ast.parse(path.read_text(encoding='utf-8'))
        except (SyntaxError, UnicodeDecodeError):
            continue
        for cls in ast.walk(tree):
            if not isinstance(cls, ast.ClassDef):
                continue
            seen = {}
            for fn in cls.body:
                if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                if _deco_names(fn) & _ALLOWED_DECOS:
                    continue
                if fn.name in seen:
                    dups.append(f'{path.relative_to(ROOT.parent)}:{fn.lineno} '
                                f'{cls.name}.{fn.name} (먼저 {seen[fn.name]}줄)')
                else:
                    seen[fn.name] = fn.lineno
    assert not dups, '클래스 안 메서드 이름 중복(뒤엣것이 앞엣것을 덮어씀):\n' + '\n'.join(dups)
