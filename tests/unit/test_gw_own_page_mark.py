"""그룹웨어 전용 화면(templates/gw/*.html)이면 gw_own_page — 공용 틀이 PM 공사 화면용 시공사 목록 미리 불러오기를
건너뛴다(2026-11-25 부장님 '느려'). PM 화면은 그대로 불러온다."""
from flask import Flask, render_template
from jinja2 import DictLoader


def test_gw_template_marked_and_base_skips_preload():
    from dashboard.blueprints.groupware import groupware_bp
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(groupware_bp)
    app.jinja_loader = DictLoader({
        'base.html': "{% if not gw_own_page %}PRELOAD{% endif %}|{% block content %}{% endblock %}",
        'gw/page.html': "{% extends 'base.html' %}{% block content %}GW{% endblock %}",
        'pm.html': "{% extends 'base.html' %}{% block content %}PM{% endblock %}",
    })
    with app.test_request_context('/'):
        assert render_template('gw/page.html') == '|GW'
        assert render_template('pm.html') == 'PRELOAD|PM'


def test_modern_base_wraps_constructor_preload():
    src = open('dashboard/templates/modern_base.html', encoding='utf-8').read()
    i = src.index('{% if not gw_own_page %}')
    j = src.index('{% endif %}', i)
    assert 'window.refreshConstructorCache()' in src[i:j]
