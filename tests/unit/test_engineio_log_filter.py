"""engineio.server 정상 소음 거름 (2026-11-10 Sentry 'Session is disconnected')."""
import logging

from dashboard import _engineio_log_filter


def _rec(msg, level=logging.ERROR):
    return logging.LogRecord('engineio.server', level, __file__, 1, msg, None, None)


def test_session_disconnected_lowered_to_warning():
    r = _rec("'Session is disconnected' DRFj5uU9MYPg4Mf3AACk (further occurrences of this error "
             "will be logged with level INFO)")
    assert _engineio_log_filter(r) is True
    assert r.levelno == logging.WARNING and r.levelname == 'WARNING'


def test_session_disconnected_info_untouched():
    r = _rec("'Session is disconnected' abc", logging.INFO)
    assert _engineio_log_filter(r) is True
    assert r.levelno == logging.INFO


def test_invalid_session_dropped():
    assert _engineio_log_filter(_rec('Invalid session abc')) is False


def test_other_errors_kept_as_error():
    r = _rec('Unexpected error in packet')
    assert _engineio_log_filter(r) is True
    assert r.levelno == logging.ERROR


def test_logger_end_to_end_handler_sees_warning():
    seen = []

    class H(logging.Handler):
        def emit(self, record):
            seen.append(record.levelno)

    lg = logging.getLogger('engineio.server.test_filter')
    lg.propagate = False
    lg.addFilter(_engineio_log_filter)
    h = H(); lg.addHandler(h)
    try:
        lg.error("'Session is disconnected' xyz")
        lg.error('Invalid session xyz')
    finally:
        lg.removeHandler(h)
    assert seen == [logging.WARNING]
