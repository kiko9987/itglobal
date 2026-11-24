"""get_logger 감싸개가 extra={'force_alert': True} 를 풀어 넣는지 — 예전엔 record.extra 로 한 겹 더 싸여
에러 슬랙 알림의 '꼭 알림'이 듣지 않았다(11-24 조용한 실패 점검)."""
import logging

from dashboard.utils.logging_config import get_logger


class _Grab(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.DEBUG)
        self.records = []

    def emit(self, record):
        self.records.append(record)


def test_extra_dict_is_flattened():
    log = get_logger('test.silent_fail.round4')
    h = _Grab()
    log.logger.addHandler(h)
    try:
        log.error('꼭 알림', extra={'force_alert': True})
        log.error('키워드로', force_alert=True)
        log.warning('None 은 무시', extra=None)
        log.info('보통')
    finally:
        log.logger.removeHandler(h)
    a, b, c, d = h.records
    assert getattr(a, 'force_alert', False) is True and not hasattr(a, 'extra')
    assert getattr(b, 'force_alert', False) is True
    assert not getattr(c, 'force_alert', False) and not getattr(d, 'force_alert', False)
