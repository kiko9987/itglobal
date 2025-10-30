"""
SSL 에러 복구 로직 통합 테스트

전문가 권장사항:
- 사업자/담당자 + 메모 동시 저장 시나리오
- SSL 에러 발생 시뮬레이션
- 재시도 성공 여부 검증
"""
import pytest
import ssl
from unittest.mock import Mock, patch, MagicMock
from dashboard.utils.google_sheets import GoogleSheetsManager, SSLRetryFailedError


class TestSSLRecovery:
    """SSL 에러 복구 로직 테스트"""

    @pytest.fixture
    def mock_service(self):
        """Mock Google Sheets 서비스"""
        service = MagicMock()
        return service

    @pytest.fixture
    def sheets_manager(self, mock_service):
        """GoogleSheetsManager 인스턴스 (mock 서비스 사용)"""
        with patch('dashboard.utils.google_sheets.GoogleSheetsManager._authenticate'):
            manager = GoogleSheetsManager()
            manager.service = mock_service
            return manager

    def test_ssl_error_triggers_reset_service(self, sheets_manager, mock_service):
        """
        SSL 에러 발생 시 reset_service() 호출 확인

        시나리오:
        1. 첫 번째 시도: SSL 에러 발생
        2. reset_service() 호출됨
        3. 두 번째 시도: 성공
        """
        # Mock request factory
        request_mock = MagicMock()

        # 첫 시도: SSL 에러, 두 번째 시도: 성공
        request_mock.execute.side_effect = [
            ssl.SSLError("TLS session timeout"),
            {"success": True, "data": "test"}
        ]

        request_factory = MagicMock(return_value=request_mock)

        # reset_service를 spy
        with patch.object(sheets_manager, 'reset_service', return_value=True) as reset_spy:
            result = sheets_manager._execute_with_retry(request_factory, "test_operation")

            # reset_service가 호출되었는지 확인
            reset_spy.assert_called_once()

            # request_factory가 2번 호출되었는지 확인 (재시도)
            assert request_factory.call_count == 2

            # 최종 결과 확인
            assert result["success"] is True

    def test_ssl_error_max_retries_raises_custom_exception(self, sheets_manager):
        """
        SSL 에러가 최대 재시도 후에도 실패하면 SSLRetryFailedError 발생

        시나리오:
        1. 3회 모두 SSL 에러 발생
        2. SSLRetryFailedError 발생
        """
        request_mock = MagicMock()
        request_mock.execute.side_effect = ssl.SSLError("Persistent TLS error")

        request_factory = MagicMock(return_value=request_mock)

        with pytest.raises(SSLRetryFailedError) as exc_info:
            sheets_manager._execute_with_retry(request_factory, "test_operation")

        # 에러 메시지 확인
        assert "SSL/TLS 연결 오류가 지속됩니다" in str(exc_info.value)

        # request_factory가 MAX_RETRIES번 호출되었는지 확인
        assert request_factory.call_count == sheets_manager.MAX_RETRIES

    def test_request_factory_creates_new_request_each_attempt(self, sheets_manager):
        """
        request_factory가 매 시도마다 새 요청을 생성하는지 확인

        시나리오:
        1. 첫 번째 시도 실패 (일반 에러)
        2. 두 번째 시도 성공
        3. request_factory가 2번 호출되어 새 요청 객체 생성
        """
        request1 = MagicMock()
        request2 = MagicMock()

        request1.execute.side_effect = Exception("Temporary error")
        request2.execute.return_value = {"success": True}

        # 매번 다른 객체 반환
        request_factory = MagicMock(side_effect=[request1, request2])

        result = sheets_manager._execute_with_retry(request_factory, "test_operation")

        # request_factory가 2번 호출됨
        assert request_factory.call_count == 2

        # 두 개의 다른 request 객체가 사용됨
        assert request1.execute.called
        assert request2.execute.called

        assert result["success"] is True

    def test_update_project_with_ssl_recovery_simulation(self, sheets_manager, mock_service):
        """
        프로젝트 업데이트 중 SSL 에러 복구 시뮬레이션

        시나리오: 사업자/담당자 변경 저장 시 SSL 에러 발생 → 복구
        """
        # Mock spreadsheets API chain
        mock_values = MagicMock()
        mock_update = MagicMock()

        # 첫 시도: SSL 에러, 두 번째 시도: 성공
        execute_mock = MagicMock(side_effect=[
            ssl.SSLError("Session expired"),
            {"updatedCells": 5}
        ])

        mock_update.execute = execute_mock
        mock_values.update.return_value = mock_update
        mock_service.spreadsheets.return_value.values.return_value = mock_values

        # reset_service mock
        with patch.object(sheets_manager, 'reset_service', return_value=True):
            # update_row 호출
            result = sheets_manager.update_row(
                sheet_id="test_sheet",
                row_number=100,
                values=["G2830-YG", "글로벌그룹", "강민석"],
                range_name="공사 현황의 사본!A{row}:AM{row}"
            )

            # 성공 확인
            assert result["updatedCells"] == 5

            # update가 2번 호출됨 (재시도)
            assert mock_values.update.call_count == 2


class TestProjectUpdateWithMemo:
    """프로젝트 업데이트 + 메모 동시 저장 통합 테스트"""

    def test_concurrent_project_and_memo_update(self):
        """
        사업자/담당자 변경 + 메모 추가 동시 시나리오

        전문가 권장: 이 플로우에 대한 통합 테스트 추가

        NOTE: 실제 API 호출 없이 로직만 검증
        """
        # TODO: Flask 테스트 클라이언트로 실제 API 엔드포인트 테스트
        # 1. PUT /api/projects/{code} - 사업자/담당자 변경
        # 2. POST /api/projects/{code}/memos - 메모 추가
        # 3. 두 작업이 트랜잭션처럼 성공/실패하는지 확인
        pass


class TestSSLMonitoring:
    """SSL 에러 모니터링 로직 테스트"""

    @pytest.fixture
    def mock_service(self):
        """Mock Google Sheets 서비스"""
        service = MagicMock()
        return service

    @pytest.fixture
    def sheets_manager(self, mock_service):
        """GoogleSheetsManager 인스턴스 (mock 서비스 사용)"""
        with patch('dashboard.utils.google_sheets.GoogleSheetsManager._authenticate'):
            manager = GoogleSheetsManager()
            manager.service = mock_service
            return manager

    def test_ssl_error_logging(self, sheets_manager, caplog):
        """
        SSL 에러 발생 시 적절한 로그가 남는지 확인

        전문가 권장: 재시도 실패 로그 모니터링
        """
        request_mock = MagicMock()
        request_mock.execute.side_effect = ssl.SSLError("Test SSL error")

        request_factory = MagicMock(return_value=request_mock)

        with pytest.raises(SSLRetryFailedError):
            sheets_manager._execute_with_retry(request_factory, "test_op")

        # 로그 확인
        assert any("[SSL_ERROR]" in record.message for record in caplog.records)
        assert any("최대 재시도 초과" in record.message for record in caplog.records)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
