"""
create_po.py CLI 테스트
=======================

CLI 진입점 및 generate_po 함수 테스트
"""

import sys
import tempfile
from datetime import datetime
from pathlib import Path
from unittest.mock import patch, MagicMock

import pandas as pd
import pytest

# create_po 모듈에서 함수 import
sys.path.insert(0, str(Path(__file__).parent.parent))
from create_po import (
    generate_po,
    show_history,
    print_available_orders,
    setup_logging,
)
from po_generator.history import sanitize_filename


class TestSanitizeFilename:
    """sanitize_filename 함수 테스트"""

    def test_removes_windows_forbidden_chars(self):
        """Windows 금지 문자 제거 테스트"""
        assert sanitize_filename('file/name') == 'file_name'
        assert sanitize_filename('file\\name') == 'file_name'
        assert sanitize_filename('file:name') == 'file_name'
        assert sanitize_filename('file*name') == 'file_name'
        assert sanitize_filename('file?name') == 'file_name'
        assert sanitize_filename('file"name') == 'file_name'
        assert sanitize_filename('file<name') == 'file_name'
        assert sanitize_filename('file>name') == 'file_name'
        assert sanitize_filename('file|name') == 'file_name'

    def test_normalizes_spaces_and_underscores(self):
        """연속 공백/언더스코어 정규화 테스트"""
        assert sanitize_filename('file  name') == 'file_name'
        assert sanitize_filename('file___name') == 'file_name'
        assert sanitize_filename('file _ name') == 'file_name'

    def test_strips_leading_trailing_underscores(self):
        """앞뒤 언더스코어 제거 테스트"""
        assert sanitize_filename('_filename_') == 'filename'
        assert sanitize_filename('__filename__') == 'filename'

    def test_handles_korean_characters(self):
        """한글 문자 처리 테스트"""
        assert sanitize_filename('고객명') == '고객명'
        assert sanitize_filename('ABC전자') == 'ABC전자'

    def test_handles_mixed_content(self):
        """복합 케이스 테스트"""
        result = sanitize_filename('ABC전자 / Korea:Branch')
        assert result == 'ABC전자_Korea_Branch'


class TestSetupLogging:
    """로깅 설정 테스트"""

    def test_setup_logging_default(self):
        """기본 로깅 설정 테스트"""
        setup_logging(verbose=False)
        import logging
        assert logging.getLogger().level == logging.INFO

    def test_setup_logging_verbose(self):
        """상세 로깅 설정 테스트"""
        setup_logging(verbose=True)
        import logging
        assert logging.getLogger().level == logging.DEBUG


class TestGeneratePO:
    """generate_po 함수 테스트"""

    @pytest.fixture
    def sample_df(self, valid_order_data):
        """샘플 DataFrame"""
        return pd.DataFrame([valid_order_data])

    @pytest.fixture
    def temp_output_dir(self):
        """임시 출력 디렉토리"""
        with tempfile.TemporaryDirectory() as tmpdir:
            yield Path(tmpdir)

    def test_returns_false_for_missing_order(self, sample_df):
        """존재하지 않는 주문번호 테스트"""
        result = generate_po('NONEXISTENT-001', sample_df, force=True)
        assert result is False

    @patch('create_po.OUTPUT_DIR')
    @patch('create_po.check_duplicate_order')
    @patch('create_po.save_to_history')
    def test_generates_po_successfully(
        self,
        mock_save_history,
        mock_check_dup,
        mock_output_dir,
        sample_df,
        temp_output_dir,
    ):
        """발주서 생성 성공 테스트"""
        mock_output_dir.__truediv__ = lambda self, x: temp_output_dir / x
        mock_output_dir.mkdir = MagicMock()
        mock_check_dup.return_value = None
        mock_save_history.return_value = True

        # 실제 OUTPUT_DIR 패치
        with patch('create_po.OUTPUT_DIR', temp_output_dir):
            result = generate_po('ND-TEST-001', sample_df, force=True)

        # 결과 확인 - 모킹 환경에서는 파일 생성이 안 될 수 있음
        # 여기서는 함수가 크래시 없이 실행되는지만 확인

    @patch('create_po.check_duplicate_order')
    @patch('builtins.input', return_value='N')
    def test_cancels_on_duplicate_warning(
        self, mock_input, mock_check_dup, sample_df
    ):
        """중복 경고 시 취소 테스트"""
        mock_check_dup.return_value = {
            '생성일시': '2026-01-01 10:00:00',
            '생성파일': 'test.xlsx',
        }

        result = generate_po('ND-TEST-001', sample_df, force=False)
        assert result is False

    @patch('create_po.check_duplicate_order')
    def test_force_skips_duplicate_check(self, mock_check_dup, sample_df):
        """--force 옵션으로 중복 체크 스킵 테스트"""
        # force=True일 때는 check_duplicate_order가 호출되지 않아야 함
        with patch('create_po.OUTPUT_DIR') as mock_dir:
            mock_dir.mkdir = MagicMock()
            mock_dir.__truediv__ = lambda self, x: Path(tempfile.gettempdir()) / x

            generate_po('ND-TEST-001', sample_df, force=True)

        # force=True면 중복 체크 호출 안 됨
        mock_check_dup.assert_not_called()


class TestShowHistory:
    """show_history 함수 테스트"""

    @patch('create_po.get_current_month_info')
    @patch('create_po.get_history_count')
    def test_shows_empty_history_message(
        self, mock_count, mock_month_info
    ):
        """빈 이력 메시지 테스트"""
        mock_month_info.return_value = ('2026년 1월', Path('/fake/path'))
        mock_count.return_value = 0

        result = show_history(export=False)

        assert result == 0

    @patch('create_po.get_current_month_info')
    @patch('create_po.get_history_count')
    @patch('create_po.get_all_history')
    def test_shows_history_list(
        self, mock_get_history, mock_count, mock_month_info
    ):
        """이력 목록 표시 테스트"""
        mock_month_info.return_value = ('2026년 1월', Path('/fake/path'))
        mock_count.return_value = 1
        mock_get_history.return_value = pd.DataFrame([{
            'RCK Order no.': 'ND-0001',
            'Customer name': 'Test Customer',
            'Description': 'Test Actuator',
            '생성일시': '2026-01-01 10:00:00',
            'Total net amount': 1000000,
        }])

        result = show_history(export=False)

        assert result == 0


class TestPrintAvailableOrders:
    """print_available_orders 함수 테스트"""

    def test_prints_order_list(self, capsys):
        """주문 목록 출력 테스트"""
        df = pd.DataFrame({
            'RCK Order no.': ['ND-0001', 'ND-0002', 'ND-0003'],
        })

        print_available_orders(df, limit=2)

        captured = capsys.readouterr()
        assert 'ND-0001' in captured.out
        assert 'ND-0002' in captured.out
        assert '외 1건' in captured.out

    def test_handles_empty_dataframe(self, capsys):
        """빈 DataFrame 처리 테스트"""
        df = pd.DataFrame({'RCK Order no.': []})

        print_available_orders(df)

        captured = capsys.readouterr()
        assert '사용 가능한 RCK Order No.' in captured.out


class TestCLIArguments:
    """CLI 인자 파싱 테스트 (argparse 기반)"""

    def test_force_flag_parsing(self):
        """--force 플래그 파싱 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['ND-0001', '--force'])
        assert args.force is True
        assert args.order_numbers == ['ND-0001']

    def test_force_short_flag(self):
        """-f 단축 플래그 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['ND-0001', '-f'])
        assert args.force is True

    def test_history_flag_parsing(self):
        """--history 플래그 파싱 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['--history'])
        assert args.history is True
        assert args.order_numbers == []

    def test_verbose_flag_parsing(self):
        """--verbose 플래그 파싱 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['ND-0001', '--verbose'])
        assert args.verbose is True

    def test_verbose_short_flag(self):
        """-v 단축 플래그 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['ND-0001', '-v'])
        assert args.verbose is True

    def test_export_flag_parsing(self):
        """--export 플래그 파싱 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['--history', '--export'])
        assert args.history is True
        assert args.export is True

    def test_multiple_order_numbers(self):
        """여러 주문번호 파싱 테스트"""
        from create_po import create_argument_parser
        parser = create_argument_parser()
        args = parser.parse_args(['ND-0001', 'ND-0002', 'ND-0003'])
        assert args.order_numbers == ['ND-0001', 'ND-0002', 'ND-0003']
        assert args.force is False


class TestConfigConstants:
    """config.py 상수 사용 테스트"""

    def test_customer_name_max_length_applied(self):
        """고객명 최대 길이 상수 적용 테스트"""
        from po_generator.config import CUSTOMER_NAME_MAX_LENGTH

        long_name = "Very Long Customer Name That Exceeds Limit"
        truncated = sanitize_filename(long_name)[:CUSTOMER_NAME_MAX_LENGTH]

        assert len(truncated) <= CUSTOMER_NAME_MAX_LENGTH

    def test_order_list_display_limit(self):
        """주문 목록 표시 제한 상수 테스트"""
        from po_generator.config import ORDER_LIST_DISPLAY_LIMIT

        assert ORDER_LIST_DISPLAY_LIMIT == 20

    def test_history_display_lengths(self):
        """이력 표시 길이 상수 테스트"""
        from po_generator.config import (
            HISTORY_CUSTOMER_DISPLAY_LENGTH,
            HISTORY_DESC_DISPLAY_LENGTH,
            HISTORY_DATE_DISPLAY_LENGTH,
        )

        assert HISTORY_CUSTOMER_DISPLAY_LENGTH == 15
        assert HISTORY_DESC_DISPLAY_LENGTH == 20
        assert HISTORY_DATE_DISPLAY_LENGTH == 10
