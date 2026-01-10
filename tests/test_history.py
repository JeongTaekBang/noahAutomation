"""
history 모듈 테스트
"""

from pathlib import Path
from datetime import datetime
from unittest.mock import patch

import pandas as pd
import pytest

from po_generator.history import (
    check_duplicate_order,
    save_to_history,
    get_history_count,
    clear_history,
)
from po_generator.config import HISTORY_FILE


class TestCheckDuplicateOrder:
    """중복 발주 체크 테스트"""

    def test_no_history_file(self, tmp_path: Path):
        """이력 파일이 없는 경우"""
        with patch('po_generator.history.HISTORY_FILE', tmp_path / 'nonexistent.xlsx'):
            result = check_duplicate_order('ND-0001')
            assert result is None

    def test_order_not_in_history(self, tmp_path: Path):
        """이력에 없는 주문번호"""
        history_file = tmp_path / 'history.xlsx'
        df = pd.DataFrame([{
            'RCK Order no.': 'ND-0001',
            '생성일시': '2024-01-01 10:00:00',
            '생성파일': 'test.xlsx',
        }])
        df.to_excel(history_file, index=False)

        with patch('po_generator.history.HISTORY_FILE', history_file):
            result = check_duplicate_order('ND-9999')
            assert result is None

    def test_order_in_history(self, tmp_path: Path):
        """이력에 있는 주문번호"""
        history_file = tmp_path / 'history.xlsx'
        df = pd.DataFrame([{
            'RCK Order no.': 'ND-0001',
            '생성일시': '2024-01-01 10:00:00',
            '생성파일': 'test.xlsx',
        }])
        df.to_excel(history_file, index=False)

        with patch('po_generator.history.HISTORY_FILE', history_file):
            result = check_duplicate_order('ND-0001')
            assert result is not None
            assert result['생성일시'] == '2024-01-01 10:00:00'
            assert result['생성파일'] == 'test.xlsx'


class TestSaveToHistory:
    """이력 저장 테스트"""

    def test_save_new_history(self, tmp_path: Path, valid_order_data: pd.Series):
        """새 이력 저장"""
        history_file = tmp_path / 'history.xlsx'
        output_file = tmp_path / 'output.xlsx'

        with patch('po_generator.history.HISTORY_FILE', history_file):
            result = save_to_history(valid_order_data, output_file)
            assert result is True
            assert history_file.exists()

            df = pd.read_excel(history_file)
            assert len(df) == 1
            assert df.iloc[0]['RCK Order no.'] == 'ND-TEST-001'

    def test_append_to_history(self, tmp_path: Path, valid_order_data: pd.Series):
        """기존 이력에 추가"""
        history_file = tmp_path / 'history.xlsx'
        output_file = tmp_path / 'output.xlsx'

        # 기존 이력 생성
        existing = pd.DataFrame([{'RCK Order no.': 'ND-0001', '생성일시': '2024-01-01'}])
        existing.to_excel(history_file, index=False)

        with patch('po_generator.history.HISTORY_FILE', history_file):
            save_to_history(valid_order_data, output_file)

            df = pd.read_excel(history_file)
            assert len(df) == 2


class TestGetHistoryCount:
    """이력 건수 조회 테스트"""

    def test_no_history_file(self, tmp_path: Path):
        """이력 파일이 없는 경우"""
        with patch('po_generator.history.HISTORY_FILE', tmp_path / 'nonexistent.xlsx'):
            count = get_history_count()
            assert count == 0

    def test_with_history(self, tmp_path: Path):
        """이력이 있는 경우"""
        history_file = tmp_path / 'history.xlsx'
        df = pd.DataFrame([
            {'RCK Order no.': 'ND-0001'},
            {'RCK Order no.': 'ND-0002'},
            {'RCK Order no.': 'ND-0003'},
        ])
        df.to_excel(history_file, index=False)

        with patch('po_generator.history.HISTORY_FILE', history_file):
            count = get_history_count()
            assert count == 3


class TestClearHistory:
    """이력 초기화 테스트"""

    def test_clear_existing_history(self, tmp_path: Path):
        """기존 이력 삭제"""
        history_file = tmp_path / 'history.xlsx'
        history_file.write_text('dummy')

        with patch('po_generator.history.HISTORY_FILE', history_file):
            result = clear_history()
            assert result is True
            assert not history_file.exists()

    def test_clear_nonexistent_history(self, tmp_path: Path):
        """없는 이력 삭제 시도"""
        with patch('po_generator.history.HISTORY_FILE', tmp_path / 'nonexistent.xlsx'):
            result = clear_history()
            assert result is True
