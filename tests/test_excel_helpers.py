"""
excel_helpers 모듈 테스트
=========================

XlConstants, prepare_template, cleanup_temp_file, xlwings_app_context 테스트
"""

from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from po_generator.excel_helpers import (
    XlConstants,
    prepare_template,
    cleanup_temp_file,
    xlwings_app_context,
)


class TestXlConstants:
    """XlConstants 클래스 테스트"""

    def test_shift_constants(self):
        """Shift 방향 상수값 검증"""
        assert XlConstants.xlShiftUp == -4162
        assert XlConstants.xlShiftDown == -4121

    def test_border_position_constants(self):
        """테두리 위치 상수값 검증"""
        assert XlConstants.xlEdgeLeft == 7
        assert XlConstants.xlEdgeTop == 8
        assert XlConstants.xlEdgeBottom == 9
        assert XlConstants.xlEdgeRight == 10
        assert XlConstants.xlInsideVertical == 11
        assert XlConstants.xlInsideHorizontal == 12

    def test_border_style_constants(self):
        """테두리 스타일 상수값 검증"""
        assert XlConstants.xlContinuous == 1
        assert XlConstants.xlThin == 2
        assert XlConstants.xlMedium == -4138


class TestPrepareTemplate:
    """prepare_template 함수 테스트"""

    def test_copies_template_to_temp(self, tmp_path):
        """템플릿을 임시 폴더로 복사"""
        # 테스트용 템플릿 생성
        template_path = tmp_path / "test_template.xlsx"
        template_path.write_text("dummy content")

        temp_template, temp_output = prepare_template(template_path, "test")

        # 임시 템플릿 파일 생성 확인
        assert temp_template.exists()
        assert "test_template_" in str(temp_template)
        assert temp_template.suffix == ".xlsx"

        # 출력 파일 경로 확인 (아직 생성되지 않음)
        assert "test_output_" in str(temp_output)
        assert temp_output.suffix == ".xlsx"

        # 정리
        temp_template.unlink()

    def test_raises_file_not_found(self, tmp_path):
        """템플릿 파일이 없으면 FileNotFoundError"""
        non_existent = tmp_path / "없는파일.xlsx"

        with pytest.raises(FileNotFoundError) as exc_info:
            prepare_template(non_existent, "test")

        assert "템플릿 파일이 없습니다" in str(exc_info.value)

    def test_prefix_in_filename(self, tmp_path):
        """접두사가 파일명에 포함됨"""
        template_path = tmp_path / "test_template.xlsx"
        template_path.write_text("dummy content")

        temp_template, temp_output = prepare_template(template_path, "my_prefix")

        assert "my_prefix_template_" in str(temp_template)
        assert "my_prefix_output_" in str(temp_output)

        # 정리
        temp_template.unlink()


class TestCleanupTempFile:
    """cleanup_temp_file 함수 테스트"""

    def test_deletes_existing_file(self, tmp_path):
        """존재하는 파일 삭제"""
        temp_file = tmp_path / "temp_file.xlsx"
        temp_file.write_text("dummy content")

        assert temp_file.exists()
        cleanup_temp_file(temp_file)
        assert not temp_file.exists()

    def test_handles_non_existent_file(self, tmp_path):
        """존재하지 않는 파일은 조용히 무시"""
        non_existent = tmp_path / "없는파일.xlsx"

        # 에러 없이 실행되어야 함
        cleanup_temp_file(non_existent)

    def test_handles_permission_error(self, tmp_path):
        """삭제 실패 시 경고만 로깅 (에러 발생 안함)"""
        temp_file = tmp_path / "temp_file.xlsx"
        temp_file.write_text("dummy content")

        with patch.object(Path, 'unlink', side_effect=PermissionError("Access denied")):
            # 에러 없이 실행되어야 함 (경고 로깅만)
            cleanup_temp_file(temp_file)


class TestXlwingsAppContext:
    """xlwings_app_context 컨텍스트 매니저 테스트"""

    def test_creates_and_closes_app(self):
        """App 생성 및 종료"""
        mock_app = MagicMock()
        mock_app.books = []

        with patch('po_generator.excel_helpers.xw.App', return_value=mock_app):
            with xlwings_app_context() as app:
                assert app is mock_app
                assert mock_app.display_alerts is False
                assert mock_app.screen_updating is False

            # 컨텍스트 종료 후 quit 호출 확인
            mock_app.quit.assert_called_once()

    def test_closes_workbooks_on_exit(self):
        """종료 시 모든 워크북 닫기"""
        mock_wb1 = MagicMock()
        mock_wb2 = MagicMock()
        mock_app = MagicMock()
        mock_app.books = [mock_wb1, mock_wb2]

        with patch('po_generator.excel_helpers.xw.App', return_value=mock_app):
            with xlwings_app_context():
                pass

            # 모든 워크북 close 호출 확인
            mock_wb1.close.assert_called_once()
            mock_wb2.close.assert_called_once()

    def test_handles_exception_in_context(self):
        """컨텍스트 내 예외 발생 시에도 정리"""
        mock_app = MagicMock()
        mock_app.books = []

        with patch('po_generator.excel_helpers.xw.App', return_value=mock_app):
            with pytest.raises(ValueError):
                with xlwings_app_context():
                    raise ValueError("Test error")

            # 예외가 발생해도 quit 호출 확인
            mock_app.quit.assert_called_once()

    def test_custom_options(self):
        """사용자 정의 옵션"""
        mock_app = MagicMock()
        mock_app.books = []

        with patch('po_generator.excel_helpers.xw.App', return_value=mock_app) as mock_app_class:
            with xlwings_app_context(
                visible=True,
                display_alerts=True,
                screen_updating=True,
            ) as app:
                # visible은 App 생성자에 전달됨
                mock_app_class.assert_called_once_with(visible=True)
                # display_alerts와 screen_updating은 속성으로 설정됨
                assert mock_app.display_alerts is True
                assert mock_app.screen_updating is True

    def test_handles_quit_failure(self):
        """quit 실패 시에도 에러 없이 종료"""
        mock_app = MagicMock()
        mock_app.books = []
        mock_app.quit.side_effect = Exception("COM error")

        with patch('po_generator.excel_helpers.xw.App', return_value=mock_app):
            # 에러 없이 종료되어야 함
            with xlwings_app_context():
                pass

    def test_handles_workbook_close_failure(self):
        """워크북 close 실패 시에도 에러 없이 진행"""
        mock_wb = MagicMock()
        mock_wb.close.side_effect = Exception("COM error")
        mock_app = MagicMock()
        mock_app.books = [mock_wb]

        with patch('po_generator.excel_helpers.xw.App', return_value=mock_app):
            # 에러 없이 종료되어야 함
            with xlwings_app_context():
                pass

            # quit은 여전히 호출되어야 함
            mock_app.quit.assert_called_once()
