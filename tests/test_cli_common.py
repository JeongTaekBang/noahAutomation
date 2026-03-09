"""
cli_common 모듈 테스트
======================

validate_output_path, generate_output_filename 테스트
특히 Path Traversal 보안 검증에 초점
"""

from pathlib import Path
from datetime import datetime

import pytest

from po_generator.cli_common import (
    validate_output_path,
    generate_output_filename,
)


class TestValidateOutputPath:
    """validate_output_path 함수 테스트"""

    def test_valid_path_in_output_dir(self, tmp_path):
        """출력 디렉토리 내의 유효한 경로 허용"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        output_file = output_dir / "test.xlsx"

        assert validate_output_path(output_file, output_dir) is True

    def test_valid_nested_subdirectory(self, tmp_path):
        """중첩 서브디렉토리 허용"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        nested = output_dir / "2026" / "01"
        nested.mkdir(parents=True)
        output_file = nested / "test.xlsx"

        assert validate_output_path(output_file, output_dir) is True

    def test_rejects_path_traversal_with_parent(self, tmp_path):
        """.. 를 사용한 경로 탈출 거부"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        # 상위 디렉토리로 탈출 시도
        output_file = output_dir / ".." / "malicious.xlsx"

        assert validate_output_path(output_file, output_dir) is False

    def test_rejects_path_outside_output_dir(self, tmp_path):
        """출력 디렉토리 외부 경로 거부"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        other_dir = tmp_path / "other"
        other_dir.mkdir()
        output_file = other_dir / "test.xlsx"

        assert validate_output_path(output_file, output_dir) is False

    def test_rejects_substring_attack(self, tmp_path):
        """부분 문자열 공격 거부 (documents가 doc_files에 포함되는 케이스)"""
        # /home/user/documents 가 /home/user/doc_files/test.xlsx에 포함되는 공격
        output_dir = tmp_path / "documents"
        output_dir.mkdir()
        other_dir = tmp_path / "doc_files"
        other_dir.mkdir()
        output_file = other_dir / "test.xlsx"

        # doc_files는 documents가 아니므로 거부되어야 함
        assert validate_output_path(output_file, output_dir) is False

    def test_rejects_absolute_path_escape(self, tmp_path):
        """절대 경로를 사용한 탈출 시도 거부"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        # 완전히 다른 절대 경로
        if Path("/tmp").exists():
            output_file = Path("/tmp/malicious.xlsx")
        else:
            output_file = Path("C:/Windows/Temp/malicious.xlsx")

        assert validate_output_path(output_file, output_dir) is False

    def test_handles_symlink_traversal(self, tmp_path):
        """심볼릭 링크를 통한 탈출 시도 (플랫폼 지원 시)"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()
        other_dir = tmp_path / "other"
        other_dir.mkdir()

        # 심볼릭 링크 생성 시도 (Windows에서는 권한 필요할 수 있음)
        try:
            symlink = output_dir / "link"
            symlink.symlink_to(other_dir)
            output_file = symlink / "test.xlsx"

            # resolve()로 실제 경로를 확인하므로 거부되어야 함
            assert validate_output_path(output_file, output_dir) is False
        except OSError:
            # 심볼릭 링크 생성 실패 시 스킵
            pytest.skip("Symlink creation not supported")


class TestGenerateOutputFilename:
    """generate_output_filename 함수 테스트"""

    def test_generates_correct_format(self, tmp_path):
        """올바른 형식의 파일명 생성"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        result = generate_output_filename(
            prefix="PO",
            order_no="ND-0001",
            customer_name="TestCustomer",  # 공백 없는 이름 사용
            output_dir=output_dir,
        )

        # 파일명 형식: PO_주문번호_고객명_YYMMDD.xlsx
        today = datetime.now().strftime("%y%m%d")
        assert result.name.startswith("PO_ND-0001_")
        assert "TestCustomer" in result.name
        assert result.name.endswith(f"_{today}.xlsx")
        assert result.parent == output_dir

    def test_sanitizes_special_characters(self, tmp_path):
        """특수문자 제거/치환"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        result = generate_output_filename(
            prefix="TS",
            order_no="ND/0001:test",
            customer_name="Customer<name>",
            output_dir=output_dir,
        )

        # 파일 시스템에서 금지된 문자가 제거됨
        filename = result.name
        assert "/" not in filename
        assert ":" not in filename
        assert "<" not in filename
        assert ">" not in filename

    def test_different_prefixes(self, tmp_path):
        """다른 접두사로 파일명 생성"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        po_file = generate_output_filename("PO", "ND-0001", "Customer", output_dir)
        ts_file = generate_output_filename("TS", "DN-0001", "Customer", output_dir)
        pi_file = generate_output_filename("PI", "PI-0001", "Customer", output_dir)

        assert po_file.name.startswith("PO_")
        assert ts_file.name.startswith("TS_")
        assert pi_file.name.startswith("PI_")

    def test_korean_customer_name(self, tmp_path):
        """한글 고객명 처리"""
        output_dir = tmp_path / "output"
        output_dir.mkdir()

        result = generate_output_filename(
            prefix="PO",
            order_no="ND-0001",
            customer_name="한글고객명",
            output_dir=output_dir,
        )

        # 한글이 그대로 유지됨
        assert "한글고객명" in result.name
