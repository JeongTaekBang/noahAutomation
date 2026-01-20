"""
결과 클래스 정의
================

문서 생성 결과 및 상태를 표현하는 데이터 클래스입니다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Any


class GenerationStatus(Enum):
    """문서 생성 상태"""
    SUCCESS = auto()           # 성공
    DUPLICATE = auto()         # 중복 발주
    NOT_FOUND = auto()         # 데이터 없음
    VALIDATION_ERROR = auto()  # 검증 오류
    FILE_ERROR = auto()        # 파일 저장 오류
    CANCELLED = auto()         # 사용자 취소


@dataclass
class DocumentResult:
    """문서 생성 결과

    Attributes:
        success: 성공 여부
        status: 생성 상태
        output_file: 생성된 파일 경로 (성공 시)
        order_no: 주문번호/문서ID
        customer_name: 고객명
        item_count: 아이템 수
        errors: 오류 목록
        warnings: 경고 목록
        message: 결과 메시지
    """
    success: bool
    status: GenerationStatus
    output_file: Path | None = None
    order_no: str = ''
    customer_name: str = ''
    item_count: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    message: str = ''

    @classmethod
    def success_result(
        cls,
        output_file: Path,
        order_no: str,
        customer_name: str,
        item_count: int = 1,
        warnings: list[str] | None = None,
    ) -> DocumentResult:
        """성공 결과 생성"""
        return cls(
            success=True,
            status=GenerationStatus.SUCCESS,
            output_file=output_file,
            order_no=order_no,
            customer_name=customer_name,
            item_count=item_count,
            warnings=warnings or [],
            message=f"문서 생성 완료: {output_file.name}",
        )

    @classmethod
    def duplicate_result(
        cls,
        order_no: str,
        previous_date: str,
        previous_file: str,
    ) -> DocumentResult:
        """중복 발주 결과 생성"""
        return cls(
            success=False,
            status=GenerationStatus.DUPLICATE,
            order_no=order_no,
            message=f"이미 발주된 건입니다. 이전 발주일: {previous_date}",
            errors=[f"이전 파일: {previous_file}"],
        )

    @classmethod
    def not_found_result(cls, order_no: str) -> DocumentResult:
        """데이터 없음 결과 생성"""
        return cls(
            success=False,
            status=GenerationStatus.NOT_FOUND,
            order_no=order_no,
            message=f"'{order_no}'를 찾을 수 없습니다.",
            errors=[f"주문번호 '{order_no}'가 데이터에 없습니다."],
        )

    @classmethod
    def validation_error_result(
        cls,
        order_no: str,
        errors: list[str],
        warnings: list[str] | None = None,
    ) -> DocumentResult:
        """검증 오류 결과 생성"""
        return cls(
            success=False,
            status=GenerationStatus.VALIDATION_ERROR,
            order_no=order_no,
            errors=errors,
            warnings=warnings or [],
            message="검증 오류가 발생했습니다.",
        )

    @classmethod
    def file_error_result(
        cls,
        order_no: str,
        error_message: str,
    ) -> DocumentResult:
        """파일 오류 결과 생성"""
        return cls(
            success=False,
            status=GenerationStatus.FILE_ERROR,
            order_no=order_no,
            errors=[error_message],
            message=f"파일 저장 실패: {error_message}",
        )

    @classmethod
    def cancelled_result(cls, order_no: str, reason: str = '') -> DocumentResult:
        """취소 결과 생성"""
        return cls(
            success=False,
            status=GenerationStatus.CANCELLED,
            order_no=order_no,
            message=f"발주 취소됨{': ' + reason if reason else ''}",
        )
