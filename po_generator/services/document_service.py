"""
문서 생성 서비스
================

PO, 거래명세표, PI 등 문서 생성을 오케스트레이션하는 서비스입니다.
CLI에서 비즈니스 로직을 분리하여 재사용 가능하게 합니다.
"""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from po_generator.config import (
    OUTPUT_DIR,
    TS_OUTPUT_DIR,
    PI_OUTPUT_DIR,
    TS_TEMPLATE_FILE,
    PI_TEMPLATE_FILE,
)
from po_generator.utils import get_value
from po_generator.validators import validate_order_data, validate_multiple_items
from po_generator.history import check_duplicate_order, save_to_history
from po_generator.excel_generator import create_po_workbook
from po_generator.ts_generator import create_ts_xlwings
from po_generator.pi_generator import create_pi_xlwings
from po_generator.cli_common import generate_output_filename, validate_output_path
from po_generator.services.result import DocumentResult, GenerationStatus
from po_generator.services.finder_service import FinderService, OrderData

logger = logging.getLogger(__name__)


class DocumentService:
    """문서 생성 오케스트레이터

    PO, 거래명세표, PI 등의 문서 생성 비즈니스 로직을 담당합니다.
    CLI에서 이 서비스를 호출하여 문서를 생성합니다.

    Usage:
        service = DocumentService()
        result = service.generate_po('ND-0001')
        if result.success:
            print(f"생성 완료: {result.output_file}")
        else:
            print(f"실패: {result.message}")
    """

    def __init__(self, finder: FinderService | None = None):
        """
        Args:
            finder: FinderService 인스턴스 (없으면 새로 생성)
        """
        self._finder = finder or FinderService()

    @property
    def finder(self) -> FinderService:
        """FinderService 인스턴스"""
        return self._finder

    def generate_po(
        self,
        order_no: str,
        force: bool = False,
        skip_history: bool = False,
    ) -> DocumentResult:
        """Purchase Order 생성

        Args:
            order_no: RCK Order No. (예: ND-0001)
            force: 중복 발주 및 검증 오류 무시
            skip_history: 이력 저장 건너뛰기

        Returns:
            DocumentResult
        """
        logger.info(f"PO 생성 시작: {order_no}")

        # 1. 중복 발주 체크
        if not force:
            dup_info = check_duplicate_order(order_no)
            if dup_info:
                return DocumentResult.duplicate_result(
                    order_no=order_no,
                    previous_date=dup_info['생성일시'],
                    previous_file=Path(dup_info['생성파일']).name,
                )

        # 2. 데이터 검색
        order_data = self._finder.find_po(order_no)
        if order_data is None:
            return DocumentResult.not_found_result(order_no)

        # 3. 데이터 검증
        if order_data.is_multi_item:
            validation = validate_multiple_items(order_data.items_df)
        else:
            validation = validate_order_data(order_data.first_item)

        if validation.has_errors and not force:
            return DocumentResult.validation_error_result(
                order_no=order_no,
                errors=validation.errors,
                warnings=validation.warnings,
            )

        # 4. 출력 디렉토리 생성
        OUTPUT_DIR.mkdir(exist_ok=True)

        # 5. 파일명 생성
        customer_name = order_data.get_value('customer_name', 'Unknown')
        output_file = generate_output_filename("PO", order_no, customer_name, OUTPUT_DIR)

        if not validate_output_path(output_file, OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=order_no,
                error_message="출력 경로 검증 실패",
            )

        # 6. 문서 생성
        try:
            wb = create_po_workbook(order_data.first_item, order_data.items_df)
            wb.save(output_file)
            logger.info(f"PO 생성 완료: {output_file}")

        except PermissionError:
            return DocumentResult.file_error_result(
                order_no=order_no,
                error_message=f"파일이 열려있거나 권한 없음: {output_file.name}",
            )
        except Exception as e:
            logger.exception("PO 생성 중 오류 발생")
            # 롤백: 부분 생성된 파일 삭제
            if output_file.exists():
                try:
                    output_file.unlink()
                except (IOError, OSError, PermissionError):
                    pass
            return DocumentResult.file_error_result(
                order_no=order_no,
                error_message=str(e),
            )

        # 7. 이력 저장
        if not skip_history:
            history_saved = save_to_history(output_file, order_no, customer_name)
            if not history_saved:
                logger.warning("이력 저장 실패 - 발주서는 정상 생성됨")

        return DocumentResult.success_result(
            output_file=output_file,
            order_no=order_no,
            customer_name=customer_name,
            item_count=order_data.item_count,
            warnings=validation.warnings,
        )

    def generate_ts(
        self,
        doc_id: str,
        doc_type: str = 'DN',
    ) -> DocumentResult:
        """거래명세표 생성

        Args:
            doc_id: DN_ID 또는 선수금_ID
            doc_type: 'DN' 또는 'ADV'

        Returns:
            DocumentResult
        """
        logger.info(f"거래명세표 생성 시작: {doc_id} (유형: {doc_type})")

        # 템플릿 확인
        if not TS_TEMPLATE_FILE.exists():
            return DocumentResult.file_error_result(
                order_no=doc_id,
                error_message=f"템플릿 파일 없음: {TS_TEMPLATE_FILE}",
            )

        # 데이터 검색
        if doc_type == 'ADV':
            result = self._finder.find_so_for_advance(doc_id)
            if result is None:
                return DocumentResult.not_found_result(doc_id)
            _, order_data = result
        else:  # DN
            order_data = self._finder.find_dn(doc_id)
            if order_data is None:
                return DocumentResult.not_found_result(doc_id)

        # 출력 디렉토리 생성
        TS_OUTPUT_DIR.mkdir(exist_ok=True)

        # 파일명 생성
        customer_name = order_data.get_value('customer_name', 'Unknown')
        prefix = "TS_ADV" if doc_type == 'ADV' else "TS"
        output_file = generate_output_filename(prefix, doc_id, customer_name, TS_OUTPUT_DIR)

        if not validate_output_path(output_file, TS_OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=doc_id,
                error_message="출력 경로 검증 실패",
            )

        # 문서 생성
        try:
            create_ts_xlwings(
                template_path=TS_TEMPLATE_FILE,
                output_path=output_file,
                order_data=order_data.first_item,
                items_df=order_data.items_df,
                doc_type=doc_type,
            )
            logger.info(f"거래명세표 생성 완료: {output_file}")

        except FileNotFoundError as e:
            return DocumentResult.file_error_result(
                order_no=doc_id,
                error_message=str(e),
            )
        except PermissionError:
            return DocumentResult.file_error_result(
                order_no=doc_id,
                error_message=f"파일이 열려있거나 권한 없음: {output_file.name}",
            )
        except Exception as e:
            logger.exception("거래명세표 생성 중 오류 발생")
            return DocumentResult.file_error_result(
                order_no=doc_id,
                error_message=str(e),
            )

        return DocumentResult.success_result(
            output_file=output_file,
            order_no=doc_id,
            customer_name=customer_name,
            item_count=order_data.item_count,
        )

    def generate_pi(self, so_id: str) -> DocumentResult:
        """Proforma Invoice 생성

        Args:
            so_id: SO_ID (예: SOO-2026-0001)

        Returns:
            DocumentResult
        """
        logger.info(f"PI 생성 시작: {so_id}")

        # 템플릿 확인
        if not PI_TEMPLATE_FILE.exists():
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message=f"템플릿 파일 없음: {PI_TEMPLATE_FILE}",
            )

        # 데이터 검색
        order_data = self._finder.find_so_export(so_id)
        if order_data is None:
            return DocumentResult.not_found_result(so_id)

        # 출력 디렉토리 생성
        PI_OUTPUT_DIR.mkdir(exist_ok=True)

        # 파일명 생성
        customer_name = order_data.get_value('customer_name', 'Unknown')
        output_file = generate_output_filename("PI", so_id, customer_name, PI_OUTPUT_DIR)

        if not validate_output_path(output_file, PI_OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message="출력 경로 검증 실패",
            )

        # 문서 생성
        try:
            create_pi_xlwings(
                template_path=PI_TEMPLATE_FILE,
                output_path=output_file,
                order_data=order_data.first_item,
                items_df=order_data.items_df,
            )
            logger.info(f"PI 생성 완료: {output_file}")

        except FileNotFoundError as e:
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message=str(e),
            )
        except PermissionError:
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message=f"파일이 열려있거나 권한 없음: {output_file.name}",
            )
        except Exception as e:
            logger.exception("PI 생성 중 오류 발생")
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message=str(e),
            )

        return DocumentResult.success_result(
            output_file=output_file,
            order_no=so_id,
            customer_name=customer_name,
            item_count=order_data.item_count,
        )
