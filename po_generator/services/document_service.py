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
    CI_OUTPUT_DIR,
    FI_OUTPUT_DIR,
    PL_OUTPUT_DIR,
    OC_OUTPUT_DIR,
    TS_TEMPLATE_FILE,
    PI_TEMPLATE_FILE,
    CI_TEMPLATE_FILE,
    FI_TEMPLATE_FILE,
    PL_TEMPLATE_FILE,
    OC_TEMPLATE_FILE,
)
from po_generator.utils import get_value
from po_generator.validators import validate_order_data, validate_multiple_items
from po_generator.history import check_duplicate_order, save_to_history
from po_generator.excel_generator import create_po_workbook
from po_generator.ts_generator import create_ts_xlwings
from po_generator.pi_generator import create_pi_xlwings
from po_generator.ci_generator import create_ci_xlwings
from po_generator.fi_generator import create_fi_xlwings
from po_generator.pl_generator import create_pl_xlwings
from po_generator.oc_generator import create_oc_xlwings
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

    def _enrich_with_model_number(self, order_data: OrderData) -> pd.DataFrame:
        """DN 아이템에 SO_해외의 Model number 컬럼 추가 (CI용)

        SO_ID + Item name으로 매칭하여 Model number를 가져옵니다.
        매칭 실패 시 원본 데이터를 그대로 반환합니다.
        """
        items_df = order_data.items_df if order_data.items_df is not None else pd.DataFrame([order_data.first_item])

        so_id = get_value(order_data.first_item, 'so_id', '')
        if not so_id:
            return items_df

        try:
            df_so = self._finder.load_so_export_data()
            so_items = df_so[df_so['SO_ID'] == so_id]
            if so_items.empty:
                return items_df

            # DN과 SO의 item_name 컬럼명 찾기
            dn_item_col = None
            for col in ('Item name', 'Item'):
                if col in items_df.columns:
                    dn_item_col = col
                    break

            so_item_col = 'Item name' if 'Item name' in so_items.columns else None
            model_col = 'Model number' if 'Model number' in so_items.columns else None

            if dn_item_col and so_item_col and model_col:
                model_map = dict(zip(so_items[so_item_col], so_items[model_col]))
                items_df = items_df.copy()
                items_df['Model number'] = items_df[dn_item_col].map(model_map)
                logger.debug(f"Model number 보강 완료: {sum(items_df['Model number'].notna())}건 매칭")

        except Exception as e:
            logger.warning(f"Model number 보강 실패: {e}")

        return items_df

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

    def generate_ci(self, dn_id: str) -> DocumentResult:
        """Commercial Invoice 생성

        Args:
            dn_id: DN_ID (예: DNO-2026-0001)

        Returns:
            DocumentResult
        """
        logger.info(f"CI 생성 시작: {dn_id}")

        # 템플릿 확인
        if not CI_TEMPLATE_FILE.exists():
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=f"템플릿 파일 없음: {CI_TEMPLATE_FILE}",
            )

        # 데이터 검색 (DN_해외 + Customer_해외 JOIN)
        order_data = self._finder.find_dn_export(dn_id)
        if order_data is None:
            return DocumentResult.not_found_result(dn_id)

        # SO_해외에서 Model number 보강
        items_df = self._enrich_with_model_number(order_data)

        # 출력 디렉토리 생성
        CI_OUTPUT_DIR.mkdir(exist_ok=True)

        # 파일명 생성
        customer_name = order_data.get_value('customer_name', 'Unknown')
        output_file = generate_output_filename("CI", dn_id, customer_name, CI_OUTPUT_DIR)

        if not validate_output_path(output_file, CI_OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message="출력 경로 검증 실패",
            )

        # 문서 생성
        try:
            create_ci_xlwings(
                template_path=CI_TEMPLATE_FILE,
                output_path=output_file,
                order_data=items_df.iloc[0],
                items_df=items_df,
            )
            logger.info(f"CI 생성 완료: {output_file}")

        except FileNotFoundError as e:
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=str(e),
            )
        except PermissionError:
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=f"파일이 열려있거나 권한 없음: {output_file.name}",
            )
        except Exception as e:
            logger.exception("CI 생성 중 오류 발생")
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=str(e),
            )

        return DocumentResult.success_result(
            output_file=output_file,
            order_no=dn_id,
            customer_name=customer_name,
            item_count=order_data.item_count,
        )

    def generate_fi(self, dn_id: str) -> DocumentResult:
        """Final Invoice 생성 (대금 청구용)

        Args:
            dn_id: DN_ID (예: DNO-2026-0001)

        Returns:
            DocumentResult
        """
        logger.info(f"FI 생성 시작: {dn_id}")

        # 템플릿 확인
        if not FI_TEMPLATE_FILE.exists():
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=f"템플릿 파일 없음: {FI_TEMPLATE_FILE}",
            )

        # 데이터 검색 (DN_해외 + Customer_해외 JOIN)
        order_data = self._finder.find_dn_export(dn_id)
        if order_data is None:
            return DocumentResult.not_found_result(dn_id)

        # 출력 디렉토리 생성
        FI_OUTPUT_DIR.mkdir(exist_ok=True)

        # 파일명 생성
        customer_name = order_data.get_value('customer_name', 'Unknown')
        output_file = generate_output_filename("FI", dn_id, customer_name, FI_OUTPUT_DIR)

        if not validate_output_path(output_file, FI_OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message="출력 경로 검증 실패",
            )

        # 문서 생성
        try:
            create_fi_xlwings(
                template_path=FI_TEMPLATE_FILE,
                output_path=output_file,
                order_data=order_data.first_item,
                items_df=order_data.items_df,
            )
            logger.info(f"FI 생성 완료: {output_file}")

        except FileNotFoundError as e:
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=str(e),
            )
        except PermissionError:
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=f"파일이 열려있거나 권한 없음: {output_file.name}",
            )
        except Exception as e:
            logger.exception("FI 생성 중 오류 발생")
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=str(e),
            )

        return DocumentResult.success_result(
            output_file=output_file,
            order_no=dn_id,
            customer_name=customer_name,
            item_count=order_data.item_count,
        )

    def generate_pl(self, dn_id: str) -> DocumentResult:
        """Packing List 생성

        Args:
            dn_id: DN_ID (예: DNO-2026-0001)

        Returns:
            DocumentResult
        """
        logger.info(f"PL 생성 시작: {dn_id}")

        if not PL_TEMPLATE_FILE.exists():
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=f"템플릿 파일 없음: {PL_TEMPLATE_FILE}",
            )

        # 데이터 검색 (DN_해외 + Customer_해외 JOIN)
        order_data = self._finder.find_dn_export(dn_id)
        if order_data is None:
            return DocumentResult.not_found_result(dn_id)

        # SO_해외에서 Model number 보강
        items_df = self._enrich_with_model_number(order_data)

        PL_OUTPUT_DIR.mkdir(exist_ok=True)

        customer_name = order_data.get_value('customer_name', 'Unknown')
        output_file = generate_output_filename("PL", dn_id, customer_name, PL_OUTPUT_DIR)

        if not validate_output_path(output_file, PL_OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message="출력 경로 검증 실패",
            )

        try:
            create_pl_xlwings(
                template_path=PL_TEMPLATE_FILE,
                output_path=output_file,
                order_data=items_df.iloc[0],
                items_df=items_df,
            )
            logger.info(f"PL 생성 완료: {output_file}")

        except FileNotFoundError as e:
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=str(e),
            )
        except PermissionError:
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=f"파일이 열려있거나 권한 없음: {output_file.name}",
            )
        except Exception as e:
            logger.exception("PL 생성 중 오류 발생")
            return DocumentResult.file_error_result(
                order_no=dn_id,
                error_message=str(e),
            )

        return DocumentResult.success_result(
            output_file=output_file,
            order_no=dn_id,
            customer_name=customer_name,
            item_count=order_data.item_count,
        )

    def generate_oc(self, so_id: str) -> DocumentResult:
        """Order Confirmation 생성

        Args:
            so_id: SO_ID (예: SOO-2026-0001)

        Returns:
            DocumentResult
        """
        logger.info(f"OC 생성 시작: {so_id}")

        if not OC_TEMPLATE_FILE.exists():
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message=f"템플릿 파일 없음: {OC_TEMPLATE_FILE}",
            )

        # 데이터 검색 (SO_해외 + Customer_해외 JOIN)
        order_data = self._finder.find_so_export_with_customer(so_id)
        if order_data is None:
            return DocumentResult.not_found_result(so_id)

        OC_OUTPUT_DIR.mkdir(exist_ok=True)

        customer_name = order_data.get_value('customer_name', 'Unknown')
        output_file = generate_output_filename("OC", so_id, customer_name, OC_OUTPUT_DIR)

        if not validate_output_path(output_file, OC_OUTPUT_DIR):
            return DocumentResult.file_error_result(
                order_no=so_id,
                error_message="출력 경로 검증 실패",
            )

        try:
            create_oc_xlwings(
                template_path=OC_TEMPLATE_FILE,
                output_path=output_file,
                order_data=order_data.first_item,
                items_df=order_data.items_df,
            )
            logger.info(f"OC 생성 완료: {output_file}")

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
            logger.exception("OC 생성 중 오류 발생")
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
