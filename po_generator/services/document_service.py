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
from po_generator.utils import get_value, resolve_column, build_weight_map
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
        """DN 아이템에 SO_해외의 Model number/Model code 컬럼 추가

        SO_ID + Line item 복합키로 매칭합니다.
        DN에 여러 SO_ID가 섞여 있어도 모든 아이템을 매칭합니다.
        매칭 실패 시 원본 데이터를 그대로 반환합니다.
        """
        items_df = order_data.items_df if order_data.items_df is not None else pd.DataFrame([order_data.first_item])

        # DN에 SO_ID가 없으면 스킵
        so_id_col = resolve_column(items_df.columns, 'so_id')
        if not so_id_col:
            return items_df

        # DN에 포함된 모든 SO_ID 수집
        so_ids = items_df[so_id_col].dropna().unique().tolist()
        if not so_ids:
            return items_df

        try:
            df_so = self._finder.load_so_export_data()
            so_items = df_so[df_so['SO_ID'].isin(so_ids)]
            if so_items.empty:
                return items_df

            # Line item 컬럼 확인
            dn_line_col = 'Line item' if 'Line item' in items_df.columns else None
            so_line_col = 'Line item' if 'Line item' in so_items.columns else None

            if not (dn_line_col and so_line_col):
                logger.debug("Line item 컬럼 없음 — Model 보강 건너뜀")
                return items_df

            # SO_ID + Line item → Model number / Model code 매핑
            model_col = 'Model number' if 'Model number' in so_items.columns else None
            model_code_col = resolve_column(so_items.columns, 'model_code')

            if not model_col and not model_code_col:
                return items_df

            # 복합키 생성 (SO_ID + Line item)
            so_items = so_items.copy()
            so_items['_join_key'] = so_items['SO_ID'].astype(str) + '_' + so_items[so_line_col].astype(str)

            items_df = items_df.copy()
            items_df['_join_key'] = items_df[so_id_col].astype(str) + '_' + items_df[dn_line_col].astype(str)

            if model_col:
                model_map = dict(zip(so_items['_join_key'], so_items[model_col]))
                items_df['Model number'] = items_df['_join_key'].map(model_map)
                logger.debug(f"Model number 보강 완료: {items_df['Model number'].notna().sum()}/{len(items_df)}건 매칭")

            if model_code_col:
                model_code_map = dict(zip(so_items['_join_key'], so_items[model_code_col]))
                items_df['Model code'] = items_df['_join_key'].map(model_code_map)
                logger.debug(f"Model code 보강 완료: {items_df['Model code'].notna().sum()}/{len(items_df)}건 매칭")

            items_df.drop(columns='_join_key', inplace=True)

        except Exception as e:
            logger.warning(f"Model number 보강 실패: {e}")

        return items_df

    def _enrich_with_weight(self, items_df: pd.DataFrame) -> pd.DataFrame:
        """Weight 시트 기반으로 Net Weight (Weight per unit) 보강

        items_df의 'Model code' 값으로 Weight 시트의 ITEM→WEIGHT 매핑을 조회하여
        'Weight per unit' 컬럼을 추가합니다.
        """
        model_code_col = resolve_column(items_df.columns, 'model_code')
        if not model_code_col:
            logger.debug("Model code 컬럼 없음 — Weight 보강 건너뜀")
            return items_df

        weight_map = build_weight_map()
        if not weight_map:
            return items_df

        items_df = items_df.copy()
        items_df['Weight per unit'] = items_df[model_code_col].map(weight_map)
        matched = items_df['Weight per unit'].notna().sum()
        logger.debug(f"Weight 보강 완료: {matched}/{len(items_df)}건 매칭")
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
        history_ok = True
        if not skip_history:
            history_ok = save_to_history(output_file, order_no, customer_name)
            if not history_ok:
                logger.warning("이력 저장 실패 - 발주서는 정상 생성됨")

        result = DocumentResult.success_result(
            output_file=output_file,
            order_no=order_no,
            customer_name=customer_name,
            item_count=order_data.item_count,
            warnings=validation.warnings,
        )
        if not history_ok:
            result.history_saved = False
            result.warnings.append("이력 저장 실패 — 발주서는 정상 생성되었으나 po_history에 기록되지 않았습니다.")
        return result

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

    def generate_fi(self, dn_id: str, rck_po: str | None = None) -> DocumentResult:
        """Final Invoice 생성 (대금 청구용)

        Args:
            dn_id: DN_ID (예: DNO-2026-0001)
            rck_po: RCK PO 번호 (지정 시 해당 발주 아이템만 포함)

        Returns:
            DocumentResult
        """
        label = f"{dn_id} / {rck_po}" if rck_po else dn_id
        logger.info(f"FI 생성 시작: {label}")

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

        # RCK PO 필터링
        if rck_po is not None:
            items_df = (
                order_data.items_df
                if order_data.items_df is not None
                else pd.DataFrame([order_data.first_item])
            )
            rck_po_col = resolve_column(items_df.columns, 'rck_po')
            if rck_po_col:
                filtered = items_df[items_df[rck_po_col] == rck_po]
                if filtered.empty:
                    return DocumentResult.not_found_result(
                        f"{dn_id} (RCK PO: {rck_po})"
                    )
                order_data = OrderData.from_result(filtered)
            else:
                logger.warning("RCK PO 컬럼을 찾을 수 없음 — 전체 아이템으로 생성")

        # 출력 디렉토리 생성
        FI_OUTPUT_DIR.mkdir(exist_ok=True)

        # 파일명 생성 (PO 지정 시 PO번호 포함)
        customer_name = order_data.get_value('customer_name', 'Unknown')
        order_label = f"{dn_id}_{rck_po}" if rck_po else dn_id
        output_file = generate_output_filename("FI", order_label, customer_name, FI_OUTPUT_DIR)

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

        # SO_해외에서 Model number/Model code 보강
        items_df = self._enrich_with_model_number(order_data)
        items_df = self._enrich_with_weight(items_df)

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
