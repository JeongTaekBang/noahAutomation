"""
데이터 조회 서비스
==================

NOAH_SO_PO_DN.xlsx에서 데이터를 조회하는 서비스입니다.
utils.py의 함수들을 래핑하여 통합된 인터페이스를 제공합니다.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

import pandas as pd

from po_generator.config import NOAH_SO_PO_DN_FILE, SO_DOMESTIC_SHEET
from po_generator.utils import (
    load_noah_po_lists,
    load_dn_data,
    load_pmt_data,
    load_so_export_data,
    load_dn_export_data,
    load_so_export_with_customer,
    find_order_data,
    find_dn_data,
    find_pmt_data,
    find_so_export_data,
    find_dn_export_data,
    load_so_for_advance,
    get_value,
)

logger = logging.getLogger(__name__)


@dataclass
class OrderData:
    """주문 데이터 래퍼

    단일/다중 아이템을 통합하여 처리합니다.

    Attributes:
        first_item: 첫 번째 아이템 (공통 정보 추출용)
        items_df: 다중 아이템인 경우 전체 DataFrame, 단일이면 None
        item_count: 아이템 수
    """
    first_item: pd.Series
    items_df: pd.DataFrame | None
    item_count: int

    @classmethod
    def from_result(
        cls,
        result: pd.Series | pd.DataFrame,
    ) -> OrderData:
        """find_*_data 결과로부터 생성"""
        if isinstance(result, pd.DataFrame):
            return cls(
                first_item=result.iloc[0],
                items_df=result,
                item_count=len(result),
            )
        else:
            return cls(
                first_item=result,
                items_df=None,
                item_count=1,
            )

    def get_value(self, key: str, default: Any = '') -> Any:
        """첫 번째 아이템에서 값 가져오기"""
        return get_value(self.first_item, key, default)

    @property
    def is_multi_item(self) -> bool:
        """다중 아이템 여부"""
        return self.item_count > 1


class FinderService:
    """데이터 조회 서비스

    NOAH_SO_PO_DN.xlsx에서 각종 데이터를 조회합니다.
    데이터 로딩은 지연 로딩(lazy loading)으로 처리합니다.
    """

    def __init__(self):
        self._po_df: pd.DataFrame | None = None
        self._dn_df: pd.DataFrame | None = None
        self._pmt_df: pd.DataFrame | None = None
        self._so_domestic_df: pd.DataFrame | None = None
        self._so_export_df: pd.DataFrame | None = None
        self._so_export_cust_df: pd.DataFrame | None = None
        self._dn_export_df: pd.DataFrame | None = None

    def load_po_data(self) -> pd.DataFrame:
        """PO 데이터 로드 (국내 + 해외)"""
        if self._po_df is None:
            logger.info("PO 데이터 로딩 중...")
            self._po_df = load_noah_po_lists()
            logger.info(f"PO 데이터 {len(self._po_df)}건 로드 완료")
        return self._po_df

    def load_dn_data(self) -> pd.DataFrame:
        """DN 데이터 로드"""
        if self._dn_df is None:
            logger.info("DN 데이터 로딩 중...")
            self._dn_df = load_dn_data()
            logger.info(f"DN 데이터 {len(self._dn_df)}건 로드 완료")
        return self._dn_df

    def load_pmt_data(self) -> pd.DataFrame:
        """PMT 데이터 로드"""
        if self._pmt_df is None:
            logger.info("PMT 데이터 로딩 중...")
            self._pmt_df = load_pmt_data()
            logger.info(f"PMT 데이터 {len(self._pmt_df)}건 로드 완료")
        return self._pmt_df

    def load_so_export_data(self) -> pd.DataFrame:
        """SO 해외 데이터 로드"""
        if self._so_export_df is None:
            logger.info("SO 해외 데이터 로딩 중...")
            self._so_export_df = load_so_export_data()
            logger.info(f"SO 해외 데이터 {len(self._so_export_df)}건 로드 완료")
        return self._so_export_df

    def find_po(self, order_no: str) -> OrderData | None:
        """PO 데이터 검색

        Args:
            order_no: RCK Order No. (예: ND-0001)

        Returns:
            OrderData 또는 None
        """
        df = self.load_po_data()
        result = find_order_data(df, order_no)
        if result is None:
            return None
        return OrderData.from_result(result)

    def find_dn(self, dn_id: str) -> OrderData | None:
        """DN 데이터 검색

        Args:
            dn_id: DN_ID (예: DN-2026-0001)

        Returns:
            OrderData 또는 None
        """
        df = self.load_dn_data()
        result = find_dn_data(df, dn_id)
        if result is None:
            return None
        return OrderData.from_result(result)

    def find_pmt(self, advance_id: str) -> OrderData | None:
        """PMT 데이터 검색

        Args:
            advance_id: 선수금_ID (예: ADV_2026-0001)

        Returns:
            OrderData 또는 None
        """
        df = self.load_pmt_data()
        result = find_pmt_data(df, advance_id)
        if result is None:
            return None
        return OrderData.from_result(result)

    def find_so_export(self, so_id: str) -> OrderData | None:
        """SO 해외 데이터 검색

        Args:
            so_id: SO_ID (예: SOO-2026-0001)

        Returns:
            OrderData 또는 None
        """
        df = self.load_so_export_data()
        result = find_so_export_data(df, so_id)
        if result is None:
            return None
        return OrderData.from_result(result)

    def load_so_export_with_customer(self) -> pd.DataFrame:
        """SO 해외 + Customer_해외 데이터 로드"""
        if self._so_export_cust_df is None:
            logger.info("SO 해외 + Customer 데이터 로딩 중...")
            self._so_export_cust_df = load_so_export_with_customer()
            logger.info(f"SO 해외 + Customer 데이터 {len(self._so_export_cust_df)}건 로드 완료")
        return self._so_export_cust_df

    def find_so_export_with_customer(self, so_id: str) -> OrderData | None:
        """SO 해외 + Customer_해외 데이터 검색 (OC용)

        Args:
            so_id: SO_ID (예: SOO-2026-0001)

        Returns:
            OrderData 또는 None
        """
        df = self.load_so_export_with_customer()
        result = find_so_export_data(df, so_id)
        if result is None:
            return None
        return OrderData.from_result(result)


    def load_dn_export_data(self) -> pd.DataFrame:
        """DN 해외 데이터 로드 (Customer_해외 JOIN 포함)"""
        if self._dn_export_df is None:
            logger.info("DN 해외 데이터 로딩 중...")
            self._dn_export_df = load_dn_export_data()
            logger.info(f"DN 해외 데이터 {len(self._dn_export_df)}건 로드 완료")
        return self._dn_export_df

    def find_dn_export(self, dn_id: str) -> OrderData | None:
        """DN 해외 데이터 검색

        Args:
            dn_id: DN_ID (예: DNO-2026-0001)

        Returns:
            OrderData 또는 None
        """
        df = self.load_dn_export_data()
        result = find_dn_export_data(df, dn_id)
        if result is None:
            return None
        return OrderData.from_result(result)

    def find_dn_export_by_customer_po(self, customer_po: str) -> OrderData | None:
        """Customer PO(발주번호)로 DN 해외 데이터 검색 (복수 DN 통합)

        여러 DN에 걸쳐 동일 Customer PO를 가진 아이템을 모두 찾아 반환합니다.

        Args:
            customer_po: 고객 발주번호 (예: 26KPO00144)

        Returns:
            OrderData 또는 None
        """
        from po_generator.utils import resolve_column

        df = self.load_dn_export_data()
        cpo_col = resolve_column(df.columns, 'customer_po')
        if cpo_col is None:
            logger.warning("Customer PO 컬럼을 찾을 수 없습니다.")
            return None

        matched = df[df[cpo_col].astype(str) == customer_po]
        if matched.empty:
            return None

        return OrderData.from_result(matched)

    def get_available_dn_export_ids(self, limit: int = 20) -> list[tuple[str, str]]:
        """사용 가능한 DN_ID (해외) 목록 반환

        Args:
            limit: 반환할 최대 개수

        Returns:
            (DN_ID, 고객명) 튜플 목록
        """
        df = self.load_dn_export_data()
        pairs = (df.dropna(subset=['DN_ID'])
                 .drop_duplicates(subset='DN_ID', keep='first')
                 .head(limit))
        return [
            (str(row['DN_ID']), str(row['Customer name']) if pd.notna(row.get('Customer name')) else '')
            for _, row in pairs.iterrows()
        ]

    def _load_so_domestic(self) -> pd.DataFrame:
        """SO_국내 원본 데이터 로드 (캐시)"""
        if self._so_domestic_df is None:
            if not NOAH_SO_PO_DN_FILE.exists():
                raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")
            with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
                self._so_domestic_df = pd.read_excel(xl, sheet_name=SO_DOMESTIC_SHEET)
        return self._so_domestic_df

    def find_so_for_advance(self, advance_id: str) -> tuple[pd.Series, OrderData] | None:
        """선수금용 SO 데이터 검색

        PMT_국내에서 선수금_ID로 SO_ID를 찾고,
        SO_국내에서 해당 SO의 모든 아이템을 반환합니다.
        캐시된 PMT/SO 데이터를 재사용하여 중복 Excel 로드를 방지합니다.

        Args:
            advance_id: 선수금_ID (예: ADV_2026-0001)

        Returns:
            (PMT 정보, SO 아이템 OrderData) 또는 None
        """
        # 캐시된 PMT 데이터에서 선수금_ID 검색
        pmt_df = self.load_pmt_data()
        pmt_mask = pmt_df['선수금_ID'] == advance_id
        if pmt_mask.sum() == 0:
            logger.warning(f"선수금_ID '{advance_id}'를 찾을 수 없습니다.")
            return None

        pmt_data = pmt_df[pmt_mask].iloc[0]
        so_id = pmt_data['SO_ID']
        logger.info(f"선수금_ID '{advance_id}' -> SO_ID: {so_id}")

        # 캐시된 SO 데이터에서 해당 SO_ID의 모든 아이템 검색
        so_df = self._load_so_domestic()
        so_items = so_df[so_df['SO_ID'] == so_id].copy()

        if len(so_items) == 0:
            logger.warning(f"SO_ID '{so_id}'에 해당하는 SO 데이터를 찾을 수 없습니다.")
            return None

        so_items['선수금_ID'] = advance_id
        so_items['_시트구분'] = '국내'
        so_items['_문서유형'] = 'ADV'

        logger.info(f"SO_ID '{so_id}': {len(so_items)}개 아이템 발견")
        return pmt_data, OrderData.from_result(so_items)

    def get_available_po_ids(self, limit: int = 20) -> list[tuple[str, str]]:
        """사용 가능한 PO_ID 목록 반환

        Args:
            limit: 반환할 최대 개수

        Returns:
            (PO_ID, 고객명) 튜플 목록
        """
        df = self.load_po_data()
        # PO_ID 컬럼 찾기
        po_col = None
        for col in ['PO_ID', 'RCK Order no.']:
            if col in df.columns:
                po_col = col
                break

        if po_col is None:
            return []

        pairs = (df.dropna(subset=[po_col])
                 .drop_duplicates(subset=po_col, keep='first')
                 .head(limit))
        return [
            (str(row[po_col]), str(row['Customer name']) if pd.notna(row.get('Customer name')) else '')
            for _, row in pairs.iterrows()
        ]

    def get_available_dn_ids(self, limit: int = 20) -> list[tuple[str, str]]:
        """사용 가능한 DN_ID 목록 반환

        Args:
            limit: 반환할 최대 개수

        Returns:
            (DN_ID, 고객명) 튜플 목록
        """
        df = self.load_dn_data()
        pairs = (df.dropna(subset=['DN_ID'])
                 .drop_duplicates(subset='DN_ID', keep='first')
                 .head(limit))
        return [
            (str(row['DN_ID']), str(row['Customer name']) if pd.notna(row.get('Customer name')) else '')
            for _, row in pairs.iterrows()
        ]

    def get_available_so_export_ids(self, limit: int = 20) -> list[tuple[str, str]]:
        """사용 가능한 SO_ID (해외) 목록 반환

        Args:
            limit: 반환할 최대 개수

        Returns:
            (SO_ID, 고객명) 튜플 목록
        """
        df = self.load_so_export_data()
        pairs = (df.dropna(subset=['SO_ID'])
                 .drop_duplicates(subset='SO_ID', keep='first')
                 .head(limit))
        return [
            (str(row['SO_ID']), str(row['Customer name']) if pd.notna(row.get('Customer name')) else '')
            for _, row in pairs.iterrows()
        ]
