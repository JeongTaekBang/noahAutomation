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

from po_generator.utils import (
    load_noah_po_lists,
    load_dn_data,
    load_pmt_data,
    load_so_export_data,
    find_order_data,
    find_dn_data,
    find_pmt_data,
    find_so_export_data,
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
        self._so_export_df: pd.DataFrame | None = None

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

    def find_so_for_advance(self, advance_id: str) -> tuple[pd.Series, OrderData] | None:
        """선수금용 SO 데이터 검색

        PMT_국내에서 선수금_ID로 SO_ID를 찾고,
        SO_국내에서 해당 SO의 모든 아이템을 반환합니다.

        Args:
            advance_id: 선수금_ID (예: ADV_2026-0001)

        Returns:
            (PMT 정보, SO 아이템 OrderData) 또는 None
        """
        result = load_so_for_advance(advance_id)
        if result is None:
            return None
        pmt_data, so_items = result
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

        result = []
        for po_id in df[po_col].dropna().unique()[:limit]:
            customer = df[df[po_col] == po_id]['Customer name'].iloc[0] if len(df[df[po_col] == po_id]) > 0 else ''
            result.append((str(po_id), str(customer) if pd.notna(customer) else ''))

        return result

    def get_available_dn_ids(self, limit: int = 20) -> list[tuple[str, str]]:
        """사용 가능한 DN_ID 목록 반환

        Args:
            limit: 반환할 최대 개수

        Returns:
            (DN_ID, 고객명) 튜플 목록
        """
        df = self.load_dn_data()
        result = []
        for dn_id in df['DN_ID'].dropna().unique()[:limit]:
            customer = df[df['DN_ID'] == dn_id]['Customer name'].iloc[0] if len(df[df['DN_ID'] == dn_id]) > 0 else ''
            result.append((str(dn_id), str(customer) if pd.notna(customer) else ''))
        return result

    def get_available_so_export_ids(self, limit: int = 20) -> list[tuple[str, str]]:
        """사용 가능한 SO_ID (해외) 목록 반환

        Args:
            limit: 반환할 최대 개수

        Returns:
            (SO_ID, 고객명) 튜플 목록
        """
        df = self.load_so_export_data()
        result = []
        for so_id in df['SO_ID'].dropna().unique()[:limit]:
            customer = df[df['SO_ID'] == so_id]['Customer name'].iloc[0] if len(df[df['SO_ID'] == so_id]) > 0 else ''
            result.append((str(so_id), str(customer) if pd.notna(customer) else ''))
        return result
