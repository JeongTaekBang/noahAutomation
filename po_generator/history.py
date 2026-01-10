"""
이력 관리 모듈
==============

발주서 생성 이력을 관리합니다.
- 중복 발주 체크
- 이력 저장
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, TypedDict

import pandas as pd

from po_generator.config import HISTORY_FILE
from po_generator.utils import get_safe_value

logger = logging.getLogger(__name__)


class DuplicateInfo(TypedDict):
    """중복 발주 정보"""
    생성일시: str
    생성파일: str


def check_duplicate_order(order_no: str) -> Optional[DuplicateInfo]:
    """중복 발주 체크

    Args:
        order_no: RCK Order No.

    Returns:
        중복인 경우 이전 발주 정보, 아니면 None
    """
    if not HISTORY_FILE.exists():
        logger.debug(f"이력 파일 없음: {HISTORY_FILE}")
        return None

    try:
        df_history = pd.read_excel(HISTORY_FILE)
    except Exception as e:
        logger.error(f"이력 파일 읽기 실패: {e}")
        return None

    mask = df_history['RCK Order no.'] == order_no
    if mask.sum() > 0:
        last_record = df_history[mask].iloc[-1]
        logger.warning(f"중복 발주 감지: {order_no}")
        return DuplicateInfo(
            생성일시=str(last_record.get('생성일시', '')),
            생성파일=str(last_record.get('생성파일', ''))
        )

    return None


def save_to_history(order_data: pd.Series, output_file: Path) -> bool:
    """발주 이력 저장

    Args:
        order_data: 주문 데이터
        output_file: 생성된 파일 경로

    Returns:
        저장 성공 여부
    """
    history_record = {
        '생성일시': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'RCK Order no.': get_safe_value(order_data, 'RCK Order no.'),
        'Customer PO': get_safe_value(order_data, 'Customer PO'),
        'Customer name': get_safe_value(order_data, 'Customer name'),
        'Item name': get_safe_value(order_data, 'Item name'),
        'Item qty': get_safe_value(order_data, 'Item qty'),
        'ICO Unit': get_safe_value(order_data, 'ICO Unit'),
        'Total ICO': get_safe_value(order_data, 'Total ICO'),
        '생성파일': str(output_file),
        '시트구분': get_safe_value(order_data, '_시트구분'),
    }

    try:
        if HISTORY_FILE.exists():
            df_history = pd.read_excel(HISTORY_FILE)
        else:
            df_history = pd.DataFrame()

        df_new = pd.DataFrame([history_record])
        df_history = pd.concat([df_history, df_new], ignore_index=True)
        df_history.to_excel(HISTORY_FILE, index=False)

        logger.info(f"이력 저장 완료: {HISTORY_FILE.name}")
        return True

    except Exception as e:
        logger.error(f"이력 저장 실패: {e}")
        return False


def get_history_count() -> int:
    """이력 건수 조회

    Returns:
        이력 건수
    """
    if not HISTORY_FILE.exists():
        return 0

    try:
        df = pd.read_excel(HISTORY_FILE)
        return len(df)
    except Exception:
        return 0


def clear_history() -> bool:
    """이력 초기화 (테스트용)

    Returns:
        성공 여부
    """
    try:
        if HISTORY_FILE.exists():
            HISTORY_FILE.unlink()
            logger.info("이력 파일 삭제됨")
        return True
    except Exception as e:
        logger.error(f"이력 삭제 실패: {e}")
        return False
