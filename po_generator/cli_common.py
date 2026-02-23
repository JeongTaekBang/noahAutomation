"""
CLI 공통 유틸리티
==================

create_po.py와 create_ts.py에서 공유하는 함수들
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime

import pandas as pd

from po_generator.config import ORDER_LIST_DISPLAY_LIMIT
from po_generator.history import sanitize_filename


def print_available_orders(df: pd.DataFrame, limit: int = ORDER_LIST_DISPLAY_LIMIT) -> None:
    """사용 가능한 주문번호 목록 출력

    Args:
        df: 주문 데이터
        limit: 출력 제한 수 (기본값: ORDER_LIST_DISPLAY_LIMIT)

    Returns:
        None
    """
    orders = df['RCK Order no.'].dropna().unique().tolist()
    print("\n사용 가능한 RCK Order No. 목록:")
    for order in orders[:limit]:
        print(f"  - {order}")
    if len(orders) > limit:
        print(f"  ... 외 {len(orders) - limit}건")


def validate_output_path(output_file: Path, output_dir: Path) -> bool:
    """출력 파일 경로 검증 (Path Traversal 방지)

    Args:
        output_file: 출력 파일 경로
        output_dir: 허용된 출력 디렉토리

    Returns:
        유효한 경로이면 True
    """
    try:
        resolved_file = output_file.resolve()
        resolved_dir = output_dir.resolve()

        # Python 3.9+ is_relative_to() 사용
        if hasattr(resolved_file, 'is_relative_to'):
            if not resolved_file.is_relative_to(resolved_dir):
                print(f"  [오류] 잘못된 파일 경로: {output_file}")
                return False
        else:
            # Python 3.8 fallback: relative_to() 사용 (실패 시 ValueError 발생)
            resolved_file.relative_to(resolved_dir)
    except (ValueError, AttributeError):
        # relative_to() 실패 = output_file이 output_dir 하위가 아님
        print(f"  [오류] 잘못된 파일 경로: {output_file}")
        return False
    return True


def generate_output_filename(
    prefix: str,
    order_no: str,
    customer_name: str,
    output_dir: Path,
) -> Path:
    """출력 파일명 생성

    Args:
        prefix: 파일명 접두사 (예: "PO", "TS")
        order_no: 주문번호
        customer_name: 고객명
        output_dir: 출력 디렉토리

    Returns:
        출력 파일 경로
    """
    today = datetime.now().strftime("%y%m%d")
    customer_name_safe = sanitize_filename(customer_name)
    order_no_safe = sanitize_filename(order_no)
    return output_dir / f"{prefix}_{order_no_safe}_{customer_name_safe}_{today}.xlsx"
