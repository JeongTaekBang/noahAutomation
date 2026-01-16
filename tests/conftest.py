"""
Pytest fixtures for NOAH PO Generator tests
"""

from datetime import datetime, timedelta
from pathlib import Path
import shutil

import pandas as pd
import pytest


@pytest.fixture(autouse=True)
def protect_templates(tmp_path, monkeypatch):
    """템플릿을 테스트로부터 보호

    테스트 중 템플릿이 생성/덮어쓰기 되지 않도록
    임시 폴더를 사용합니다.
    """
    import po_generator.config as config
    import po_generator.template_engine as template_engine

    # 원본 템플릿 경로 저장
    original_template_dir = config.TEMPLATE_DIR

    # 임시 템플릿 디렉토리 설정
    test_template_dir = tmp_path / "templates"
    test_template_dir.mkdir(exist_ok=True)

    # 원본 PO 템플릿이 있으면 복사
    if config.PO_TEMPLATE_FILE.exists():
        shutil.copy(config.PO_TEMPLATE_FILE, test_template_dir / "purchase_order.xlsx")

    # config 모듈의 경로를 임시 경로로 변경
    monkeypatch.setattr(config, 'TEMPLATE_DIR', test_template_dir)
    monkeypatch.setattr(config, 'PO_TEMPLATE_FILE', test_template_dir / "purchase_order.xlsx")

    # template_engine도 업데이트
    monkeypatch.setattr(template_engine, 'TEMPLATE_DIR', test_template_dir)
    monkeypatch.setattr(template_engine, 'PO_TEMPLATE_FILE', test_template_dir / "purchase_order.xlsx")

    yield

    # 테스트 후 정리는 pytest의 tmp_path가 자동 처리


@pytest.fixture
def valid_order_data() -> pd.Series:
    """유효한 주문 데이터"""
    return pd.Series({
        'RCK Order no.': 'ND-TEST-001',
        'Customer name': 'Test Customer',
        'Customer PO': 'CPO-12345',
        'Item name': 'Test Actuator',
        'Item qty': 2,
        'Model': 'NA-100',
        'ICO Unit': 1000000,
        'Sales Unit Price': 1000000,  # 거래명세표용 판매단가
        'Total ICO': 2000000,
        'Power supply': 'AC220V-1Ph-50Hz',
        'ALS': 'Y',
        'Requested delivery date': datetime.now() + timedelta(days=30),
        'Incoterms': 'EXW',
        '_시트구분': '국내',
    })


@pytest.fixture
def invalid_order_data_missing_fields() -> pd.Series:
    """필수 필드 누락 주문 데이터"""
    return pd.Series({
        'RCK Order no.': 'ND-TEST-002',
        'Customer name': '',  # 누락
        'Customer PO': 'CPO-12345',
        'Item name': 'Test Actuator',
        'Item qty': 2,
        'Model': '',  # 누락
        'ICO Unit': 1000000,
        '_시트구분': '국내',
    })


@pytest.fixture
def invalid_order_data_zero_ico() -> pd.Series:
    """ICO Unit이 0인 주문 데이터"""
    return pd.Series({
        'RCK Order no.': 'ND-TEST-003',
        'Customer name': 'Test Customer',
        'Customer PO': 'CPO-12345',
        'Item name': 'Test Actuator',
        'Item qty': 2,
        'Model': 'NA-100',
        'ICO Unit': 0,  # 0
        '_시트구분': '국내',
    })


@pytest.fixture
def invalid_order_data_past_delivery() -> pd.Series:
    """납기일이 과거인 주문 데이터"""
    return pd.Series({
        'RCK Order no.': 'ND-TEST-004',
        'Customer name': 'Test Customer',
        'Customer PO': 'CPO-12345',
        'Item name': 'Test Actuator',
        'Item qty': 2,
        'Model': 'NA-100',
        'ICO Unit': 1000000,
        'Requested delivery date': datetime.now() - timedelta(days=10),  # 과거
        '_시트구분': '국내',
    })


@pytest.fixture
def order_data_urgent_delivery() -> pd.Series:
    """납기일이 촉박한 주문 데이터"""
    return pd.Series({
        'RCK Order no.': 'ND-TEST-005',
        'Customer name': 'Test Customer',
        'Customer PO': 'CPO-12345',
        'Item name': 'Test Actuator',
        'Item qty': 2,
        'Model': 'NA-100',
        'ICO Unit': 1000000,
        'Requested delivery date': datetime.now() + timedelta(days=3),  # 3일 후
        '_시트구분': '국내',
    })


@pytest.fixture
def multiple_items_df(valid_order_data: pd.Series) -> pd.DataFrame:
    """다중 아이템 DataFrame"""
    item1 = valid_order_data.copy()
    item1['Item name'] = 'Actuator A'
    item1['Item qty'] = 1
    item1['ICO Unit'] = 500000
    item1['Sales Unit Price'] = 500000

    item2 = valid_order_data.copy()
    item2['Item name'] = 'Actuator B'
    item2['Item qty'] = 3
    item2['ICO Unit'] = 750000
    item2['Sales Unit Price'] = 750000

    return pd.DataFrame([item1, item2])
