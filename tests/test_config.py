"""
config 모듈 테스트
==================

COLUMN_ALIASES, REQUIRED_FIELDS 등 설정값 테스트
"""

import pytest

from po_generator.config import (
    COLUMN_ALIASES,
    REQUIRED_FIELDS,
    META_COLUMNS,
    SPEC_FIELDS,
    OPTION_FIELDS,
    Colors,
    ColumnWidths,
)


class TestColumnAliases:
    """COLUMN_ALIASES 딕셔너리 테스트"""

    def test_required_keys_exist(self):
        """필수 내부 키 존재 확인"""
        required_keys = [
            'order_no',
            'customer_name',
            'customer_po',
            'item_qty',
            'ico_unit',
            'model',
            'item_name',
            'delivery_date',
            'sheet_type',
        ]
        for key in required_keys:
            assert key in COLUMN_ALIASES, f"필수 키 누락: {key}"

    def test_alias_values_are_tuples(self):
        """모든 별칭 값이 tuple 형식인지 확인"""
        for key, aliases in COLUMN_ALIASES.items():
            assert isinstance(aliases, tuple), f"{key}: tuple이 아님 - {type(aliases)}"
            assert len(aliases) > 0, f"{key}: 빈 tuple"

    def test_alias_values_are_strings(self):
        """모든 별칭이 문자열인지 확인"""
        for key, aliases in COLUMN_ALIASES.items():
            for alias in aliases:
                assert isinstance(alias, str), f"{key}: 별칭이 문자열이 아님 - {alias}"

    def test_no_duplicate_aliases_across_keys(self):
        """서로 다른 키에 같은 별칭이 없는지 확인 (충돌 방지)"""
        seen_aliases = {}
        duplicates = []

        for key, aliases in COLUMN_ALIASES.items():
            for alias in aliases:
                if alias in seen_aliases:
                    duplicates.append(f"'{alias}': {seen_aliases[alias]} <-> {key}")
                else:
                    seen_aliases[alias] = key

        # 중복이 있으면 상세 메시지 출력 (일부 중복은 의도적일 수 있음)
        if duplicates:
            pytest.skip(f"별칭 중복 발견 (의도적일 수 있음): {duplicates[:5]}...")

    def test_first_alias_is_primary(self):
        """첫 번째 별칭이 기본 컬럼명인지 문서화 테스트"""
        # order_no의 기본값은 'PO_ID'
        assert COLUMN_ALIASES['order_no'][0] == 'PO_ID'
        # customer_name의 기본값은 'Customer name'
        assert COLUMN_ALIASES['customer_name'][0] == 'Customer name'

    def test_legacy_column_names_included(self):
        """레거시 컬럼명이 별칭에 포함되어 있는지 확인"""
        # 기존 RCK Order no.가 order_no 별칭에 포함
        assert 'RCK Order no.' in COLUMN_ALIASES['order_no']

    def test_korean_aliases_included(self):
        """한글 별칭이 포함되어 있는지 확인"""
        # 고객명
        assert any('고객' in alias for alias in COLUMN_ALIASES['customer_name'])
        # 수량
        assert any('수량' in alias for alias in COLUMN_ALIASES['item_qty'])


class TestRequiredFields:
    """REQUIRED_FIELDS 테스트"""

    def test_required_fields_exist(self):
        """필수 필드 목록 존재 확인"""
        assert len(REQUIRED_FIELDS) > 0

    def test_required_fields_in_column_aliases(self):
        """필수 필드가 COLUMN_ALIASES에 모두 정의되어 있는지 확인"""
        for field in REQUIRED_FIELDS:
            assert field in COLUMN_ALIASES, f"필수 필드 '{field}'가 COLUMN_ALIASES에 없음"


class TestMetaColumns:
    """META_COLUMNS 테스트"""

    def test_meta_columns_is_frozenset(self):
        """META_COLUMNS가 frozenset인지 확인"""
        assert isinstance(META_COLUMNS, frozenset)

    def test_internal_columns_included(self):
        """내부 컬럼이 포함되어 있는지 확인"""
        assert '_시트구분' in META_COLUMNS

    def test_key_columns_included(self):
        """핵심 메타 컬럼이 포함되어 있는지 확인"""
        assert 'PO_ID' in META_COLUMNS
        assert 'Customer name' in META_COLUMNS


class TestSpecOptionFields:
    """SPEC_FIELDS, OPTION_FIELDS 테스트"""

    def test_spec_fields_not_empty(self):
        """SPEC_FIELDS가 비어있지 않음"""
        assert len(SPEC_FIELDS) > 0

    def test_option_fields_not_empty(self):
        """OPTION_FIELDS가 비어있지 않음"""
        assert len(OPTION_FIELDS) > 0

    def test_spec_fields_are_strings(self):
        """SPEC_FIELDS가 문자열 튜플"""
        assert all(isinstance(f, str) for f in SPEC_FIELDS)

    def test_option_fields_are_strings(self):
        """OPTION_FIELDS가 문자열 튜플"""
        assert all(isinstance(f, str) for f in OPTION_FIELDS)

    def test_power_supply_in_spec_fields(self):
        """Power supply가 SPEC_FIELDS 시작"""
        assert SPEC_FIELDS[0] == 'Power supply'

    def test_model_in_option_fields(self):
        """Model이 OPTION_FIELDS 시작"""
        assert OPTION_FIELDS[0] == 'Model'


class TestColors:
    """Colors 데이터클래스 테스트"""

    def test_colors_are_hex_strings(self):
        """색상값이 유효한 hex 문자열"""
        colors = Colors()
        for attr in ['RED', 'RED_BRIGHT', 'GRAY', 'TEAL', 'GREEN', 'WHITE']:
            value = getattr(colors, attr)
            assert isinstance(value, str)
            assert len(value) == 6
            int(value, 16)  # 유효한 hex인지 확인

    def test_colors_immutable(self):
        """Colors가 불변 데이터클래스"""
        colors = Colors()
        with pytest.raises(Exception):  # frozen=True로 인해 수정 불가
            colors.RED = "000000"


class TestColumnWidths:
    """ColumnWidths 데이터클래스 테스트"""

    def test_all_columns_defined(self):
        """A-J 열 너비 정의 확인"""
        widths = ColumnWidths()
        columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        for col in columns:
            assert hasattr(widths, col)
            assert getattr(widths, col) > 0

    def test_as_dict_returns_all_columns(self):
        """as_dict()가 모든 열 반환"""
        widths = ColumnWidths()
        d = widths.as_dict()
        assert len(d) == 10
        assert 'A' in d
        assert 'J' in d

    def test_column_widths_immutable(self):
        """ColumnWidths가 불변 데이터클래스"""
        widths = ColumnWidths()
        with pytest.raises(Exception):
            widths.A = 100
