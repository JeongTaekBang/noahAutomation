"""
config 설정 폴백 테스트
=======================

배포판은 user_settings.py 없이 noah_config.ini로 경로를 잡는다.
우선순위(user_settings.py → noah_config.ini → 기본값)가 지켜지는지 검증한다.

특히 중요한 것: 개발 PC에 ini가 생겨도 기존 동작이 바뀌면 안 된다.
`OUTPUT_BASE_DIR = None`처럼 "None이 의미 있는 설정"을 ini가 덮어쓰면 출력 폴더가
조용히 바뀌므로, user_settings.py에 이름이 있으면 값이 None이어도 그것이 최종값이다.
"""

import sys
import types

import pytest

from po_generator import config


@pytest.fixture
def fake_user_settings(monkeypatch):
    """user_settings 모듈을 테스트용으로 교체"""
    def _install(**attrs):
        module = types.ModuleType("user_settings")
        for name, value in attrs.items():
            setattr(module, name, value)
        monkeypatch.setitem(sys.modules, "user_settings", module)
        return module
    return _install


@pytest.fixture
def no_user_settings(monkeypatch):
    """user_settings.py가 없는 상태 (배포판)

    sys.modules에 None을 넣으면 import 시 ImportError가 발생한다.
    """
    monkeypatch.setitem(sys.modules, "user_settings", None)


@pytest.fixture
def ini_values(monkeypatch):
    """noah_config.ini 값 주입"""
    def _install(**values):
        monkeypatch.setattr(config, "_INI_VALUES", dict(values))
    return _install


class TestLoadUserSetting:
    """_load_user_setting 우선순위"""

    def test_user_settings_wins_over_ini(self, fake_user_settings, ini_values):
        """user_settings.py가 ini보다 우선"""
        fake_user_settings(DATA_FOLDER=r"C:\from_user_settings")
        ini_values(data_folder=r"C:\from_ini")

        assert config._load_user_setting('DATA_FOLDER', None) == r"C:\from_user_settings"

    def test_user_settings_none_still_wins(self, fake_user_settings, ini_values):
        """user_settings.py의 None도 유효한 설정 — ini가 덮어쓰지 않는다"""
        fake_user_settings(OUTPUT_BASE_DIR=None)
        ini_values(output_base_dir=r"C:\from_ini")

        assert config._load_user_setting('OUTPUT_BASE_DIR', 'default') is None

    def test_ini_used_when_no_user_settings(self, no_user_settings, ini_values):
        """배포판 — user_settings.py가 없으면 ini 사용"""
        ini_values(data_folder=r"C:\from_ini")

        assert config._load_user_setting('DATA_FOLDER', None) == r"C:\from_ini"

    def test_ini_used_when_attribute_missing(self, fake_user_settings, ini_values):
        """user_settings.py는 있지만 해당 항목이 없으면 ini 사용"""
        fake_user_settings(SOMETHING_ELSE=1)
        ini_values(data_folder=r"C:\from_ini")

        assert config._load_user_setting('DATA_FOLDER', None) == r"C:\from_ini"

    def test_default_when_nothing_set(self, no_user_settings, ini_values):
        """둘 다 없으면 기본값"""
        ini_values()

        assert config._load_user_setting('DATA_FOLDER', 'fallback') == 'fallback'

    def test_ini_ignored_for_unmapped_key(self, no_user_settings, ini_values):
        """ini가 지원하지 않는 설정은 기본값 (공급자 정보 등은 ini로 바꾸지 않는다)"""
        ini_values(vat_rate_domestic='0.5')

        assert config._load_user_setting('VAT_RATE_DOMESTIC', 0.1) == 0.1


class TestReadIni:
    """_read_ini 파싱"""

    def test_reads_paths_section(self, tmp_path, monkeypatch):
        ini = tmp_path / "noah_config.ini"
        ini.write_text(
            "[paths]\n"
            "data_folder = C:\\NOAH\n"
            "output_base_dir = C:\\NOAH\n",
            encoding='utf-8',
        )
        monkeypatch.setattr(config, "_INI_FILE", ini)

        assert config._read_ini() == {
            'data_folder': r"C:\NOAH",
            'output_base_dir': r"C:\NOAH",
        }

    def test_missing_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "_INI_FILE", tmp_path / "nope.ini")

        assert config._read_ini() == {}

    def test_missing_section(self, tmp_path, monkeypatch):
        ini = tmp_path / "noah_config.ini"
        ini.write_text("[other]\nfoo = bar\n", encoding='utf-8')
        monkeypatch.setattr(config, "_INI_FILE", ini)

        assert config._read_ini() == {}

    def test_empty_values_dropped(self, tmp_path, monkeypatch):
        """빈 값은 '설정 안 함'으로 취급 — 빈 문자열이 경로로 쓰이면 안 된다"""
        ini = tmp_path / "noah_config.ini"
        ini.write_text("[paths]\ndata_folder =\noutput_base_dir =   \n", encoding='utf-8')
        monkeypatch.setattr(config, "_INI_FILE", ini)

        assert config._read_ini() == {}

    def test_percent_in_path(self, tmp_path, monkeypatch):
        """경로에 %가 있어도 깨지지 않는다 (interpolation=None)"""
        ini = tmp_path / "noah_config.ini"
        ini.write_text("[paths]\ndata_folder = C:\\100%_data\n", encoding='utf-8')
        monkeypatch.setattr(config, "_INI_FILE", ini)

        assert config._read_ini() == {'data_folder': r"C:\100%_data"}

    def test_utf8_bom(self, tmp_path, monkeypatch):
        """BOM이 붙어도 읽힌다

        메모장으로 ini를 편집하면 UTF-8 BOM이 붙는다. BOM을 처리하지 않으면
        첫 섹션 헤더가 '\ufeff[paths]'가 되어 설정 전체가 조용히 무시된다.
        (배포판 검증 중 실제로 발생)
        """
        ini = tmp_path / "noah_config.ini"
        ini.write_text("[paths]\ndata_folder = C:\\NOAH\n", encoding='utf-8-sig')
        monkeypatch.setattr(config, "_INI_FILE", ini)

        assert config._read_ini() == {'data_folder': r"C:\NOAH"}

    def test_malformed_file(self, tmp_path, monkeypatch):
        """깨진 ini로 전체가 죽지 않는다"""
        ini = tmp_path / "noah_config.ini"
        ini.write_text("이건 ini가 아닙니다\n= = =\n", encoding='utf-8')
        monkeypatch.setattr(config, "_INI_FILE", ini)

        assert config._read_ini() == {}
