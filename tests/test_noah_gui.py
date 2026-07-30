"""
noah_gui.py 테스트
===================

문서 종류 정의(`DOC_TYPES`)와 CLI 명령 조립(`build_command`)만 검증합니다.
Tk() 창은 만들지 않습니다 — 위젯 배선은 사람이 눈으로 보는 편이 빠르고,
여기서 잡고 싶은 건 **조용히 틀리는 것들**이다.

가장 중요한 건 `test_doc_scripts_are_packaged`다. GUI에 문서 종류를 추가하고
`build_portable_gui.APP_FILES`에 넣는 걸 잊으면, 개발 PC에서는 멀쩡한데
배포판에서만 그 버튼이 죽는다 (실제로 납기현황이 그랬다).
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

import noah_gui
from po_generator import config

PROJECT_ROOT = Path(noah_gui.__file__).resolve().parent


def _load_build_module():
    """cli_dist/build_portable_gui.py 로드 (패키지가 아니라 경로로 연다)

    모듈 레벨은 상수 정의뿐이라 임포트 부작용이 없다.
    """
    path = PROJECT_ROOT / "cli_dist" / "build_portable_gui.py"
    spec = importlib.util.spec_from_file_location("_build_portable_gui_probe", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def args_of(doc_key: str, ids: list[str], options: dict) -> list[str]:
    """build_command 결과에서 파이썬 실행 파일을 뺀 인자만 (경로는 PC마다 다르다)"""
    return noah_gui.build_command(doc_key, ids, options)[1:]


# === 문서 종류 정의 ========================================================

def test_doc_keys_are_unique():
    keys = [d['key'] for d in noah_gui.DOC_TYPES]
    assert len(keys) == len(set(keys))


def test_doc_entries_have_uniform_schema():
    """DOC_TYPES는 고정 레코드 표다 — 모든 항목이 같은 키를 가져야 한다

    선택 키를 허용하면 소비처마다 기본값을 다시 새겨야 하고(`.get(..., True)`),
    한 곳이 기본값을 틀리게 쓰면 조용히 동작이 갈린다.
    """
    expected = {'key', 'label', 'script', 'id_label', 'hint', 'out_attr', 'options', 'multi'}
    for doc in noah_gui.DOC_TYPES:
        assert set(doc) == expected, f"{doc.get('key')}: {set(doc) ^ expected}"


def test_doc_scripts_exist():
    """DOC_TYPES가 가리키는 CLI가 실제로 있는가 (배포 빌드 verify()와 같은 판정)"""
    assert noah_gui.missing_scripts() == []


def test_doc_scripts_are_packaged():
    """DOC_TYPES의 CLI가 전부 배포판에 들어가는가 — 이 테스트가 이번 버그의 회귀 방지"""
    build = _load_build_module()
    not_packaged = [d['script'] for d in noah_gui.DOC_TYPES
                    if d['script'] not in build.APP_FILES]
    assert not not_packaged, (
        f"배포판 APP_FILES에 빠진 CLI: {not_packaged} — "
        "받는 사람 PC에서 해당 버튼만 실패합니다."
    )


# === 배포판 버전 (BUILD_INFO.txt) ==========================================

def test_read_build_info_parses_version(tmp_path):
    info = tmp_path / "BUILD_INFO.txt"
    info.write_text(
        "version: 2026.07.30+326b87e\nbuilt: 2026-07-30 15:12\npackages:\n  pandas 2.3.3\n",
        encoding='utf-8',
    )
    assert noah_gui.read_build_info(info) == '2026.07.30+326b87e'


def test_read_build_info_missing_file_is_none(tmp_path):
    """개발 PC에는 BUILD_INFO가 없다 — None이 정상 동작(타이틀에 버전 생략)"""
    assert noah_gui.read_build_info(tmp_path / "없음.txt") is None


def test_read_build_info_without_version_line_is_none(tmp_path):
    info = tmp_path / "BUILD_INFO.txt"
    info.write_text("built: 2026-07-30\n", encoding='utf-8')
    assert noah_gui.read_build_info(info) is None


# === 배포판 핀 대조 헬퍼 ====================================================

def test_parse_pins_reads_pinned_lines():
    build = _load_build_module()
    pins = build.parse_pins("# 주석\npandas==2.3.3\n\net_xmlfile == 2.0.0  # 인라인 주석\n")
    assert pins == {'pandas': '2.3.3', 'et-xmlfile': '2.0.0'}


def test_parse_pins_rejects_ranges():
    """`>=`가 끼어들면 재현성 약속이 조용히 깨진다 — 파싱 단계에서 막아야 한다"""
    build = _load_build_module()
    with pytest.raises(ValueError):
        build.parse_pins("pandas>=2.0\n")


def test_canon_absorbs_spelling_differences():
    build = _load_build_module()
    assert build.canon('et_xmlfile') == 'et-xmlfile'
    assert build.canon('Python-Dateutil') == 'python-dateutil'


def test_requirements_pins_direct_deps():
    """실제 requirements.txt가 전부 핀이고 직접 의존성 4개를 담는가

    transitive 전수는 빌드의 3자 대조가 강제한다 — 여기 박으면 정당한
    의존성 변경마다 테스트가 깨지므로 직접 의존성만 본다.
    """
    build = _load_build_module()
    req = (PROJECT_ROOT / "cli_dist" / "requirements.txt").read_text(encoding='utf-8')
    pins = build.parse_pins(req)  # 범위 지정이 섞이면 여기서 ValueError
    assert {'pandas', 'openpyxl', 'xlwings', 'pywin32'} <= set(pins)


def test_doc_output_dirs_exist_in_config():
    """[출력 폴더 열기]가 참조하는 config 상수가 실제로 있는가"""
    missing = [d['out_attr'] for d in noah_gui.DOC_TYPES
               if not hasattr(config, str(d['out_attr']))]
    assert not missing


# === 명령 조립 =============================================================

def test_po_command():
    assert args_of('po', ['ND-0001', 'ND-0002'], {}) == \
        ['create_po.py', 'ND-0001', 'ND-0002']


def test_po_force():
    assert args_of('po', ['ND-0001'], {'force': True}) == \
        ['create_po.py', 'ND-0001', '--force']


def test_ts_defaults_to_no_mail():
    """메일 의도는 항상 명령에 남는다 (로그에 무엇을 눌렀는지 보이게)"""
    assert args_of('ts', ['DND-2026-0001'], {'merge': False, 'mail': False}) == \
        ['create_ts.py', 'DND-2026-0001', '--no-mail']


def test_ts_merge_and_mail():
    assert args_of('ts', ['DND-2026-0001'], {'merge': True, 'mail': True}) == \
        ['create_ts.py', 'DND-2026-0001', '--merge', '--mail']


def test_fi_po_mode():
    assert args_of('fi', ['26KPO00144'], {'fi_mode': 'po'}) == \
        ['create_fi.py', '--po', '26KPO00144']


def test_fi_po_mode_without_ids_lists_available():
    assert args_of('fi', [], {'fi_mode': 'po'}) == ['create_fi.py', '--po']


def test_fi_dn_mode():
    assert args_of('fi', ['DNO-2026-0001'], {'fi_mode': 'dn'}) == \
        ['create_fi.py', 'DNO-2026-0001']


# === 납기현황 ==============================================================

def test_ds_customer_lookup():
    cmd = args_of('ds', ['615-81-88675'],
                  {'ds_mode': 'customer', 'ds_all': False, 'mail': False})
    assert cmd == ['delivery_status.py', '615-81-88675', '--no-mail']


def test_ds_all_and_mail():
    cmd = args_of('ds', ['엔이에스'],
                  {'ds_mode': 'customer', 'ds_all': True, 'mail': True})
    assert cmd == ['delivery_status.py', '엔이에스', '--all', '--mail']


def test_ds_list_mode_ignores_input():
    """목록 모드는 입력란을 쓰지 않는다 — 남아 있는 글자가 인자로 새면 argparse가 죽는다"""
    cmd = args_of('ds', ['615-81-88675'],
                  {'ds_mode': 'list', 'ds_all': True, 'mail': True})
    assert cmd == ['delivery_status.py', '--list']


@pytest.mark.parametrize(('doc_key', 'ids', 'options', 'expected'), [
    ('ds', [], {'ds_mode': 'list'}, True),
    ('ds', ['615-81-88675'], {'ds_mode': 'customer'}, False),
    ('fi', [], {'fi_mode': 'po'}, True),
    ('fi', ['26KPO00144'], {'fi_mode': 'po'}, False),
    ('fi', [], {'fi_mode': 'dn'}, False),
    ('po', [], {}, False),
])
def test_list_mode_selected(doc_key, ids, options, expected):
    assert noah_gui.list_mode_selected(doc_key, ids, options) is expected


def test_ds_is_marked_single_entry():
    """납기현황 CLI는 거래처 하나만 받는다 — GUI가 이 표식으로 여러 줄 입력을 막는다

    표식이 빠지면 두 줄 입력이 그대로 넘어가 argparse가
    `unrecognized arguments`로 죽고, 사용자는 이유를 알 수 없다.
    """
    assert noah_gui.DOC_BY_KEY['ds']['multi'] is False


def test_ds_command_round_trips_through_cli_parser():
    """GUI가 만든 ds 명령을 실제 CLI 파서가 받는가 — multi=False 미러의 원본 대조

    `multi: False`는 delivery_status.py의 `customer`(nargs='?')를 미러링한 값이다.
    CLI 쪽 인자 정의가 바뀌어 미러가 어긋나면 여기서 잡는다
    (DOC_TYPES ↔ APP_FILES를 대조하는 위 테스트와 같은 패턴).
    """
    import delivery_status  # 무겁다(pandas) — 이 대조 테스트에서만 쓴다

    parser = delivery_status.create_argument_parser()

    argv = args_of('ds', ['615-81-88675'],
                   {'ds_mode': 'customer', 'ds_all': True, 'mail': True})[1:]
    args = parser.parse_args(argv)
    assert args.customer == '615-81-88675'

    argv = args_of('ds', ['무시됨'], {'ds_mode': 'list'})[1:]
    assert parser.parse_args(argv).show_list is True

    with pytest.raises(SystemExit):  # 두 곳은 CLI도 거부한다 — GUI 차단이 미리 막는 그 오류
        parser.parse_args(['거래처1', '거래처2'])
