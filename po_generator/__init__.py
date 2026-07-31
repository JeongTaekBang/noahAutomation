"""
NOAH Purchase Order Auto-Generator Package
==========================================

RCK Order No.를 입력하면 NOAH_PO_Lists.xlsx에서 해당 데이터를 읽어
자동으로 발주서(Purchase Order + Description)를 생성합니다.
"""

__version__ = "2.5.0"

# 여기서 하위 모듈을 재export하지 않는다 (의도적으로 비워 둔 것이다).
#
# 예전에는 config/utils/validators/history/excel_generator를 전부 eager import 했다.
# 그러면 `from po_generator.config import DB_FILE` 한 줄에도 패키지 __init__이 먼저 돌아
# **pandas·openpyxl·xlwings(→pywin32/COM)가 통째로 딸려온다.**
#
# 대시보드는 config.DB_FILE과 db_schema 세 개만 쓰는데(둘 다 표준 라이브러리만 의존),
# 그 때문에 Excel COM 라이브러리를 로드하고 배포판에도 실어야 했다. 읽기 전용 대시보드가
# 엑셀 라이브러리 import 실패로 안 뜨는 건 말이 안 된다.
#
# 재export되던 이름(create_po_workbook 등)을 패키지 루트에서 가져다 쓰는 코드는
# 저장소 전체에 0건이었다 — `from po_generator import config`처럼 하위 모듈을
# 직접 가져오는 형태뿐이고, 그건 이 파일이 비어 있어도 그대로 동작한다.
#
# 필요한 것은 `from po_generator.<모듈> import ...`로 직접 가져올 것.
