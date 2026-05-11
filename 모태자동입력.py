"""
모태펀드 자동입력기
====================
엑셀 파일의 데이터를 읽어 한글(.hwpx) 문서에 자동으로 입력합니다.

사용법 1 - 드래그앤드롭 또는 인수 전달:
  모태자동입력.exe 엑셀파일.xlsx 한글파일.hwpx

사용법 2 - 그냥 실행하면 경로 입력 안내:
  모태자동입력.exe
"""

import sys
import json
import time
from pathlib import Path

# PyInstaller exe 경로 처리
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys._MEIPASS)
else:
    _BASE_DIR = Path(__file__).parent

sys.path.insert(0, str(_BASE_DIR))


def print_line():
    print("-" * 60)


def ask_path(prompt: str, must_exist: bool = True) -> Path:
    """파일 경로 입력 받기 (따옴표/공백 처리)."""
    while True:
        raw = input(prompt).strip().strip('"').strip("'")
        if not raw:
            print("  경로를 입력해주세요.")
            continue
        p = Path(raw)
        if must_exist and not p.exists():
            print(f"  파일을 찾을 수 없습니다: {p}")
            continue
        return p


def main():
    print()
    print("=" * 60)
    print("   모태펀드 자동입력기")
    print("=" * 60)
    print()

    # ── 파일 경로 결정 ────────────────────────────────────────────────
    args = sys.argv[1:]

    # 드래그앤드롭 또는 인수로 전달된 경우
    excel_path = None
    hwpx_path  = None

    for a in args:
        p = Path(a.strip('"').strip("'"))
        if p.suffix.lower() in ('.xlsx', '.xlsm', '.xlsb', '.xls'):
            excel_path = p
        elif p.suffix.lower() == '.hwpx':
            hwpx_path = p

    # 대화형 입력
    if excel_path is None:
        print("[1단계] 엑셀 파일 경로를 입력하세요.")
        print("  (파일을 이 창에 드래그해도 됩니다)")
        excel_path = ask_path("  엑셀 파일: ")

    if hwpx_path is None:
        print()
        print("[2단계] 한글(.hwpx) 파일 경로를 입력하세요.")
        hwpx_path = ask_path("  한글 파일: ")

    print()
    print_line()
    print(f"  엑셀: {excel_path.name}")
    print(f"  한글: {hwpx_path.name}")
    print_line()

    # 출력 경로
    out_path = hwpx_path.parent / (hwpx_path.stem + "_완성.hwpx")

    # ── 1단계: 엑셀 데이터 추출 ──────────────────────────────────────
    print()
    print("[1/2] 엑셀에서 펀드 데이터 추출 중...")
    print("  Excel 창이 잠깐 열렸다 닫힙니다. 조작하지 마세요.")
    print()

    try:
        from extract_excel import ExcelExtractor
        ext = ExcelExtractor(str(excel_path))

        try:
            fund_data = ext.read_all(show_progress=True)
        except Exception as e1:
            print(f"\n  백그라운드 실행 실패: {e1}")
            print("  Excel 창 표시 모드로 재시도...")
            fund_data = ext._read_all_visible()

    except Exception as e:
        print(f"\n오류: 엑셀 데이터 추출 실패")
        print(f"  {e}")
        print()
        print("확인사항:")
        print("  - Microsoft Excel(데스크탑 버전)이 설치되어 있어야 합니다")
        print("  - 엑셀 파일이 다른 프로그램에서 열려있으면 닫아주세요")
        input("\n엔터를 누르면 종료합니다...")
        sys.exit(1)

    print(f"\n  완료: {len(fund_data)}개 펀드 유형 추출")

    # ── 2단계: HWPX 자동 입력 ────────────────────────────────────────
    print()
    print("[2/2] 한글 문서에 데이터 입력 중...")

    try:
        from excel_to_hwpx import fill_hwpx, load_mapping
        from extract_excel import Section1Row, Section2Row, Section3Row, FundData as FD

        # fund_data를 dict[str, FundData]로 변환
        from dataclasses import asdict
        fund_data_obj = {}
        for ft, fd in fund_data.items():
            fund_data_obj[ft] = fd

        mapping = load_mapping()
        if not mapping.get("sections"):
            print("\n오류: mapping.json을 찾을 수 없습니다.")
            print(f"  찾는 위치: {_BASE_DIR / 'mapping.json'}")
            input("\n엔터를 누르면 종료합니다...")
            sys.exit(1)

        fill_hwpx(str(hwpx_path), fund_data_obj, str(out_path), mapping)

    except Exception as e:
        import traceback
        print(f"\n오류: 한글 문서 입력 실패")
        print(f"  {e}")
        traceback.print_exc()
        input("\n엔터를 누르면 종료합니다...")
        sys.exit(1)

    # ── 완료 ─────────────────────────────────────────────────────────
    print()
    print("=" * 60)
    print("   완료!")
    print(f"   저장 위치: {out_path}")
    print("=" * 60)
    print()
    print("한글에서 파일을 열어 내용을 확인하세요.")
    print()
    input("엔터를 누르면 종료합니다...")


if __name__ == "__main__":
    main()
