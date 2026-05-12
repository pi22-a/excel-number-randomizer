"""
모태펀드 자동입력기
====================
엑셀 파일의 데이터를 읽어 한글(.hwpx) 문서에 자동으로 입력합니다.

[모드 A] 엑셀 + 한글 → 완성 한글  (Excel 설치 필요)
  모태자동입력.exe 엑셀파일.xlsx 한글파일.hwpx

[모드 B] JSON + 한글 → 완성 한글  (Excel 불필요)
  모태자동입력.exe fund_data.json 한글파일.hwpx

[모드 C] 엑셀만 → JSON 저장  (다른 PC로 데이터 이동 시)
  모태자동입력.exe 엑셀파일.xlsx
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


def ask_yes(prompt: str) -> bool:
    ans = input(prompt + " [y/n]: ").strip().lower()
    return ans in ('y', 'yes', '')


def load_fund_data_from_json(json_path: Path) -> dict:
    """JSON 파일에서 fund_data dict[str, FundData] 복원."""
    from extract_excel import ExcelExtractor, FundData, Section1Row, Section2Row, Section3Row

    raw = ExcelExtractor.load_json(str(json_path))
    fund_data = {}
    for ft, d in raw.items():
        fd = FundData(fund_type=d["fund_type"], unit=d["unit"])
        for key, val in d.items():
            if key in ("s1_rows", "s2_rows", "s3_rows"):
                continue
            if hasattr(fd, key):
                setattr(fd, key, val)
        fd.s1_rows = [Section1Row(**r) for r in d.get("s1_rows", [])]
        fd.s2_rows = [Section2Row(**r) for r in d.get("s2_rows", [])]
        fd.s3_rows = [Section3Row(**r) for r in d.get("s3_rows", [])]
        fund_data[ft] = fd
    return fund_data


def extract_from_excel(excel_path: Path) -> dict:
    """Excel 자동화로 fund_data 추출. 실패 시 visible 모드 재시도."""
    from extract_excel import ExcelExtractor

    ext = ExcelExtractor(str(excel_path))
    try:
        return ext.read_all(show_progress=True)
    except Exception as e1:
        print(f"\n  백그라운드 실행 실패: {e1}")
        print("  Excel 창 표시 모드로 재시도...")
        return ext._read_all_visible()


def main():
    print()
    print("=" * 60)
    print("   모태펀드 자동입력기")
    print("=" * 60)
    print()

    # ── 파일 경로 결정 ────────────────────────────────────────────────
    args = sys.argv[1:]

    excel_path = None
    json_path  = None
    hwpx_path  = None

    for a in args:
        p = Path(a.strip('"').strip("'"))
        if p.suffix.lower() in ('.xlsx', '.xlsm', '.xlsb', '.xls'):
            excel_path = p
        elif p.suffix.lower() == '.json':
            json_path = p
        elif p.suffix.lower() == '.hwpx':
            hwpx_path = p

    # ── 모드 판단 ─────────────────────────────────────────────────────
    # 아무것도 드롭 안 한 경우 → 대화형 안내
    if excel_path is None and json_path is None:
        print("사용 방법을 선택하세요:")
        print()
        print("  [1] 엑셀 + 한글 → 완성 한글  (Excel 설치 필요)")
        print("  [2] JSON + 한글 → 완성 한글  (Excel 불필요)")
        print("  [3] 엑셀 → JSON 저장          (데이터 이동용)")
        print()
        choice = input("번호 입력 (1/2/3): ").strip()

        if choice == "1":
            print()
            print("[엑셀 파일] 경로를 입력하세요.")
            excel_path = ask_path("  엑셀 파일: ")
        elif choice == "2":
            print()
            print("[JSON 파일] 경로를 입력하세요.")
            json_path = ask_path("  JSON 파일: ")
        elif choice == "3":
            print()
            print("[엑셀 파일] 경로를 입력하세요.")
            excel_path = ask_path("  엑셀 파일: ")
            hwpx_path = None  # HWPX 없이 JSON만 저장
        else:
            print("잘못된 입력입니다. 종료합니다.")
            input("\n엔터를 누르면 종료합니다...")
            sys.exit(1)

    # 모드 C: 엑셀만 → JSON 저장 (HWPX 없이)
    # 드래그로 엑셀만 떨어뜨린 경우에도 선택지 제공
    if excel_path is not None and hwpx_path is None and json_path is None:
        print(f"  엑셀: {excel_path.name}")
        print()
        print("한글 파일 없이 엑셀만 입력됐습니다.")
        print("  [1] 한글 파일 경로를 추가로 입력")
        print("  [2] JSON으로 저장 (나중에 Excel 없는 환경에서 사용)")
        print()
        sub = input("선택 (1/2): ").strip()
        if sub == "1":
            hwpx_path = ask_path("  한글 파일: ")
        else:
            pass  # JSON 저장 모드 진행

    # HWPX 경로 없으면 입력 요청 (JSON 저장 모드 제외)
    if hwpx_path is None and (excel_path is not None or json_path is not None):
        # JSON 저장 모드인지 확인 (엑셀만 있고 HWPX 입력 안 한 경우는 스킵)
        if json_path is not None:
            print()
            print("[한글 파일] 경로를 입력하세요.")
            hwpx_path = ask_path("  한글 파일: ")
        elif excel_path is not None:
            # 아직 HWPX 없으면 JSON 저장 모드로 확정됨 (위에서 처리됨)
            pass

    # ── 상태 출력 ─────────────────────────────────────────────────────
    print()
    print_line()
    if excel_path:
        print(f"  엑셀: {excel_path.name}")
    if json_path:
        print(f"  JSON: {json_path.name}")
    if hwpx_path:
        print(f"  한글: {hwpx_path.name}")
    if hwpx_path:
        out_path = hwpx_path.parent / (hwpx_path.stem + "_완성.hwpx")
        print(f"  출력: {out_path.name}")
    print_line()

    # ── 데이터 추출 / 로드 ────────────────────────────────────────────
    fund_data = None

    if json_path is not None:
        # 모드 B: JSON에서 로드
        print()
        print("[1/2] JSON 파일에서 펀드 데이터 로드 중...")
        try:
            fund_data = load_fund_data_from_json(json_path)
            print(f"  완료: {len(fund_data)}개 펀드 유형")
        except Exception as e:
            print(f"\n오류: JSON 로드 실패 — {e}")
            input("\n엔터를 누르면 종료합니다...")
            sys.exit(1)

    elif excel_path is not None:
        # 모드 A 또는 C: Excel 자동화
        print()
        print("[1/2] 엑셀에서 펀드 데이터 추출 중...")
        print("  Excel 창이 잠깐 열렸다 닫힙니다. 조작하지 마세요.")
        print()
        try:
            fund_data = extract_from_excel(excel_path)
            print(f"\n  완료: {len(fund_data)}개 펀드 유형 추출")
        except Exception as e:
            print(f"\n오류: 엑셀 데이터 추출 실패")
            print(f"  {e}")
            print()
            print("확인사항:")
            print("  - Microsoft Excel(데스크탑 버전)이 설치되어 있어야 합니다")
            print("  - 엑셀 파일이 다른 프로그램에서 열려있으면 닫아주세요")
            input("\n엔터를 누르면 종료합니다...")
            sys.exit(1)

        # 모드 C: JSON 저장 후 종료 (HWPX 없는 경우)
        if hwpx_path is None:
            json_save_path = excel_path.parent / (excel_path.stem + "_data.json")
            print()
            print(f"[2/2] JSON으로 저장 중: {json_save_path.name}")
            try:
                from extract_excel import ExcelExtractor
                ext = ExcelExtractor(str(excel_path))
                ext.export_json(str(json_save_path), fund_data)
                print()
                print("=" * 60)
                print("   JSON 저장 완료!")
                print(f"   저장 위치: {json_save_path}")
                print("=" * 60)
                print()
                print("이 JSON 파일을 Excel 없는 PC로 복사한 뒤,")
                print("한글 파일과 함께 이 프로그램에 드래그하면 됩니다.")
            except Exception as e:
                print(f"\n오류: JSON 저장 실패 — {e}")
            input("\n엔터를 누르면 종료합니다...")
            return

    # ── HWPX 입력 ────────────────────────────────────────────────────
    print()
    print("[2/2] 한글 문서에 데이터 입력 중...")

    try:
        from excel_to_hwpx import fill_hwpx, load_mapping

        mapping = load_mapping()
        if not mapping.get("sections"):
            print("\n오류: mapping.json을 찾을 수 없습니다.")
            print(f"  찾는 위치: {_BASE_DIR / 'mapping.json'}")
            input("\n엔터를 누르면 종료합니다...")
            sys.exit(1)

        fill_hwpx(str(hwpx_path), fund_data, str(out_path), mapping)

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
