"""
엑셀 데이터를 읽어 HWPX 한글 문서에 자동으로 입력하는 메인 프로그램.

사용법:
  # 1단계: 엑셀에서 전체 펀드 데이터 추출 (Excel 설치 필요, 1회만 실행)
  python excel_to_hwpx.py --extract <엑셀파일.xlsx> [--json fund_data.json]

  # 2단계: HWPX에 데이터 입력
  python excel_to_hwpx.py --fill <입력.hwpx> [--json fund_data.json] [--out 출력.hwpx]

  # 한 번에 실행 (Excel + HWPX 모두 있을 때)
  python excel_to_hwpx.py --excel <파일.xlsx> --hwpx <파일.hwpx> [--out 출력.hwpx]

매핑 정보는 mapping.json에서 관리합니다.
analyze_hwpx.py를 먼저 실행하여 mapping.json을 생성하세요.
"""

import sys
import json
import argparse
from pathlib import Path

from extract_excel import ExcelExtractor, FundData
from hwpx_utils import HwpxEditor


# ── PyInstaller exe 실행 시 경로 처리 ─────────────────────────────────────
# exe로 빌드되면 __file__ 대신 sys._MEIPASS 사용
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys._MEIPASS)
else:
    _BASE_DIR = Path(__file__).parent


# ── 숫자 포매팅 ────────────────────────────────────────────────────────────

def fmt_num(value: float, decimals: int = 0, comma: bool = True) -> str:
    """숫자를 문자열로 포맷. 예: 1234567.8 → '1,234,568'"""
    if decimals == 0:
        result = int(round(value))
        return f"{result:,}" if comma else str(result)
    else:
        result = round(value, decimals)
        s = f"{result:,.{decimals}f}" if comma else f"{result:.{decimals}f}"
        return s

def fmt_pct(value: float, decimals: int = 1) -> str:
    """비율을 퍼센트 문자열로. 예: 0.712 → '71.2%'"""
    return f"{value * 100:.{decimals}f}%"


# ── 매핑 구조 ─────────────────────────────────────────────────────────────
#
# mapping.json 구조 (analyze_hwpx.py가 템플릿 생성, 사람이 채워 넣음):
#
# {
#   "section_id_field": "fund_type",   // 섹션 식별에 쓸 FundData 필드
#   "sections": [
#     {
#       "fund_type": "마이크로VC",
#       "section_keyword": "마이크로VC편",  // HWPX 안에서 섹션 식별 키워드
#       "text_replacements": [
#         {
#           "description": "총 조합수",
#           "pattern": r"(\d+)개 조합",    // 정규식
#           "field": "s1_total_count",      // FundData 필드명
#           "format": "int"                 // int / float1 / float2 / pct1 / pct2
#         },
#         ...
#       ],
#       "table_cells": [
#         {
#           "description": "섹션1 2004년 조합수",
#           "table_idx": 0,   // 섹션 내 표 인덱스 (0-based)
#           "row": 1,
#           "col": 1,
#           "data_path": "s1_rows.0.count",  // FundData 경로
#           "format": "int"
#         },
#         ...
#       ]
#     },
#     ...
#   ]
# }
#
# ※ 실제 HWP 파일을 받기 전까지 이 매핑은 비어 있습니다.
#   analyze_hwpx.py를 실행하여 구조를 파악한 후 채워 넣으세요.


MAPPING_PATH = _BASE_DIR / "mapping.json"


def load_mapping() -> dict:
    if not MAPPING_PATH.exists():
        return {"sections": []}
    with open(MAPPING_PATH, encoding='utf-8') as f:
        return json.load(f)


def get_field_value(fd: FundData, data_path: str) -> float:
    """
    'field' 또는 'field.idx.subfield' 형태의 경로로 FundData에서 값 추출.
    예: 's1_total_count' → fd.s1_total_count
        's1_rows.0.count' → fd.s1_rows[0].count
        's2_rows.1.invested' → fd.s2_rows[1].invested
    """
    parts = data_path.split('.')
    obj = fd
    for part in parts:
        if isinstance(obj, list):
            obj = obj[int(part)]
        elif hasattr(obj, part):
            obj = getattr(obj, part)
        elif isinstance(obj, dict):
            obj = obj[part]
        else:
            raise ValueError(f"경로 '{data_path}'에서 '{part}' 접근 실패")
    return obj


def apply_format(value: float, fmt: str) -> str:
    """포맷 문자열에 따라 값을 문자열로 변환."""
    if fmt == "int":
        return fmt_num(value, 0)
    elif fmt == "float1":
        return fmt_num(value, 1)
    elif fmt == "float2":
        return fmt_num(value, 2)
    elif fmt == "pct1":
        return fmt_pct(value, 1)
    elif fmt == "pct2":
        return fmt_pct(value, 2)
    else:
        return str(value)


# ── 핵심 처리 함수 ─────────────────────────────────────────────────────────

def fill_hwpx(hwpx_path: str, fund_data: dict, output_path: str,
              mapping: dict = None):
    """
    fund_data (dict[str, FundData]) 와 mapping 정보를 이용해
    HWPX의 값을 교체하고 output_path로 저장.
    """
    if mapping is None:
        mapping = load_mapping()

    editor = HwpxEditor(hwpx_path)
    editor.load()

    sections_map = mapping.get("sections", [])
    if not sections_map:
        print("경고: mapping.json에 섹션 정보가 없습니다.")
        print("  → analyze_hwpx.py를 먼저 실행하여 매핑을 설정하세요.")
        editor.save(output_path)
        return

    total_replaced = 0

    for sec_cfg in sections_map:
        fund_type    = sec_cfg.get("fund_type", "")
        keyword      = sec_cfg.get("section_keyword", "")
        sec_file     = sec_cfg.get("section_file", "")   # 예: "Contents/section0.xml"

        if fund_type not in fund_data:
            print(f"  건너뜀: '{fund_type}' — 엑셀 데이터 없음")
            continue

        fd = fund_data[fund_type]
        print(f"  처리: {fund_type} (파일: {sec_file or '자동'}, 키워드: {keyword})")

        # 텍스트 교체
        for tr in sec_cfg.get("text_replacements", []):
            try:
                val = get_field_value(fd, tr["data_path"])
                fmt_val = apply_format(val, tr.get("format", "int"))
                n = editor.replace_in_paragraphs(
                    section_keyword=keyword,
                    pattern=tr["pattern"],
                    replacement=fmt_val,
                    section_file=tr.get("section_file", sec_file),
                )
                total_replaced += n
            except Exception as e:
                print(f"    경고: 텍스트 교체 실패 ({tr.get('description','?')}) — {e}")

        # 표 셀 교체
        for tc in sec_cfg.get("table_cells", []):
            try:
                val = get_field_value(fd, tc["data_path"])
                fmt_val = apply_format(val, tc.get("format", "int"))
                ok = editor.set_table_cell(
                    table_idx=tc["table_idx"],
                    row=tc["row"],
                    col=tc["col"],
                    value=fmt_val,
                    section_keyword=keyword,
                    section_file=tc.get("section_file", sec_file),
                )
                if not ok:
                    print(f"    경고: 표 셀 교체 실패 ({tc.get('description','?')})")
                else:
                    total_replaced += 1
            except Exception as e:
                print(f"    경고: 표 셀 교체 실패 ({tc.get('description','?')}) — {e}")

    print(f"\n총 {total_replaced}개 항목 교체 완료")
    editor.save(output_path)


# ── 매핑 템플릿 생성 ───────────────────────────────────────────────────────

def generate_mapping_template(fund_types: list[str], output_path: str = None):
    """
    analyze_hwpx.py의 분석 결과를 바탕으로 매핑 템플릿 JSON 생성.
    실제 section_keyword와 table_idx는 분석 후 직접 입력해야 함.
    """
    template = {
        "_comment": "analyze_hwpx.py로 구조 파악 후 section_keyword와 table_idx를 채우세요.",
        "sections": []
    }

    for ft in fund_types:
        sec = {
            "fund_type": ft,
            "section_keyword": f"{ft}편",  # 실제 키워드로 수정 필요
            "text_replacements": [
                {
                    "description": "총 조합수 (섹션1 요약)",
                    "pattern": r"PLACEHOLDER_COUNT개 조합",
                    "data_path": "s1_total_count",
                    "format": "int"
                },
                {
                    "description": "총 결성액 (섹션1 요약)",
                    "pattern": r"PLACEHOLDER_FORMED억원 결성",
                    "data_path": "s1_total_formed",
                    "format": "float1"
                },
                {
                    "description": "총 투자액 (섹션1 요약)",
                    "pattern": r"PLACEHOLDER_INVESTED억원 투자",
                    "data_path": "s1_total_invested",
                    "format": "float1"
                },
            ],
            "table_cells": [
                {
                    "description": "섹션2 해산완료 조합수",
                    "table_idx": 0,  # 실제 표 인덱스로 수정 필요
                    "row": 0,
                    "col": 1,
                    "data_path": "s2_rows.0.count",
                    "format": "int"
                },
                {
                    "description": "섹션2 운용중 조합수",
                    "table_idx": 0,
                    "row": 1,
                    "col": 1,
                    "data_path": "s2_rows.1.count",
                    "format": "int"
                },
            ]
        }
        template["sections"].append(sec)

    out = output_path or str(MAPPING_PATH)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(template, f, ensure_ascii=False, indent=2)
    print(f"매핑 템플릿 생성: {out}")
    print("  → section_keyword와 table_cells를 실제 값으로 수정하세요.")


# ── CLI ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="엑셀 데이터를 HWPX 한글 문서에 자동 입력"
    )
    parser.add_argument("--extract", metavar="EXCEL",
                        help="엑셀 파일에서 전체 펀드 데이터 추출 (Excel 필요)")
    parser.add_argument("--fill", metavar="HWPX",
                        help="HWPX 파일에 데이터 입력")
    parser.add_argument("--excel", metavar="EXCEL",
                        help="엑셀 파일 (--extract + --fill 통합)")
    parser.add_argument("--hwpx", metavar="HWPX",
                        help="HWPX 파일 (--fill과 함께 사용)")
    parser.add_argument("--json", metavar="JSON", default="fund_data.json",
                        help="중간 JSON 파일 경로 (기본: fund_data.json)")
    parser.add_argument("--out", metavar="OUTPUT",
                        help="출력 HWPX 파일 경로")
    parser.add_argument("--mapping", metavar="MAPPING",
                        help="매핑 JSON 파일 경로 (기본: mapping.json)")
    parser.add_argument("--gen-mapping", action="store_true",
                        help="매핑 템플릿 JSON 생성 후 종료")

    args = parser.parse_args()

    # 매핑 템플릿 생성 모드
    if args.gen_mapping:
        generate_mapping_template(ExcelExtractor.FUND_TYPES)
        return

    # 엑셀 추출 모드
    if args.extract or args.excel:
        excel_path = args.extract or args.excel
        ext = ExcelExtractor(excel_path)
        print(f"엑셀 데이터 추출 중: {excel_path}")
        data = ext.read_all()
        ext.export_json(args.json, data)

    # HWPX 입력 모드
    hwpx_path = args.fill or args.hwpx
    if hwpx_path:
        json_path = args.json
        out_path  = args.out or hwpx_path.replace('.hwpx', '_filled.hwpx')

        print(f"\nHWPX 입력 시작: {hwpx_path}")
        fund_data_raw = ExcelExtractor.load_json(json_path)

        # JSON → dict[str, FundData] 변환
        # (JSON은 dict 형태, FundData로 역변환)
        from extract_excel import Section1Row, Section2Row, Section3Row
        fund_data = {}
        for ft, d in fund_data_raw.items():
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

        mapping = None
        if args.mapping:
            with open(args.mapping, encoding='utf-8') as f:
                mapping = json.load(f)

        fill_hwpx(hwpx_path, fund_data, out_path, mapping)

    if not any([args.extract, args.excel, args.fill, args.hwpx, args.gen_mapping]):
        parser.print_help()


if __name__ == "__main__":
    main()
