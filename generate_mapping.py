"""
generate_mapping.py - mapping.json 완전 자동 생성

usage: python generate_mapping.py <randomized.hwpx> [--out mapping.json]

HWPX 파일의 실제 표 구조를 읽어 table_cells 매핑을 자동 생성합니다.
text_replacements 패턴은 HWPX 텍스트에서 자동 추출을 시도하고,
실패한 경우 "TODO" 마커를 남깁니다.
"""

import sys
import json
import zipfile
import re
import argparse
from pathlib import Path
from lxml import etree


# ─────────────────────────────────────────────────────────────────────────────
# HWPX 로딩 유틸
# ─────────────────────────────────────────────────────────────────────────────

SECTION_RE = re.compile(r'^Contents/section\d+\.xml$', re.IGNORECASE)
TABLE_TAGS = {'tbl', 'table'}
ROW_TAGS   = {'tr', 'row'}
CELL_TAGS  = {'tc', 'td', 'cell'}


def _local(tag: str) -> str:
    return tag.split('}')[-1].lower() if '}' in tag else tag.lower()


def _all_text(elem) -> str:
    return ''.join(elem.itertext())


def load_sections(hwpx_path: str) -> dict:
    """HWPX ZIP에서 section XML을 파싱. {filename: etree_root}"""
    sections = {}
    with zipfile.ZipFile(hwpx_path) as zf:
        for name in sorted(zf.namelist()):
            if SECTION_RE.match(name):
                data = zf.read(name)
                try:
                    root = etree.fromstring(data)
                    sections[name] = root
                    print(f"  로드: {name}  ({len(data):,} bytes)")
                except etree.XMLSyntaxError as e:
                    print(f"  경고: {name} 파싱 실패 — {e}")
    return sections


def get_tables(root) -> list:
    """섹션 루트에서 표 노드 목록 반환."""
    return [n for n in root.iter() if _local(n.tag) in TABLE_TAGS]


def get_data_rows(tbl) -> list:
    """표에서 행 노드 목록 반환 (헤더 포함)."""
    return [n for n in tbl if _local(n.tag) in ROW_TAGS]


def get_row_cell_count(tbl) -> int:
    """표의 데이터 행 수 (row 0 = 헤더 제외)."""
    rows = get_data_rows(tbl)
    return max(0, len(rows) - 1)  # 헤더 제외


# ─────────────────────────────────────────────────────────────────────────────
# 섹션1 / 섹션2 / 섹션3 컬럼 → FundData 필드 매핑
# ─────────────────────────────────────────────────────────────────────────────

# (col_idx, field_name, format)
S1_COLS = [
    (1, 'count',       'int'),
    (2, 'formed',      'float1'),
    (3, 'moat',        'float1'),
    (4, 'moat_ratio',  'pct1'),
    (5, 'companies',   'int'),
    (6, 'invested',    'float1'),
    (7, 'per_company', 'float1'),
]

S2_COLS = S1_COLS  # 동일한 열 구조

S3_COLS = [
    (1, 'invested',   'float1'),
    (2, 'capacity',   'float1'),
    # col 3 없음: HWPX 섹션3 표는 3열(0=연도, 1=투자금액, 2=투자여력(누계))만 존재
    # cumulative(누적투자금액)은 별도 열 없이 표기되지 않으므로 제거
]


# ─────────────────────────────────────────────────────────────────────────────
# 펀드 설정 (HWPX 분석 결과 기반)
#
# 형식: (fund_type_excel, section_file_key, section_keyword,
#        s1_table_indices, s2_table_index, s3_table_index)
#
# section_file_key: "s0" = Contents/section0.xml
#                   "s1" = Contents/section1.xml
#
# s1_table_indices: 리스트 (복합 펀드는 여러 개)
# s2/s3_table_index: 정수 (None이면 해당 섹션 없음)
#
# 세컨더리(복합), IP직접+특허, 과기(SaaS 등) 같은 복합 섹션은
# 각 Excel 펀드 유형이 별도 s1 표를 가짐.
# s2/s3는 복합 섹션 내 공통 표이므로 첫 번째 펀드만 쓰고
# 나머지는 None 처리 (중복 덮어쓰기 방지).
# ─────────────────────────────────────────────────────────────────────────────

FUND_CONFIGS = [
    # ── section0.xml ────────────────────────────────────────────────────────
    # fund_type_excel,              file_key, keyword,          s1_indices,  s2,   s3
    ("창업초기",                    "s0", "창업초기(전체)",      [6],         7,    8),
    ("창업초기(벤처투자조합등)",    "s0", "창업초기(벤처",       [14],        15,   16),
    ("창업초기(개인투자조합)",      "s0", "개인투자조합",        [22],        23,   24),
    ("청년창업",                    "s0", "청년창업",            [29],        30,   31),
    ("마이크로VC",                  "s0", "마이크로VC",          [37],        38,   39),
    # 엔젤 모펀드(44,45,46), 엔젤투자매칭(51,52,53) → ExcelExtractor.FUND_TYPES에 없음 → 생략
    ("재기지원",                    "s0", "재기지원",            [59],        60,   61),
    ("4차산업혁명",                 "s0", "4차산업혁명",         [66],        67,   68),
    ("조선업구조개선",              "s0", "조선업구조개선",      [73],        74,   75),
    ("지방기업",                    "s0", "지방기업",            [80],        81,   82),
    ("여성벤처",                    "s0", "여성벤처",            [88],        89,   90),
    ("소셜임팩트",                  "s0", "소셜임팩트",          [96],        97,   98),
    ("기술지주",                    "s0", "기술지주",            [103],       104,  105),
    ("혁신성장",                    "s0", "혁신성장",            [110],       111,  112),
    ("스케일업",                    "s0", "스케일업",            [117],       118,  119),
    ("M&A",                         "s0", "M&A",                 [124],       125,  126),
    # 세컨더리(복합): s1 3개, s2/s3 공유 → 첫 번째만 s2/s3 기록
    ("세컨더리",                    "s0", "세컨더리",            [132],       135,  136),
    ("엔젤세컨더리",                "s0", "세컨더리",            [133],       None, None),
    ("LP지분유동화",                "s0", "세컨더리",            [134],       None, None),
    ("소재부품장비",                "s0", "소재부품장비",        [142],       143,  144),

    # ── section1.xml ────────────────────────────────────────────────────────
    ("버팀목",                      "s1", "버팀목",              [1],         2,    3),
    ("해외진출",                    "s1", "해외진출",            [8],         9,    10),
    ("IP·문화산업",                 "s1", "IP",                  [16],        17,   18),
    ("문화일반, FI유치",            "s1", "문화일반",            [24],        25,   26),
    ("수출",                        "s1", "수출",                [32],        33,   34),
    ("신기술",                      "s1", "신기술",              [40],        41,   42),
    ("지역·청년·일자리",            "s1", "지역",                [48],        49,   50),
    ("M&A·세컨더리",                "s1", "M&A",                 [56],        57,   58),
    ("아시아문화중심도시육성",      "s1", "아시아문화",          [63],        64,   65),
    ("스포츠산업",                  "s1", "스포츠산업",          [70],        71,   72),
    ("스포츠프로젝트",              "s1", "스포츠프로젝트",      [78],        79,   80),
    ("관광기업육성",                "s1", "관광기업",            [84],        85,   86),
    ("한국영화 메인투자",           "s1", "한국영화 메인",       [92],        93,   94),
    ("중저예산한국영화",            "s1", "중저예산",            [100],       101,  102),
    ("한국영화 개봉촉진",           "s1", "개봉촉진",            [108],       109,  110),
    ("국토교통혁신",                "s1", "국토교통",            [114],       115,  116),
    ("도시재생",                    "s1", "도시재생",            [121],       122,  123),
    ("대학창업",                    "s1", "대학창업",            [128],       129,  130),
    ("사회적기업",                  "s1", "사회적기업",          [135],       136,  137),
    # IP직접+특허: s1/s2/s3 각 2개
    ("IP직접투자",                  "s1", "IP직접",              [143],       144,  145),
    ("특허기술사업화",              "s1", "특허기술",            [146],       147,  148),
    # 과기(SaaS 등): s1/s2/s3 각 4개 → 첫 번째만 s2/s3
    ("SaaS",                        "s1", "SaaS",                [156],       157,  158),
    ("사이버보안",                  "s1", "사이버보안",          [159],       160,  161),
    ("메타버스",                    "s1", "메타버스",            [162],       163,  164),
    ("공공기술사업화",              "s1", "공공기술사업화",      [165],       166,  167),
    ("보건",                        "s1", "보건",                [178],       179,  180),
    ("미래환경산업",                "s1", "미래환경",            [186],       187,  188),
]

FILE_KEY_MAP = {
    "s0": "Contents/section0.xml",
    "s1": "Contents/section1.xml",
}


# ─────────────────────────────────────────────────────────────────────────────
# 텍스트 패턴 추출 (섹션 헤더 표에서 자동 추출 시도)
# ─────────────────────────────────────────────────────────────────────────────

# 결성 및 투자 요약문 패턴 예시:
#   "'05~'26년, 163개 조합 66,307억원 결성, 24,018억원 투자"
# 각 숫자를 캡처하는 정규식을 자동으로 만들기 위해
# 실제 HWPX 텍스트에서 숫자를 (\d[\d,\.]*) 로 치환

NUM_PAT = re.compile(r'\d[\d,.]*')


def make_pattern(text: str) -> str:
    """텍스트에서 숫자를 (\\d[\\d,.]*) 캡처그룹으로 치환하여 정규식 패턴 생성."""
    escaped = re.escape(text)
    # re.escape 후 숫자 부분을 패턴으로 복원
    result = re.sub(r'(\\ |\\\d)+[\d,\\.]*', r'(\\d[\\d,.]*)', escaped)
    # 더 간단하게: 숫자 위치만 바꿈
    return NUM_PAT.sub(r'(\\\\d[\\\\d,.]*)', text)


def extract_header_text(tbl) -> str:
    """헤더 표의 전체 텍스트 추출."""
    return _all_text(tbl).strip()


def build_text_replacements(header_text: str, fund_type: str) -> list:
    """
    헤더 표 텍스트에서 text_replacements 목록 생성.
    숫자를 포함한 문장을 찾아 패턴/data_path 매핑 생성.
    """
    reps = []

    # 패턴 1: 총 조합수 / 결성액 / 투자액
    # 예: "163개 조합 66,307억원 결성, 24,018억원 투자"
    m = re.search(r'(\d[\d,]*)개\s*조합\s*([\d,]+(?:\.\d+)?)억원\s*결성[,\s]*([\d,]+(?:\.\d+)?)억원\s*투자', header_text)
    if m:
        reps.append({
            "description": "총 조합수 (결성·투자 요약문)",
            "pattern": r"(\d[\d,]*)개\s*조합",
            "data_path": "s1_total_count",
            "format": "int",
        })
        reps.append({
            "description": "총 결성액 (결성·투자 요약문)",
            "pattern": r"(\d[\d,]*)억원\s*결성",
            "data_path": "s1_total_formed",
            "format": "float1",
        })
        reps.append({
            "description": "총 투자액 (결성·투자 요약문)",
            "pattern": r"(\d[\d,]*)억원\s*투자",
            "data_path": "s1_total_invested",
            "format": "float1",
        })

    # 패턴 2: 26년 선정 (선정) '26년 N개 조합 M억원 규모 선정
    m2 = re.search(r"'26년\s+(\d[\d,]*)개\s*조합\s*([\d,]+(?:\.\d+)?)억원\s*규모\s*선정", header_text)
    if m2:
        reps.append({
            "description": "'26년 선정 조합수",
            "pattern": r"'26년\s+(\d[\d,]*)개\s*조합",
            "data_path": "s1_year26_count",
            "format": "int",
        })
        reps.append({
            "description": "'26년 선정 결성액",
            "pattern": r"(\d[\d,]*)억원\s*규모\s*선정",
            "data_path": "s1_year26_formed",
            "format": "float1",
        })

    # 패턴 3: 현재 결성 현황 N개 조합 M억원 결성
    m3 = re.search(r'(\d[\d,]*)개\s*조합.*?(\d[\d,]*)억원.*?결성', header_text)
    if not reps and m3:
        # 앞서 패턴1이 없었던 경우에만
        reps.append({
            "description": "현재 결성 조합수",
            "pattern": r"(\d[\d,]*)개\s*조합",
            "data_path": "s1_cur_count",
            "format": "int",
        })

    if not reps:
        # 자동 추출 실패 → TODO 표시
        reps.append({
            "_TODO": "패턴 수동 입력 필요",
            "description": "총 조합수",
            "pattern": "TODO_PATTERN",
            "data_path": "s1_total_count",
            "format": "int",
        })

    return reps


# ─────────────────────────────────────────────────────────────────────────────
# 표 셀 매핑 생성
# ─────────────────────────────────────────────────────────────────────────────

def build_s1_cells(s1_table_idx: int, tbl, row_offset: int, section_file: str) -> list:
    """
    섹션1 (연도별) 표의 table_cells 생성.
    row_offset: 데이터 행이 여러 s1 표로 나뉠 때 Excel 인덱스 오프셋.
    """
    cells = []
    rows = get_data_rows(tbl)
    # row 0 = 헤더, row 1..N = 데이터
    n_data = len(rows) - 1
    for i in range(n_data):
        excel_row_idx = row_offset + i     # s1_rows 리스트 인덱스
        hwpx_row = i + 1                   # HWPX 표 행 (0-based, 0=헤더)
        for col_idx, field_name, fmt in S1_COLS:
            cells.append({
                "description": f"섹션1 행{excel_row_idx} {field_name}",
                "table_idx": s1_table_idx,
                "row": hwpx_row,
                "col": col_idx,
                "data_path": f"s1_rows.{excel_row_idx}.{field_name}",
                "format": fmt,
                "section_file": section_file,
            })
    return cells


def build_s2_cells(s2_table_idx: int, tbl, section_file: str) -> list:
    """섹션2 (조합상태) 표의 table_cells 생성."""
    cells = []
    # row 1=해산완료, row 2=운용중, row 3=진행
    s2_status_map = [(1, 0, "해산완료"), (2, 1, "운용중"), (3, 2, "진행")]
    for hwpx_row, list_idx, status in s2_status_map:
        for col_idx, field_name, fmt in S2_COLS:
            cells.append({
                "description": f"섹션2 {status} {field_name}",
                "table_idx": s2_table_idx,
                "row": hwpx_row,
                "col": col_idx,
                "data_path": f"s2_rows.{list_idx}.{field_name}",
                "format": fmt,
                "section_file": section_file,
            })
    return cells


def build_s3_cells(s3_table_idx: int, tbl, section_file: str) -> list:
    """섹션3 (연도별 투자여력) 표의 table_cells 생성."""
    cells = []
    rows = get_data_rows(tbl)
    n_data = len(rows) - 1
    for i in range(n_data):
        hwpx_row = i + 1
        for col_idx, field_name, fmt in S3_COLS:
            cells.append({
                "description": f"섹션3 행{i} {field_name}",
                "table_idx": s3_table_idx,
                "row": hwpx_row,
                "col": col_idx,
                "data_path": f"s3_rows.{i}.{field_name}",
                "format": fmt,
                "section_file": section_file,
            })
    return cells


# ─────────────────────────────────────────────────────────────────────────────
# 메인 생성 로직
# ─────────────────────────────────────────────────────────────────────────────

def generate_mapping(hwpx_path: str) -> dict:
    print(f"\nHWPX 로딩: {hwpx_path}")
    sections = load_sections(hwpx_path)

    # 섹션별 표 목록 캐시
    section_tables = {fname: get_tables(root) for fname, root in sections.items()}
    for fname, tbls in section_tables.items():
        print(f"  {fname}: 표 {len(tbls)}개")

    mapping = {"sections": []}

    for config in FUND_CONFIGS:
        fund_type, file_key, keyword, s1_indices, s2_idx, s3_idx = config
        section_file = FILE_KEY_MAP[file_key]

        if section_file not in section_tables:
            print(f"  [주의] {fund_type}: {section_file} 없음 -- 건너뜀")
            continue

        tbls = section_tables[section_file]
        total_tables = len(tbls)

        sec_entry = {
            "fund_type":      fund_type,
            "section_file":   section_file,
            "section_keyword": keyword,
            "text_replacements": [],
            "table_cells":    [],
        }

        # ── 섹션1 표 (연도별) ────────────────────────────────────────────────
        row_offset = 0
        for s1_tbl_idx in s1_indices:
            if s1_tbl_idx >= total_tables:
                print(f"  [주의] {fund_type}: 섹션1 표 인덱스 {s1_tbl_idx} 범위 초과 (총 {total_tables}개)")
                continue
            tbl = tbls[s1_tbl_idx]
            cells = build_s1_cells(s1_tbl_idx, tbl, row_offset, section_file)
            sec_entry["table_cells"].extend(cells)
            n_rows = get_row_cell_count(tbl)
            row_offset += n_rows

        # ── 섹션2 표 (조합상태) ─────────────────────────────────────────────
        if s2_idx is not None:
            if s2_idx >= total_tables:
                print(f"  [주의] {fund_type}: 섹션2 표 인덱스 {s2_idx} 범위 초과")
            else:
                tbl2 = tbls[s2_idx]
                cells2 = build_s2_cells(s2_idx, tbl2, section_file)
                sec_entry["table_cells"].extend(cells2)

        # ── 섹션3 표 (투자여력) ─────────────────────────────────────────────
        if s3_idx is not None:
            if s3_idx >= total_tables:
                print(f"  [주의] {fund_type}: 섹션3 표 인덱스 {s3_idx} 범위 초과")
            else:
                tbl3 = tbls[s3_idx]
                cells3 = build_s3_cells(s3_idx, tbl3, section_file)
                sec_entry["table_cells"].extend(cells3)

        # ── 헤더 표에서 텍스트 패턴 추출 ────────────────────────────────────
        # 헤더 표 = s1_indices[0] - 1 (s1 바로 앞)
        header_idx = s1_indices[0] - 1
        if 0 <= header_idx < total_tables:
            header_text = extract_header_text(tbls[header_idx])
            sec_entry["text_replacements"] = build_text_replacements(header_text, fund_type)
            sec_entry["_header_text_preview"] = header_text[:200]
        else:
            sec_entry["text_replacements"] = [{
                "_TODO": "헤더 표 인덱스 범위 초과 — 수동 입력",
                "description": "총 조합수",
                "pattern": "TODO",
                "data_path": "s1_total_count",
                "format": "int",
            }]

        n_cells = len(sec_entry["table_cells"])
        n_reps  = len(sec_entry["text_replacements"])
        print(f"  {fund_type:30s}  표셀 {n_cells:4d}개  텍스트교체 {n_reps}개")

        mapping["sections"].append(sec_entry)

    return mapping


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="HWPX 분석 결과로 mapping.json 자동 생성"
    )
    parser.add_argument("hwpx", help="분석할 HWPX 파일 (randomized 버전 가능)")
    parser.add_argument("--out", default=None,
                        help="출력 경로 (기본: 스크립트 위치의 mapping.json)")
    args = parser.parse_args()

    out_path = args.out or str(Path(__file__).parent / "mapping.json")

    mapping = generate_mapping(args.hwpx)

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    total_sections = len(mapping["sections"])
    total_cells    = sum(len(s["table_cells"]) for s in mapping["sections"])
    total_reps     = sum(len(s["text_replacements"]) for s in mapping["sections"])

    print(f"\n[완료] 저장: {out_path}")
    print(f"  펀드 섹션: {total_sections}개")
    print(f"  표 셀 매핑: {total_cells:,}개")
    print(f"  텍스트 교체: {total_reps:,}개")
    print()

    # TODO 항목 확인
    todos = [s["fund_type"] for s in mapping["sections"]
             if any("TODO" in str(tr) for tr in s.get("text_replacements", []))]
    if todos:
        print(f"  [주의] 텍스트 교체 패턴 수동 확인 필요 ({len(todos)}개):")
        for ft in todos:
            print(f"    - {ft}")


if __name__ == "__main__":
    main()
