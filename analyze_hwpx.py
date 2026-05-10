"""
[회사에서 실행] HWPX 파일 구조를 분석하여 mapping.json 템플릿을 자동 생성.

사용법:
  python analyze_hwpx.py <파일.hwpx>

실행하면:
  1. HWPX 내부 구조 출력 (섹션 헤더, 표 목록, 텍스트 패턴)
  2. mapping_template.json 생성 → 이걸 mapping.json으로 채워 넣으면 됨
  3. 어떤 부분을 수동으로 채워야 하는지 안내

이 스크립트는 더미 HWPX (randomize_hwpx.py로 생성한 파일)로 실행하세요.
"""

import sys
import re
import json
import zipfile
from pathlib import Path
from collections import defaultdict
from lxml import etree

SECTION_RE = re.compile(r'^Contents/section\d+\.xml$', re.IGNORECASE)

# 펀드 유형 목록 (펀드순서 시트 기준)
FUND_TYPES = [
    "창업초기", "창업초기(벤처투자조합등)", "창업초기(개인투자조합)",
    "청년창업", "마이크로VC", "재기지원", "4차산업혁명",
    "조선업구조개선", "지방기업", "여성벤처", "소셜임팩트",
    "기술지주", "혁신성장", "스케일업", "M&A", "세컨더리",
    "엔젤세컨더리", "LP지분유동화", "소재부품장비",
    "버팀목", "해외진출",
    "IP·문화산업", "문화일반, FI유치", "수출", "신기술",
    "지역·청년·일자리", "M&A·세컨더리", "아시아문화중심도시육성",
    "한국영화 메인투자", "중저예산한국영화", "한국영화 개봉촉진",
    "관광기업육성", "스포츠산업", "스포츠프로젝트",
    "국토교통혁신", "도시재생", "대학창업", "사회적기업",
    "SaaS", "사이버보안", "메타버스", "공공기술사업화",
    "IP직접투자", "특허기술사업화", "미래환경산업", "보건",
]

# 엑셀 섹션1 연도 목록
YEARS = [str(y) + "년" for y in range(2004, 2027)] + ["합계"]

# 숫자 패턴
NUMBER_RE = re.compile(r'-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?')


def get_all_text(elem) -> str:
    return ''.join(elem.itertext())


def load_sections(hwpx_path: str) -> dict:
    """HWPX에서 섹션 XML 파싱. {filename: root_element}"""
    sections = {}
    with zipfile.ZipFile(hwpx_path, 'r') as zf:
        for name in zf.namelist():
            if SECTION_RE.match(name):
                data = zf.read(name)
                try:
                    root = etree.fromstring(data)
                    sections[name] = root
                except etree.XMLSyntaxError:
                    pass
    return sections


def find_fund_type_locations(sections: dict) -> dict:
    """
    각 펀드 유형이 어느 섹션 파일에 있는지, 어떤 키워드로 등장하는지 탐색.
    반환: {fund_type: {'file': ..., 'keyword': ..., 'candidates': [...]}}
    """
    results = {}

    for ft in FUND_TYPES:
        candidates = []
        for fname, root in sections.items():
            full_text = get_all_text(root)
            # 직접 일치
            for variant in [ft, ft + "편", ft + " 편", ft + "펀드"]:
                if variant in full_text:
                    candidates.append({'file': fname, 'keyword': variant})
        if candidates:
            results[ft] = candidates

    return results


def find_tables_in_section(root) -> list:
    """섹션 XML에서 모든 표를 찾아 구조 반환."""
    tables = []
    tbl_tags = ('tbl', 'table')
    row_tags = ('tr', 'row')
    cell_tags = ('td', 'tc', 'cell')

    idx = 0
    for node in root.iter():
        local = node.tag.split('}')[-1].lower()
        if local in tbl_tags:
            rows_data = []
            for child in node:
                child_local = child.tag.split('}')[-1].lower()
                if child_local in row_tags:
                    cells = []
                    for cell in child:
                        c_local = cell.tag.split('}')[-1].lower()
                        if c_local in cell_tags:
                            cells.append(get_all_text(cell).strip())
                    if cells:
                        rows_data.append(cells)
            if rows_data:
                tables.append({
                    'table_idx': idx,
                    'rows': len(rows_data),
                    'cols': max(len(r) for r in rows_data),
                    'preview': rows_data[:4],
                })
                idx += 1

    return tables


def find_text_patterns(root, fund_type: str) -> list:
    """
    펀드 유형과 관련된 텍스트 단락에서
    엑셀 데이터가 들어갈 것 같은 숫자 패턴을 찾아 반환.
    """
    patterns = []
    for node in root.iter():
        text = get_all_text(node).strip()
        if not text or len(text) > 200:
            continue
        # 숫자가 포함된 짧은 텍스트 단락
        if NUMBER_RE.search(text) and len(text) < 100:
            patterns.append(text)
    # 중복 제거, 너무 많으면 앞 30개
    seen = set()
    unique = []
    for p in patterns:
        if p not in seen:
            seen.add(p)
            unique.append(p)
    return unique[:30]


def generate_mapping_template(fund_locations: dict,
                               section_tables: dict,
                               section_patterns: dict) -> dict:
    """
    분석 결과를 바탕으로 mapping.json 템플릿 생성.
    사람이 직접 채워야 할 부분은 FILL_IN으로 표시.
    """
    template = {
        "_comment": [
            "이 파일을 mapping.json으로 복사 후 FILL_IN 부분을 실제 값으로 수정하세요.",
            "section_keyword: 해당 섹션을 식별하는 고유 텍스트 (예: '마이크로VC편')",
            "table_idx: 해당 섹션 내에서 표의 순서 (0-based)",
            "row/col: 표 내 위치 (0-based)",
            "data_path: FundData 필드 경로 (extract_excel.py 참조)",
            "format: int / float1 / float2 / pct1 / pct2",
        ],
        "sections": []
    }

    for ft in FUND_TYPES:
        if ft not in fund_locations:
            # 파일에서 못 찾은 유형
            sec = {
                "fund_type": ft,
                "_found": False,
                "section_keyword": "FILL_IN",
                "text_replacements": [],
                "table_cells": [],
            }
        else:
            candidates = fund_locations[ft]
            best = candidates[0]
            fname = best['file']

            # 해당 섹션의 텍스트 패턴
            patterns = section_patterns.get(fname, [])
            # 해당 섹션의 표
            tables = section_tables.get(fname, [])

            sec = {
                "fund_type": ft,
                "_found": True,
                "_file": fname,
                "_keyword_candidates": [c['keyword'] for c in candidates],
                "section_keyword": best['keyword'],

                # ── 텍스트 교체 (패턴은 실제 파일 보고 수정 필요) ──────────
                "text_replacements": [
                    {
                        "description": "총 조합수",
                        "pattern": "FILL_IN (예: r'(\\d+)개 조합')",
                        "data_path": "s1_total_count",
                        "format": "int"
                    },
                    {
                        "description": "총 결성액",
                        "pattern": "FILL_IN",
                        "data_path": "s1_total_formed",
                        "format": "float1"
                    },
                    {
                        "description": "총 투자액",
                        "pattern": "FILL_IN",
                        "data_path": "s1_total_invested",
                        "format": "float1"
                    },
                ],

                # ── 표 셀 교체 (table_idx, row, col은 아래 _tables 보고 수정) ─
                "_tables_found": [
                    {
                        "table_idx": t['table_idx'],
                        "rows": t['rows'],
                        "cols": t['cols'],
                        "preview_row0": t['preview'][0] if t['preview'] else [],
                    }
                    for t in tables[:5]
                ],

                # 섹션2 조합 상태 표 셀 (예시)
                "table_cells": [
                    {
                        "description": "섹션2 해산완료 조합수 → table_idx/row/col FILL_IN",
                        "table_idx": "FILL_IN",
                        "row": "FILL_IN",
                        "col": "FILL_IN",
                        "data_path": "s2_rows.0.count",
                        "format": "int"
                    },
                    {
                        "description": "섹션2 해산완료 결성액",
                        "table_idx": "FILL_IN",
                        "row": "FILL_IN",
                        "col": "FILL_IN",
                        "data_path": "s2_rows.0.formed",
                        "format": "float1"
                    },
                    {
                        "description": "섹션2 운용중 조합수",
                        "table_idx": "FILL_IN",
                        "row": "FILL_IN",
                        "col": "FILL_IN",
                        "data_path": "s2_rows.1.count",
                        "format": "int"
                    },
                    {
                        "description": "섹션2 운용중 결성액",
                        "table_idx": "FILL_IN",
                        "row": "FILL_IN",
                        "col": "FILL_IN",
                        "data_path": "s2_rows.1.formed",
                        "format": "float1"
                    },
                ],

                # 텍스트 패턴 힌트 (실제 파일에서 발견된 패턴)
                "_text_patterns_found": patterns[:10],
            }

        template["sections"].append(sec)

    return template


def print_report(fund_locations: dict, section_tables: dict):
    """분석 결과를 콘솔에 보기 좋게 출력."""
    print("\n" + "="*60)
    print("HWPX 구조 분석 리포트")
    print("="*60)

    print(f"\n[1] 발견된 펀드 유형: {len(fund_locations)}/{len(FUND_TYPES)}개")
    print("-"*40)
    for ft in FUND_TYPES:
        if ft in fund_locations:
            c = fund_locations[ft][0]
            print(f"  ✓  {ft:20s}  파일={c['file']}  키워드='{c['keyword']}'")
        else:
            print(f"  ✗  {ft:20s}  (미발견 — 직접 확인 필요)")

    print(f"\n[2] 섹션별 표 목록")
    print("-"*40)
    for fname, tables in section_tables.items():
        if tables:
            print(f"\n  {fname}")
            for t in tables:
                preview0 = str(t['preview'][0])[:60] if t['preview'] else ''
                print(f"    표[{t['table_idx']}]: {t['rows']}행 × {t['cols']}열  →  {preview0}")

    print("\n[3] 다음 작업")
    print("-"*40)
    print("  1. mapping_template.json 확인")
    print("  2. section_keyword: 실제 키워드 확인 후 수정")
    print("  3. table_idx / row / col: _tables_found 참고하여 FILL_IN 채우기")
    print("  4. text_replacements pattern: 실제 텍스트 보고 정규식 작성")
    print("  5. mapping_template.json → mapping.json 으로 복사 후 사용")


def main():
    if len(sys.argv) < 2:
        print("사용법: python analyze_hwpx.py <더미파일.hwpx>")
        print()
        print("  더미 HWPX 파일을 분석하여 mapping_template.json을 생성합니다.")
        print("  randomize_hwpx.py로 생성한 더미 파일을 사용하세요.")
        sys.exit(1)

    hwpx_path = sys.argv[1]
    out_dir = Path(hwpx_path).parent
    template_path = out_dir / "mapping_template.json"

    print(f"분석 중: {hwpx_path}")

    sections = load_sections(hwpx_path)
    print(f"  섹션 XML: {len(sections)}개")

    fund_locations = find_fund_type_locations(sections)

    section_tables = {}
    section_patterns = {}
    for fname, root in sections.items():
        section_tables[fname] = find_tables_in_section(root)
        section_patterns[fname] = find_text_patterns(root, "")

    print_report(fund_locations, section_tables)

    # 템플릿 생성
    template = generate_mapping_template(fund_locations, section_tables, section_patterns)
    with open(template_path, 'w', encoding='utf-8') as f:
        json.dump(template, f, ensure_ascii=False, indent=2)

    print(f"\n매핑 템플릿 저장: {template_path}")
    print("\n다음 단계:")
    print("  1. mapping_template.json 열어서 FILL_IN 항목 수정")
    print("  2. 수정 완료 후 mapping.json으로 파일명 변경")
    print("  3. python excel_to_hwpx.py --fill <hwpx> 실행")


if __name__ == "__main__":
    main()
