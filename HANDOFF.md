# 인수인계 문서 — 엑셀→한글 자동 입력 프로그램

> 이 문서를 클로드 코드 새 세션에 그대로 붙여넣으면 작업을 이어받을 수 있습니다.
> 최종 업데이트: 2026-05-10

---

## 프로젝트 전체 목적

회사 보안상 실제 파일을 외부 AI에 직접 줄 수 없어서,
아래 2단계 프로그램을 만들고 있습니다:

**1단계 (완료)** — 숫자 랜덤 변환기
- `randomize_excel.py` : xlsx/xlsm/xlsb 숫자 더미 변환
- `randomize_hwpx.py`  : hwpx 숫자 더미 변환 (lxml 기반)

**2단계 (진행 중)** — 엑셀→한글 자동 입력기
- 매달 엑셀 데이터 갱신 시 한글 문서를 자동으로 채우는 프로그램
- **현재 상태: 코드 뼈대 완성, 매핑 정보(mapping.json)만 미완성**
- 매핑은 실제 HWPX 파일을 받아야 채울 수 있음

---

## 환경

- OS: Windows 11
- Python: `C:\Users\pi22a\AppData\Local\Programs\Python\Python313\python.exe` (3.13.12)
- 프로젝트 경로: `C:\Users\pi22a\통계\`
- GitHub: https://github.com/pi22-a/excel-number-randomizer
- 설치된 패키지: openpyxl, pyxlsb, lxml, xlwings

---

## 파일 구조

```
C:\Users\pi22a\통계\
├── randomize_excel.py     # [완성] xlsx/xlsm/xlsb 숫자 랜덤 변환
├── randomize_hwpx.py      # [완성] hwpx 숫자 랜덤 변환
│
├── extract_excel.py       # [완성] 엑셀 → Python 딕셔너리 추출
├── hwpx_utils.py          # [완성] HWPX 표/텍스트 조작 유틸
├── excel_to_hwpx.py       # [완성 - 뼈대] 메인 실행 프로그램
├── analyze_hwpx.py        # [완성] HWPX 구조 분석 → mapping_template.json 생성
│
├── mapping.json           # [미완성 - 회사에서 채워야 함] 엑셀↔HWPX 매핑 정보
├── fund_data.json         # [생성 필요] 엑셀 추출 데이터 (extract_excel.py 실행 후 생성)
│
├── build_patch.py         # PyInstaller 빌드 스크립트
├── HANDOFF.md             # 이 문서
└── 배포/
    └── 숫자변환기.exe
```

---

## 엑셀 파일 구조 (분석 완료)

### 파일명
`10. DATA - 펀드별 참고자료 작업_(202603).xlsx`
(더미 파일명: `..._randomized.xlsx`)

### 핵심 시트: 펀드별참고자료

- **B1셀** = 펀드 유형명 입력 → 시트 전체 SUMIFS 재계산
- **섹션1** (행5~37): 결성 및 투자 현황
- **섹션2** (행7~9, 열K~R): 조합 상태별 집계
- **섹션3** (행7~19, 열W~Z): 연도별 투자금액 및 투자여력

### 셀 위치 정리

| 섹션 | 데이터 | 셀 위치 |
|------|--------|---------|
| 섹션1 요약 | 총 조합수 | B5 |
| 섹션1 요약 | 총 결성액 | D5 |
| 섹션1 요약 | 총 투자액 | F5 |
| 섹션1 요약 | 26년 선정 조합수 | B6 |
| 섹션1 요약 | 26년 선정 결성액 | D6 |
| 섹션1 요약 | 26년 모태 | G6 |
| 섹션1 요약 | 현재(N월) 결성수 | A7 |
| 섹션1 요약 | 현재(N월) 결성액 | C7 |
| 섹션1 요약 | 현재(N월) 투자액 | E7 |
| 섹션1 요약 | 결성 진행중 수 | A8 |
| 섹션1 표 헤더 | 행10 | A~H10 |
| 섹션1 표 데이터 | 2004년~2026년 | 행11~33 (A~H열) |
| 섹션1 합계 | 행34 | A~H34 |
| 섹션1 소계1 | 조성소계 | C~F36 |
| 섹션1 소계2 | 선정및결성소계 | C~F37 |
| 섹션2 헤더 | 행6 | K~R6 |
| 섹션2 해산완료 | 행7 | L~R7 |
| 섹션2 운용중 | 행8 | L~R8 |
| 섹션2 진행 | 행9 | L~R9 |
| 섹션3 헤더 | 행6 | W~Z6 |
| 섹션3 <=2015~<=2026 | 행7~18 | W~Z열 |
| 섹션3 합계(누적) | 행19 | W~Z19 |

### 펀드 유형 목록 (48개, 펀드순서 시트)

```python
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
```

---

## 코드 설명

### extract_excel.py — 엑셀 데이터 추출

```python
from extract_excel import ExcelExtractor

ext = ExcelExtractor("파일.xlsx")

# 모드1: 현재 저장된 펀드 유형 1개만 읽기 (빠름, Excel 불필요)
fd = ext.read_current()

# 모드2: 전체 48개 유형 순회 (Excel 자동화, xlwings 사용)
data = ext.read_all()   # dict[펀드유형명, FundData]
ext.export_json("fund_data.json", data)
```

FundData 필드:
```
fund_type, unit
s1_total_count, s1_total_formed, s1_total_invested   # 섹션1 총계
s1_year26_count, s1_year26_formed, s1_year26_moat   # 26년 선정
s1_cur_count, s1_cur_formed, s1_cur_invested         # 현재(N월)
s1_pending_count                                      # 결성 진행중
s1_rows    : List[Section1Row]  # 연도별 (year, count, formed, moat, moat_ratio, companies, invested, per_company)
s1_sub_formed, s1_sub_moat, s1_sub_companies, s1_sub_invested     # 조성소계
s1_sub2_formed, s1_sub2_moat, s1_sub2_companies, s1_sub2_invested # 선정및결성소계
s2_rows    : List[Section2Row]  # 조합상태 (status, count, formed, moat, moat_ratio, companies, invested, per_company)
s3_rows    : List[Section3Row]  # 연도별 투자여력 (year, invested, capacity, cumulative)
s3_total_invested                                     # 합계(누적)
```

### hwpx_utils.py — HWPX 조작

```python
from hwpx_utils import HwpxEditor

editor = HwpxEditor("파일.hwpx")
editor.load()

# 구조 확인
headers = editor.find_section_headers()   # 섹션 헤더 후보
tables  = editor.find_tables()            # 표 목록

# 텍스트 교체 (section_keyword가 있는 XML 파일에서만 교체)
editor.replace_in_paragraphs("마이크로VC", r"(\d+)개 조합", "42")
editor.replace_text_exact("1,234", "5,678")

# 표 셀 교체 (0-based)
editor.set_table_cell(table_idx=0, row=1, col=2, value="999")

editor.save("출력.hwpx")
```

### analyze_hwpx.py — HWPX 구조 분석 (회사에서 실행)

```bash
python analyze_hwpx.py 더미파일.hwpx
```
→ 콘솔에 섹션 헤더, 표 구조 출력
→ `mapping_template.json` 자동 생성

### excel_to_hwpx.py — 메인 실행 프로그램

```bash
# 엑셀 데이터 추출 (1회, Excel 필요)
python excel_to_hwpx.py --extract 엑셀파일.xlsx --json fund_data.json

# HWPX 채우기 (매달 실행)
python excel_to_hwpx.py --fill 입력.hwpx --json fund_data.json --out 출력.hwpx
```

---

## 회사에서 해야 할 작업 순서

### Step 1: 더미 파일 생성 (1회)
```bash
# HWP → HWPX 변환 (HOP 사용)
# https://golbin.github.io/hop/

# 숫자 랜덤 변환
python randomize_hwpx.py 실제파일.hwpx 더미파일.hwpx
```

### Step 2: HWPX 구조 분석 (1회)
```bash
python analyze_hwpx.py 더미파일.hwpx
# → mapping_template.json 생성
```

### Step 3: 매핑 정보 채우기 (1회)
1. `mapping_template.json` 열기
2. `section_keyword`: 실제 섹션 헤더 텍스트 확인 후 입력
3. `table_idx / row / col`: 분석 결과 `_tables_found` 보고 입력
4. `text_replacements pattern`: 실제 텍스트 패턴 맞게 정규식 작성
5. 파일을 `mapping.json`으로 저장

### Step 4: 엑셀 데이터 추출 (1회)
```bash
python excel_to_hwpx.py --extract 실제엑셀.xlsx --json fund_data.json
```

### Step 5: 매달 실행
```bash
# 새 엑셀 파일로 데이터 갱신 후
python excel_to_hwpx.py --extract 새엑셀.xlsx --json fund_data.json
python excel_to_hwpx.py --fill 한글파일.hwpx --json fund_data.json --out 결과.hwpx
```

---

## mapping.json 작성 가이드

```json
{
  "sections": [
    {
      "fund_type": "마이크로VC",
      "section_keyword": "마이크로VC편",   ← HWPX에서 이 텍스트로 섹션 위치 특정
      "text_replacements": [
        {
          "description": "총 조합수",
          "pattern": "(\\d+)개 조합",       ← 정규식 (교체할 숫자 부분을 포함)
          "data_path": "s1_total_count",    ← FundData 필드명
          "format": "int"                   ← int/float1/float2/pct1/pct2
        }
      ],
      "table_cells": [
        {
          "description": "섹션2 해산완료 조합수",
          "table_idx": 0,    ← 해당 섹션 내 표 순서 (0부터)
          "row": 1,          ← 표 내 행 (0부터)
          "col": 1,          ← 표 내 열 (0부터)
          "data_path": "s2_rows.0.count",
          "format": "int"
        }
      ]
    }
  ]
}
```

**data_path 예시:**
- `s1_total_count` → 섹션1 총 조합수
- `s1_rows.3.formed` → 섹션1 4번째 연도(2007년)의 결성액
- `s2_rows.0.count` → 섹션2 해산완료 조합수
- `s2_rows.1.invested` → 섹션2 운용중 투자금액
- `s3_rows.2.capacity` → 섹션3 3번째 연도 투자여력

---

*이 문서를 클로드 코드 새 세션에 그대로 붙여넣으면 됩니다.*
