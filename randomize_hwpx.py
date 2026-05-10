"""
HWPX 파일의 숫자를 연도 제외하고 랜덤으로 바꾸는 스크립트
지원 형식: .hwpx (HWP Open XML)
사용법: python randomize_hwpx.py <입력파일.hwpx> [출력파일.hwpx]

.hwp 파일은 HOP으로 먼저 .hwpx로 변환 후 사용하세요:
  https://golbin.github.io/hop/
"""

import sys
import re
import zipfile
import math
import random
from lxml import etree


YEAR_MIN = 2004
YEAR_MAX = 2040

# HWPX 섹션 XML 패턴 (실제 문서 텍스트가 들어있는 파일)
SECTION_RE = re.compile(r'^Contents/section\d+\.xml$', re.IGNORECASE)

# 숫자 패턴: 정수/소수, 음수, 천단위 콤마 포함
#   예: 1,234,567  /  -3.14  /  100  /  0.5
NUMBER_RE = re.compile(r'-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?')


# ── 핵심 변환 로직 (randomize_excel.py와 동일) ─────────────────────────────

def is_year(value: float) -> bool:
    return value == int(value) and YEAR_MIN <= int(value) <= YEAR_MAX


def randomize_number(value: float):
    if value == 0:
        return 0
    is_int = (value == int(value))
    sign = -1 if value < 0 else 1
    abs_val = abs(value)
    magnitude = 10 ** math.floor(math.log10(abs_val))
    rand_val = random.uniform(magnitude, magnitude * 10)
    if is_int:
        return sign * round(rand_val)
    decimal_places = len(str(abs_val).rstrip('0').split('.')[-1]) if '.' in str(abs_val) else 0
    return sign * round(rand_val, decimal_places)


# ── 텍스트 내 숫자 치환 ────────────────────────────────────────────────────

def replace_in_text(text: str, stats: dict) -> str:
    """문자열 안의 숫자 패턴만 찾아서 교체 (연도 제외)"""
    def replacer(m: re.Match) -> str:
        raw = m.group()
        # 천단위 콤마 제거 후 float 변환
        try:
            val = float(raw.replace(',', ''))
        except ValueError:
            return raw

        if is_year(val):
            stats['kept'] += 1
            return raw  # 원본 유지

        new_val = randomize_number(val)
        stats['changed'] += 1

        # 천단위 콤마 형식 복원 여부 판단
        has_comma = ',' in raw
        is_int_val = (val == int(val))

        if is_int_val:
            result = int(new_val)
            if has_comma:
                return f"{result:,}"
            return str(result)
        else:
            decimal_places = len(str(abs(val)).rstrip('0').split('.')[-1]) if '.' in str(abs(val)) else 0
            result = round(float(new_val), decimal_places)
            return str(result)

    return NUMBER_RE.sub(replacer, text)


# ── XML 처리: 텍스트 노드만 순회 ─────────────────────────────────────────

def process_xml_bytes(data: bytes, stats: dict) -> bytes:
    """XML 바이트를 파싱해서 텍스트 노드의 숫자만 교체 후 반환"""
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        # XML 파싱 실패 시 원본 반환
        return data

    for node in root.iter():
        if node.text:
            node.text = replace_in_text(node.text, stats)
        if node.tail:
            node.tail = replace_in_text(node.tail, stats)

    return etree.tostring(root, encoding='utf-8', xml_declaration=True)


# ── 메인 처리 ─────────────────────────────────────────────────────────────

def print_progress(current: int, total: int, name: str):
    pct = int(current / total * 100) if total > 0 else 0
    bar_len = 30
    filled = int(bar_len * pct / 100)
    bar = '#' * filled + '-' * (bar_len - filled)
    label = name.split('/')[-1]  # 파일명만 표시
    print(f"\r  [{bar}] {pct:3d}%  |  {label}  ({current}/{total})", end='', flush=True)


def process_hwpx(input_path: str, output_path: str) -> tuple[int, int]:
    stats = {'changed': 0, 'kept': 0}

    with zipfile.ZipFile(input_path, 'r') as zin:
        all_names = zin.namelist()
        section_names = [n for n in all_names if SECTION_RE.match(n)]
        total = len(section_names)

        if total == 0:
            print("  경고: 처리할 섹션 XML을 찾지 못했습니다.")
            print(f"  ZIP 내 파일 목록: {all_names[:20]}")

        with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            done = 0
            for name in all_names:
                data = zin.read(name)

                if SECTION_RE.match(name):
                    done += 1
                    print_progress(done, total, name)
                    data = process_xml_bytes(data, stats)

                zout.writestr(zin.getinfo(name), data)

    print()  # 줄바꿈
    return stats['changed'], stats['kept']


def main():
    if len(sys.argv) < 2:
        print("사용법: python randomize_hwpx.py <입력파일.hwpx> [출력파일.hwpx]")
        print()
        print(".hwp 파일은 먼저 HOP으로 .hwpx 변환이 필요합니다:")
        print("  https://golbin.github.io/hop/")
        sys.exit(1)

    input_path = sys.argv[1]
    if not input_path.lower().endswith('.hwpx'):
        print("오류: .hwpx 파일만 지원합니다.")
        print(".hwp 파일은 HOP(https://golbin.github.io/hop/)으로 먼저 변환하세요.")
        sys.exit(1)

    output_path = (
        sys.argv[2] if len(sys.argv) >= 3
        else input_path[:-5] + '_randomized.hwpx'
    )

    print(f"입력 파일: {input_path}")
    print(f"출력 파일: {output_path}")
    print("처리 중...")

    try:
        changed, kept = process_hwpx(input_path, output_path)
    except FileNotFoundError:
        print(f"오류: 파일을 찾을 수 없습니다 — {input_path}")
        sys.exit(1)
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print("완료!")
    print(f"  변경된 숫자: {changed:,}개")
    print(f"  유지된 연도: {kept:,}개  ({YEAR_MIN}~{YEAR_MAX} 범위)")


if __name__ == "__main__":
    main()
