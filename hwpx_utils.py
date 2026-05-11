"""
HWPX 파일에서 섹션/표를 찾고 값을 교체하는 유틸리티.

HWPX = ZIP 아카이브 (내부 XML, Contents/section*.xml에 본문)
"""

import re
import zipfile
import copy
from lxml import etree
from typing import Optional


# ── HWPX 네임스페이스 ────────────────────────────────────────────────────────
# 실제 파일을 받기 전까지는 일반적인 HWP Open XML 네임스페이스 사용.
# analyze_hwpx.py 실행 후 실제 네임스페이스로 업데이트 필요.
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'ht': 'http://www.hancom.co.kr/hwpml/2011/table',
}

SECTION_RE = re.compile(r'^Contents/section\d+\.xml$', re.IGNORECASE)


# ── 텍스트 추출 헬퍼 ────────────────────────────────────────────────────────

def _get_all_text(elem) -> str:
    """엘리먼트 하위의 모든 텍스트를 이어붙여 반환."""
    return ''.join(elem.itertext())


def _iter_text_nodes(root):
    """루트 하위의 모든 텍스트 노드를 (node, 'text'|'tail') 형태로 순회."""
    for node in root.iter():
        if node.text and node.text.strip():
            yield node, 'text'
        if node.tail and node.tail.strip():
            yield node, 'tail'


# ── HwpxEditor ──────────────────────────────────────────────────────────────

class HwpxEditor:
    """
    HWPX 파일을 로드하고 텍스트/표 값을 교체한 뒤 저장하는 클래스.

    사용 흐름:
        editor = HwpxEditor("input.hwpx")
        editor.load()

        # 섹션 목록 확인
        headers = editor.find_section_headers()

        # 표 목록 확인
        tables = editor.find_tables()

        # 값 교체
        editor.replace_in_paragraphs("마이크로VC", r"(\d[\d,.]*)개 조합", "42개 조합")  # noqa
        editor.set_table_cell(table_idx=0, row=1, col=2, value="1,234")

        editor.save("output.hwpx")
    """

    def __init__(self, hwpx_path: str):
        self.hwpx_path = hwpx_path
        # {filename: (bytes_original, etree_root)}
        self._sections: dict[str, tuple[bytes, etree._Element]] = {}
        self._zip_entries: dict[str, bytes] = {}   # 섹션 외 파일들

    # ── 로드 ─────────────────────────────────────────────────────────────────

    def load(self):
        """HWPX 파일을 파싱하여 메모리에 로드."""
        self._sections.clear()
        self._zip_entries.clear()

        with zipfile.ZipFile(self.hwpx_path, 'r') as zf:
            for name in zf.namelist():
                data = zf.read(name)
                if SECTION_RE.match(name):
                    try:
                        root = etree.fromstring(data)
                        self._sections[name] = root
                    except etree.XMLSyntaxError as e:
                        print(f"  경고: {name} XML 파싱 실패 — {e}")
                        self._zip_entries[name] = data
                else:
                    self._zip_entries[name] = data

        print(f"로드 완료: 섹션 {len(self._sections)}개")

    # ── 섹션 헤더 탐색 ──────────────────────────────────────────────────────

    def find_section_headers(self, min_len: int = 2, max_len: int = 30) -> list[dict]:
        """
        각 섹션 XML에서 짧은 텍스트 단락(제목 후보)을 추출.
        반환: [{'file': ..., 'text': ..., 'xpath': ...}, ...]
        """
        results = []
        for fname, root in self._sections.items():
            for node in root.iter():
                text = _get_all_text(node).strip()
                if min_len <= len(text) <= max_len and '\n' not in text:
                    # 태그명 마지막 부분만 (네임스페이스 제거)
                    tag = node.tag.split('}')[-1] if '}' in node.tag else node.tag
                    results.append({
                        'file': fname,
                        'tag': tag,
                        'text': text,
                    })
        # 중복 제거
        seen = set()
        unique = []
        for r in results:
            key = (r['file'], r['text'])
            if key not in seen:
                seen.add(key)
                unique.append(r)
        return unique

    # ── 표 탐색 ─────────────────────────────────────────────────────────────

    def find_tables(self) -> list[dict]:
        """
        모든 섹션에서 표를 찾아 구조 정보 반환.
        반환: [{'file': ..., 'table_idx': ..., 'rows': ..., 'cols': ...,
                'preview': [[cell_text, ...], ...]}, ...]

        지원 태그 (네임스페이스 무관):
          표: Tbl, Table, tbl, table
          행: Tr, Row, tr, row
          셀: Tc, Td, Cell, tc, td, cell
        """
        TABLE_TAGS = {'tbl', 'table'}
        ROW_TAGS   = {'tr', 'row'}
        CELL_TAGS  = {'tc', 'td', 'cell'}

        results = []
        table_global_idx = 0

        for fname, root in self._sections.items():
            tbl_roots = [n for n in root.iter()
                         if n.tag.split('}')[-1].lower() in TABLE_TAGS]

            for tbl in tbl_roots:
                rows_data = []
                for row_node in tbl:
                    row_local = row_node.tag.split('}')[-1].lower()
                    if row_local not in ROW_TAGS:
                        continue
                    row_cells = []
                    for cell_node in row_node:
                        cell_local = cell_node.tag.split('}')[-1].lower()
                        if cell_local in CELL_TAGS:
                            row_cells.append(_get_all_text(cell_node).strip())
                    if row_cells:
                        rows_data.append(row_cells)

                if rows_data:
                    results.append({
                        'file': fname,
                        'table_idx': table_global_idx,
                        'rows': len(rows_data),
                        'cols': max(len(r) for r in rows_data),
                        'preview': rows_data[:3],
                    })
                    table_global_idx += 1

        return results

    # ── 텍스트 단락 교체 ────────────────────────────────────────────────────

    def replace_in_paragraphs(self,
                               section_keyword: str,
                               pattern: str,
                               replacement: str,
                               flags: int = 0,
                               section_file: str = "") -> int:
        """
        section_keyword를 포함하는 파일의 텍스트에서 pattern을 replacement로 교체.

        section_file: 직접 파일명 지정 (예: 'Contents/section0.xml'), 지정 시 keyword 무시
        section_keyword: 섹션을 식별하는 키워드 (예: "마이크로VC")
        pattern: 교체할 정규식 패턴
        replacement: 교체할 문자열 (re.sub에 전달)
        반환: 교체 횟수
        """
        count = 0
        compiled = re.compile(pattern, flags)

        for fname, root in self._sections.items():
            if section_file:
                if fname != section_file:
                    continue
            else:
                # 섹션 키워드가 이 XML 파일에 존재하는지 확인
                full_text = _get_all_text(root)
                if section_keyword and section_keyword not in full_text:
                    continue

            for node, attr in _iter_text_nodes(root):
                orig = getattr(node, attr)
                new, n = compiled.subn(replacement, orig)
                if n > 0:
                    setattr(node, attr, new)
                    count += n

        return count

    def replace_text_exact(self, old_text: str, new_text: str) -> int:
        """
        모든 섹션에서 old_text를 new_text로 단순 치환.
        반환: 교체 횟수
        """
        count = 0
        for fname, root in self._sections.items():
            for node, attr in _iter_text_nodes(root):
                orig = getattr(node, attr)
                if old_text in orig:
                    setattr(node, attr, orig.replace(old_text, new_text))
                    count += 1
        return count

    # ── 표 셀 교체 ─────────────────────────────────────────────────────────

    def _get_table_roots(self) -> list[tuple[str, etree._Element]]:
        """(파일명, 표루트엘리먼트) 목록 반환."""
        result = []
        for fname, root in self._sections.items():
            for n in root.iter():
                local = n.tag.split('}')[-1].lower()
                if local in ('tbl', 'table'):
                    result.append((fname, n))
        return result

    def set_table_cell(self,
                       table_idx: int,
                       row: int,
                       col: int,
                       value: str,
                       section_keyword: str = "",
                       section_file: str = "") -> bool:
        """
        전체 표 중 table_idx번째 표의 (row, col) 셀 텍스트를 value로 교체.
        row, col은 0-based.
        section_file을 지정하면 해당 파일(예: 'Contents/section0.xml')의 표만 탐색.
        section_file이 없으면 section_keyword로 파일을 필터링.
        반환: 성공 여부
        """
        tables = self._get_table_roots()

        if section_file:
            tables = [(f, t) for f, t in tables if f == section_file]
        elif section_keyword:
            tables = [(f, t) for f, t in tables
                      if section_keyword in _get_all_text(self._sections[f])]

        if table_idx >= len(tables):
            return False

        _, tbl = tables[table_idx]

        # 행 순회
        row_nodes = [n for n in tbl if n.tag.split('}')[-1].lower() in ('tr', 'row')]
        if row >= len(row_nodes):
            return False

        row_node = row_nodes[row]
        cell_nodes = [n for n in row_node if n.tag.split('}')[-1].lower() in ('td', 'tc', 'cell')]
        if col >= len(cell_nodes):
            return False

        cell = cell_nodes[col]
        # 셀 내 첫 번째 텍스트 노드를 교체
        for text_node in cell.iter():
            if text_node.text and text_node.text.strip():
                text_node.text = value
                return True

        # 텍스트 노드가 없으면 직접 설정
        cell.text = value
        return True

    def get_table_cell(self, table_idx: int, row: int, col: int,
                       section_keyword: str = "",
                       section_file: str = "") -> Optional[str]:
        """표 셀 텍스트 읽기 (디버깅용)."""
        tables = self._get_table_roots()
        if section_file:
            tables = [(f, t) for f, t in tables if f == section_file]
        elif section_keyword:
            tables = [(f, t) for f, t in tables
                      if section_keyword in _get_all_text(self._sections[f])]
        if table_idx >= len(tables):
            return None
        _, tbl = tables[table_idx]
        row_nodes = [n for n in tbl if n.tag.split('}')[-1].lower() in ('tr', 'row')]
        if row >= len(row_nodes):
            return None
        row_node = row_nodes[row]
        cell_nodes = [n for n in row_node if n.tag.split('}')[-1].lower() in ('td', 'tc', 'cell')]
        if col >= len(cell_nodes):
            return None
        return _get_all_text(cell_nodes[col]).strip()

    # ── 저장 ─────────────────────────────────────────────────────────────────

    def save(self, output_path: str):
        """수정된 내용을 새 HWPX 파일로 저장."""
        with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            # 원본 ZIP의 메타 정보 보존을 위해 원본을 다시 열어서 info 복사
            with zipfile.ZipFile(self.hwpx_path, 'r') as zin:
                for name in zin.namelist():
                    info = zin.getinfo(name)
                    if name in self._sections:
                        root = self._sections[name]
                        data = etree.tostring(root, encoding='utf-8', xml_declaration=True)
                    else:
                        data = self._zip_entries.get(name, zin.read(name))
                    zout.writestr(info, data)

        print(f"저장 완료: {output_path}")

    # ── 디버깅 헬퍼 ─────────────────────────────────────────────────────────

    def print_structure(self, max_text_len: int = 60):
        """섹션별 텍스트 내용을 요약 출력 (디버깅용)."""
        for fname, root in self._sections.items():
            print(f"\n=== {fname} ===")
            tables = [n for n in root.iter() if n.tag.split('}')[-1].lower() in ('tbl', 'table')]
            print(f"  표 개수: {len(tables)}")
            full = _get_all_text(root)
            preview = full[:300].replace('\n', ' ')
            print(f"  텍스트 미리보기: {preview}...")
