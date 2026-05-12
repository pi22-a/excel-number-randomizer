"""
외부망 → 내부망 전송용 패키지 생성기
실행: python pack_for_internal.py
결과: 배포/내부망패키지.txt  (이 파일 하나만 내부망으로 전송)
"""
import base64
import zipfile
import io
import pathlib

HERE = pathlib.Path(__file__).parent

# 내부망으로 보낼 파일 목록
FILES = [
    "모태자동입력.py",
    "extract_excel.py",
    "excel_to_hwpx.py",
    "hwpx_utils.py",
    "mapping.json",
]

# 메모리 안에서 ZIP 생성
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
    for fname in FILES:
        fpath = HERE / fname
        if not fpath.exists():
            print(f"  [경고] 파일 없음: {fname}")
            continue
        zf.write(fpath, fname)
        size = fpath.stat().st_size
        print(f"  포함: {fname}  ({size:,} bytes)")

    # 내부망 설치/실행 안내 스크립트도 포함
    readme = """내부망 실행 방법
================

1. 이 파일들을 원하는 폴더에 압축 해제 (이미 됐으면 건너뜀)

2. 패키지 설치 (최초 1회):
   pip install xlwings lxml openpyxl pywin32

   인터넷 안 될 경우 → 관리자에게 패키지 요청 또는 아래 오프라인 방법 사용

3. 실행:
   python 모태자동입력.py

   또는 엑셀/한글 파일 경로를 직접 전달:
   python 모태자동입력.py 엑셀파일.xlsx 한글파일.hwpx
""".encode("utf-8")
    zf.writestr("실행방법.txt", readme)

# base64 인코딩
zip_bytes = buf.getvalue()
encoded = base64.b64encode(zip_bytes).decode("utf-8")

# 저장
out_dir = HERE / "배포"
out_dir.mkdir(exist_ok=True)
out_path = out_dir / "내부망패키지.txt"
out_path.write_text(encoded, encoding="utf-8")

print()
print(f"생성 완료: {out_path}")
print(f"파일 크기: {out_path.stat().st_size / 1024:.1f} KB")
print()
print("이 txt 파일 하나만 내부망으로 전송하세요.")
