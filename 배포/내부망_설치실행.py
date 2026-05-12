"""
내부망 설치 및 실행 스크립트
================================
1. '내부망패키지.txt' 와 같은 폴더에 이 파일을 놓으세요.
2. python 내부망_설치실행.py

표준 라이브러리만 사용 (추가 설치 불필요)
"""
import base64
import zipfile
import io
import subprocess
import sys
import os
from pathlib import Path

HERE = Path(__file__).parent

print("=" * 50)
print("  모태펀드 자동입력기 - 내부망 설치")
print("=" * 50)
print()

# ── 1. 패키지 파일 찾기 ───────────────────────────────────
pkg_file = HERE / "내부망패키지.txt"
if not pkg_file.exists():
    print(f"오류: '내부망패키지.txt' 파일이 없습니다.")
    print(f"  찾는 위치: {HERE}")
    input("\n엔터를 누르면 종료...")
    sys.exit(1)

# ── 2. 압축 해제 ──────────────────────────────────────────
print("[1/2] 파일 압축 해제 중...")
encoded = pkg_file.read_text(encoding="utf-8")
zip_bytes = base64.b64decode(encoded)

with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
    for name in zf.namelist():
        if name == "실행방법.txt":
            continue
        zf.extract(name, HERE)
        print(f"  추출: {name}")

print()

# ── 3. 패키지 설치 ────────────────────────────────────────
print("[2/2] 필요 패키지 설치 중...")
packages = ["xlwings", "lxml", "openpyxl", "pywin32"]

result = subprocess.run(
    [sys.executable, "-m", "pip", "install"] + packages,
    capture_output=True, text=True
)

if result.returncode == 0:
    print("  설치 완료!")
else:
    print("  [주의] 설치 중 문제 발생:")
    print(result.stderr[-500:])
    print()
    print("  인터넷이 안 되는 경우 관리자에게 아래 패키지 설치 요청:")
    for p in packages:
        print(f"    - {p}")

print()
print("=" * 50)
print("  설치 완료! 이제 실행하세요:")
print(f"  python 모태자동입력.py")
print("=" * 50)
print()
input("엔터를 누르면 종료...")
