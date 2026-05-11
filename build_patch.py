"""
숫자변환기.exe 빌드 스크립트
대상: randomize_excel.py → 배포/숫자변환기.exe

실행:
  python build_patch.py
  (가상환경 자동 생성 후 최소 패키지만 포함하여 빌드)
"""
import sys
import subprocess
import shutil
import pathlib

_HERE    = pathlib.Path(__file__).parent.resolve()
_VENV    = _HERE / "venv_autofill"
_PYTHON  = _VENV / "Scripts" / "python.exe"

# ── 가상환경 부트스트랩 ────────────────────────────────────────────────────
# 이미 venv_autofill로 실행 중이 아니면 자동 생성 후 재실행
if pathlib.Path(sys.prefix).resolve() != _VENV.resolve():
    if not _PYTHON.exists():
        print("가상환경 생성 중... (최초 1회)")
        subprocess.run([sys.executable, "-m", "venv", str(_VENV)], check=True)
        print("패키지 설치 중...")
        subprocess.run([
            str(_PYTHON), "-m", "pip", "install", "--quiet",
            "pyinstaller", "openpyxl", "pyxlsb", "lxml", "xlwings", "pywin32"
        ], check=True)
        print("설치 완료\n")
    print("가상환경으로 빌드 시작...")
    subprocess.run([str(_PYTHON), str(__file__)], check=True)
    sys.exit(0)

# ── 이하 가상환경 내에서 실행 ──────────────────────────────────────────────

# Python 3.13 memoryview 버그 우회
def _patched_copyfile(src, dst, *, follow_symlinks=True):
    pathlib.Path(dst).write_bytes(pathlib.Path(src).read_bytes())

shutil.copyfile = _patched_copyfile

sys.argv = [
    'pyinstaller',
    '--onefile',
    '--console',
    '--name', '숫자변환기',
    '--distpath', '배포',
    '--workpath', 'build_tmp',
    '--specpath', 'build_tmp',
    '--optimize', '2',
    '--exclude-module', 'tkinter',
    '--exclude-module', 'matplotlib',
    '--exclude-module', 'numpy',
    '--exclude-module', 'pandas',
    '--exclude-module', 'scipy',
    '--exclude-module', 'PIL',
    '--exclude-module', 'unittest',
    '--exclude-module', 'xmlrpc',
    '--exclude-module', 'pydoc',
    '--exclude-module', 'doctest',
    '--clean',
    'randomize_excel.py',
]

from PyInstaller.__main__ import run
run()
