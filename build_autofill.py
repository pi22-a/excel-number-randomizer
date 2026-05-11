"""
모태자동입력.exe 빌드 스크립트
대상: 모태자동입력.py → 배포/모태자동입력.exe

실행:
  python build_autofill.py
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

def _patched_copyfile(src, dst, *, follow_symlinks=True):
    pathlib.Path(dst).write_bytes(pathlib.Path(src).read_bytes())

shutil.copyfile = _patched_copyfile

_MAPPING = str(_HERE / 'mapping.json')

sys.argv = [
    'pyinstaller',
    '--onefile',
    '--console',
    '--name', '모태자동입력',
    '--distpath', '배포',
    '--workpath', 'build_tmp',
    '--specpath', 'build_tmp',

    # mapping.json을 exe 내부에 번들
    '--add-data', f'{_MAPPING};.',

    # xlwings 전체 수집 (COM 자동화)
    '--collect-all', 'xlwings',

    # lxml 전체 수집
    '--collect-all', 'lxml',

    # win32com / COM 관련
    '--hidden-import', 'win32com',
    '--hidden-import', 'win32com.client',
    '--hidden-import', 'win32com.server',
    '--hidden-import', 'pythoncom',
    '--hidden-import', 'pywintypes',
    '--hidden-import', 'win32api',
    '--hidden-import', 'win32con',

    '--hidden-import', 'openpyxl',
    '--hidden-import', 'openpyxl.styles',
    '--hidden-import', 'openpyxl.utils',

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

    '--optimize', '2',
    '--clean',
    '모태자동입력.py',
]

from PyInstaller.__main__ import run
run()
