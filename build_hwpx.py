"""
shutil.copyfile의 Python 3.13 memoryview 버그를 우회해서 PyInstaller 빌드를 실행하는 스크립트
대상: randomize_hwpx.py → 한글숫자변환기.exe
"""
import shutil
import pathlib

# memoryview write 버그 우회
def _patched_copyfile(src, dst, *, follow_symlinks=True):
    pathlib.Path(dst).write_bytes(pathlib.Path(src).read_bytes())

shutil.copyfile = _patched_copyfile

import sys
sys.argv = [
    'pyinstaller',
    '--onefile',
    '--console',
    '--name', '한글숫자변환기',
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
    'randomize_hwpx.py',
]

from PyInstaller.__main__ import run
run()
