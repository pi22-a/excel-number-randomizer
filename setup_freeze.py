from cx_Freeze import setup, Executable

build_options = {
    "packages": ["openpyxl", "pyxlsb", "et_xmlfile"],
    "excludes": [
        "tkinter", "matplotlib", "numpy", "pandas", "scipy",
        "PIL", "unittest", "email", "html", "http", "urllib",
        "xmlrpc", "pydoc", "doctest", "test", "distutils",
        "setuptools", "pkg_resources", "lxml",
    ],
    "include_files": [],
    "optimize": 2,
    "build_exe": "dist_cx/randomize_excel",
}

setup(
    name="randomize_excel",
    version="1.0",
    options={"build_exe": build_options},
    executables=[
        Executable(
            "randomize_excel.py",
            target_name="randomize_excel.exe",
            base="console",
        )
    ],
)
