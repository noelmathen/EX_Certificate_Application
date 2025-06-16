#!/usr/bin/env python3
"""
setup.py  –  Build script for proEXy desktop application
Run with:
    python setup.py build_exe
"""

import sys
from pathlib import Path
from cx_Freeze import setup, Executable

# ─── Paths ────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.resolve()
ASSETS_DIR = BASE_DIR / "assets"
SRC_DIR    = BASE_DIR / "src"

# ─── Build options ───────────────────────────────────────────────────────
build_exe_options = {
    # all the Python packages your app uses
    "packages": [
        "tkinter", "logging", "threading", "queue", "datetime", "webbrowser",
        "os", "sys", "pathlib", "subprocess", "shutil", "glob", "re",
        "fitz",       # PyMuPDF
        "camelot",    # camelot-py
        "pandas",     # dataframes + Excel
        "openpyxl",   # Excel formatting
        "numpy"
    ],
    # copy your assets and source modules into the executable bundle
    "include_files": [
        (str(ASSETS_DIR), "assets"),
        (str(SRC_DIR),    "src"),
    ],
    
    # output folder for the frozen app
    "build_exe": "exe_build",
    # apply byte-code optimization
    "optimize": 2,
}

# ─── Executable definition ───────────────────────────────────────────────
base = "Win32GUI" if sys.platform == "win32" else None
icon_path = ASSETS_DIR / "EX_logo.ico"
executables = [
    Executable(
        script      = "main.py",
        base        = base,
        target_name = "proEXy.exe" if sys.platform == "win32" else "proEXy",
        icon        = str(icon_path) if icon_path.exists() else None
    )
]

# ─── Setup call ──────────────────────────────────────────────────────────
setup(
    name             = "proEXy",
    version          = "1.0.0",
    description      = "Professional EX Certificate Processor",
    long_description = (
        "proEXy is a desktop application for processing EX certificates "
        "from Proserv and Oman sources into structured Excel files."
    ),
    author           = "Your Name",
    author_email     = "your.email@example.com",
    url              = "https://github.com/noelmathen/EX_Certificate_Application",
    license          = "MIT",
    options          = {"build_exe": build_exe_options},
    executables      = executables,
    install_requires = [  # these mirror your requirements.txt
        "PyMuPDF>=1.23.0",
        "camelot-py[cv]>=0.10.1",
        "pandas>=1.5.0",
        "openpyxl>=3.0.10",
        "numpy>=1.21.0",
        "requests==2.32.4"
    ]
)
