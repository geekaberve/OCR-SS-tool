import sys
from cx_Freeze import setup, Executable
import os
import paddle
import paddleocr
import platform
import sys
import distutils.sysconfig

# Get paths
paddle_path = os.path.dirname(paddle.__file__)
paddleocr_path = os.path.dirname(paddleocr.__file__)
paddle_model_path = os.path.expanduser('~/.paddleocr/whl')

# Get Python DLL
python_version = platform.python_version_tuple()[:2]
python_dll = f'python{python_version[0]}{python_version[1]}.dll'
python_dll_path = os.path.join(os.path.dirname(sys.executable), python_dll)

# Get all required DLLs
dll_path = os.path.join(os.path.dirname(sys.executable), 'DLLs')
dlls_to_include = []
if os.path.exists(dll_path):
    for file in os.listdir(dll_path):
        if file.endswith('.dll'):
            dlls_to_include.append((os.path.join(dll_path, file), file))

# Add Python DLL to the list if it exists
if os.path.exists(python_dll_path):
    dlls_to_include.append((python_dll_path, python_dll))

# Base include files
include_files = [
    ("icons", "icons"),
    ("OCR_Modules", "OCR_Modules"),
    (paddleocr_path, "paddleocr"),
    (paddle_path, "paddle"),
    ("tesseract_binary", "tesseract_binary"),
    ("tessdata", "tessdata"),
    (os.path.join(paddle_model_path, 'det'), 'paddleocr/whl/det'),
    (os.path.join(paddle_model_path, 'cls'), 'paddleocr/whl/cls'),
    (os.path.join(paddle_model_path, 'rec'), 'paddleocr/whl/rec'),
]

# Add DLLs to include files
include_files.extend(dlls_to_include)

build_exe_options = {
    "packages": [
        "tkinter", 
        "PIL", 
        "paddle", 
        "paddleocr", 
        "pytesseract", 
        "numpy",
        "keyboard", 
        "win32clipboard",
        "os",
        "sys",
        "platform"
    ],
    "excludes": ["cv2.gapi.wip", "cv2.mat_wrapper"],
    "include_files": include_files,
    "includes": ["cv2"],
    "zip_include_packages": ["*"],
    "zip_exclude_packages": [],
    "include_msvcr": True,  # Include Microsoft Visual C++ runtime
}

# GUI base
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable(
        "main.py",
        base=base,
        icon="icons/icon.ico",
        target_name="OCR_Tool.exe",
        shortcut_name="OCR Tool",
        shortcut_dir="DesktopFolder"
    )
]

setup(
    name="OCR_Tool",
    version="1.0",
    description="OCR to Excel Tool",
    options={"build_exe": build_exe_options},
    executables=executables
) 