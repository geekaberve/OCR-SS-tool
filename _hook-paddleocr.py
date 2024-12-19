from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import os
import paddleocr

# Get paddleocr installation directory
paddleocr_root = os.path.dirname(paddleocr.__file__)

# Collect all data files from paddleocr
datas = [
    # Main package files
    (os.path.join(paddleocr_root, '__init__.py'), 'paddleocr'),
    (os.path.join(paddleocr_root, 'paddleocr.py'), 'paddleocr'),
    
    # Tools directory - ensure it's in the correct location
    (os.path.join(paddleocr_root, 'tools'), 'paddleocr/tools'),
]

# Add any .py files from tools directory
tools_dir = os.path.join(paddleocr_root, 'tools')
for root, dirs, files in os.walk(tools_dir):
    for file in files:
        if file.endswith('.py'):
            full_path = os.path.join(root, file)
            # Preserve directory structure under tools
            rel_path = os.path.relpath(os.path.dirname(full_path), tools_dir)
            target_dir = os.path.join('paddleocr/tools', rel_path)
            datas.append((full_path, target_dir))

# Collect all paddleocr submodules
hiddenimports = collect_submodules('paddleocr') 