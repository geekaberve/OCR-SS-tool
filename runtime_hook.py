import os
import sys
import shutil

def setup_runtime_environment():
    if getattr(sys, 'frozen', False):
        # Running in PyInstaller bundle
        runtime_dir = os.path.join(os.getenv('LOCALAPPDATA'), 'OCR_Tool', 'runtime')
        os.makedirs(runtime_dir, exist_ok=True)
        
        # Set environment variables
        os.environ['PADDLE_OCR_HOME'] = runtime_dir
        os.environ['PYTHONPATH'] = runtime_dir
        
        # Create necessary directories
        for dir_name in ['paddleocr', 'paddle', 'tools']:
            os.makedirs(os.path.join(runtime_dir, dir_name), exist_ok=True)

setup_runtime_environment() 