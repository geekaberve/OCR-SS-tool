import os
import sys
import tempfile
import shutil
import atexit
import glob
import ctypes
from ctypes import windll
import logging
from datetime import datetime

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_temp_dirs():
    """Get all PyInstaller temp directories"""
    temp_dir = tempfile._get_default_tempdir()
    mei_pattern = os.path.join(temp_dir, '_MEI*')
    return glob.glob(mei_pattern)

def cleanup_temp_dirs():
    """Clean up all PyInstaller temp directories with proper error handling"""
    try:
        # Get all PyInstaller temp directories
        temp_dirs = get_temp_dirs()
        
        for dir_path in temp_dirs:
            try:
                if os.path.exists(dir_path):
                    # Try to remove read-only flags if any
                    for root, dirs, files in os.walk(dir_path):
                        for d in dirs:
                            try:
                                os.chmod(os.path.join(root, d), 0o777)
                            except:
                                pass
                        for f in files:
                            try:
                                os.chmod(os.path.join(root, f), 0o777)
                            except:
                                pass
                    
                    # Remove directory
                    shutil.rmtree(dir_path, ignore_errors=True)
                    
                    # Double check if directory is removed
                    if os.path.exists(dir_path):
                        # If still exists, try using system commands
                        if sys.platform == 'win32':
                            os.system(f'rd /s /q "{dir_path}"')
            except Exception as e:
                print(f"Error cleaning temp directory {dir_path}: {str(e)}")
                
    except Exception as e:
        print(f"Error in cleanup process: {str(e)}")

def ensure_temp_access():
    """Ensure the application has access to temp directories"""
    temp_dir = tempfile._get_default_tempdir()
    try:
        # Try to create a test file
        test_file = os.path.join(temp_dir, 'test_access.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
    except Exception as e:
        print(f"Warning: Limited temp directory access: {str(e)}")
        if not is_admin():
            print("Application may need administrative privileges")

# Register cleanup function
atexit.register(cleanup_temp_dirs)

# Ensure temp directory access on startup
ensure_temp_access() 