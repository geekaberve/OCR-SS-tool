import os
import shutil
from pathlib import Path
import paddleocr
import paddle

def copy_paddleocr_files(ocr_files_dir):
    # PaddleOCR model paths
    paddle_model_path = Path.home() / '.paddleocr/whl'
    for model_dir in ['det', 'cls', 'rec']:
        src_dir = paddle_model_path / model_dir
        dest_dir = ocr_files_dir / 'paddleocr' / 'whl' / model_dir
        if src_dir.exists():
            shutil.copytree(src_dir, dest_dir, dirs_exist_ok=True)
            print(f"Copied PaddleOCR model directory: {model_dir}")

def copy_paddle_files(ocr_files_dir):
    # Paddle package path
    paddle_path = Path(paddle.__file__).parent
    dest_dir = ocr_files_dir / 'paddle'
    shutil.copytree(paddle_path, dest_dir, dirs_exist_ok=True)
    print("Copied Paddle package")

def copy_tesseract_files(ocr_files_dir):
    # Tesseract installation directory
    tesseract_install_dir = Path(r"C:\Program Files\Tesseract-OCR")
    dest_dir = ocr_files_dir / 'tesseract'
    tessdata_dest_dir = dest_dir / 'tessdata'
    
    if tesseract_install_dir.exists():
        try:
            # Copy tesseract.exe
            shutil.copy2(tesseract_install_dir / "tesseract.exe", dest_dir / "tesseract.exe")
            print("Copied tesseract.exe")
            
            # Copy DLLs
            for file in tesseract_install_dir.glob("*.dll"):
                try:
                    shutil.copy2(file, dest_dir / file.name)
                    print(f"Copied {file.name}")
                except PermissionError:
                    print(f"Permission denied for {file.name}, skipping...")
            
            # Copy tessdata - only copy necessary files, skip configs directory
            tessdata_src_dir = tesseract_install_dir / "tessdata"
            tessdata_dest_dir.mkdir(parents=True, exist_ok=True)
            
            # List of file extensions to copy
            valid_extensions = {'.traineddata', '.cube.', '.cube.params', '.cube.size', '.cube.nn'}
            
            for file in tessdata_src_dir.glob("*"):
                # Skip the configs directory and only copy necessary files
                if file.is_file() and any(ext in file.name for ext in valid_extensions):
                    try:
                        shutil.copy2(file, tessdata_dest_dir / file.name)
                        print(f"Copied tessdata file: {file.name}")
                    except PermissionError:
                        print(f"Permission denied for {file.name}, skipping...")
                else:
                    print(f"Skipping {file.name}")
                    
        except PermissionError as e:
            print(f"Permission error: {e}")
            print("Try running the script with administrator privileges")
        except Exception as e:
            print(f"Error copying Tesseract files: {e}")

def main():
    project_dir = Path(__file__).parent
    ocr_files_dir = project_dir / 'ocr_files'
    
    # Create directory structure
    (ocr_files_dir / 'paddleocr' / 'whl').mkdir(parents=True, exist_ok=True)
    (ocr_files_dir / 'paddle').mkdir(parents=True, exist_ok=True)
    (ocr_files_dir / 'tesseract' / 'tessdata').mkdir(parents=True, exist_ok=True)
    
    # Copy files
    copy_paddleocr_files(ocr_files_dir)
    copy_paddle_files(ocr_files_dir)
    copy_tesseract_files(ocr_files_dir)

if __name__ == "__main__":
    main() 