import os
import shutil
from pathlib import Path

def copy_tesseract_files():
    # Source Tesseract installation directory
    tesseract_install_dir = r"C:\Program Files\Tesseract-OCR"
    
    # Your project's tesseract directory
    project_dir = os.path.dirname(os.path.abspath(__file__))
    tesseract_binary_dir = os.path.join(project_dir, "tesseract_binary")
    tessdata_dir = os.path.join(tesseract_binary_dir, "tessdata")
    
    # Create directories if they don't exist
    os.makedirs(tesseract_binary_dir, exist_ok=True)
    os.makedirs(tessdata_dir, exist_ok=True)
    
    try:
        # Copy tesseract.exe
        shutil.copy2(
            os.path.join(tesseract_install_dir, "tesseract.exe"),
            os.path.join(tesseract_binary_dir, "tesseract.exe")
        )
        print("Copied tesseract.exe")
        
        # Copy all DLL files
        for file in os.listdir(tesseract_install_dir):
            if file.endswith('.dll'):
                shutil.copy2(
                    os.path.join(tesseract_install_dir, file),
                    os.path.join(tesseract_binary_dir, file)
                )
                print(f"Copied {file}")
        
        # Copy eng.traineddata
        shutil.copy2(
            os.path.join(tesseract_install_dir, "tessdata", "eng.traineddata"),
            os.path.join(tessdata_dir, "eng.traineddata")
        )
        print("Copied eng.traineddata")
        
        print("\nAll files copied successfully!")
        
    except Exception as e:
        print(f"Error copying files: {str(e)}")

if __name__ == "__main__":
    copy_tesseract_files() 