@echo off
echo Building OCR Tool...

:: Create and activate virtual environment
python -m venv venv
call venv\Scripts\activate.bat

:: Install requirements
pip install -r requirements.txt

:: Uninstall PyInstaller and reinstall fresh without cache
pip uninstall -y pyinstaller
pip install --no-cache-dir pyinstaller

:: Create ICO file from PNG
python create_ico.py

:: Clean previous build
rmdir /s /q build dist
mkdir "dist"

:: Build executable
echo Building executable...
pyinstaller --clean build.spec

:: Verify the executable was created
if not exist "dist\OCR_to_Excel_Tool.exe" (
    echo Error: Build failed - executable not created
    exit /b 1
)

:: Create NSIS installer
echo Creating installer...
"C:\Program Files (x86)\NSIS\makensis.exe" installer.nsi

echo Build complete!
if exist "OCR_Tool_Setup.exe" (
    echo Installer created successfully at "OCR_Tool_Setup.exe"
) else (
    echo Error: Installer creation failed
)
pause