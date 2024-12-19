RequestExecutionLevel admin

!include "MUI2.nsh"
!include "FileFunc.nsh"
!include LogicLib.nsh

SetCompressor /FINAL /SOLID lzma

# Define application information
!define APPNAME "OCR Tool"
!define PUBLISHER "Your Company Name"
!define VERSION "1.0.0"

Name "${APPNAME}"
OutFile "OCR_Tool_Setup.exe"
InstallDir "$PROGRAMFILES64\${APPNAME}"

!define MUI_ICON "icons\icon.ico"
!define MUI_UNICON "icons\icon.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Section "MainSection" SEC01
    SetShellVarContext all
    SetOutPath "$INSTDIR"
    
    # Create temp directory in AppData
    CreateDirectory "$LOCALAPPDATA\OCR_Tool\temp"
    AccessControl::GrantOnFile "$LOCALAPPDATA\OCR_Tool" "(S-1-1-0)" "FullAccess"
    
    # Copy main executable
    File "dist\OCR_Tool.exe"
    
    # Create shortcuts
    CreateDirectory "$SMPROGRAMS\${APPNAME}"
    CreateShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\OCR_Tool.exe"
    CreateShortCut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\OCR_Tool.exe"
    
    # Create uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    
    # Register application
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" \
                     "DisplayName" "${APPNAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" \
                     "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OCR Tool" \
                     "DisplayIcon" "$INSTDIR\OCR_Tool.exe"
SectionEnd

Section "Uninstall"
    SetShellVarContext all
    
    # Remove application files
    Delete "$INSTDIR\OCR_Tool.exe"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir "$INSTDIR"
    
    # Remove temp directory
    RMDir /r "$LOCALAPPDATA\OCR_Tool"
    
    # Remove shortcuts
    Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
    Delete "$DESKTOP\${APPNAME}.lnk"
    RMDir "$SMPROGRAMS\${APPNAME}"
    
    # Remove registry keys
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
SectionEnd 